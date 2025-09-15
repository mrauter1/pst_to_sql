unit mail.ComWorker;

{
  COM Sync Worker (Delphi 10 Seattle compatible)
  ----------------------------------------------
  Purpose:
    Run Outlook-driven synchronization (TMailFullSyncReconciler) on a dedicated
    STA thread every X seconds. The first cycle performs a FULL scan. Subsequent
    cycles perform incremental scans filtered by LastModificationTime.

  Reuse:
    - mail.Core      : TOutlookCOM, TIngestor, DetectCurrentPstFolder
    - mail.FullSync  : TMailFullSyncReconciler
    - mail.Data      : TDb
    - mail.TypesUtils: TCliOptions, TLogger

  Notes:
    - Outlook OOM requires a single-threaded apartment. This worker thread
      initializes COM with COINIT_APARTMENTTHREADED and owns all COM objects.
    - Interval selection: if IntervalSeconds passed to the constructor is 0,
      Batch (TCliOptions.Batch) is interpreted as "seconds between cycles".
      If neither is > 0, a default of 300 seconds is used.
    - PST attachment is done once per worker lifetime while initialized; the same
      Outlook session and DB connection are reused across cycles for efficiency.
    - If Classic Outlook is not running (no active OLE server), the thread will
      wait and re-check periodically until it becomes active, without launching it.
    - Resilience: any error during a cycle is caught and logged, never propagated
      to the main thread; the worker tears down COM/DB/Outlook safely and retries
      at the next scheduled execution by reinitializing.
}

interface

uses
  System.SysUtils, System.Classes, System.SyncObjs, System.Math, System.Diagnostics,
  System.Variants, FireDac.Stan.Intf,
  Winapi.Windows, Winapi.ActiveX,
  System.Win.ComObj,          // OleCheck, GetActiveOleObject, EOleError types
  Mail.TypesUtils,            // TCliOptions, TLogger, helpers
  mail.Core,                  // TOutlookCOM, TIngestor, DetectCurrentPstFolder
  mail.FullSync,              // TMailFullSyncReconciler
  mail.Data,
  mail.Logger;                 // TDb

type
  { Runs COM synchronization in the background at a fixed interval. }
  TMailComWorker = class(TThread)
  private
    FCs: TCriticalSection;

    // immutable configuration
    FOpts: TCliOptions;
    FConnectionParams: TStringList;
    FIntervalSec: Cardinal;

    // lifetime-owned resources
    FStopEvent: TEvent;       // signaled to stop promptly
    FWakeEvent: TEvent;       // external trigger for "run now"
    FLogger: TLogger;         // kept across reinitializations
    FDb: TDb;                 // created when initialized
    FOlk: TOutlookCOM;        // created when initialized
    FIngestor: TIngestor;     // created when initialized

    // state
    FComInitialized: Boolean;
    FInitialized: Boolean;    // DB/Outlook/Ingestor successfully set up
    FFirstRunDone: Boolean;
    FLastRun: TDateTime;
    FLastError: string;

    // helpers
    function  DefaultIntervalSec: Cardinal;
    function  EffectiveIntervalSec: Cardinal;

    procedure InitLoggerIfNeeded;
    function  IsOutlookComActive: Boolean;
    function  WaitForOutlookActive(const PollMs: Cardinal = 5000): Boolean;

    procedure InitComIfNeeded;
    procedure InitComponents;     // DB + Outlook + Ingestor + PST attach
    procedure TearDownComponents; // Free DB/Outlook/Ingestor (keeps logger)
    procedure UninitComIfNeeded;

    procedure EnsureEffectivePstPath(var Opts: TCliOptions);
    procedure DoOneCycle;
    function  WaitForNextCycle(const TimeoutMs: Cardinal): Boolean;

    function  ShouldReinitAfter(const E: Exception): Boolean;

    // thread-safe accessors/mutators for public properties
    function  GetIntervalSeconds: Cardinal;
    function  GetLastError: string;
    function  GetLastRun: TDateTime;
    procedure SetIntervalSeconds(const Value: Cardinal);
    procedure SetLastRun(const Value: TDateTime);
    procedure SetLastError(const Value: string);
    procedure LogProfile(const S: string);

  protected
    procedure Execute; override;

  public
    { Create the worker.  }
    constructor Create(const Opts: TCliOptions; AConnectionParams: TStringList = nil; IntervalSeconds: Cardinal = 150); overload;
    constructor Create(const Opts: TCliOptions; IntervalSeconds: Cardinal = 150); overload;
    constructor Create(AConnectionParams: TStringList); overload;
    destructor Destroy; override;

    { Request an immediate cycle (does not change the fixed schedule). }
    procedure TriggerNow;

    procedure LogInfo(const S: string);
    procedure LogWarn(const S: string);
    procedure LogError(const S: string);

    { Properties for inspection (read from any thread). }
    property IntervalSeconds: Cardinal read GetIntervalSeconds write SetIntervalSeconds;
    property LastRun: TDateTime read GetLastRun;
    property LastError: string read GetLastError;
  end;

function StarTMailComWorker(const Opts: TCliOptions; IntervalSeconds: Cardinal = 0): TMailComWorker;
procedure RunComSyncWorkerConsole(const Opts: TCliOptions; IntervalSeconds: Cardinal = 0);

implementation

var
  GStopEvent: TEvent = nil;

{ ========================= Construction / Destruction ========================= }

constructor TMailComWorker.Create(const Opts: TCliOptions; AConnectionParams: TStringList = nil; IntervalSeconds: Cardinal = 150);
begin
  inherited Create(True); // suspended, we will Resume after initialization

  FCs := TCriticalSection.Create;

  FreeOnTerminate := False;

  FOpts := Opts;
  FOpts.LogPath:= 'ingest.log';
  SetIntervalSeconds(IntervalSeconds); // protect write via CS

  FStopEvent := TEvent.Create(nil, True{manual reset}, False, '');
  FWakeEvent := TEvent.Create(nil, True{manual reset}, False, '');

  FLogger := nil;
  FDb := nil;
  FOlk := nil;
  FIngestor := nil;
  FComInitialized := False;
  FInitialized := False;
  FFirstRunDone := False;

  // Safe here (thread not yet running)
  FLastRun := 0;
  FLastError := '';

  FConnectionParams := AConnectionParams;
end;

destructor TMailComWorker.Destroy;
begin
  if Self.Started then
  begin
      Terminate;
    // Signal termination and wake any waits
    if Assigned(FStopEvent) then FStopEvent.SetEvent;
    if Assigned(FWakeEvent) then FWakeEvent.SetEvent;

    // Wait for the thread to finish its cleanup
    try
      WaitFor;
    except
      // swallow any wait exceptions on teardown
    end;
  end;

  // Final cleanup
  try
    TearDownComponents;
    UninitComIfNeeded;
  except
    // ignore errors on shutdown
  end;

  FreeAndNil(FStopEvent);
  FreeAndNil(FWakeEvent);
  FreeAndNil(FLogger);

  FCs.Free;
  inherited;
end;

{ ========================= Public API ========================= }

procedure TMailComWorker.TriggerNow;
begin
  if Assigned(FWakeEvent) then
    FWakeEvent.SetEvent;
end;

{ ========================= Logging helpers ========================= }

procedure TMailComWorker.InitLoggerIfNeeded;
begin
  if FLogger = nil then
  begin
    if Assigned(FDB) then
      FLogger := TLogger.Create(FOpts.LogPath, FDB.Conn.Params)
    else
      FLogger := TLogger.Create(FOpts.LogPath, nil);

    FLogger.EnabledKinds:= [lkWarn, lkError, lkProfile];
  end;
end;

procedure TMailComWorker.LogInfo(const S: string);
begin
  InitLoggerIfNeeded;
  if Assigned(FLogger) then
    FLogger.Info(S);
end;

procedure TMailComWorker.LogWarn(const S: string);
begin
  InitLoggerIfNeeded;
  if Assigned(FLogger) then
    FLogger.Warn(S);
end;

procedure TMailComWorker.LogError(const S: string);
begin
  InitLoggerIfNeeded;
  if Assigned(FLogger) then
    FLogger.Error(S);
end;

procedure TMailComWorker.LogProfile(const S: string);
begin
  InitLoggerIfNeeded;
  if Assigned(FLogger) then
    FLogger.Profile(S);
end;

{ ========================= Internal helpers ========================= }

constructor TMailComWorker.Create(
  AConnectionParams: TStringList);
var
  FOpt: TCliOptions;
begin
  FOpt:= Default(TCliOptions); // FOpt should be empty when not running through CLI
  Create(FOpt, AConnectionParams, 150);
end;

constructor TMailComWorker.Create(const Opts: TCliOptions;
  IntervalSeconds: Cardinal);
begin
  Create(Opts, nil, IntervalSeconds);
end;

function TMailComWorker.DefaultIntervalSec: Cardinal;
const
  C_DEFAULT = 300; // 5 minutes
  C_MIN     = 5;   // safety lower bound
var
  S: Cardinal;
  Curr: Cardinal;
begin
  Curr := GetIntervalSeconds; // protected read

  if Curr > 0 then
    S := Curr
  else if FOpts.Batch > 0 then
    S := Cardinal(FOpts.Batch)
  else
    S := C_DEFAULT;

  if S < C_MIN then
    S := C_MIN;

  Result := S;
end;

function TMailComWorker.EffectiveIntervalSec: Cardinal;
begin
  Result := DefaultIntervalSec;
end;

function TMailComWorker.IsOutlookComActive: Boolean;
var
  V: OleVariant;
begin
  try
    V := GetActiveOleObject('Outlook.Application'); // raises if not running
    Result := IsVariantAssigned(V);
  except
    Result := False;
  end;
end;

function TMailComWorker.WaitForOutlookActive(const PollMs: Cardinal): Boolean;
begin
  Result := IsOutlookComActive;
  if Result then Exit(True);

  LogWarn('Outlook COM not active; waiting for Classic Outlook to be opened...');
  while not Terminated do
  begin
    if IsOutlookComActive then
    begin
      LogInfo('Detected active Outlook COM.');
      Exit(True);
    end;

    if (FStopEvent.WaitFor(PollMs) = wrSignaled) or Terminated then
      Exit(False);
  end;
  Result := False;
end;

procedure TMailComWorker.InitComIfNeeded;
begin
  if not FComInitialized then
  begin
    OleCheck(CoInitializeEx(nil, COINIT_APARTMENTTHREADED));
    FComInitialized := True;
  end;
end;

procedure TMailComWorker.UninitComIfNeeded;
begin
  if FComInitialized then
  begin
    CoUninitialize;
    FComInitialized := False;
  end;
end;

procedure TMailComWorker.EnsureEffectivePstPath(var Opts: TCliOptions);
var
  DetPst, DetRoot, DetFolderPath: string;
begin
  // Resolve "active" or empty PST path using the current Outlook session
  if (Trim(Opts.PSTPath) = '') or SameText(Trim(Opts.PSTPath), 'active') then
  begin
    if not DetectCurrentPstFolder(FOlk.Outlook, DetPst, DetRoot, DetFolderPath) then
      raise Exception.Create('Unable to detect the active Outlook folder/store via COM. ' +
                             'Open Classic Outlook, select a folder in the target PST, or pass --pst=<path>.');

    if DetPst = '' then
      raise Exception.Create('Detected active store has no local file path (likely an online mailbox). ' +
                             'Please pass a PST path explicitly with --pst=<path>.');

    LogInfo(Format('Detected active store "%s" at: %s', [DetRoot, DetPst]));
    if DetFolderPath <> '' then
      LogInfo('Active folder: ' + DetFolderPath);

    Opts.PSTPath := DetPst;
  end;
end;

procedure TMailComWorker.InitComponents;
begin
  // COM must be initialized at this point
  if not FComInitialized then
    InitComIfNeeded;

  // DB connection (respects .env if configured)
  if FDb = nil then
  begin
    if Assigned(FConnectionParams) then
      FDB := TDb.Create(FConnectionParams)
    else
      FDb := TDb.Create(FOpts.ConnOverride, FOpts.EnvFile);
  end;

  if Assigned(FLogger) then
    FreeAndNil(FLogger);

  // Recreate here so we are sure to get DB connection
  InitLoggerIfNeeded;

  FLogger.Info('Checking if outlook is active.');

  // Wait until Classic Outlook is actually running (do NOT auto-launch it)
  if not WaitForOutlookActive(5000) then
    raise EAbort.Create('Stop requested while waiting for Outlook.');

  FLogger.Info('Outlook active, initiating components.');

  // Outlook COM (bind to running instance now)
  if FOlk = nil then
  begin
    FOlk := TOutlookCOM.Create;
    FOlk.Initialize;
  end;

  // Resolve PST path if needed, then construct the ingestor facade and attach PST
  if FIngestor = nil then
  begin
    EnsureEffectivePstPath(FOpts);
    FIngestor := TIngestor.Create(FDb, FOlk, FLogger, FOpts);
    FIngestor.LoadPst;
    LogInfo(Format('Attached PST "%s" (pst_id=%d).', [FIngestor.RootDisplay, FIngestor.PstId]));
  end;

  FInitialized := True;
end;

procedure TMailComWorker.TearDownComponents;
begin
  // Important: free high-level objects first, then Outlook/DB, then COM.
  if Assigned(FIngestor) then
  begin
    try
      FIngestor.UnloadPST;
    except
      // ignore detach failures in teardown
    end;
  end;

  FreeAndNil(FIngestor);
  FreeAndNil(FOlk);
  FreeAndNil(FDb);

  FInitialized := False;
end;

function TMailComWorker.WaitForNextCycle(const TimeoutMs: Cardinal): Boolean;
var
  Handles: array[0..1] of THandle;
  r: DWORD;
begin
  // Reset the wake event before we go to sleep; if someone calls TriggerNow
  // after this, the event will wake us promptly.
  if Assigned(FWakeEvent) then
    FWakeEvent.ResetEvent;

  Handles[0] := FStopEvent.Handle;
  Handles[1] := FWakeEvent.Handle;

  r := WaitForMultipleObjects(2, @Handles[0], False{wait any}, TimeoutMs);
  Result := (r <> WAIT_TIMEOUT);
  // Return True when woken (stop or trigger), False when elapsed.
end;

procedure TMailComWorker.SetIntervalSeconds(const Value: Cardinal);
begin
  FCs.Enter;
  try
    FIntervalSec := Value;
  finally
    FCs.Leave;
  end;
end;

procedure TMailComWorker.SetLastRun(const Value: TDateTime);
begin
  FCs.Enter;
  try
    FLastRun := Value;
  finally
    FCs.Leave;
  end;
end;

procedure TMailComWorker.SetLastError(const Value: string);
begin
  FCs.Enter;
  try
    FLastError := Value;
  finally
    FCs.Leave;
  end;
end;

function TMailComWorker.ShouldReinitAfter(const E: Exception): Boolean;
var
  MsgLow: string;
begin
  Result:= True;
  // Re-initialize on any COM-related failure or when Outlook likely went away.
  Result := (E is EOleSysError) or (E is EOleError);
  if not Result then
  begin
    MsgLow := LowerCase(E.Message);
    Result :=
      (Pos('rpc_e_', MsgLow) > 0) or
      (Pos('class not registered', MsgLow) > 0) or
      (Pos('call was rejected by callee', MsgLow) > 0) or
      (Pos('server busy', MsgLow) > 0) or
      (Pos('disconnectedcontext', MsgLow) > 0) or
      (Pos('outlook', MsgLow) > 0);
  end;
end;

procedure TMailComWorker.DoOneCycle;
var
  Reco: TMailFullSyncReconciler;
  Cutoff: TDateTime;
  SW: TStopwatch;
  FLogPrefix: String;
begin
//  LogInfo('Starting sync cycle...');
  FLogPrefix:= '';

  Reco := TMailFullSyncReconciler.Create(FIngestor);
  try
    SW := TStopwatch.StartNew;

    if not FFirstRunDone then
    begin
      FLogPrefix:= 'Full Sync finished in';
      // First run: full scan (no cutoff)
      LogInfo('First run: FULL reconcile (no cutoff).');
      Reco.IterationMode:= imTable; // Table mode is faster for full scan
      Reco.Execute(False{full mode}, 0{cutoff});
      FFirstRunDone := True;
    end
    else
    begin
      // Incremental: only items updated since the last known LastModificationTime
      Cutoff := UTCDateToLocal(FDb.GetScanCutoffLastModUtc(FIngestor.PstId));
//      if Cutoff > 0 then
//        Cutoff := Cutoff - (10.0 / (24.0 * 60.0 * 60.0)); // subtract 10 seconds as a safety window

//      LogInfo(Format('Incremental reconcile with cutoff (LastModificationTime) >= %s',
//        [DateToStrLocalNoSeconds(Cutoff)]));
      FLogPrefix:= 'Incremental Sync finished in';

      Reco.IterationMode:= imRestrict; // Restrict is faster for incremental scan (50 or less items);
      // Use full reconcile mode with a cutoff to capture new/updated items efficiently.
      Reco.Execute(False{full mode}, Cutoff);
    end;

    SW.Stop;
    LogProfile(Format(FLogPrefix+' %d ms (%.3f s).',
      [SW.ElapsedMilliseconds, SW.Elapsed.TotalSeconds]));

    // Protected updates for public-readable state
    SetLastRun(Now); // local clock
    SetLastError('');
  finally
    Reco.Free;
  end;
end;

{ ========================= Thread main ========================= }

procedure TMailComWorker.Execute;
var
  intervalMs: Cardinal;
  woke: Boolean;
  needReinit: Boolean;
begin
  // Top-level guard: never let an exception escape the thread
  try
    InitLoggerIfNeeded;
    LogInfo('COM sync worker starting...');

    // Main service loop
    while not Terminated do
    begin
      // Ensure COM + components are initialized; retry on failures without
      // propagating exceptions to the main thread.
      if not FInitialized then
      begin
        try
          InitComIfNeeded;
          InitComponents;
          LogInfo('Initialization successful.');
        except
          on E: Exception do
          begin
            SetLastError(E.Message);
            LogError('Initialization failed: ' + E.Message);

            // Best-effort cleanup before retrying
            try
              TearDownComponents;
            except end;
            try
              UninitComIfNeeded;
            except end;

            // Wait a short backoff before retrying initialization
            if FStopEvent.WaitFor(3000) = wrSignaled then
              Break;

            Continue; // attempt to init again
          end;
        end;
      end;

      // Run one cycle and make the worker resilient to any error.
      needReinit := False;
      try
        // Outlook may have been closed between cycles; check proactively.
        if not IsOutlookComActive then
          raise Exception.Create('Outlook COM is not active.');

        DoOneCycle;
      except
        on E: Exception do
        begin
          SetLastError(E.Message);
          LogError('Sync cycle error: ' + E.Message);
          needReinit := ShouldReinitAfter(E);
        end;
      end;

      if Terminated then
        Break;

      // If a COM-related error occurred, tear everything down and re-initialize
      // on the next loop iteration (safer than attempting to continue).
      if needReinit then
      begin
        LogWarn('Reinitializing COM/Outlook/DB after failure...');
        try
          TearDownComponents;
        except
          // ignore
        end;
        try
          UninitComIfNeeded;
        except
          // ignore
        end;

        // Small pause before next re-init attempt
        if FStopEvent.WaitFor(2000) = wrSignaled then
          Break;

        Continue;
      end;

      // Normal wait for the next cycle (or an external trigger)
      intervalMs := EffectiveIntervalSec * 1000;
      woke := WaitForNextCycle(intervalMs);
      if Terminated then
        Break;

      // If we were woken by TriggerNow, just continue; if by timeout, also continue.
      if woke and Assigned(FWakeEvent) then
        FWakeEvent.ResetEvent;
    end;

  except
    on E: Exception do
    begin
      // Last resort catch — still never propagate to the main thread.
      SetLastError(E.Message);
      LogError('Unexpected thread error: ' + E.Message);
    end;
  end;

  // Final teardown (idempotent)
  try
    TearDownComponents;
  except end;
  try
    UninitComIfNeeded;
  except end;

  LogInfo('COM sync worker stopped.');
end;

function TMailComWorker.GetIntervalSeconds: Cardinal;
begin
  FCs.Enter;
  try
    Result := FIntervalSec;
  finally
    FCs.Leave;
  end;
end;

function TMailComWorker.GetLastError: string;
begin
  FCs.Enter;
  try
    Result := FLastError;
  finally
    FCs.Leave;
  end;
end;

function TMailComWorker.GetLastRun: TDateTime;
begin
  FCs.Enter;
  try
    Result := FLastRun;
  finally
    FCs.Leave;
  end;
end;

function ConsoleCtrlHandler(CtrlType: DWORD): BOOL; stdcall;
begin
  case CtrlType of
    CTRL_C_EVENT, CTRL_BREAK_EVENT, CTRL_CLOSE_EVENT:
      begin
        if Assigned(GStopEvent) then
          GStopEvent.SetEvent;
        Result := True;
      end;
  else
    Result := False;
  end;
end;

function StarTMailComWorker(const Opts: TCliOptions; IntervalSeconds: Cardinal): TMailComWorker;
begin
  // Creates and starts the worker thread. The thread owns COM, DB and Outlook objects.
  // Keep the instance alive for as long as you want the sync to run.
  Result := TMailComWorker.Create(Opts, IntervalSeconds);
  Result.Start;
end;

procedure RunComSyncWorkerConsole(const Opts: TCliOptions; IntervalSeconds: Cardinal);
var
  Worker: TMailComWorker;
begin
  // Blocking runner suitable for console tools. It:
  //  - starts the worker
  //  - pumps queued calls (CheckSynchronize)
  //  - exits cleanly on Ctrl+C / console close
  Worker := StarTMailComWorker(Opts, IntervalSeconds);
  try
    GStopEvent := TEvent.Create(nil, True{manual reset}, False, '');
    SetConsoleCtrlHandler(@ConsoleCtrlHandler, True);
    try
      while WaitForSingleObject(GStopEvent.Handle, 100) = WAIT_TIMEOUT do
        CheckSynchronize(50);
    finally
      SetConsoleCtrlHandler(@ConsoleCtrlHandler, False);
      FreeAndNil(GStopEvent);
    end;
  finally
    // Frees the thread (its destructor signals termination and waits for a clean shutdown)
    Worker.Free;
  end;
end;

end.

