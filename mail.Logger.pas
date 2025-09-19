unit mail.Logger;

interface

uses
  System.Classes, System.SyncObjs, mail.TypesUtils,
  FireDAC.Comp.Client, FireDAC.Stan.Intf, Data.DB; // for TFDConnection

type
  // Choose where persistent logs go; console/debug output is always emitted.
  TLogTarget = (ltFile, ltDb);
  TLogTargets = set of TLogTarget;

  // Log type filtering (new)
  TLogKind  = (lkInfo, lkWarn, lkError, lkProfile);
  TLogKinds = set of TLogKind;

  TLogger = class
  private
    FLogPath: String;
    FLock: TCriticalSection;
    FWarnedNoFile: Boolean;

    // === New fields for DB logging and routing ===
    FTargets: TLogTargets;
    FConnectionParams: TFDConnectionDefParams;
    FDbConn: TFDConnection;
    FDbQuery: TFDQuery;
    FDbReady: Boolean;
    FUserMoniker: string;

    // NEW: which kinds are enabled
    FEnabledKinds: TLogKinds;

    procedure InitCore(const LogPath: string; aConnectionParams: TFDConnectionDefParams; const Targets: TLogTargets);
    procedure EmitConsole(const S: string);
    procedure EnsureDbReady;
    procedure DbInsertLine(const Msg: string);
    function  BuildUserMoniker: string;

    // Internal writer that actually emits the line (no filtering here)
    procedure WriteRaw(const Level, Msg: string);

    // Helpers for filtering
    function  ShouldLog(const Kind: TLogKind): Boolean; inline;
    function  TryParseKind(const Level: string; out Kind: TLogKind): Boolean;
  public
    // Backward-compatible: file logging (and console/debug output).
    constructor Create(const LogPath: string); overload;

    // New: receive an existing FireDAC connection; default to file+db.
    constructor Create(aConnectionParams: TFDConnectionDefParams); overload;

    // New: receive connection and explicit targets.
    constructor Create(const LogPath: string; aConnectionParams: TFDConnectionDefParams); overload;
    constructor Create(const LogPath: string; aConnectionParams: TFDConnectionDefParams; const Targets: TLogTargets); overload;

    destructor Destroy; override;

    // NOTE: Line keeps backward compatibility. If Level is one of
    // INFO/WARN/ERROR/PROFILE, filtering is applied; otherwise it is emitted.
    procedure Line(const Kind: TLogKind; const Msg: string);

    procedure Info(const Msg: string);
    procedure Warn(const Msg: string);
    procedure Error(const Msg: string);

    // NEW: profile-level logging (subject to EnabledKinds)
    procedure Profile(const Msg: string);

    // NEW: runtime flag to select which kinds are logged
    property EnabledKinds: TLogKinds read FEnabledKinds write FEnabledKinds;
  end;

implementation

uses
  DateUtils, System.SysUtils, Winapi.Windows;

{ ======================= TLogger core initialization ======================= }

procedure TLogger.InitCore(const LogPath: string; aConnectionParams: TFDConnectionDefParams; const Targets: TLogTargets);
begin
  FWarnedNoFile := False;
  FLock := TCriticalSection.Create;

  // Targets & DB state
  FTargets := Targets;

  FConnectionParams:= aConnectionParams;

  FDbQuery := nil;
  FDbReady := False;

  FLogPath:= LogPath;

  FEnabledKinds := [lkInfo, lkWarn, lkError, lkProfile];

  // Cache user moniker once
  FUserMoniker := BuildUserMoniker;
end;

constructor TLogger.Create(const LogPath: string);
begin
  inherited Create;
  // Default behavior preserved: write to file (if possible) + console/debug output
  InitCore(LogPath, nil, [ltFile]);
end;

constructor TLogger.Create(aConnectionParams: TFDConnectionDefParams);
begin
  inherited Create;
  // If a connection is provided, enable both file and db by default
  if Assigned(aConnectionParams) then
    InitCore('', aConnectionParams, [ltDb])
  else
    InitCore('', nil, [ltFile]);
end;

constructor TLogger.Create(const LogPath: string; aConnectionParams: TFDConnectionDefParams; const Targets: TLogTargets);
begin
  inherited Create;
  InitCore(LogPath, aConnectionParams, Targets);
end;

constructor TLogger.Create(const LogPath: string; aConnectionParams: TFDConnectionDefParams);
var
  LTargets: TLogTargets;
begin
  inherited Create;
  LTargets := [];
  if LogPath <> '' then
    LTargets := LTargets + [ltFile];
  if Assigned(aConnectionParams) then
    LTargets := LTargets + [ltDb];

  InitCore(LogPath, aConnectionParams, LTargets);
end;

destructor TLogger.Destroy;
begin
  if Assigned(FDbQuery) then
    FDbQuery.Free;

  if Assigned(FDBConn) then
    FDbConn.Free;

  FLock.Free;
  inherited;
end;

{ ======================= Helpers ======================= }

procedure TLogger.EmitConsole(const S: string);
begin
  // In GUI apps, avoid WriteLn; send to the debugger instead.
  if IsConsole then
    Writeln(S)
  else
    OutputDebugString(PChar(S));
end;

function TLogger.BuildUserMoniker: string;
var
  pc, usr: string;
  buf: array[0..MAX_COMPUTERNAME_LENGTH] of WideChar;
  sz: DWORD;
  ubuf: array[0..257] of WideChar;
  usz: DWORD;
begin
  // First try environment variables (fast & simple)
  pc  := GetEnvironmentVariable('COMPUTERNAME');
  usr := GetEnvironmentVariable('USERNAME');

  // Fallback to API if needed
  if pc = '' then
  begin
    sz := Length(buf);
    if GetComputerNameW(@buf[0], sz) then
      SetString(pc, PWideChar(@buf[0]), sz);
  end;

  if usr = '' then
  begin
    usz := Length(ubuf);
    if GetUserNameW(@ubuf[0], usz) then
      SetString(usr, PWideChar(@ubuf[0]), usz - 1); // API includes trailing #0 in length
  end;

  if pc = '' then pc := 'unknown-pc';
  if usr = '' then usr := 'unknown-user';
  Result := pc + '/' + usr;
end;

procedure TLogger.EnsureDbReady;
begin
  if FDbReady or not (ltDb in FTargets) then
    Exit;

  if not Assigned(FDBConn) then
  begin
    if not Assigned(FConnectionParams) then
      Exit;

    FDbConn  := TFDConnection.Create(nil);
    FDbConn.TxOptions.AutoCommit := True;
    FDBConn.Params.Assign(FConnectionParams);
  end;

  if not FDbConn.Connected then
    FDbConn.Connected:= True; // respect caller-managed connection state

  if FDbQuery = nil then
  begin
    FDbQuery := TFDQuery.Create(nil);
    FDbQuery.Connection := FDbConn;
  end;

  // Create table if missing (schema "mail" must exist beforehand)
  // Bracket [log] to avoid clashes with potential identifiers.
  try
    FDbQuery.SQL.Text :=
      'IF OBJECT_ID(''mail.log'') IS NULL ' +
      'BEGIN ' +
      '  CREATE TABLE mail.[log] (' +
      '    id INT IDENTITY (1,1) CONSTRAINT pk_mail_log PRIMARY KEY,' +
      '    user_name VARCHAR(255) NULL,' +
      '    [message] VARCHAR(MAX) NULL' +
      '  ) ' +
      'END';
    FDbQuery.ExecSQL;
  except
    // If creation fails, we still try to insert; errors will be surfaced then.
  end;

  FDbReady := True;
end;

procedure TLogger.DbInsertLine(const Msg: string);
begin
  EnsureDbReady;

  if (FDbConn = nil) or not (ltDb in FTargets) then
    Exit;

  if not FDbReady then
    Exit;

  try
    FDbQuery.SQL.Text := 'INSERT INTO mail.[log](user_name, [message]) VALUES (:u, :m)';
    FDbQuery.Params.Clear;
    FDbQuery.Params.Add.Name := 'u';
    FDbQuery.ParamByName('u').DataType := ftWideString;
    FDbQuery.ParamByName('u').AsString := FUserMoniker;

    // Use WideMemo to allow large text; SQL Server will NVARCHAR->VARCHAR convert if needed.
    FDbQuery.Params.Add.Name := 'm';
    FDbQuery.ParamByName('m').DataType := ftWideMemo;
    FDbQuery.ParamByName('m').AsWideMemo := Msg;

    FDbQuery.ExecSQL;
  except
    // Do not raise; logging must never break the caller.
    // Best-effort: emit a debug hint that DB logging failed.
    EmitConsole('WARN DB log insert failed; continuing without DB logging.');
  end;
end;

function TLogger.ShouldLog(const Kind: TLogKind): Boolean;
begin
  Result := Kind in FEnabledKinds;
end;

function TLogger.TryParseKind(const Level: string; out Kind: TLogKind): Boolean;
var
  U: string;
begin
  U := UpperCase(Trim(Level));
  if U = 'INFO' then begin Kind := lkInfo;    Exit(True); end;
  if U = 'WARN' then begin Kind := lkWarn;    Exit(True); end;
  if U = 'ERROR' then begin Kind := lkError;  Exit(True); end;
  if U = 'PROFILE' then begin Kind := lkProfile; Exit(True); end;
  Result := False;
end;

procedure TLogger.WriteRaw(const Level, Msg: string);
var
  L: string;
  Writer: TStreamWriter;
begin
  L := FormatDateTime('yyyy-mm-dd hh:nn:ss.zzz', Now) + ' ' + Level + ' ' + Msg;

  FLock.Enter;
  try
    // Console/debug stream (always emitted; pick suitable sink)
    EmitConsole(L);

    // File target: open → write → close (avoid keeping the file locked)
    if (ltFile in FTargets) and (FLogPath <> '') then
    begin
      try
        Writer := TStreamWriter.Create(FLogPath, True, TEncoding.UTF8);
        try
          Writer.WriteLine(L);
          Writer.Flush;
        finally
          Writer.Free; // release the file handle immediately
        end;
      except
        if not FWarnedNoFile then
        begin
          EmitConsole('WARN log file not available; continuing with console/debug logs.');
          FWarnedNoFile := True;
        end;
      end;
    end;

    // Database target
    if ltDb in FTargets then
      DbInsertLine(L);
  finally
    FLock.Leave;
  end;
end;

{ ======================= Public API ======================= }

procedure TLogger.Line(const Kind: TLogKind; const Msg: string);
var
  FStrKind: String;
begin
  if not ShouldLog(Kind) then
    Exit;

  case Kind of
    lkInfo: FStrKind:= 'INFO';
    lkWarn: FStrKind:= 'WARN';
    lkError: FStrKind:= 'ERROR';
    lkProfile: FStrKind:= 'PROFILE';
  else
    FStrKind:= 'LOG';
  end;

  WriteRaw(FStrKind, Msg);
end;

procedure TLogger.Info(const Msg: string);
begin
  Line(lkInfo, Msg);
end;

procedure TLogger.Warn(const Msg: string);
begin
  Line(lkWarn, Msg);
end;

procedure TLogger.Error(const Msg: string);
begin
  Line(lkError, Msg);
  {$IFDEF DEBUG}
  OutputDebugString(PChar(Msg));
  {$ENDIF}
end;

procedure TLogger.Profile(const Msg: string);
begin
  Line(lkProfile, Msg);
end;

end.

