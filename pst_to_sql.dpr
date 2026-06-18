program pst_to_sql;

{$APPTYPE CONSOLE}

uses
  System.SysUtils,
  System.IOUtils,
  System.StrUtils,
  Winapi.ActiveX,
  mail.data in 'mail.data.pas',
  mail.TypesUtils in 'mail.TypesUtils.pas',
  mail.Core in 'mail.Core.pas' {/  mail.ComWorker in 'mail.ComWorker.pas';},
  mail.ComWorker in 'mail.ComWorker.pas',
  mail.FullSync in 'mail.FullSync.pas',
  mail.Logger in 'mail.Logger.pas';

const
  cDefaultEnvFile = '.env';
  cDefaultLogFile = 'ingest.log';
  cDefaultBatch   = 500;

procedure PrintGuardStateResult(const R: TOutlookGuardResult);
  function GuardModeName(const Mode: TOutlookObjectModelGuardMode): string;
  begin
    case Mode of
      emPolicyControlled:
        Result := 'PolicyControlled';
      emMachineObjectModelGuard:
        Result := 'MachineObjectModelGuard';
      emNoExplicitSetting:
        Result := 'NoExplicitSetting';
    else
      Result := 'Unknown';
    end;
  end;
begin
  Writeln('Mode: ', Ord(R.Mode), ' (', GuardModeName(R.Mode), ')');
  Writeln('Description: ', R.Description);
  if R.Mode = emMachineObjectModelGuard then
  begin
    Writeln('ObjectModelGuardValue: ', R.ObjectModelGuardValue);
    Writeln('ObjectModelGuardMeaning: ', R.ObjectModelGuardMeaning);
    Writeln('ObjectModelGuardPath: ', R.ObjectModelGuardPath);
  end
  else if R.Mode = emPolicyControlled then
  begin
    Writeln('AdminSecurityMode: ', R.AdminSecurityMode);
    Writeln('PromptOOMSend: ', R.PromptOOMSend);
    Writeln('PromptOOMAddressInformationAccess: ', R.PromptOOMAddressInformationAccess);
    Writeln('PromptOOMAddressBookAccess: ', R.PromptOOMAddressBookAccess);
    Writeln('PromptOOMSaveAs: ', R.PromptOOMSaveAs);
  end;
end;

procedure CheckGuardState;
var
  FOutlookVersion: String;
  R: TOutlookGuardResult;
begin
  FOutlookVersion:= GetLatestOutlookOfficeVersion;
  Writeln('Latest Outlook Version: ', FOutlookVersion);
  if FOutlookVersion = '' then
  begin
    Writeln('ERROR: Could not detect installed Outlook Office version.');
    Readln;
    Exit;
  end;
  R := ReadEffectiveOutlookObjectModelGuard(FOutlookVersion);

  PrintGuardStateResult(R);
  Writeln('GetOutlookObjectModelGuardValue: ', GetOutlookObjectModelGuardValue(FOutlookVersion));

  Readln;
end;

function PowerShellDoubleQuoted(const S: string): string;
begin
  Result := StringReplace(S, '"', '`"', [rfReplaceAll]);
end;

procedure PrintSetGuardStateAdminCommands(const WriteResults: TArray<TOutlookGuardWriteResult>);
var
  I: Integer;
  BasePath: string;
begin
  if Length(WriteResults) = 0 then
    Exit;

  Writeln('Run PowerShell as Administrator and execute:');
  for I := 0 to High(WriteResults) do
  begin
    BasePath := PowerShellDoubleQuoted(WriteResults[I].RegistryPath);
    Writeln('$base = "' + BasePath + '"');
    if Pos('\ClickToRun\', WriteResults[I].RegistryPath) > 0 then
    begin
      Writeln('if (Test-Path $base) {');
      Writeln('  New-ItemProperty -Path $base -Name "ObjectModelGuard" -Value 2 -PropertyType DWord -Force | Out-Null');
      Writeln('} else {');
      Writeln('  Write-Error "ClickToRun key not found: $base"');
      Writeln('}');
    end
    else
    begin
      Writeln('New-Item -Path $base -Force | Out-Null');
      Writeln('New-ItemProperty -Path $base -Name "ObjectModelGuard" -Value 2 -PropertyType DWord -Force | Out-Null');
    end;
  end;
end;

procedure PrintGuardWriteResults(const Caption: string; const WriteResults: TArray<TOutlookGuardWriteResult>);
var
  I: Integer;
begin
  Writeln(Caption);
  for I := 0 to High(WriteResults) do
  begin
    if WriteResults[I].Success then
      Writeln('  OK   ', WriteResults[I].RegistryPath)
    else
      Writeln('  FAIL ', WriteResults[I].RegistryPath, ' - ', WriteResults[I].ErrorMessage);
  end;
end;

function HasGuardWriteFailures(const WriteResults: TArray<TOutlookGuardWriteResult>): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 0 to High(WriteResults) do
    if not WriteResults[I].Success then
      Exit(True);
end;

procedure SetGuardState;
var
  FOutlookVersion: string;
  AfterState: TOutlookGuardResult;
  PolicyWriteResults, MachineWriteResults: TArray<TOutlookGuardWriteResult>;
  PolicyWriteFailures, MachineWriteFailures: Boolean;
begin
  FOutlookVersion := GetLatestOutlookOfficeVersion;
  Writeln('Latest Outlook Version: ', FOutlookVersion);

  if FOutlookVersion = '' then
  begin
    Writeln('ERROR: Could not detect installed Outlook Office version. No registry key was written.');
    Halt(1);
  end;

  Writeln('Writing current-user Outlook policy keys. Run this as the same Windows user/session that runs Outlook.');

  try
    SetOutlookObjectModelGuardPolicyCurrentUser(FOutlookVersion, PolicyWriteResults);
  except
    on E: Exception do
    begin
      Writeln('ERROR: Failed to prepare current-user Outlook policy writes: ' + E.Message);
      Halt(1);
    end;
  end;

  PrintGuardWriteResults('Current-user Outlook policy write results:', PolicyWriteResults);
  PolicyWriteFailures := HasGuardWriteFailures(PolicyWriteResults);

  try
    SetOutlookObjectModelGuardMachine(FOutlookVersion, 2, MachineWriteResults);
  except
    on E: Exception do
    begin
      Writeln('WARNING: Failed to prepare machine ObjectModelGuard=2 fallback writes: ' + E.Message);
      SetLength(MachineWriteResults, 0);
    end;
  end;

  PrintGuardWriteResults('Machine ObjectModelGuard fallback write results:', MachineWriteResults);
  MachineWriteFailures := HasGuardWriteFailures(MachineWriteResults);

  if PolicyWriteFailures then
  begin
    Writeln('ERROR: One or more current-user Outlook policy registry writes failed.');
    Halt(1);
  end;

  if MachineWriteFailures then
  begin
    Writeln('WARNING: One or more HKLM ObjectModelGuard fallback registry writes failed.');
    PrintSetGuardStateAdminCommands(MachineWriteResults);
  end;

  AfterState := ReadEffectiveOutlookObjectModelGuard(FOutlookVersion);
  Writeln('Effective guard state after write:');
  PrintGuardStateResult(AfterState);

  if not IsOutlookObjectModelGuardNeverWarn(AfterState) then
  begin
    Writeln('ERROR: Outlook guard settings were written but are not the effective Never Warn/Auto Approve setting.');
    if MachineWriteFailures then
      PrintSetGuardStateAdminCommands(MachineWriteResults);
    Halt(1);
  end;
end;

function UnquoteIfQuoted(const S: string): string;
begin
  Result := S;
  if (Length(Result) >= 2) and
     (((Result[1] = '"') and (Result[Length(Result)] = '"')) or
      ((Result[1] = '''') and (Result[Length(Result)] = ''''))) then
    Result := Result.Substring(1, Result.Length - 2);
end;

function GetCmdValue(const Name, DefaultValue: string): string;
var
  i: Integer;
  key, s, nextArg: string;
begin
  Result := DefaultValue;
  key := '--' + Name;

  i := 1;
  while i <= ParamCount do
  begin
    s := ParamStr(i);

    // Form: --name=value
    if StartsText(key + '=', s) then
      Exit(UnquoteIfQuoted(Copy(s, Length(key) + 2, MaxInt)));

    // Form: --name value   OR a bare boolean flag (--name)
    if SameText(s, key) then
    begin
      if (i < ParamCount) then
      begin
        nextArg := ParamStr(i + 1);
        // If the next token is not another flag, treat it as the value
        if not StartsText('--', nextArg) then
          Exit(UnquoteIfQuoted(nextArg));
      end;
      // Bare flag -> "true"
      Exit('true');
    end;

    Inc(i);
  end;
end;

function HasFlag(const Name: string): Boolean;
var
  i: Integer;
begin
  for i := 1 to ParamCount do
    if SameText(ParamStr(i), '--' + Name) then
      Exit(True);
  Result := False;
end;

function ParseCli(out Opt: TCliOptions): Boolean;
var
  BatchStr: string;
begin
  FillChar(Opt, SizeOf(Opt), 0);
  Opt.PSTPath      := GetCmdValue('pst', '');
  Opt.ConnOverride := GetCmdValue('conn', '');
  Opt.EnvFile      := GetCmdValue('env-file', cDefaultEnvFile);
  Opt.LogPath      := GetCmdValue('log', cDefaultLogFile);
  Opt.AttachMode   := ParseAttMode(GetCmdValue('attachments', 'none'));
  Opt.QueryMode    := GetCmdValue('query','false').ToLower = 'true';
  Opt.FullSync     := GetCmdValue('fullsync','false').ToLower = 'true';
  Opt.MovesOnly    := GetCmdValue('movesonly','false').ToLower = 'true';
  Opt.Incremental  := GetCmdValue('incremental','false').ToLower = 'true';
  Opt.UseRestrict  := GetCmdValue('userestrict','false').ToLower = 'true';
  BatchStr         := GetCmdValue('batch', IntToStr(cDefaultBatch));
  if not TryStrToInt(BatchStr, Opt.Batch) then
    Opt.Batch := cDefaultBatch;
  // --pst is optional; will detect active store/folder when absent
  Result := True;
end;

procedure PrintUsage;
begin
  Writeln('Usage: pst_to_sql [--pst=<path>|active] [--attachments=none|meta|meta-hash|bytes] [--batch=N] [--fullsync]');
  Writeln('       [--conn="<FireDAC params>"] [--env-file=.env] [--log=ingest.log] [--query] [--movesonly] [--incremental] [--userestrict]');
  Writeln('       [--guardstate] [--setguardstate]');
end;

var
  Opt: TCliOptions;
  MessagesIngested: Int64;
begin
  ReportMemoryLeaksOnShutdown := False;

  if HasFlag('help') or HasFlag('h') then
  begin
    PrintUsage;
    Exit;
  end;

  if GetCmdValue('setguardstate','false').ToLower = 'true' then
  begin
    SetGuardState;
    Exit;
  end;

  if GetCmdValue('guardstate','false').ToLower = 'true' then
  begin
    CheckGuardState;
    Exit;
  end;

  CoInitialize(nil);
  try
    if not ParseCli(Opt) then
      raise Exception.Create('Failed to parse command line options.');

    if Opt.PSTPath = '' then
    begin
      WriteLn('Digite o caminho para o arquivo PST ou deixe em branco para utilizar o arquivo PST ativo no Outlook. ');
      Readln(Opt.PSTPath);
      Opt.PSTPath := Trim(Opt.PSTPath);
    end;

    if (Opt.PSTPath <> '') and (not SameText(Opt.PSTPath, 'active')) then
    begin
      Opt.PSTPath := TPath.GetFullPath(Opt.PSTPath);
      if not TFile.Exists(Opt.PSTPath) then
        raise Exception.Create('PST not found: ' + Opt.PSTPath);
    end;

    if OPt.FullSync or OPt.QueryMode or Opt.MovesOnly or opt.Incremental then
      RunIngest(Opt, MessagesIngested)
    else
      RunComSyncWorkerConsole(Opt, 300);
  except
    on E: Exception do
    begin
      Writeln('ERROR ' + E.ClassName + ': ' + E.Message);
      Halt(1);
    end;
  end;
  CoUninitialize;
end.


