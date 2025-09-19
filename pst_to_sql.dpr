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
  mail.Logger in 'mail.Logger.pas',
  mail.OutlookEvents in 'mail.OutlookEvents.pas';

const
  cDefaultEnvFile = '.env';
  cDefaultLogFile = 'ingest.log';
  cDefaultBatch   = 500;

procedure CheckGuardState;
var
  FOutlookVersion: String;
  R: TOutlookGuardResult;
begin
  FOutlookVersion:= GetLatestOutlookOfficeVersion;
  Writeln('Latest Outlook Version: ', FOutlookVersion);
  R := ReadEffectiveOutlookObjectModelGuard(FOutlookVersion);

  Writeln('Mode: ', Ord(R.Mode));
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

  Readln;
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
  Writeln('       [--guardstate]');
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


