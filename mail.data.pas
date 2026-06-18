unit mail.data;

interface

uses
  System.SysUtils, Data.DB,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param, FireDAC.Stan.Error,
  FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf, FireDAC.Stan.Async,
  FireDAC.Stan.Pool, FireDAC.Stan.Def, FireDAC.DApt, FireDAC.UI.Intf,
  FireDAC.Phys, FireDAC.Phys.MSSQL, FireDAC.Phys.MSSQLDef, FireDAC.Comp.Client,
  Mail.TypesUtils, System.Classes;

type
  TDb = class
  private
    FConn: TFDConnection;
    FQ: TFDQuery;

    procedure LoadEnvFromExecutableFolder(const EnvFile: string);

    procedure EnsureParam(const AName: string; ADataType: TFieldType);
    procedure SetParam(const AName: string; ADataType: TFieldType; const AValue: Variant);
    procedure BindParamsByName(const Names: array of string; const Values: array of Variant);
    procedure UpdatePSTLastUpdatedData(const pst_id: Integer);
    procedure LoadConnectionParamFromEnv;
    procedure LoadConnectionParamsFromConnOverride(const AConnOverride: String);
    procedure ConnectionAfterOpen(Sender: TObject);

    property Q: TFDQuery read FQ;
  public
    constructor Create; overload;
    constructor Create(const ConnOverride, EnvFile: string); overload;
    constructor Create(const AConnectionParams: TStringList); overload;
    destructor Destroy; override;

    procedure Commit;
    procedure Rollback;

    function ExecScalarIntN(const SQL: string; const Names: array of string; const Values: array of Variant): Integer;
    procedure ExecSQLN(const SQL: string; const Names: array of string; const Values: array of Variant);

    function EnsurePstRow(const PstPath, RootDisplay: string): Integer;
    function EnsureFolderRow(const PstId: Integer; const ParentId: Variant; const Name, FullPath: string; const Depth: Integer): Integer;

    function MessageExistsByMsgId(const PstId: Integer; const MsgId: string): Integer;
    function MessageExistsByEntryId(const PstId: Integer; const EntryID: string): Integer;

    function InsertMessageReturnId(
      const PstId, FolderId: Integer;
      const InternetMessageId, OutlookEntryID, Subject, SenderName, SenderEmail,
            DisplayTo, DisplayCc: string;
      const SentUTC, ReceivedUTC, CreatedUTC: TDateTime;
      const TransportHeaders, BodyText, BodyHtml: string;
      const SizeBytes: Variant; const LastModUTC: TDateTime; const SearchKey: TBytes): Integer;

    procedure UpdateMessageCoreFields(
      const MessageId: Integer;
      const InternetMessageId, OutlookEntryID, Subject, SenderName, SenderEmail,
            DisplayTo, DisplayCc: string;
      const SentUTC, ReceivedUTC, CreatedUTC: TDateTime;
      const TransportHeaders, BodyText, BodyHtml: string;
      const LastModUtc: TDateTime; const SearchKey: TBytes);

    procedure InsertRecipient(const MessageId, Kind: Integer; const DisplayName, Email: string);

    // UPDATED: now uses ArrayDML + set-based insert under the hood (single-message batch)
    procedure SaveRecipients(const MessageId: Integer; const RecRows: TArray<TArray<Variant>>); overload;

    procedure SaveRecipients(const MessageIds: TArray<Integer>; const RecRowsPerMessage: TArray<TArray<TArray<Variant>>>); overload;

    procedure InsertAttachment(const MessageId: Integer; const FileName, MimeType: string;
      const SizeBytes: Variant; const Sha256Hex: string; const Content: TByteStr);

    procedure OpenHeuristicCandidatesNoSentUtc(
      const PstId: Integer; const Subject, SenderEmail: string);

    function IsDuplicateKeyError(const E: EFDDBEngineException): Boolean;

    { --------- Data accessors / utilities moved from mail.FullSync --------- }
    procedure UpdateMessageEnrichment(const MessageId: Integer; const InternetMessageId, EntryId, DisplayTo, DisplayCc, Headers: string);
    function HasRecipientForMessage(const MessageId: Integer): Boolean;
    function GetScanCutoffLastModUtc(const PstId: Integer): TDateTime;
    function TryGetMessageLastModUtc(const MessageId: Integer; out LastModUtc: TDateTime): Boolean;

    // NEW: bulk helpers used by FullSync (moved here to keep DB logic in data layer)
    function ResolveMessageIdsBySearchKeys(const PstId: Integer; const Keys: TArray<TBytes>): TArray<Integer>;
    function BatchMessagesHaveRecipients(const MsgIds: TArray<Integer>): TArray<Boolean>;

    property Conn: TFDConnection read FConn;
  end;

procedure OpenOrAutoRollback(AQuery: TFDQuery; ASql: String='');

const
  OPT_SAVE_BODY_HTML = False;

implementation

uses
  System.StrUtils, System.IOUtils, FireDAC.Comp.DataSet, Variants;

procedure OpenOrAutoRollback(AQuery: TFDQuery; ASql: String='');
begin
  try
    if ASql<>'' then
      AQuery.Open(ASql)
    else
      AQuery.Open;

  except
    if AQuery.Connection.InTransaction then
      AQuery.Connection.Rollback;

    raise;
  end;

end;

function VIfThen(AValue: Boolean; ATrue: Variant; AFalse: Variant): Variant;
begin
  if AValue then
    Result := ATrue
  else
    Result := AFalse;
end;

function EnvToBool(const S: string): Boolean;
var
  V: string;
begin
  V := LowerCase(Trim(S));
  Result := (V = 'yes') or (V = 'true') or (V = '1') or (V = 'y');
end;

procedure TDb.ConnectionAfterOpen(Sender: TObject);
begin
  FConn.ExecSQL('SET IMPLICIT_TRANSACTIONS OFF; SET XACT_ABORT ON; SET LOCK_TIMEOUT 30000;');
end;

constructor TDb.Create;
begin
  inherited Create;
  FConn := TFDConnection.Create(nil);
  FConn.LoginPrompt := False;
  FConn.AfterConnect:= ConnectionAfterOpen;

  FConn.TxOptions.AutoCommit := True;      // commit after every statement
  FConn.TxOptions.AutoStart  := False;     // do not wrap SELECTs in long txns
  FConn.TxOptions.Isolation  := xiReadCommitted; // relies on RCSI (see §1)
//  FConn.TxOptions.DisconnectAction := xdRollback;

  FQ := TFDQuery.Create(nil);
  FQ.Connection := FConn;
end;

constructor TDb.Create(const ConnOverride, EnvFile: string);
begin
  Create;

  if ConnOverride <> '' then
  begin
    LoadConnectionParamsFromConnOverride(ConnOverride);
  end
  else
  begin
    LoadEnvFromExecutableFolder(EnvFile);
    LoadConnectionParamFromEnv;
  end;

  FConn.Connected := True;
end;

constructor TDb.Create(const AConnectionParams: TStringList);
begin
  Create;

  FConn.Params.Assign(AConnectionParams);

  FConn.Connected:= True;
end;

procedure TDb.LoadEnvFromExecutableFolder(const EnvFile: string);
var
  BaseDir, Candidate, ExplicitPath: string;
  HasLoaded: Boolean;
begin
  HasLoaded := False;
  BaseDir := ExtractFilePath(ParamStr(0));

  if EnvFile <> '' then
  begin
    if TPath.IsPathRooted(EnvFile) then
      ExplicitPath := EnvFile
    else
      ExplicitPath := TPath.Combine(BaseDir, EnvFile);

    if TFile.Exists(ExplicitPath) then
    begin
      LoadEnvFile(ExplicitPath);
      HasLoaded := True;
    end;
  end;

  if not HasLoaded then
  begin
    Candidate := TPath.Combine(BaseDir, '.env');
    if TFile.Exists(Candidate) then
      LoadEnvFile(Candidate);
  end;
end;

procedure TDb.LoadConnectionParamFromEnv;
var
  ConnStr: string;
  Pairs: TArray<string>;
  P: string;
  Eq: Integer;
begin
  ConnStr := GetEnv('DB_CONN_STRING', '');
  if ConnStr <> '' then
  begin
    FConn.Params.Clear;
    Pairs := ConnStr.Split([';'], TStringSplitOptions.ExcludeEmpty);
    for P in Pairs do
    begin
      Eq := Pos('=', P);
      if Eq > 0 then
        FConn.Params.Values[Trim(Copy(P, 1, Eq - 1))] := Trim(Copy(P, Eq + 1, MaxInt));
    end;
    if FConn.Params.Values['DriverID'] = '' then
      FConn.Params.Values['DriverID'] := 'MSSQL';
  end
  else
  begin
    FConn.Params.Clear;
    FConn.Params.Values['DriverID'] := 'MSSQL';
    FConn.Params.Values['Server']   := GetEnv('SQL_SERVER', '');
    FConn.Params.Values['Database'] := GetEnv('SQL_DATABASE', '');
    if (FConn.Params.Values['Server'] = '') or (FConn.Params.Values['Database'] = '') then
      raise Exception.Create('Missing SQL_SERVER or SQL_DATABASE and no DB_CONN_STRING provided.');

    if EnvToBool(GetEnv('SQL_TRUSTED_CONNECTION', 'no')) then
      FConn.Params.Values['OSAuthent'] := 'Yes'
    else
    begin
      FConn.Params.Values['User_Name'] := GetEnv('SQL_USERNAME', '');
      FConn.Params.Values['Password']  := GetEnv('SQL_PASSWORD', '');
      if (FConn.Params.Values['User_Name'] = '') or (FConn.Params.Values['Password'] = '') then
        raise Exception.Create('For SQL auth, set SQL_USERNAME and SQL_PASSWORD (or use SQL_TRUSTED_CONNECTION=Yes).');
    end;

    if GetEnv('SQL_ENCRYPT', '') <> '' then
      FConn.Params.Values['Encrypt'] := GetEnv('SQL_ENCRYPT', 'yes');
    if GetEnv('SQL_TRUST_SERVER_CERTIFICATE', '') <> '' then
      FConn.Params.Values['ODBCAdvanced'] := 'TrustServerCertificate=yes';
  end;
end;

procedure TDB.LoadConnectionParamsFromConnOverride(const AConnOverride: String);
var
  Pairs: TArray<string>;
  P: string;
  Eq: Integer;
begin
  FConn.Params.Clear;
  Pairs := AConnOverride.Split([';'], TStringSplitOptions.ExcludeEmpty);
  for P in Pairs do
  begin
    Eq := Pos('=', P);
    if Eq > 0 then
      FConn.Params.Values[Trim(Copy(P, 1, Eq - 1))] := Trim(Copy(P, Eq + 1, MaxInt));
  end;
  if FConn.Params.Values['DriverID'] = '' then
    FConn.Params.Values['DriverID'] := 'MSSQL';

end;

destructor TDb.Destroy;
begin
  FQ.Free;
  FConn.Free;
  inherited;
end;

function TDb.GetScanCutoffLastModUtc(const PstId: Integer): TDateTime;
var
  hasMax: Boolean;
begin
  Result := 0;

  FQ.Close;
  FQ.SQL.Text:= 'SELECT MAX(last_modification_time) FROM mail.message WHERE pst_id = :pid';
  FQ.Params.Clear;
  SetParam('pid', ftInteger, PstId);
  FQ.Open;

  hasMax := not FQ.Fields[0].IsNull;
  if hasMax then
    Result := FQ.Fields[0].AsDateTime
  else
    Result := 0;

  FQ.Close;
end;

procedure TDb.Commit;
begin
  if FConn.InTransaction then
    FConn.Commit;
end;

procedure TDb.Rollback;
begin
  if FConn.InTransaction then
    FConn.Rollback;
end;

procedure TDb.EnsureParam(const AName: string; ADataType: TFieldType);
var
  Pm: TFDParam;
begin
  Pm := FQ.Params.FindParam(AName);
  if Pm = nil then
    Pm := FQ.Params.CreateParam(ADataType, AName, ptInput)
  else
    Pm.DataType := ADataType;
end;

procedure TDb.SetParam(const AName: string; ADataType: TFieldType; const AValue: Variant);
var
  Pm: TFDParam;
begin
  EnsureParam(AName, ADataType);
  Pm := FQ.Params.ParamByName(AName);
  if VarIsNull(AValue) or VarIsEmpty(AValue) then
    Pm.Clear
  else
    Pm.Value := AValue;
end;

procedure TDb.BindParamsByName(const Names: array of string; const Values: array of Variant);
var
  i: Integer;
  Pm: TFDParam;
begin
  if Length(Names) <> Length(Values) then
    raise Exception.CreateFmt('Parameter name/value length mismatch (%d vs %d)', [Length(Names), Length(Values)]);
  for i := 0 to High(Names) do
  begin
    Pm := FQ.Params.FindParam(Names[i]);
    if Pm = nil then
      Pm := FQ.Params.CreateParam(ftUnknown, Names[i], ptInput);
    Pm.Value := Values[i];
  end;
end;

function TDb.ExecScalarIntN(const SQL: string; const Names: array of string; const Values: array of Variant): Integer;
begin
  FQ.Close;
  FQ.SQL.Text := SQL;
  FQ.Params.Clear;
  BindParamsByName(Names, Values);
  FQ.Open;
  if not FQ.Eof then
    Result := FQ.Fields[0].AsInteger
  else
    Result := 0;
  FQ.Close;
end;

procedure TDb.ExecSQLN(const SQL: string; const Names: array of string; const Values: array of Variant);
begin
  FQ.Close;
  FQ.SQL.Text := SQL;
  FQ.Params.Clear;
  BindParamsByName(Names, Values);
  FQ.ExecSQL;
end;

procedure TDb.UpdatePSTLastUpdatedData(const pst_id: Integer);
begin
  FQ.Close;
  FQ.SQL.Text := 'UPDATE mail.pst_file SET last_updated_date = SYSUTCDATETIME() WHERE pst_id = :id';
  FQ.Params.Clear;
  SetParam('id', ftInteger, pst_id);
  FQ.ExecSQL;
end;

function TDb.EnsurePstRow(const PstPath, RootDisplay: string): Integer;
begin
  Result := ExecScalarIntN(
    'SELECT pst_id FROM mail.pst_file WHERE LOWER(pst_path) = LOWER(:p)',
    ['p'], [PstPath]);

  if Result <> 0 then
  begin
    if RootDisplay <> '' then
    begin
      FQ.Close;
      FQ.SQL.Text := 'UPDATE mail.pst_file SET root_display = :r WHERE pst_id = :id';
      FQ.Params.Clear;
      SetParam('r',  ftWideString, RootDisplay);
      SetParam('id', ftInteger,    Result);
      FQ.ExecSQL;
    end;

    UpdatePSTLastUpdatedData(Result);
    Exit;
  end;

  FQ.Close;
  FQ.SQL.Text :=
    'INSERT INTO mail.pst_file(pst_path,root_display) OUTPUT INSERTED.pst_id VALUES (:p,:r)';
  FQ.Params.Clear;
  SetParam('p', ftWideString, PstPath);
  SetParam('r', ftWideString, RootDisplay);
  FQ.Open;
  if not FQ.Eof then
    Result := FQ.Fields[0].AsInteger
  else
    raise Exception.Create('Failed to insert mail.pst_file');
  FQ.Close;

  UpdatePSTLastUpdatedData(Result);

  Commit;
end;

function TDb.EnsureFolderRow(const PstId: Integer; const ParentId: Variant;
  const Name, FullPath: string; const Depth: Integer): Integer;
begin
  Result := ExecScalarIntN(
    'SELECT folder_id FROM mail.folder WHERE pst_id = :pid AND LOWER(full_path) = LOWER(:fp)',
    ['pid','fp'], [PstId, FullPath]);

  if Result <> 0 then
  begin
    FQ.Close;
    FQ.SQL.Text :=
      'UPDATE mail.folder SET ' +
      '  parent_folder_id = COALESCE(parent_folder_id, :par), ' +
      '  name = COALESCE(name, :nm), ' +
      '  depth = COALESCE(depth, :dp) ' +
      'WHERE folder_id = :id';
    FQ.Params.Clear;
    SetParam('par', ftInteger,    ParentId);
    SetParam('nm',  ftWideString, Name);
    SetParam('dp',  ftInteger,    Depth);
    SetParam('id',  ftInteger,    Result);
    FQ.ExecSQL;

    Exit;
  end;

  try
    FQ.Close;
    FQ.SQL.Text :=
      'INSERT INTO mail.folder(pst_id,parent_folder_id,name,full_path,depth) ' +
      'OUTPUT INSERTED.folder_id VALUES (:pid,:par,:nm,:fp,:dp)';
    FQ.Params.Clear;
    SetParam('pid', ftInteger,    PstId);
    SetParam('par', ftInteger,    ParentId);
    SetParam('nm',  ftWideString, Name);
    SetParam('fp',  ftWideString, FullPath);
    SetParam('dp',  ftInteger,    Depth);
    FQ.Open;
    if not FQ.Eof then
      Result := FQ.Fields[0].AsInteger
    else
      raise Exception.Create('Failed to insert mail.folder');
    FQ.Close;
  except
    on E: EFDDBEngineException do
    begin
      Result := ExecScalarIntN(
        'SELECT folder_id FROM mail.folder WHERE pst_id = :pid AND LOWER(full_path) = LOWER(:fp)',
        ['pid','fp'], [PstId, FullPath]);
      if Result = 0 then
        raise;
    end;
  end;
end;

function TDb.MessageExistsByMsgId(const PstId: Integer; const MsgId: string): Integer;
begin
  Result := 0;
  if MsgId = '' then Exit;

  FQ.Close;
  FQ.SQL.Text :=
    'SELECT m.message_id FROM mail.message m ' +
    'WHERE m.pst_id = :pid AND LOWER(COALESCE(m.internet_message_id, '''')) = LOWER(:mid)';
  FQ.Params.Clear;
  SetParam('pid', ftInteger,    PstId);
  SetParam('mid', ftWideString, MsgId);
  FQ.Open;
  if not FQ.Eof then
    Result := FQ.Fields[0].AsInteger;
  FQ.Close;
end;

function TDb.MessageExistsByEntryId(const PstId: Integer; const EntryID: string): Integer;
begin
  Result := 0;
  if EntryID = '' then Exit;

  FQ.Close;
  FQ.SQL.Text :=
    'SELECT m.message_id FROM mail.message m ' +
    'WHERE m.pst_id = :pid AND LOWER(COALESCE(m.outlook_entry_id, '''')) = LOWER(:eid)';
  FQ.Params.Clear;
  SetParam('pid', ftInteger,    PstId);
  SetParam('eid', ftWideString, EntryID);
  FQ.Open;
  if not FQ.Eof then
    Result := FQ.Fields[0].AsInteger;
  FQ.Close;
end;

function TDb.InsertMessageReturnId(
  const PstId, FolderId: Integer;
  const InternetMessageId, OutlookEntryID, Subject, SenderName, SenderEmail,
        DisplayTo, DisplayCc: string;
  const SentUTC, ReceivedUTC, CreatedUTC: TDateTime;
  const TransportHeaders, BodyText, BodyHtml: string;
  const SizeBytes: Variant; const LastModUTC: TDateTime; const SearchKey: TBytes): Integer;
begin
  FQ.Close;
  FQ.SQL.Text :=
    'INSERT INTO mail.message ' +
    '(pst_id, folder_id, internet_message_id, outlook_entry_id, subject, sender_name, sender_email, ' +
    ' display_to, display_cc, sent_utc, received_utc, created_utc, ' +
    ' transport_headers, body_text, body_html, size_bytes, last_modification_time, search_key) ' +
    'VALUES (:pid,:fid,:mid,:eid,:subj,:sname,:semail,:dto,:dcc,:sent,:recv,:crt,:hdr,:btxt,:bhtm,:sz,:lmod,:sk);'+
    'SELECT CAST(SCOPE_IDENTITY() AS INT);';
  FQ.Params.Clear;

  SetParam('pid',    ftInteger,    PstId);
  SetParam('fid',    ftInteger,    FolderId);
  SetParam('mid',    ftWideString, VIfThen(InternetMessageId <> '', InternetMessageId, Null));
  SetParam('eid',    ftWideString, VIfThen(OutlookEntryID <> '', OutlookEntryID, Null));
  SetParam('subj',   ftWideString, Subject);
  SetParam('sname',  ftWideString, VIfThen(SenderName <> '', SenderName, Null));
  SetParam('semail', ftWideString, VIfThen(SenderEmail <> '', SenderEmail, Null));
  SetParam('dto',    ftWideString, VIfThen(DisplayTo <> '', DisplayTo, Null));
  SetParam('dcc',    ftWideString, VIfThen(DisplayCc <> '', DisplayCc, Null));
  SetParam('sent',   ftDateTime,   VIfThen(SentUTC <> 0, SentUTC, Null));
  SetParam('recv',   ftDateTime,   VIfThen(ReceivedUTC <> 0, ReceivedUTC, Null));
  SetParam('crt',    ftDateTime,   VIfThen(CreatedUTC <> 0, CreatedUTC, Null));
  SetParam('hdr',    ftWideMemo,   VIfThen(TransportHeaders <> '', TransportHeaders, Null));
  SetParam('btxt',   ftWideMemo,   VIfThen(BodyText <> '', BodyText, Null));
  if OPT_SAVE_BODY_HTML then
    SetParam('bhtm', ftWideMemo,   VIfThen(BodyHtml <> '', BodyHtml, Null))
  else
    SetParam('bhtm', ftWideMemo,   Null);
  SetParam('sz',     ftInteger,    VIfThen(VarIsNull(SizeBytes), Null, SizeBytes));
  SetParam('lmod',   ftDateTime,   VIfThen(LastModUTC <> 0, LastModUTC, Null));

  EnsureParam('sk', ftVarBytes);
  if Length(SearchKey) = 0 then
    FQ.ParamByName('sk').Clear
  else
    FQ.ParamByName('sk').SetData(PByte(@SearchKey[0]), LongWord(Length(SearchKey)));

  FQ.Open;
  if not FQ.Eof then
    Result := FQ.Fields[0].AsInteger
  else
    raise Exception.Create('Failed to insert mail.message');
  FQ.Close;
end;

procedure TDb.SaveRecipients(const MessageId: Integer; const RecRows: TArray<TArray<Variant>>);
var
  N, i: Integer;
  Q, Ins: TFDQuery;
begin
  N := Length(RecRows);
  if N = 0 then
    Exit;

  // New fast path: ArrayDML into a temp table, then set-based insert with de-duplication.
  Q   := TFDQuery.Create(nil);
  Ins := TFDQuery.Create(nil);
  try
    try
      Q.Connection   := FConn;
      Ins.Connection := FConn;

      // Temp table to stage recipients for this message
      Q.SQL.Text := 'CREATE TABLE #rec (kind INT NOT NULL, display_name NVARCHAR(400) NULL, email NVARCHAR(320) NULL)';
      Q.ExecSQL;

      // ArrayDML into #rec
      Ins.SQL.Text := 'INSERT INTO #rec(kind, display_name, email) VALUES (:k, :d, :e)';
      Ins.Params.Clear;
      Ins.Params.Add.Name := 'k';
      Ins.ParamByName('k').DataType := ftInteger;
      Ins.Params.Add.Name := 'd';
      Ins.ParamByName('d').DataType := ftWideString;
      Ins.Params.Add.Name := 'e';
      Ins.ParamByName('e').DataType := ftWideString;

      Ins.Params.ArraySize := N;
      for i := 0 to N - 1 do
      begin
        // kind
        if VarIsNull(RecRows[i][0]) or VarIsEmpty(RecRows[i][0]) then
          Ins.ParamByName('k').AsIntegers[i] := 1
        else
          Ins.ParamByName('k').AsIntegers[i] := Integer(RecRows[i][0]);

        // display_name
        if VarToStrDefSafe(RecRows[i][1], '') <> '' then
          Ins.ParamByName('d').AsStrings[i] := VarToStrDefSafe(RecRows[i][1], '')
        else
          Ins.ParamByName('d').Clear(i);

        // email
        if VarToStrDefSafe(RecRows[i][2], '') <> '' then
          Ins.ParamByName('e').AsStrings[i] := VarToStrDefSafe(RecRows[i][2], '')
        else
          Ins.ParamByName('e').Clear(i);
      end;
      Ins.Execute(N, 0);

      // Insert distinct staged rows that do not already exist for this message (NULL-safe email match, case-insensitive)
      Q.SQL.Text :=
        'WITH x AS (' + sLineBreak +
        '  SELECT kind, display_name, email,' + sLineBreak +
        '         ROW_NUMBER() OVER (PARTITION BY kind, LOWER(COALESCE(email, '''')) ORDER BY (SELECT 0)) rn' + sLineBreak +
        '  FROM #rec' + sLineBreak +
        ')' + sLineBreak +
        'INSERT INTO mail.recipient(message_id, kind, display_name, email)' + sLineBreak +
        'SELECT :m, x.kind, x.display_name, x.email' + sLineBreak +
        'FROM x' + sLineBreak +
        'WHERE x.rn = 1' + sLineBreak +
        '  AND NOT EXISTS (' + sLineBreak +
        '        SELECT 1 FROM mail.recipient r' + sLineBreak +
        '        WHERE r.message_id = :m' + sLineBreak +
        '          AND r.kind = x.kind' + sLineBreak +
        '          AND ( (r.email IS NULL AND x.email IS NULL)' + sLineBreak +
        '             OR (r.email IS NOT NULL AND x.email IS NOT NULL AND LOWER(r.email) = LOWER(x.email)) )' + sLineBreak +
        '      )';
      Q.Params.Clear;
      Q.Params.Add.Name := 'm';
      Q.ParamByName('m').DataType := ftInteger;
      Q.ParamByName('m').AsInteger := MessageId;
      Q.ExecSQL;

    except
      // Conservative fallback to the previous row-by-row path if anything goes wrong.
      on E: Exception do
      begin
        try
          // best-effort cleanup of temp table if it exists
          Q.SQL.Text := 'IF OBJECT_ID(''tempdb..#rec'') IS NOT NULL DROP TABLE #rec';
          try Q.ExecSQL; except end;
        except end;

        // Fallback to the old logic (preserves behavior & avoids breaking existing flows)
        for i := 0 to N - 1 do
          InsertRecipient(
            MessageId,
            VarToIntDefSafe(RecRows[i][0], 1),
            VarToStrDefSafe(RecRows[i][1], ''),
            VarToStrDefSafe(RecRows[i][2], ''));
      end;
    end;
  finally
    // Drop temp table explicitly
    try Q.SQL.Text := 'DROP TABLE #rec'; Q.ExecSQL; except end;
    Ins.Free;
    Q.Free;
  end;


end;

// NEW: multi-message batch variant
procedure TDb.SaveRecipients(const MessageIds: TArray<Integer>;
  const RecRowsPerMessage: TArray<TArray<TArray<Variant>>>);
const
  CHUNK = 2000;
var
  Q, Ins: TFDQuery;
  i, total, done, nThis, iThis: Integer;
  mIdx, rIdx: Integer;
  row: TArray<Variant>;

  procedure DropTemp;
  begin
    try
      Q.SQL.Text := 'IF OBJECT_ID(''tempdb..#rec2'') IS NOT NULL DROP TABLE #rec2';
      Q.ExecSQL;
    except
      // ignore
    end;
  end;

begin
  if Length(MessageIds) <> Length(RecRowsPerMessage) then
    raise Exception.Create('SaveRecipients: parameter length mismatch (MessageIds vs RecRowsPerMessage).');

  total := 0;
  for i := 0 to High(MessageIds) do
    Inc(total, Length(RecRowsPerMessage[i]));
  if total = 0 then
    Exit;

  Q   := TFDQuery.Create(nil);
  Ins := TFDQuery.Create(nil);
  try
    try
      Q.Connection   := FConn;
      Ins.Connection := FConn;

      // Staging table for all recipients (across many messages)
      Q.SQL.Text :=
        'CREATE TABLE #rec2(' +
        '  message_id INT NOT NULL,' +
        '  kind INT NOT NULL,' +
        '  display_name NVARCHAR(400) NULL,' +
        '  email NVARCHAR(320) NULL' +
        ')';
      Q.ExecSQL;

      // ArrayDML into the staging table
      Ins.SQL.Text := 'INSERT INTO #rec2(message_id, kind, display_name, email) VALUES (:m,:k,:d,:e)';
      Ins.Params.Clear;
      Ins.Params.Add.Name := 'm'; Ins.ParamByName('m').DataType := ftInteger;
      Ins.Params.Add.Name := 'k'; Ins.ParamByName('k').DataType := ftInteger;
      Ins.Params.Add.Name := 'd'; Ins.ParamByName('d').DataType := ftWideString;
      Ins.Params.Add.Name := 'e'; Ins.ParamByName('e').DataType := ftWideString;

      done := 0;
      mIdx := 0; rIdx := 0;

      while done < total do
      begin
        nThis := CHUNK;
        if nThis > (total - done) then
          nThis := total - done;

        Ins.Params.ArraySize := nThis;
        iThis := 0;

        while iThis < nThis do
        begin
          // advance to next message that still has rows
          while (mIdx < Length(MessageIds)) and (rIdx >= Length(RecRowsPerMessage[mIdx])) do
          begin
            Inc(mIdx);
            rIdx := 0;
          end;
          if mIdx >= Length(MessageIds) then
            Break; // safety

          row := RecRowsPerMessage[mIdx][rIdx];

          // message_id
          Ins.ParamByName('m').AsIntegers[iThis] := MessageIds[mIdx];

          // kind
          if VarIsNull(row[0]) or VarIsEmpty(row[0]) then
            Ins.ParamByName('k').AsIntegers[iThis] := 1
          else
            Ins.ParamByName('k').AsIntegers[iThis] := Integer(row[0]);

          // display_name
          if VarToStrDefSafe(row[1], '') <> '' then
            Ins.ParamByName('d').AsStrings[iThis] := VarToStrDefSafe(row[1], '')
          else
            Ins.ParamByName('d').Clear(iThis);

          // email
          if VarToStrDefSafe(row[2], '') <> '' then
            Ins.ParamByName('e').AsStrings[iThis] := VarToStrDefSafe(row[2], '')
          else
            Ins.ParamByName('e').Clear(iThis);

          Inc(rIdx);
          Inc(iThis);
          Inc(done);
        end;

        if iThis > 0 then
          Ins.Execute(iThis, 0);
      end;

      // One set-based insert with per-(message_id, kind, email) dedup + NOT EXISTS against target
      Q.SQL.Text :=
        'WITH x AS (' + sLineBreak +
        '  SELECT message_id, kind, display_name, email,' + sLineBreak +
        '         ROW_NUMBER() OVER (PARTITION BY message_id, kind, LOWER(COALESCE(email, ''''))' + sLineBreak +
        '                             ORDER BY (SELECT 0)) rn' + sLineBreak +
        '  FROM #rec2' + sLineBreak +
        ')' + sLineBreak +
        'INSERT INTO mail.recipient(message_id, kind, display_name, email)' + sLineBreak +
        'SELECT x.message_id, x.kind, x.display_name, x.email' + sLineBreak +
        'FROM x' + sLineBreak +
        'WHERE x.rn = 1' + sLineBreak +
        '  AND NOT EXISTS (' + sLineBreak +
        '        SELECT 1 FROM mail.recipient r' + sLineBreak +
        '        WHERE r.message_id = x.message_id' + sLineBreak +
        '          AND r.kind = x.kind' + sLineBreak +
        '          AND ( (r.email IS NULL AND x.email IS NULL)' + sLineBreak +
        '             OR (r.email IS NOT NULL AND x.email IS NOT NULL AND LOWER(r.email) = LOWER(x.email)) )' + sLineBreak +
        '      )';
      Q.ExecSQL;

    except
      on E: Exception do
      begin
        // Clean up temp table if it exists
        try DropTemp; except end;

        // Fallback to safe row-by-row path (keeps behavior intact)
        for i := 0 to High(MessageIds) do
          Self.SaveRecipients(MessageIds[i], RecRowsPerMessage[i]);
      end;
    end;
  finally
    DropTemp;
    Ins.Free;
    Q.Free;
  end;


end;

procedure TDb.UpdateMessageCoreFields(
  const MessageId: Integer;
  const InternetMessageId, OutlookEntryID, Subject, SenderName, SenderEmail,
        DisplayTo, DisplayCc: string;
  const SentUTC, ReceivedUTC, CreatedUTC: TDateTime;
  const TransportHeaders, BodyText, BodyHtml: string;
  const LastModUtc: TDateTime; const SearchKey: TBytes);
begin
  FQ.Close;
  FQ.SQL.Text :=
    'UPDATE mail.message SET ' +
    '  internet_message_id   = :mid, ' +
    '  outlook_entry_id      = :eid, ' +
    '  subject               = :subj, ' +
    '  sender_name           = :sname, ' +
    '  sender_email          = :semail, ' +
    '  display_to            = :dto, ' +
    '  display_cc            = :dcc, ' +
    '  sent_utc              = :sent, ' +
    '  received_utc          = :recv, ' +
    '  created_utc           = :crt, ' +
    '  transport_headers     = :hdr, ' +
    '  body_text             = :btxt, ' +
    '  body_html             = :bhtm, ' +
    '  last_modification_time= :lmod, ' +
    '  search_key            = :sk ' +
    'WHERE message_id = :id';
  FQ.Params.Clear;
  SetParam('mid',  ftWideString, VIfThen(InternetMessageId <> '', InternetMessageId, Null));
  SetParam('eid',  ftWideString, VIfThen(OutlookEntryID <> '', OutlookEntryID, Null));
  SetParam('subj', ftWideString, Subject);
  SetParam('sname',ftWideString, VIfThen(SenderName <> '', SenderName, Null));
  SetParam('semail',ftWideString,VIfThen(SenderEmail <> '', SenderEmail, Null));
  SetParam('dto',  ftWideString, VIfThen(DisplayTo <> '', DisplayTo, Null));
  SetParam('dcc',  ftWideString, VIfThen(DisplayCc <> '', DisplayCc, Null));
  SetParam('sent', ftDateTime,   VIfThen(SentUTC <> 0, SentUTC, Null));
  SetParam('recv', ftDateTime,   VIfThen(ReceivedUTC <> 0, ReceivedUTC, Null));
  SetParam('crt',  ftDateTime,   VIfThen(CreatedUTC <> 0, CreatedUTC, Null));
  SetParam('hdr',  ftWideMemo,   VIfThen(TransportHeaders <> '', TransportHeaders, Null));
  SetParam('btxt', ftWideMemo,   VIfThen(BodyText <> '', BodyText, Null));
  if OPT_SAVE_BODY_HTML then
    SetParam('bhtm', ftWideMemo, VIfThen(BodyHtml <> '', BodyHtml, Null))
  else
    SetParam('bhtm', ftWideMemo, Null);
  SetParam('id',   ftInteger,    MessageId);
  SetParam('lmod', ftDateTime,   VIfThen(LastModUtc <> 0, LastModUtc, Null));

  EnsureParam('sk', ftVarBytes);
  if Length(SearchKey) = 0 then
    FQ.ParamByName('sk').Clear
  else
    FQ.ParamByName('sk').SetData(PByte(@SearchKey[0]), LongWord(Length(SearchKey)));

  FQ.ExecSQL;
end;

procedure TDb.InsertRecipient(const MessageId, Kind: Integer; const DisplayName, Email: string);
begin
  // Retained for compatibility and as fallback.
  FQ.Close;
  FQ.SQL.Text :=
    'SELECT TOP 1 1 FROM mail.recipient ' +
    'WHERE message_id = :m AND kind = :k ' +
    '  AND ( (email IS NULL AND :e IS NULL) OR ' +
    '        (email IS NOT NULL AND :e IS NOT NULL AND LOWER(email) = LOWER(:e)) )';
  FQ.Params.Clear;
  SetParam('m', ftInteger,    MessageId);
  SetParam('k', ftInteger,    Kind);
  SetParam('e', ftWideString, VIfThen(Email <> '', Email, Null));
  FQ.Open;
  if not FQ.Eof then
  begin
    FQ.Close;
    Exit;
  end;
  FQ.Close;

  FQ.Close;
  FQ.SQL.Text := 'INSERT INTO mail.recipient(message_id,kind,display_name,email) VALUES (:m,:k,:d,:e)';
  FQ.Params.Clear;
  SetParam('m', ftInteger,    MessageId);
  SetParam('k', ftInteger,    Kind);
  SetParam('d', ftWideString, VIfThen(DisplayName <> '', DisplayName, Null));
  SetParam('e', ftWideString, VIfThen(Email <> '', Email, Null));
  FQ.ExecSQL;
end;

procedure TDb.InsertAttachment(const MessageId: Integer; const FileName, MimeType: string;
  const SizeBytes: Variant; const Sha256Hex: string; const Content: TByteStr);
begin
  FQ.Close;
  FQ.SQL.Text :=
    'INSERT INTO mail.attachment(message_id,file_name,mime_type,size_bytes,sha256,content) ' +
    'VALUES (:m,:f,:mt,:sz,:sha,:cnt)';
  FQ.Params.Clear;

  SetParam('m',   ftInteger,    MessageId);
  SetParam('f',   ftWideString, FileName);
  SetParam('mt',  ftWideString, VIfThen(MimeType <> '', MimeType, Null));
  SetParam('sz',  ftInteger,    VIfThen(VarIsNull(SizeBytes), Null, SizeBytes));
  SetParam('sha', ftWideString, VIfThen(Sha256Hex <> '', Sha256Hex, Null));

  EnsureParam('cnt', ftBlob);
  if Length(Content) > 0 then
    FQ.Params.ParamByName('cnt').AsByteStr := Content
  else
    FQ.Params.ParamByName('cnt').Clear;

  FQ.ExecSQL;
end;

procedure TDb.OpenHeuristicCandidatesNoSentUtc(
  const PstId: Integer; const Subject, SenderEmail: string);
begin
  FQ.Close;
  FQ.SQL.Text :=
    'SELECT TOP 50 message_id, display_to, display_cc ' +
    'FROM mail.message ' +
    'WHERE pst_id = :pid ' +
    '  AND LOWER(COALESCE(subject, '''')) = LOWER(:s) ' +
    '  AND LOWER(COALESCE(sender_email, '''')) = LOWER(:se) ' +
    '  AND sent_utc IS NULL';
  FQ.Params.Clear;
  SetParam('pid',  ftInteger,  PstId);
  SetParam('s',    ftWideString, Subject);
  SetParam('se',   ftWideString, SenderEmail);
  FQ.Open;
end;

function TDb.IsDuplicateKeyError(const E: EFDDBEngineException): Boolean;
begin
  Result := False;
  if E = nil then Exit;
  if E.ErrorCount > 0 then
  begin
    Result :=
      (E.Errors[0].Kind = ekUKViolated) or
      (E.Errors[0].ErrorCode = 2601) or
      (E.Errors[0].ErrorCode = 2627);
  end;
  if not Result then
  begin
    Result :=
      (Pos('2601', E.Message) > 0) or
      (Pos('2627', E.Message) > 0) or
      (Pos('UX_message_msgid_hash', E.Message) > 0);
  end;
end;

procedure TDb.UpdateMessageEnrichment(const MessageId: Integer; const InternetMessageId, EntryId, DisplayTo, DisplayCc, Headers: string);
begin
  ExecSQLN(
    'UPDATE mail.message SET ' +
    '  internet_message_id = COALESCE(internet_message_id, :mid), ' +
    '  outlook_entry_id     = COALESCE(outlook_entry_id, :eid), ' +
    '  display_to           = COALESCE(display_to, :dto), ' +
    '  display_cc           = COALESCE(display_cc, :dcc), ' +
    '  transport_headers    = COALESCE(transport_headers, :hdr), ' +
    '  com_enriched_at      = SYSUTCDATETIME() ' +
    'WHERE message_id = :id',
    ['mid','eid','dto','dcc','hdr','id'],
    [VIfThen(InternetMessageId <> '', InternetMessageId, Null),
     VIfThen(EntryId <> '', EntryId, Null),
     VIfThen(DisplayTo <> '', DisplayTo, Null),
     VIfThen(DisplayCc <> '', DisplayCc, Null),
     VIfThen(Headers <> '', Headers, Null),
     MessageId]);
end;

function TDb.HasRecipientForMessage(const MessageId: Integer): Boolean;
begin
  Result := ExecScalarIntN(
              'SELECT TOP 1 1 FROM mail.recipient WHERE message_id = :m',
              ['m'], [MessageId]) <> 0;
end;

function TDb.TryGetMessageLastModUtc(const MessageId: Integer; out LastModUtc: TDateTime): Boolean;
begin
  LastModUtc := 0;
  FQ.Close;
  FQ.SQL.Text := '  SELECT last_modification_time AS lm FROM mail.message WHERE message_id = :id';
  FQ.Params.Clear;
  SetParam('id', ftInteger, MessageId);
  FQ.Open;
  Result := not FQ.Fields[0].IsNull;
  if Result then
    LastModUtc := FQ.Fields[0].AsDateTime;
  FQ.Close;
end;

{ ===== New bulk helpers (moved from mail.FullSync) ===== }

function TDb.ResolveMessageIdsBySearchKeys(const PstId: Integer; const Keys: TArray<TBytes>): TArray<Integer>;
var
  N, i, idx, mid: Integer;
  Q, Ins: TFDQuery;
begin
  N := Length(Keys);
  SetLength(Result, N);
  for i := 0 to N - 1 do
    Result[i] := 0;
  if N = 0 then Exit;

  Q   := TFDQuery.Create(nil);
  Ins := TFDQuery.Create(nil);
  try
    Q.Connection   := FConn;
    Ins.Connection := FConn;

    Q.SQL.Text := 'CREATE TABLE #sk(idx INT NOT NULL, search_key VARBINARY(900) NOT NULL)';
    Q.ExecSQL;

    Ins.SQL.Text := 'INSERT INTO #sk(idx, search_key) VALUES (:i, :rk)';
    Ins.Params.Clear;
    Ins.Params.Add.Name := 'i';
    Ins.ParamByName('i').DataType := ftInteger;
    Ins.Params.Add.Name := 'rk';
    Ins.ParamByName('rk').DataType := ftVarBytes;

    Ins.Params.ArraySize := N;
    for i := 0 to N - 1 do
    begin
      Ins.ParamByName('i').AsIntegers[i] := i;
      if Length(Keys[i]) = 0 then
        Ins.ParamByName('rk').Clear(i)
      else
        Ins.ParamByName('rk').SetData(PByte(@Keys[i][0]), LongWord(Length(Keys[i])), i);
    end;
    Ins.Execute(N, 0);

    Q.SQL.Text :=
      'SELECT s.idx, m.message_id ' +
      'FROM #sk s JOIN mail.message m ON m.pst_id = :pid AND m.search_key = s.search_key';
    Q.Params.Clear;
    Q.Params.Add.Name := 'pid';
    Q.ParamByName('pid').DataType := ftInteger;
    Q.ParamByName('pid').AsInteger := PstId;

    Q.Open;
    while not Q.Eof do
    begin
      idx := Q.Fields[0].AsInteger;
      mid := Q.Fields[1].AsInteger;
      if (idx >= 0) and (idx < N) then
        Result[idx] := mid;
      Q.Next;
    end;
    Q.Close;
  finally
    try Q.SQL.Text := 'DROP TABLE #sk'; Q.ExecSQL; except end;
    Q.Free;
    Ins.Free;
  end;
end;

function TDb.BatchMessagesHaveRecipients(const MsgIds: TArray<Integer>): TArray<Boolean>;
var
  N, i, idx, cnt: Integer;
  Q, Ins: TFDQuery;
begin
  N := Length(MsgIds);
  SetLength(Result, N);
  for i := 0 to N - 1 do
    Result[i] := False;
  if N = 0 then Exit;

  Q   := TFDQuery.Create(nil);
  Ins := TFDQuery.Create(nil);
  try
    Q.Connection   := FConn;
    Ins.Connection := FConn;

    Q.SQL.Text := 'CREATE TABLE #ids(idx INT NOT NULL, message_id INT NOT NULL)';
    Q.ExecSQL;

    Ins.SQL.Text := 'INSERT INTO #ids(idx, message_id) VALUES (:i, :m)';
    Ins.Params.Clear;
    Ins.Params.Add.Name := 'i';
    Ins.ParamByName('i').DataType := ftInteger;
    Ins.Params.Add.Name := 'm';
    Ins.ParamByName('m').DataType := ftInteger;

    Ins.Params.ArraySize := N;
    for i := 0 to N - 1 do
    begin
      Ins.ParamByName('i').AsIntegers[i] := i;
      Ins.ParamByName('m').AsIntegers[i] := MsgIds[i];
    end;
    Ins.Execute(N, 0);

    Q.SQL.Text :=
      'SELECT d.idx, COUNT(r.message_id) AS c ' +
      'FROM #ids d LEFT JOIN mail.recipient r ON r.message_id = d.message_id ' +
      'GROUP BY d.idx';
    Q.Open;
    while not Q.Eof do
    begin
      idx := Q.Fields[0].AsInteger;
      cnt := Q.Fields[1].AsInteger;
      if (idx >= 0) and (idx < N) then
        Result[idx] := cnt > 0;
      Q.Next;
    end;
    Q.Close;

  finally
    try Q.SQL.Text := 'DROP TABLE #ids'; Q.ExecSQL; except end;
    Q.Free;
    Ins.Free;
  end;
end;

end.

