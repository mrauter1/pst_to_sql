unit mail.FullSync;

{
  Fast full-folder reconcile for local PSTs using Outlook OOM + SQL presence map.

  ADDITION (minimal): iteration strategy switch
  ---------------------------------------------
  You can choose how folder items are iterated:
    - imTable    : MAPIFolder.GetTable (rowset)   [existing default]
    - imRestrict : Items.Restrict (real items)

  Notes:
    * imRestrict is typically faster when the filter is selective (few or zero hits),
      because it avoids "row → item" rebinds. imTable is best for scanning many rows
      with a small set of properties and rebinding only a few.
    * Code maintains Delphi 10 Seattle compatibility.
}

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, System.Generics.Defaults, System.Variants,
  System.StrUtils, System.Math,
  System.Diagnostics,
  Winapi.ActiveX, System.Win.ComObj,
  FireDAC.Comp.Client, FireDAC.Stan.Intf, FireDAC.Stan.Param, FireDAC.Stan.Error,
  Mail.TypesUtils, mail.Core, mail.Data, mail.Logger, Data.DB;

type
// Batch buffers for ArrayDML upserts
  TMsgBatchRow = record
    FolderId: Integer;
    InternetMessageId: string;
    OutlookEntryId: string;
    Subject: string;
    SenderName: string;
    SenderEmail: string;
    DisplayTo: string;
    DisplayCc: string;
    Headers: string;
    BodyText: string;
    BodyHtml: string;
    SentUtc: TDateTime;
    RecvUtc: TDateTime;
    CreateUtc: TDateTime;
    LastModUtc: TDateTime;
    SearchKey: TBytes;
    Recipients: TArray<TArray<Variant>>;
    MessageClass: string;
  end;

  type
    TScanMode = (smFull, smMovesOnly);

    // NEW: iteration strategy
    TIterMode = (imTable, imRestrict);

    TCollectedRow = record
      EntryId    : string;
      RecKey     : TBytes;
      LastModUtc : TDateTime; // 0 when not requested
      MessageClass: string;   // '' when not requested
    end;

  TMailFullSyncReconciler = class
  private
    FIngestor: TIngestor;
    FFoldersToIgnore: TArray<string>;

    // Prepared DML & transaction state
    FQUpdate: TFDQuery;
    FQInsert: TFDQuery;

    // New prepared full DML (ArrayDML for full insert/update)
    FQInsertFull: TFDQuery;
    FQUpdateFull: TFDQuery;

    // Cached bytes comparer
    FBytesCmp: IEqualityComparer<TBytes>;

    FInsertRows: TList<TMsgBatchRow>;
    FUpdateRows: TList<TMsgBatchRow>;

    FDebug: Boolean;

    // NEW: selected iteration mode (default imTable)
    FIterMode: TIterMode;

    procedure BuildFolderTable(const Folder: OleVariant;
      const ColumnsToAdd: array of string; out Table: OleVariant; ACutOffDate: TDateTime = 0);
    procedure UpsertMessageFull(const FolderId: Integer; const EntryId: string;
      const RecKey: TBytes; const StoreId, FolderFullPath: string;
      const TreatAsUpdate: Boolean);

    procedure MsgBatchRowToParams(const AMessageRow: TMsgBatchRow; AQuery: TFDQuery;
      AIndex: Integer);

    // Loads the PST, attach if needed and store it to the database if needed. Returns the Pst ID;
    function LoadPst: Integer;
    function GetDB: TDb;
    function GetLogger: TLogger;
    function GetOlk: TOutlookCOM;
    function FolderWasUpdated(Folder: OleVariant; FolderId: Integer): Boolean;
    procedure CollectFolderRows(const Folder: OleVariant; Mode: TScanMode;
      ACutOffDate: TDateTime; out Rows: TArray<TCollectedRow>;
      out TotalFound: Integer);

    // NEW: Restrict-based collector
    procedure CollectFolderRowsViaRestrict(const Folder: OleVariant; Mode: TScanMode;
      ACutOffDate: TDateTime; out Rows: TArray<TCollectedRow>; out TotalFound: Integer);

    procedure ProcessFolderCommon(const Folder: OleVariant;
      const ParentFolderId: Variant; const ParentPath: string;
      const Depth: Integer; ACutOffDate: TDateTime; Mode: TScanMode);
    procedure ReconcileRows(const FolderId: Integer; const StoreId,
      FullPath: string; const Rows: TArray<TCollectedRow>; Mode: TScanMode);

    const BATCH_COMMIT_SIZE = 2000;
    const PR_SEARCH_KEY     = 'http://schemas.microsoft.com/mapi/proptag/0x300B0102';
    const olUserItems       = 0;
    const LM_EPSILON        = 1.0 / (24.0 * 60.0 * 60.0);

    // reasonable ArrayDML flush sizes (lower than commit-size to keep memory/friction low)
    const BATCH_UPSERT_FLUSH = 512;

    function  VariantByteArrayToBytes(const V: OleVariant): TBytes;
    function  HexToBytes(const Hex: string): TBytes; // fallback only
    function  TimesDiffer(const A, B: TDateTime): Boolean;

    // comparer for TBytes keys
    function  GetBytesComparer: IEqualityComparer<TBytes>;
    procedure EnsurePreparedDml;
    procedure EnsurePreparedFullDml;

    procedure EnsureFolderRow(const Folder: OleVariant; const ParentFolderId: Variant;
      const ParentPath: string; const Depth: Integer; out FolderId: Integer; out FullPath: string);

    procedure LoadDbPresenceByRecKeyBin(const Keys: TArray<TBytes>;
      out PresentIdx: TArray<Boolean>; out DbLastModUtc: TArray<TDateTime>; out DbFolderIds: TArray<Integer>);

    procedure ExecUpdateRehomePrepared(const FolderId: Integer;
      const EntryIds: TArray<string>; const RecKeys: TArray<TBytes>; const LastModsUtc: TArray<TDateTime>);

    // ===== New: ArrayDML helpers for full upsert path =====
    function  ResolveMsgIdsBySearchKeys(const Keys: TArray<TBytes>): TArray<Integer>;
    function  BatchHasRecipients(const MsgIds: TArray<Integer>): TArray<Boolean>;
    procedure FlushInsertBatch;
    procedure FlushUpdateBatch;
    procedure FlushBatches; // Save the rows to be inserted / updated to the database

    procedure ProcessFolder(const Folder: OleVariant; const ParentFolderId: Variant;
      const ParentPath: string; const Depth: Integer; ACutOffDate: TDateTime = 0);

    procedure ProcessFolderMovesOnly(const Folder: OleVariant;
      const ParentFolderId: Variant; const ParentPath: string;
      const Depth: Integer; ACutOffDate: TDateTime = 0);

    function  IsValidFolder(const Folder: OleVariant): Boolean;
  public
    constructor Create(const AIngestor: TIngestor; const ADebugMode: Boolean = False; AIterMode: TIterMode = imTable);
    destructor Destroy; override;

    // Execute a scan of the pst and save it to the database.
    // When AMovesOnlyScan is false, the function will add new items and update folder changes. Updated messages will not be updated.
    // Set AMovesOnlyScan to True for a fast full scan of the pst, to track new items and/or folder/PST moves (not tracked by message.lastModificationDate)
    // To track and update updated items too, set aMovesOnlyScan to true (Slower).
    procedure Execute(AMovesOnlyScan: Boolean = False; ACutOffDate: TDateTime = 0);

    property DB: TDb read GetDB;
    property Olk: TOutlookCOM read GetOlk;
    property Logger: TLogger read GetLogger;

    // NEW: expose iteration mode
    property IterationMode: TIterMode read FIterMode write FIterMode;
  end;

implementation

{ ============================= Helpers ============================= }

constructor TMailFullSyncReconciler.Create(const AIngestor: TIngestor; const ADebugMode: Boolean = False; AIterMode: TIterMode = imTable);
begin
  inherited Create;
  FIngestor:= AIngestor;
  FQUpdate := nil;
  FQInsert := nil;
  FQInsertFull := nil;
  FQUpdateFull := nil;
  FBytesCmp := GetBytesComparer; // cache once
  FInsertRows := TList<TMsgBatchRow>.Create;
  FUpdateRows := TList<TMsgBatchRow>.Create;
  FDebug:= ADebugMode;
  FIterMode := AIterMode;
  FFoldersToIgnore := TArray<string>.Create('trash', 'spam', 'junk', 'deleted', 'lixo', 'excluídos');
end;

destructor TMailFullSyncReconciler.Destroy;
begin
  FlushBatches; // ensure nothing is left pending
  FreeAndNil(FQUpdate);
  FreeAndNil(FQInsert);
  FreeAndNil(FQInsertFull);
  FreeAndNil(FQUpdateFull);
  FreeAndNil(FInsertRows);
  FreeAndNil(FUpdateRows);
  inherited;
end;

procedure TMailFullSyncReconciler.BuildFolderTable(const Folder: OleVariant;
  const ColumnsToAdd: array of string; out Table: OleVariant; ACutOffDate: TDateTime = 0);
var
  FFilter: string;
  i: Integer;
begin
  FFilter := '';
  if ACutOffDate > 0 then
    FFilter := BuildLastModJETFilter(ACutOffDate);

  Table := Unassigned;
  try
    Table := Folder.GetTable(FFilter, olUserItems);
    if IsVariantAssigned(Table) then
    begin
      // Reset and add only the columns requested by the caller
      try Table.Columns.RemoveAll; except end;
      for i := Low(ColumnsToAdd) to High(ColumnsToAdd) do
        try
          if ColumnsToAdd[i] <> '' then
            Table.Columns.Add(ColumnsToAdd[i]);
        except
          // ignore individual add failures; continue adding the rest
        end;
    end;
  except
    on E: Exception do
      Writeln(Format('Error on BuildFolderTable. %s: %s', [E.ClassName, E.Message]));
  end;
end;


function TMailFullSyncReconciler.VariantByteArrayToBytes(const V: OleVariant): TBytes;
var
  lb, hb, i: Integer;
begin
  SetLength(Result, 0);
  if not (VarIsArray(V) and (VarArrayDimCount(V) = 1)) then Exit;
  lb := VarArrayLowBound(V, 1);
  hb := VarArrayHighBound(V, 1);
  SetLength(Result, hb - lb + 1);
  for i := 0 to Length(Result) - 1 do
    Result[i] := Byte(V[lb + i]);
end;

function TMailFullSyncReconciler.HexToBytes(const Hex: string): TBytes;
var
  i, n: Integer;
  h: string;
  b: Byte;
begin
  h := Trim(Hex);
  if (Length(h) mod 2) <> 0 then
    SetLength(h, Length(h) - 1);
  n := Length(h) div 2;
  SetLength(Result, n);
  for i := 0 to n - 1 do
  begin
    b := StrToIntDef('$' + Copy(h, i * 2 + 1, 2), 0);
    Result[i] := b;
  end;
end;

function TMailFullSyncReconciler.TimesDiffer(const A, B: TDateTime): Boolean;
begin
  if (A = 0) xor (B = 0) then Exit(True);
  if (A = 0) and (B = 0) then Exit(False);
  Result := Abs(A - B) > LM_EPSILON;
end;

function TMailFullSyncReconciler.GetBytesComparer: IEqualityComparer<TBytes>;
begin
  Result := TEqualityComparer<TBytes>.Construct(
    function(const L, R: TBytes): Boolean
    var
      i: Integer;
    begin
      if Length(L) <> Length(R) then Exit(False);
      for i := 0 to Length(L) - 1 do
        if L[i] <> R[i] then Exit(False);
      Result := True;
    end,
    function(const V: TBytes): Integer
    var
      h: Cardinal;
      i: Integer;
    begin
      // FNV-1a 32-bit
      h := 2166136261;
      for i := 0 to Length(V) - 1 do
      begin
        h := h xor V[i];
        h := h * 16777619;
      end;
      Result := Integer(h);
    end
  );
end;

function TMailFullSyncReconciler.GetDB: TDb;
begin
  Result:= FIngestor.DB;
end;

function TMailFullSyncReconciler.GetLogger: TLogger;
begin
  Result:= FIngestor.Logger;
end;

function TMailFullSyncReconciler.GetOlk: TOutlookCOM;
begin
  Result:= FIngestor.Olk;
end;

procedure TMailFullSyncReconciler.EnsureFolderRow(const Folder: OleVariant;
  const ParentFolderId: Variant; const ParentPath: string; const Depth: Integer;
  out FolderId: Integer; out FullPath: string);
var
  name: string;
begin
  name := VarToStrDefSafe(Folder.Name, '');
  if name = '' then
    name := '(unnamed)';

  if ParentPath <> '' then
    FullPath := ParentPath + '/' + name
  else
    FullPath := name;

  FolderId := Db.EnsureFolderRow(FIngestor.PstId, ParentFolderId, name, FullPath, Depth);
end;

procedure TMailFullSyncReconciler.EnsurePreparedDml;
begin
  if FQUpdate = nil then
  begin
    FQUpdate := TFDQuery.Create(nil);
    FQUpdate.Connection := Db.Conn;
    FQUpdate.SQL.Text :=
      'UPDATE mail.message ' +
      'SET folder_id = :f, outlook_entry_id = :eid, ' +
      '    search_key = :rk, ' +
      '    last_modification_time = :lmod, ' +
      '    is_missing_from_pst = 0 ' +
      'WHERE pst_id = :pid AND search_key = :rk';

    FQUpdate.ParamByName('f').DataType     := ftInteger;
    FQUpdate.ParamByName('eid').DataType   := ftWideString;
    FQUpdate.ParamByName('rk').DataType    := ftVarBytes;
    FQUpdate.ParamByName('pid').DataType   := ftInteger;
    FQUpdate.ParamByName('lmod').DataType  := ftDateTime;
    FQUpdate.Prepare;
  end;

  if FQInsert = nil then
  begin
    FQInsert := TFDQuery.Create(nil);
    FQInsert.Connection := Db.Conn;
    FQInsert.SQL.Text :=
      'INSERT INTO mail.message (' +
      '  pst_id, folder_id, outlook_entry_id, search_key, last_modification_time, message_class, is_missing_from_pst' +
      ') VALUES (' +
      '  :pid, :f, :eid, :rk, :lmod, :mclass, 0' +
      ')';

    FQInsert.ParamByName('pid').DataType    := ftInteger;
    FQInsert.ParamByName('f').DataType      := ftInteger;
    FQInsert.ParamByName('eid').DataType    := ftWideString;
    FQInsert.ParamByName('rk').DataType     := ftVarBytes;
    FQInsert.ParamByName('lmod').DataType   := ftDateTime;
    // --- NEW ---
    FQInsert.ParamByName('mclass').DataType := ftWideString;
    FQInsert.Prepare;
  end;
end;

procedure TMailFullSyncReconciler.EnsurePreparedFullDml;

  procedure CreateParams(AQuery: TFDQuery);
  begin
    with AQuery do
    begin
      ParamByName('sk').DataType    := ftVarBytes;

      ParamByName('mclass').DataType    := ftWideString;
      ParamByName('mid').DataType    := ftWideString;
      ParamByName('eid').DataType   := ftWideString;
      ParamByName('subj').DataType   := ftWideString;
      ParamByName('semail').DataType := ftWideString;
      ParamByName('sname').DataType := ftWideString;

      ParamByName('sent').DataType  := ftDateTime;
      ParamByName('recv').DataType  := ftDateTime;
      ParamByName('crt').DataType   := ftDateTime;
      ParamByName('lmod').DataType  := ftDateTime;

      ParamByName('dto').DataType   := ftWideMemo;
      ParamByName('dcc').DataType   := ftWideMemo;
      ParamByName('hdr').DataType   := ftWideMemo;
      ParamByName('btxt').DataType  := ftWideMemo;
      ParamByName('bhtm').DataType  := ftWideMemo;

      ParamByName('sz').DataType    := ftInteger;

    end;
  end;

begin
  if FQInsertFull = nil then
  begin
    FQInsertFull := TFDQuery.Create(nil);
    FQInsertFull.Connection := Db.Conn;

    //ATTENTION: DUE TO A KNOWN SQL SERVER BUG WITH ARRAY DML, ALL MAX FIELDS (WIDEMEMO) MUST BE SET AT THE VERY END OF THE COLUMN/PARAMETER LIST
    FQInsertFull.SQL.Text :=
      'INSERT INTO mail.message ' +
      '(pst_id, folder_id, internet_message_id, outlook_entry_id, subject, sender_name, sender_email, ' +
      ' sent_utc, received_utc, created_utc, ' +
      ' size_bytes, last_modification_time, search_key, message_class, ' +
      ' display_to, display_cc, transport_headers, body_text, body_html) ' +
      'VALUES (:pid,:fid,:mid,:eid,:subj,:sname,:semail,:sent,:recv,:crt,:sz,:lmod,:sk,:mclass,:dto,:dcc,:hdr,:btxt,:bhtm)';
    FQInsertFull.ParamByName('pid').DataType   := ftInteger;
    FQInsertFull.ParamByName('fid').DataType   := ftInteger;
    CreateParams(FQInsertFull);

    // ArrayDMLSize must be set to 1 to avoid SQL Server Driver Bug with Blob and Memo Filelds
    FQInsertFull.ResourceOptions.ArrayDMLSize := 1;

    FQInsertFull.Prepare;
  end;

  if FQUpdateFull = nil then
  begin
    FQUpdateFull := TFDQuery.Create(nil);
    FQUpdateFull.Connection := Db.Conn;
    //ATTENTION: DUE TO A KNOWN SQL SERVER BUG WITH ARRAY DML, ALL MAX FIELDS (WIDEMEMO) MUST BE SET AT THE VERY END OF THE COLUMN/PARAMETER LIST
    FQUpdateFull.SQL.Text :=
      'UPDATE mail.message SET ' +
      '  internet_message_id   = :mid, ' +
      '  outlook_entry_id      = :eid, ' +
      '  subject               = :subj, ' +
      '  sender_name           = :sname, ' +
      '  sender_email          = :semail, ' +
      '  sent_utc              = :sent, ' +
      '  received_utc          = :recv, ' +
      '  created_utc           = :crt, ' +
      '  size_bytes            = :sz, ' +
      '  last_modification_time= :lmod, ' +
      '  search_key            = :sk, ' +
      '  message_class         = :mclass, ' +
      '  display_to            = :dto, ' +
      '  display_cc            = :dcc, ' +
      '  transport_headers     = :hdr, ' +
      '  body_text             = :btxt, ' +
      '  body_html             = :bhtm ' +
      'WHERE message_id = :id';
    with FQUpdateFull.Params do
    begin
      ParamByName('id').DataType    := ftInteger;
    end;
    CreateParams(FQUpdateFull);

    // ArrayDMLSize must be set to 1 to avoid SQL Server Driver Bug with Blob and Memo Filelds
    FQUpdateFull.ResourceOptions.ArrayDMLSize := 1;
    FQUpdateFull.Prepare;
  end;
end;

function TMailFullSyncReconciler.IsValidFolder(const Folder: OleVariant): Boolean;
var
  path, fname: string;
  i: Integer;
begin
  // Read folder path and name defensively
  path := '';
  fname := '';
  try path := VarToStrDefSafe(Folder.FolderPath, ''); except path := ''; end;
  try fname := VarToStrDefSafe(Folder.Name, ''); except fname := ''; end;

  // Ignore Outlook's virtual search folder container
  if (path <> '') and ContainsText(path, 'Search Folders') then
    Exit(False);

  // Ignore configured trash/spam-like folders (by name or full path)
  for i := Low(FFoldersToIgnore) to High(FFoldersToIgnore) do
    if (FFoldersToIgnore[i] <> '') and
       (ContainsText(fname, FFoldersToIgnore[i]) or ContainsText(path, FFoldersToIgnore[i])) then
      Exit(False);

  // If we cannot read a path, treat as real to avoid dropping valid roots
  if path = '' then
    Exit(True);

  Result := True;
end;

procedure TMailFullSyncReconciler.LoadDbPresenceByRecKeyBin(
  const Keys: TArray<TBytes>; out PresentIdx: TArray<Boolean>; out DbLastModUtc: TArray<TDateTime>; out DbFolderIds: TArray<Integer>);
var
  total, i: Integer;
  SelQ: TFDQuery;
  rowKey: TBytes;
  lm: TDateTime;
  fid: Integer;
  AllIndex: TDictionary<TBytes, Integer>;
  idx: Integer;
  InsQ: TFDQuery;
  chunk, done, nThis: Integer;
begin
  total := Length(Keys);
  SetLength(PresentIdx, total);
  SetLength(DbLastModUtc, total);
  SetLength(DbFolderIds, total);
  if total = 0 then Exit;

  for i := 0 to total - 1 do
  begin
    PresentIdx[i]  := False;
    DbLastModUtc[i]:= 0;
    DbFolderIds[i] := 0;
  end;

  AllIndex := TDictionary<TBytes, Integer>.Create(total, FBytesCmp);
  try
    for i := 0 to total - 1 do
      if not AllIndex.ContainsKey(Keys[i]) then
        AllIndex.Add(Keys[i], i);

    // Preferred path on SQL Server: temp table + join
    begin
      SelQ := TFDQuery.Create(nil);
      InsQ := TFDQuery.Create(nil);
      try
        SelQ.Connection := Db.Conn;
        InsQ.Connection := Db.Conn;

        // create temp table
        SelQ.SQL.Text := 'CREATE TABLE #probe (search_key VARBINARY(900) NOT NULL)';
        SelQ.ExecSQL;

        // bulk insert all keys with ArrayDML in chunks
        InsQ.SQL.Text := 'INSERT INTO #probe (search_key) VALUES (:rk)';
        InsQ.Params.Clear;
        InsQ.Params.Add.Name := 'rk';
        InsQ.ParamByName('rk').DataType := ftVarBytes;

        chunk := 2000;
        done := 0;
        while done < total do
        begin
          nThis := Min(chunk, total - done);
          InsQ.Params.ArraySize := nThis;
          for i := 0 to nThis - 1 do
          begin
            if Length(Keys[done + i]) = 0 then
              InsQ.ParamByName('rk').Clear(i)
            else
              InsQ.ParamByName('rk').SetData(PByte(@Keys[done + i][0]),
                                             LongWord(Length(Keys[done + i])), i);
          end;
          InsQ.Execute(nThis, 0);
          done := done + nThis;
        end;

        // join back to message table
        SelQ.SQL.Text :=
          'SELECT m.search_key AS sk, m.last_modification_time AS lm, m.folder_id AS fid ' +
          'FROM #probe p JOIN mail.message m ' +
          '  ON m.pst_id = :pid AND m.search_key = p.search_key';
        SelQ.Params.Clear;
        SelQ.Params.Add.Name := 'pid';
        SelQ.ParamByName('pid').DataType := ftInteger;
        SelQ.ParamByName('pid').AsInteger := FIngestor.PstId;

        SelQ.Open;
        while not SelQ.Eof do
        begin
          rowKey := SelQ.FieldByName('sk').AsBytes;
          lm     := 0; if not SelQ.FieldByName('lm').IsNull then lm := SelQ.FieldByName('lm').AsDateTime;
          fid    := 0; if not SelQ.FieldByName('fid').IsNull then fid := SelQ.FieldByName('fid').AsInteger;

          if AllIndex.TryGetValue(rowKey, idx) then
          begin
            PresentIdx[idx]  := True;
            DbLastModUtc[idx]:= lm;
            DbFolderIds[idx] := fid;
          end;

          SelQ.Next;
        end;
        SelQ.Close;
      finally
        // drop temp table explicitly
        try SelQ.SQL.Text := 'DROP TABLE #probe'; SelQ.ExecSQL; except end;
        SelQ.Free;
        InsQ.Free;
      end;
      Exit; // done via temp table
    end;

    // (fallback omitted; SQL Server path used by this application)
  finally
    AllIndex.Free;
  end;
end;

function TMailFullSyncReconciler.LoadPst: Integer;
begin

end;

procedure TMailFullSyncReconciler.ExecUpdateRehomePrepared(
  const FolderId: Integer; const EntryIds: TArray<string>; const RecKeys: TArray<TBytes>; const LastModsUtc: TArray<TDateTime>);
var
  N, i: Integer;
begin
  N := Length(RecKeys);
  if N = 0 then Exit;
  EnsurePreparedDml;

  FQUpdate.Params.ArraySize := N;

  for i := 0 to N - 1 do
  begin
    FQUpdate.ParamByName('f').AsIntegers[i]    := FolderId;
    FQUpdate.ParamByName('eid').AsStrings[i]   := EntryIds[i];
    FQUpdate.ParamByName('pid').AsIntegers[i]  := FIngestor.PstId;

    if Length(RecKeys[i]) = 0 then
      FQUpdate.ParamByName('rk').Clear(i)
    else
      FQUpdate.ParamByName('rk').SetData(PByte(@RecKeys[i][0]), LongWord(Length(RecKeys[i])), i);

    if (i <= High(LastModsUtc)) and (LastModsUtc[i] <> 0) then
      FQUpdate.ParamByName('lmod').AsDateTimes[i] := LastModsUtc[i]
    else
      FQUpdate.ParamByName('lmod').Clear(i);
  end;

  FQUpdate.Execute(N, 0);
end;

function TMailFullSyncReconciler.ResolveMsgIdsBySearchKeys(const Keys: TArray<TBytes>): TArray<Integer>;
begin
  // DB responsibility moved to TDb; keep thin wrapper to minimize diff and risk
  Result := Db.ResolveMessageIdsBySearchKeys(FIngestor.PstId, Keys);
end;

function TMailFullSyncReconciler.BatchHasRecipients(const MsgIds: TArray<Integer>): TArray<Boolean>;
begin
  // DB responsibility moved to TDb
  Result := Db.BatchMessagesHaveRecipients(MsgIds);
end;

procedure TMailFullSyncReconciler.MsgBatchRowToParams(
  const AMessageRow: TMsgBatchRow; AQuery: TFDQuery; AIndex: Integer);
  procedure SetStr(const Name, V: string);
  begin
    if V <> '' then
      AQuery.ParamByName(Name).AsStrings[AIndex] := V;
  end;

  procedure SetMemo(const Name, V: string);
  begin
    if V <> '' then
      AQuery.ParamByName(Name).AsWideMemos[AIndex] := V
    else
      AQuery.ParamByName(Name).Clear(AIndex);
      // ftWideMemo
  end;

  procedure SetDT(const Name: string; const V: TDateTime);
  begin
    if V <> 0 then
      AQuery.ParamByName(Name).AsDateTimes[AIndex] := V;
  end;

  procedure SetBytes(const Name: string; const B: TBytes);
  begin
    if Length(B) > 0 then
      AQuery.ParamByName(Name).SetData(PByte(@B[0]), LongWord(Length(B)), AIndex);
  end;

begin
  if (AQuery = nil) then
    Exit;

  // Always write search_key and message_class
  SetBytes('sk',   AMessageRow.SearchKey);
  SetStr('mclass', AMessageRow.MessageClass);
  SetDT('lmod', AMessageRow.LastModUtc);

  // Only set the full set of fields for mail items
  if not StartsText('ipm.note', AMessageRow.MessageClass.ToLower) then
    Exit;

  // Core strings
  SetStr('mid',    AMessageRow.InternetMessageId);
  SetStr('eid',    AMessageRow.OutlookEntryId);
  SetStr('subj',   AMessageRow.Subject);
  SetStr('sname',  AMessageRow.SenderName);
  SetStr('semail', AMessageRow.SenderEmail);

  // Times
  SetDT('sent', AMessageRow.SentUtc);
  SetDT('recv', AMessageRow.RecvUtc);
  SetDT('crt',  AMessageRow.CreateUtc);

  // Large text
  SetMemo('dto',    AMessageRow.DisplayTo);
  SetMemo('dcc',    AMessageRow.DisplayCc);
  SetMemo('hdr',  AMessageRow.Headers);
  SetMemo('btxt', AMessageRow.BodyText);
  AQuery.ParamByName('sz').Clear(AIndex);
//  SetMemo('bhtm', AMessageRow.BodyHtml);
end;

procedure TMailFullSyncReconciler.FlushInsertBatch;
var
  N, i: Integer;
  Keys: TArray<TBytes>;
  Ids:  TArray<Integer>;
  S: string;
  MsgIds      : TArray<Integer>;
  RecPerMsg   : TArray<TArray<TArray<Variant>>>;
  MsgIdsList  : TList<Integer>;
  RecList     : TList<TArray<TArray<Variant>>>;
begin
  N := FInsertRows.Count;
  if N = 0 then Exit;

  EnsurePreparedFullDml;

  FQInsertFull.Params.ArraySize := N;
  for i := 0 to N - 1 do
  begin
    FQInsertFull.Params.ClearValues(i);
    FQInsertFull.ParamByName('pid').AsIntegers[i] := FIngestor.PstId;
    FQInsertFull.ParamByName('fid').AsIntegers[i] := FInsertRows[i].FolderId;

    MsgBatchRowToParams(FInsertRows[I], FQInsertFull, I);
  end;

  FQInsertFull.Execute(N, 0);

  // Resolve IDs and persist recipients with the new ArrayDML path in TDb.SaveRecipients
  SetLength(Keys, N);
  for i := 0 to N - 1 do
    Keys[i] := FInsertRows[i].SearchKey;

  Ids := ResolveMsgIdsBySearchKeys(Keys);

  // Save recipients in one ArrayDML call (across all inserted messages)

  MsgIdsList := TList<Integer>.Create;
  RecList    := TList<TArray<TArray<Variant>>>.Create;
  try
    for i := 0 to N - 1 do
      if (Ids[i] <> 0) and (Length(FInsertRows[i].Recipients) > 0) then
      begin
        MsgIdsList.Add(Ids[i]);
        RecList.Add(FInsertRows[i].Recipients);
      end;

    if MsgIdsList.Count > 0 then
    begin
      SetLength(MsgIds, MsgIdsList.Count);
      SetLength(RecPerMsg, RecList.Count);
      for i := 0 to MsgIdsList.Count - 1 do
      begin
        MsgIds[i]    := MsgIdsList[i];
        RecPerMsg[i] := RecList[i];
      end;
      Db.SaveRecipients(MsgIds, RecPerMsg);
    end;
  finally
    RecList.Free;
    MsgIdsList.Free;
  end;

  FInsertRows.Clear;

  Db.Commit;
end;

procedure TMailFullSyncReconciler.FlushUpdateBatch;
var
  N, i: Integer;
  Keys: TArray<TBytes>;
  Ids:  TArray<Integer>;
  HasRec: TArray<Boolean>;
  S: string;
  MsgIds      : TArray<Integer>;
  RecPerMsg   : TArray<TArray<TArray<Variant>>>;
  MsgIdsList  : TList<Integer>;
  RecList     : TList<TArray<TArray<Variant>>>;
begin
  N := FUpdateRows.Count;
  if N = 0 then Exit;

  EnsurePreparedFullDml;

  SetLength(Keys, N);
  for i := 0 to N - 1 do
    Keys[i] := FUpdateRows[i].SearchKey;
  Ids := ResolveMsgIdsBySearchKeys(Keys);

  FQUpdateFull.Params.ArraySize := N;

  for i := 0 to N - 1 do
  begin
    FQUpdateFull.Params.ClearValues(i);
    if Ids[i] = 0 then
    begin
      FQUpdateFull.ParamByName('id').Clear(i);
      Continue;
    end;

    FQUpdateFull.ParamByName('id').AsIntegers[i] := Ids[i];

    MsgBatchRowToParams(FUpdateRows[I], FQUpdateFull, I);
  end;

  FQUpdateFull.Execute(N, 0);

  // Batch check recipient presence; backfill only where missing (ArrayDML)
  HasRec := BatchHasRecipients(Ids);

  MsgIdsList := TList<Integer>.Create;
  RecList    := TList<TArray<TArray<Variant>>>.Create;
  try
    for i := 0 to N - 1 do
      if (Ids[i] <> 0) and (not HasRec[i]) and (Length(FUpdateRows[i].Recipients) > 0) then
      begin
        MsgIdsList.Add(Ids[i]);
        RecList.Add(FUpdateRows[i].Recipients);
      end;

    if MsgIdsList.Count > 0 then
    begin
      SetLength(MsgIds, MsgIdsList.Count);
      SetLength(RecPerMsg, RecList.Count);
      for i := 0 to MsgIdsList.Count - 1 do
      begin
        MsgIds[i]    := MsgIdsList[i];
        RecPerMsg[i] := RecList[i];
      end;
      Db.SaveRecipients(MsgIds, RecPerMsg);
    end;
  finally
    RecList.Free;
    MsgIdsList.Free;
  end;


  FUpdateRows.Clear;
end;

procedure TMailFullSyncReconciler.FlushBatches;
begin
  // flush updates first to ensure any rehomes/upgrades are written before inserts (no strict dependency, but tidy)
  FlushUpdateBatch;
  FlushInsertBatch;
end;

{ ============================= UpsertMessageFull (now batched) ============================= }

procedure TMailFullSyncReconciler.UpsertMessageFull(
  const FolderId: Integer; const EntryId: string; const RecKey: TBytes;
  const StoreId, FolderFullPath: string; const TreatAsUpdate: Boolean);
var
  Item: OleVariant;
  InternetMessageId, OutlookEntryId, Subject, SenderName, SenderEmail: string;
  DisplayTo, DisplayCc, Headers, BodyText, BodyHtml: string;
  SentUtc, RecvUtc, CreateUtc, LastModNew: TDateTime;
  SearchKeyRead: TBytes;
  RecRows: TArray<TArray<Variant>>;
  R: TMsgBatchRow;
  // --- NEW ---
  MsgClass: string;
begin
  // Resolve COM item
  Item := Unassigned;
  try
    if StoreId <> '' then
      Item := Olk.Session.GetItemFromID(EntryId, StoreId)
    else
      Item := Olk.Session.GetItemFromID(EntryId);
  except
    Item := Unassigned;
  end;
  if not IsVariantAssigned(Item) then
    Exit;

  // Extract full message data once
  ReadMessageCore(Item,
    InternetMessageId, OutlookEntryId, Subject, SenderName, SenderEmail,
    DisplayTo, DisplayCc, Headers, BodyText, BodyHtml, MsgClass,
    SentUtc, RecvUtc, CreateUtc, LastModNew, SearchKeyRead);

  if Length(SearchKeyRead) = 0 then
    SearchKeyRead := RecKey;

  try
    GetRecipients(Item, Headers, RecRows);
  except
    SetLength(RecRows, 0);
    Logger.Warn(Format('Recipient extraction error at %s / %s', [FolderFullPath, Subject]));
  end;

  FillChar(R, SizeOf(R), 0);
  R.FolderId          := FolderId;
  R.InternetMessageId := InternetMessageId;
  R.OutlookEntryId    := OutlookEntryId;
  R.Subject           := Subject;
  R.SenderName        := SenderName;
  R.SenderEmail       := SenderEmail;
  R.DisplayTo         := DisplayTo;
  R.DisplayCc         := DisplayCc;
//  R.Headers           := Headers;
  R.Headers           := '';
  R.BodyText          := BodyText;
  R.BodyHtml          := '';
  R.SentUtc           := SentUtc;
  R.RecvUtc           := RecvUtc;
  R.CreateUtc         := CreateUtc;
  R.LastModUtc        := LastModNew;
  R.SearchKey         := SearchKeyRead;
  R.Recipients        := RecRows;
  R.MessageClass      := MsgClass;

  if TreatAsUpdate then
    FUpdateRows.Add(R)
  else
    FInsertRows.Add(R);

  if (FInsertRows.Count >= BATCH_UPSERT_FLUSH) or (FUpdateRows.Count >= BATCH_UPSERT_FLUSH) then
    FlushBatches;
end;

function TMailFullSyncReconciler.FolderWasUpdated(Folder: OleVariant; FolderId: Integer): Boolean;
var
  commitUtc, lastSeen: TDateTime;
begin
  Exit(True); // PST does not support PR_LOCAL_COMMIT_TIME_MAX tag
end;

{ ============================= Folder processing ============================= }

procedure TMailFullSyncReconciler.ProcessFolder(const Folder: OleVariant; const ParentFolderId: Variant;
  const ParentPath: string; const Depth: Integer; ACutOffDate: TDateTime = 0);
begin
  ProcessFolderCommon(Folder, ParentFolderId, ParentPath, Depth, ACutOffDate, smFull);
end;

procedure TMailFullSyncReconciler.ProcessFolderMovesOnly(
  const Folder: OleVariant; const ParentFolderId: Variant;
  const ParentPath: string; const Depth: Integer; ACutOffDate: TDateTime = 0);
begin
  ProcessFolderCommon(Folder, ParentFolderId, ParentPath, Depth, ACutOffDate, smMovesOnly);
end;

// NEW: Restrict-based collector (used when FIterMode = imRestrict)
procedure TMailFullSyncReconciler.CollectFolderRowsViaRestrict(
  const Folder: OleVariant; Mode: TScanMode; ACutOffDate: TDateTime;
  out Rows: TArray<TCollectedRow>; out TotalFound: Integer);
var
  Items, View, It, PA, vSk: OleVariant;
  Filter: string;
  r: TCollectedRow;
  I, ItemCount: Integer;
begin
  SetLength(Rows, 0);
  TotalFound := 0;

  // Build JET filter on LastModificationTime when requested (local time semantics, kept consistent)
  Filter := '';
  if (Mode = smFull) and (ACutOffDate > 0) then
    Filter := BuildLastModJETFilter(ACutOffDate);

  Items := Unassigned;
  try Items := Folder.Items; except Items := Unassigned; end;
  if not IsVariantAssigned(Items) then
    Exit;

  // Apply Restrict only when a filter exists; otherwise iterate all
  try
    if Filter <> '' then
      View := Items.Restrict(Filter)
    else
      View := Items;
  except
    View := Unassigned;
  end;

  if not IsVariantAssigned(View) then
    Exit;

  ItemCount:= View.Count;

  for I := 1 to ItemCount do
  begin
    it:= view.Item(i);
    if not IsVariantAssigned(it) then
      Exit;

    FillChar(r, SizeOf(r), 0);
    r.EntryId      := '';
    r.MessageClass := '';
    r.LastModUtc   := 0;
    SetLength(r.RecKey, 0);

    // EntryID
    try r.EntryId := VarToStrDefSafe(It.EntryID, ''); except r.EntryId := ''; end;

    // Search key via PropertyAccessor (PT_BINARY)
    try
      PA := It.PropertyAccessor;
      vSk := PA.GetProperty(PR_SEARCH_KEY);
      r.RecKey := VariantByteArrayToBytes(vSk);
    except
      SetLength(r.RecKey, 0);
    end;

    if Mode = smFull then
    begin
      try r.LastModUtc := ToUTC(It.LastModificationTime); except r.LastModUtc := 0; end;
      try r.MessageClass := VarToStrDefSafe(It.MessageClass, ''); except r.MessageClass := ''; end;
    end;

    if Length(r.RecKey) > 0 then
    begin
      Rows := Rows + [r];
      Inc(TotalFound);
    end;
  end;
end;

procedure TMailFullSyncReconciler.CollectFolderRows(
  const Folder: OleVariant; Mode: TScanMode; ACutOffDate: TDateTime;
  out Rows: TArray<TCollectedRow>; out TotalFound: Integer);
var
  Tbl, Row, Cols, Col, Vals, vSk, vLmod: OleVariant;
  idxEntry, idxSk, idxLmod, idxMsgClass: Integer;
  nCols, i: Integer;
  colName: string;
  r: TCollectedRow;
begin
  // Switch between GetTable (existing) and Restrict (new)
  if FIterMode = imRestrict then
  begin
    CollectFolderRowsViaRestrict(Folder, Mode, ACutOffDate, Rows, TotalFound);
    Exit;
  end;

  // ===== Original GetTable path (unchanged) =====
  SetLength(Rows, 0);
  TotalFound := 0;

  // Choose minimal column set per mode
  if Mode = smFull then
    BuildFolderTable(Folder,
      ['EntryID', PR_SEARCH_KEY, PR_LAST_MODIFICATION_TIME, PR_MESSAGE_CLASS],
      Tbl, ACutOffDate)
  else
    BuildFolderTable(Folder,
      ['EntryID', PR_SEARCH_KEY],
      Tbl, ACutOffDate);

  if not IsVariantAssigned(Tbl) then
    Exit;

  idxEntry := -1; idxSk := -1; idxLmod := -1; idxMsgClass := -1;

  // Resolve column indexes (robust to ordering)
  try
    Cols := Tbl.Columns;
    try nCols := Cols.Count; except nCols := 0; end;
    for i := 1 to nCols do
    begin
      Col := Cols.Item(i);
      colName := VarToStrDefSafe(Col.Name, '');
      if SameText(colName, 'EntryID') then
        idxEntry := i - 1
      else if SameText(colName, PR_SEARCH_KEY) then
        idxSk := i - 1
      else if SameText(colName, PR_LAST_MODIFICATION_TIME) then
        idxLmod := i - 1
      else if SameText(colName, PR_MESSAGE_CLASS) then
        idxMsgClass := i - 1;
    end;
  except
    // ignore
  end;

  try Tbl.MoveToStart; except end;

  while (not Tbl.EndOfTable) do
  begin
    Row := Tbl.GetNextRow;
    if not IsVariantAssigned(Row) then Break;

    FillChar(r, SizeOf(r), 0);
    r.EntryId     := '';
    r.MessageClass:= '';
    r.LastModUtc  := 0;
    SetLength(r.RecKey, 0);

    try Vals := Row.GetValues; except Vals := Unassigned; end;

    if VarIsArray(Vals) then
    begin
      if idxEntry >= 0 then
        r.EntryId := VarToStrDefSafe(Vals[idxEntry], '');

      if idxSk >= 0 then
        r.RecKey := VariantByteArrayToBytes(Vals[idxSk]);

      if (Mode = smFull) and (idxLmod >= 0) then
      begin
        vLmod := Vals[idxLmod];
        try r.LastModUtc := vLmod; except r.LastModUtc := 0; end;
      end;

      if (Mode = smFull) and (idxMsgClass >= 0) then
        r.MessageClass := VarToStrDefSafe(Vals[idxMsgClass], '');
    end;

    if r.EntryId = '' then
      try r.EntryId := VarToStrDefSafe(Row.Item('EntryID'), ''); except r.EntryId := ''; end;

    if (Length(r.RecKey) = 0) and (idxSk < 0) then
    begin
      // Fallback: BinaryToString hex → bytes
      try vSk := Row.BinaryToString(PR_SEARCH_KEY); except vSk := Unassigned; end;
      r.RecKey := HexToBytes(VarToStrDefSafe(vSk, ''));
    end;

    if (r.LastModUtc = 0) and (Mode = smFull) and (idxLmod < 0) then
    begin
      // LastModificationTime fallback only in full mode
      try r.LastModUtc := ToUTC(Row.Item('LastModificationTime')); except r.LastModUtc := 0; end;
    end;

    if Length(r.RecKey) > 0 then
    begin
      Rows := Rows + [r];
      Inc(TotalFound);
    end;
  end;
end;

procedure TMailFullSyncReconciler.ReconcileRows(
  const FolderId: Integer; const StoreId, FullPath: string;
  const Rows: TArray<TCollectedRow>; Mode: TScanMode);
var
  totalFound: Integer;
  Present: TArray<Boolean>;
  DbLastModUtc: TArray<TDateTime>;
  DbFolderIds: TArray<Integer>;

  AllIndexByKey: TDictionary<TBytes, Integer>;
  RehomeIdx: TDictionary<TBytes, Boolean>;
  NewIdx   : TDictionary<TBytes, Boolean>;
  UpdatedIdx: TDictionary<TBytes, Boolean>;

  Keys: TArray<TBytes>;
  EntryIds: TArray<string>;
  Lmods: TArray<TDateTime>;

  idx, j, n: Integer;
  kKey: TBytes;

  ReEids, NewEids: TArray<string>;
  ReKeys, NewKeys: TArray<TBytes>;
  ReLmods, NewLmods: TArray<TDateTime>;
  shownPath: string;
begin
  totalFound := Length(Rows);
  if totalFound = 0 then
  begin
    Logger.Info(Format('Folder "%s": found=%d', [FullPath, totalFound]));
    Exit;
  end;

  // Build parallel arrays to probe presence
  SetLength(Keys, totalFound);
  SetLength(EntryIds, totalFound);
  SetLength(Lmods, totalFound);
  for idx := 0 to totalFound - 1 do
  begin
    Keys[idx]    := Rows[idx].RecKey;
    EntryIds[idx]:= Rows[idx].EntryId;
    Lmods[idx]   := Rows[idx].LastModUtc;
  end;

  LoadDbPresenceByRecKeyBin(Keys, Present, DbLastModUtc, DbFolderIds);

  AllIndexByKey := TDictionary<TBytes, Integer>.Create(totalFound, FBytesCmp);
  RehomeIdx     := TDictionary<TBytes, Boolean>.Create(totalFound, FBytesCmp);
  NewIdx        := TDictionary<TBytes, Boolean>.Create(totalFound, FBytesCmp);
  if Mode = smFull then
    UpdatedIdx := TDictionary<TBytes, Boolean>.Create(totalFound, FBytesCmp)
  else
    UpdatedIdx := nil;
  try
    for idx := 0 to totalFound - 1 do
      if not AllIndexByKey.ContainsKey(Keys[idx]) then
        AllIndexByKey.Add(Keys[idx], idx);

    for idx := 0 to totalFound - 1 do
    begin
      if Present[idx] then
      begin
        if DbFolderIds[idx] <> FolderId then
          RehomeIdx.AddOrSetValue(Keys[idx], True);
        if (Mode = smFull) and TimesDiffer(Lmods[idx], DbLastModUtc[idx]) then
          UpdatedIdx.AddOrSetValue(Keys[idx], True);
      end
      else
        NewIdx.AddOrSetValue(Keys[idx], True);
    end;

    // Rehomes (LM differs by mode)
    if RehomeIdx.Count > 0 then
    begin
      SetLength(ReEids, RehomeIdx.Count);
      SetLength(ReKeys, RehomeIdx.Count);
      SetLength(ReLmods, RehomeIdx.Count);
      j := 0;
      for kKey in RehomeIdx.Keys do
        if AllIndexByKey.TryGetValue(kKey, idx) then
        begin
          ReEids[j]  := EntryIds[idx];
          ReKeys[j]  := Keys[idx];
          // Preserve DB LM in moves-only; use scanned LM in full mode
          if Mode = smFull then
            ReLmods[j] := Lmods[idx]
          else
            ReLmods[j] := DbLastModUtc[idx];
          Inc(j);
        end;
      if j < Length(ReEids) then
      begin
        SetLength(ReEids, j); SetLength(ReKeys, j); SetLength(ReLmods, j);
      end;
      ExecUpdateRehomePrepared(FolderId, ReEids, ReKeys, ReLmods);
    end;

    // Inserts
    for kKey in NewIdx.Keys do
      if AllIndexByKey.TryGetValue(kKey, idx) then
        UpsertMessageFull(FolderId, EntryIds[idx], Keys[idx], StoreId, FullPath, False);

    // Updates (full mode only)
    if (Mode = smFull) and (UpdatedIdx.Count > 0) then
      for kKey in UpdatedIdx.Keys do
        if AllIndexByKey.TryGetValue(kKey, idx) then
          UpsertMessageFull(FolderId, EntryIds[idx], Keys[idx], StoreId, FullPath, True);

    // Flush batches produced by UpsertMessageFull
    FlushBatches;

    // Log (keep original wording)
    shownPath := FullPath;
    try shownPath := VarToStrDefSafe(OleVariant(FolderId){dummy}, FullPath); except end; // keep FullPath on log
    if Mode = smFull then
      Logger.Info(Format('Folder "%s": found=%d, rehome=%d, new=%d, updated=%d',
        [FullPath, totalFound, RehomeIdx.Count, NewIdx.Count, UpdatedIdx.Count]))
    else
      Logger.Info(Format('Folder "%s": found=%d, rehome=%d, new=%d',
        [FullPath, totalFound, RehomeIdx.Count, NewIdx.Count]));

  finally
    AllIndexByKey.Free;
    RehomeIdx.Free;
    NewIdx.Free;
    if Assigned(UpdatedIdx) then UpdatedIdx.Free;
  end;
end;

procedure TMailFullSyncReconciler.ProcessFolderCommon(
  const Folder: OleVariant; const ParentFolderId: Variant;
  const ParentPath: string; const Depth: Integer;
  ACutOffDate: TDateTime; Mode: TScanMode);
var
  Subs, Sub: OleVariant;
  FolderId: Integer;
  FullPath: string;
  StoreId: string;
  childCount, i: Integer;
  Rows: TArray<TCollectedRow>;
  totalFound: Integer;
begin
  if not IsValidFolder(Folder) then
    Exit;

  EnsureFolderRow(Folder, ParentFolderId, ParentPath, Depth, FolderId, FullPath);

  StoreId := '';
  try StoreId := VarToStrDefSafe(Folder.StoreID, ''); except end;

  if FolderWasUpdated(Folder, FolderId) then
  begin
    CollectFolderRows(Folder, Mode, ACutOffDate, Rows, totalFound);
    if totalFound > 0 then
      ReconcileRows(FolderId, StoreId, FullPath, Rows, Mode);
  end;

  // Recurse
  Subs := Unassigned;
  try Subs := Folder.Folders; except Subs := Unassigned; end;

  childCount := 0;
  if IsVariantAssigned(Subs) then
    try childCount := VarToIntDefSafe(Subs.Count, 0); except childCount := 0; end;

  for i := 1 to childCount do
  begin
    try
      Sub := Subs.Item(i);
      if Mode = smFull then
        ProcessFolderCommon(Sub, FolderId, FullPath, Depth + 1, ACutOffDate, smFull)
      else
        ProcessFolderCommon(Sub, FolderId, FullPath, Depth + 1, ACutOffDate, smMovesOnly);
    except
      on E: Exception do
      begin
        if Mode = smFull then
          Logger.Warn(Format('Subfolder processing error at "%s": %s', [FullPath, E.Message]))
        else
          Logger.Warn(Format('Subfolder (moves-only) error at "%s": %s', [FullPath, E.Message]));
      end;
    end;
  end;
end;

{ ============================= Public entry ============================= }

procedure TMailFullSyncReconciler.Execute(AMovesOnlyScan: Boolean = False; ACutOffDate: TDateTime = 0);
var
  SW: TStopwatch;
begin
  if FIngestor.PstId = 0 then
    FIngestor.LoadPst;

  SW := TStopwatch.StartNew;
  try
    if AMovesOnlyScan then
    begin
      Logger.Info(Format('Starting MOVES-ONLY scan at root "%s" (iter=%s)',
        [FIngestor.RootDisplay, IfThen(FIterMode = imTable, 'Table', 'Restrict')]));
      ProcessFolderMovesOnly(FIngestor.RootFolder, Null, '', 0, ACutOffDate);
    end
    else
    begin
      Logger.Info(Format('Starting full reconcile at root "%s" (iter=%s)',
        [FIngestor.RootDisplay, IfThen(FIterMode = imTable, 'Table', 'Restrict')]));
      ProcessFolder(FIngestor.RootFolder, Null, '', 0, ACutOffDate);
    end;
    // final safety flush after traversal (should be empty already)
    FlushBatches;
  finally
    SW.Stop;
    Logger.Info(Format('ProcessFolder completed in %d ms (%.3f s).',
      [SW.ElapsedMilliseconds, SW.Elapsed.TotalSeconds]));

    try
      if DB.Conn.InTransaction then
        DB.Conn.Commit;
    except
    end;
  end;
  Logger.Info('Full reconcile complete.');
end;

end.

