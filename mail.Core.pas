unit mail.Core;

interface

uses
  Mail.TypesUtils, FireDAC.Stan.Param, System.Classes,
  System.SysUtils, System.Variants, System.StrUtils, System.RegularExpressions,
  System.IOUtils, System.Generics.Collections, System.Generics.Defaults,
  Winapi.Windows, System.Win.ComObj, System.Types,
  FireDAC.Stan.Error, Data.DB, mail.Logger,
  Mail.Data, mail.OutlookEvents;

{ Public facade:
  Creates logger/DB/Outlook components, runs traversal + optional enrichment,
  and returns the number of ingested messages. Raises exceptions on failure. }
procedure RunIngest(const Opt: TCliOptions; out MessagesIngested: Int64);

// Detect the current active Outlook folder and its PST/Store.
// Returns True on success and fills PstPath (Store.FilePath), RootDisplay (store root name),
// and FolderFullPath (Outlook-style path like "\\Store Display\Inbox\Sub").
function DetectCurrentPstFolder(const Outlook: OleVariant;
  out PstPath, RootDisplay, FolderFullPath: string): Boolean;

type
  TOutlookCOM = class
  private
    FOutlook: OleVariant;
    FSession: OleVariant;

  public

    constructor Create;
    procedure Initialize;
    procedure EnsureClassicComAvailable;

    function AttachIfNeeded(const PSTPath: string; out Preexisting: Boolean;
      out TargetStore, RootFolder: OleVariant; out RootDisplay: string): Boolean;

    procedure DetachStoreIfNeeded(const TargetStore: OleVariant; const Preexisting: Boolean);

    property Outlook: OleVariant read FOutlook;
    property Session: OleVariant read FSession;

  end;

  // record to carry attachment data between extraction and persistence
  TAttachmentRow = record
    FileName: string;
    Mime: string;
    SizeBytes: Variant;
    ShaHex: string;
    Content: TByteStr;
  end;

  TIngestor = class
  private
    FDb: TDb;
    FOlk: TOutlookCOM;
    FLogger: TLogger;
    FOpts: TCliOptions;
    FScanCutoffModUtc: TDateTime;
    FPstId: Integer;

    FTargetStore: OleVariant;
    FRootFolder: OleVariant;

    FPreAttachedPST: Boolean;
    FRootDisplay: String;

    { ---- New private helpers for readability/maintainability ---- }

    // Extract attachments according to AttachMode; hashes/bytes as needed.
    procedure GetAttachments(const Item: OleVariant; out OutRows: TArray<TAttachmentRow>);

    // Persist attachments using existing DB method (keeps binary correctness).
    procedure SaveAttachments(const MessageId: Integer; const Rows: TArray<TAttachmentRow>);

    procedure PrintMessageInfo(const Item: OleVariant);
    procedure ClearPSTValues;
    function EnsureFolderPath(const OutlookPath: string; out FolderId: Integer;
      out FolderPath: string): Boolean;
    function EnsureFolderFromVariant(const Folder: OleVariant; out FolderId: Integer;
      out FolderPath: string): Boolean;
    function EnsureFolderByEntryId(const FolderEntryId, StoreId: string; out FolderId: Integer;
      out FolderPath: string; out Folder: OleVariant): Boolean;
    procedure UpsertSingleMessage(const EntryId, StoreId, SourceFolderId: string);
    procedure HandleMoveEvent(const Event: TOutlookItemEvent);
  public
    constructor Create(ADb: TDb; AOlk: TOutlookCOM; ALogger: TLogger; const Opts: TCliOptions);

    procedure InteractiveSearch(AGetTable: Boolean = True);  // continuous prompt (JET via Items.Restrict)

    function LoadPst: Integer;
    procedure UnloadPST;

    procedure ProcessEvents(const Events: TArray<TOutlookItemEvent>);

    property TargetStore: OleVariant read FTargetStore;
    property RootFolder: OleVariant read FRootFolder;
    property RootDisplay: String read FRootDisplay;
    property PreAttachedPST: Boolean read FPreAttachedPST;
    property PstId: Integer read FPstId;

    property DB: TDB read FDb;
    property Olk: TOutlookCOM read FOlk;
    property Logger: TLogger read FLogger;
    property Opts: TCliOptions read FOpts;
    property ScanCutoffModUtc: TDateTime read FScanCutoffModUtc;
  end;

function TryGetCurrentPstRoot(const AOutlookCom: TOutlookCom; const APSTPath: String; out RootFolder: OleVariant; out RootPath: string): Boolean;
function TryGetFolderPath(const Folder: OleVariant): string;

implementation

uses
  mail.FullSync;

{ ---------- COM-based active folder/store detection ---------- }

function GetFolderFromActiveExplorer(const Outlook: OleVariant; out Folder: OleVariant): Boolean;
var
  Exp: OleVariant;
begin
  Result := False;
  Folder := Unassigned;
  try
    Exp := Outlook.ActiveExplorer;
    if not IsVariantAssigned(Exp) then Exit;
    Folder := Exp.CurrentFolder; // MAPIFolder
    Result := IsVariantAssigned(Folder);
  except
  end;
end;

function GetFolderFromActiveInspector(const Outlook: OleVariant; out Folder: OleVariant): Boolean;
var
  Insp, Item: OleVariant;
begin
  Result := False;
  Folder := Unassigned;
  try
    Insp := Outlook.ActiveInspector;
    if not IsVariantAssigned(Insp) then Exit;
    Item := Insp.CurrentItem; // MailItem or other
    if not IsVariantAssigned(Item) then Exit;
    Folder := Item.Parent; // MAPIFolder
    Result := IsVariantAssigned(Folder);
  except
  end;
end;

function TryGetFolderPath(const Folder: OleVariant): string;
begin
  Result := '';
  try
    Result := VarToStrDefSafe(Folder.FolderPath, '');
    if Result = '' then
      Result := VarToStrDefSafe(Folder.Name, '');
  except
  end;
end;

function TryGetStoreForFolder(const Session, Folder: OleVariant; out Store: OleVariant): Boolean;
var
  Stores, S: OleVariant;
  i, n: Integer;
  StoreID, Sid: string;
begin
  Result := False;
  Store := Unassigned;

  try
    Store := Folder.Store;
    if IsVariantAssigned(Store) then
      Exit(True);
  except
  end;

  try
    StoreID := VarToStrDefSafe(Folder.StoreID, '');
  except
    StoreID := '';
  end;
  if StoreID = '' then Exit(False);

  try
    Stores := Session.Stores;
    n := 0;
    try n := Stores.Count; except n := 0; end;
    for i := 1 to n do
    begin
      S := Stores.Item(i);
      Sid := VarToStrDefSafe(S.StoreID, '');
      if (Sid <> '') and SameText(Sid, StoreID) then
      begin
        Store := S;
        Exit(True);
      end;
    end;
  except
  end;
end;

function TryGetStoreFilePathAndRoot(const Store: OleVariant; out FilePath, RootName: string): Boolean;
var
  Root: OleVariant;
begin
  Result := False;
  FilePath := '';
  RootName := '';
  try
    FilePath := VarToStrDefSafe(Store.FilePath, '');
  except
    FilePath := '';
  end;
  try
    Root := Store.GetRootFolder;
    if IsVariantAssigned(Root) then
      RootName := VarToStrDefSafe(Root.Name, '');
  except
  end;
  Result := (FilePath <> '') or (RootName <> '');
end;

function DetectCurrentPstFolder(const Outlook: OleVariant;
  out PstPath, RootDisplay, FolderFullPath: string): Boolean;
var
  Session, Folder, Store: OleVariant;
  HaveFolder, HaveStore: Boolean;
  FP, RN: string;
begin
  Result := False;
  PstPath := '';
  RootDisplay := '';
  FolderFullPath := '';

  if not IsVariantAssigned(Outlook) then
    Exit(False);

  HaveFolder := GetFolderFromActiveExplorer(Outlook, Folder);
  if not HaveFolder then
    HaveFolder := GetFolderFromActiveInspector(Outlook, Folder);
  if not HaveFolder then
    Exit(False);

  FolderFullPath := TryGetFolderPath(Folder);

  Session := Unassigned;
  try
    Session := Outlook.Session;
  except
    try
      Session := Outlook.GetNamespace('MAPI');
    except
      Session := Unassigned;
    end;
  end;
  if not IsVariantAssigned(Session) then
    Exit(False);

  HaveStore := TryGetStoreForFolder(Session, Folder, Store);
  if not HaveStore then
    Exit(False);

  FP := '';
  RN := '';
  if not TryGetStoreFilePathAndRoot(Store, FP, RN) then
    Exit(False);

  PstPath := FP;
  RootDisplay := RN;
  Result := (PstPath <> '') or (FolderFullPath <> '');
end;

{ ---------- Outlook COM ---------- }

constructor TOutlookCOM.Create;
begin
  inherited Create;
  FOutlook := Unassigned;
  FSession := Unassigned;
end;

procedure TOutlookCOM.EnsureClassicComAvailable;
begin
  try
    FOutlook := CreateOleObject('Outlook.Application');
    FSession := FOutlook.GetNamespace('MAPI');
  except
    on E: Exception do
      raise Exception.Create(
        'Outlook COM automation is unavailable. ' +
        'This tool requires Classic Outlook (the "new Outlook for Windows" does not support COM automation). ' +
        'Install Classic Outlook and try again. Underlying error: ' + E.Message);
  end;
end;

procedure TOutlookCOM.Initialize;
begin
  EnsureClassicComAvailable;
end;

function TOutlookCOM.AttachIfNeeded(const PSTPath: string; out Preexisting: Boolean;
  out TargetStore, RootFolder: OleVariant; out RootDisplay: string): Boolean;
var
  Stores, st: OleVariant;
  i, nStores: Integer;
  StorePath: string;
begin
  Preexisting := False;
  TargetStore := Unassigned;
  RootFolder  := Unassigned;
  RootDisplay := '';

  Stores := FSession.Stores;
  nStores := 0;
  try nStores := Stores.Count; except nStores := 0; end;

  for i := 1 to nStores do
  begin
    st := Stores.Item(i);
    StorePath := VarToStrDefSafe(st.FilePath, '');
    if SameText(StorePath, PSTPath) then
    begin
      Preexisting := True;
      TargetStore := st;
      Break;
    end;
  end;

  if not IsVariantAssigned(TargetStore) then
  begin
    try
      FSession.AddStoreEx(PSTPath, OL_STORE_UNICODE);
    except
      on E: Exception do
        raise Exception.Create('Failed to attach PST via Outlook: ' + E.Message);
    end;

    Stores := FSession.Stores;
    nStores := 0;
    try nStores := Stores.Count; except nStores := 0; end;
    for i := 1 to nStores do
    begin
      st := Stores.Item(i);
      StorePath := VarToStrDefSafe(st.FilePath, '');
      if SameText(StorePath, PSTPath) then
      begin
        TargetStore := st;
        Break;
      end;
    end;
  end;

  if not IsVariantAssigned(TargetStore) then
    raise Exception.Create('Outlook PST store not found/attachable for path: ' + PSTPath);

  RootFolder  := TargetStore.GetRootFolder;
  RootDisplay := VarToStrDefSafe(RootFolder.Name, '');
  if RootDisplay = '' then
    RootDisplay := VarToStrDefSafe(TargetStore.DisplayName, '');

  Result := True;
end;

procedure TOutlookCOM.DetachStoreIfNeeded(const TargetStore: OleVariant; const Preexisting: Boolean);
var
  Root: OleVariant;
begin
  if not Preexisting then
  begin
    try
      Root := TargetStore.GetRootFolder;
      FSession.RemoveStore(Root);
    except
    end;
  end;
end;

{ ---------- Ingestor ---------- }

constructor TIngestor.Create(ADb: TDb; AOlk: TOutlookCOM; ALogger: TLogger; const Opts: TCliOptions);
begin
  inherited Create;
  FDb     := ADb;
  FOlk    := AOlk;
  FLogger := ALogger;
  FOpts   := Opts;
  FScanCutoffModUtc := 0; // explicit default
end;

procedure TIngestor.GetAttachments(const Item: OleVariant; out OutRows: TArray<TAttachmentRow>);
var
  Attachments, Attachment, AttPropAccessor: OleVariant;
  j: Integer;
  FileName, FileNameSave, Mime, TmpDir, TmpPath: string;
  SizeBytesVar: Variant;
  ContentRaw: TByteStr;
  ShaHex: string;
  SavedToFile: Boolean;
  AttachMethod: Integer;
  AttachBinaryVariant: OleVariant;
begin
  SetLength(OutRows, 0);
  if FOpts.AttachMode = amNone then Exit;

  try
    Attachments := Item.Attachments;
  except
    Attachments := Unassigned;
  end;

  if not IsVariantAssigned(Attachments) then Exit;
  if Attachments.Count = 0 then Exit;

  TmpDir := TPath.Combine(TPath.GetTempPath, cAttachTempPrefix + IntToStr(GetCurrentProcessId));
  ForceDirectories(TmpDir);
  try
    for j := 1 to Attachments.Count do
    begin
      Attachment := Attachments.Item(j);
      FileName := VarToStrDefSafe(Attachment.FileName, Format('attachment_%d', [j]));
      FileNameSave := SanitizeFileName(FileName);
      SizeBytesVar := Unassigned;
      try SizeBytesVar := Attachment.Size; except end;

      Mime := '';
      try Mime := VarToStrDefSafe(Attachment.PropertyAccessor.GetProperty(PR_ATTACH_MIME_TAG), ''); except end;

      ContentRaw := '';
      ShaHex := '';
      SavedToFile := False;

      if (FOpts.AttachMode = amMetaHash) or (FOpts.AttachMode = amBytes) then
      begin
        TmpPath := TPath.Combine(TmpDir, Format('%d_%s', [j, FileNameSave]));
        try
          Attachment.SaveAsFile(TmpPath);
          ContentRaw := ReadAllBytesRaw(TmpPath);
          SavedToFile := True;
        except
          SavedToFile := False;
        end;

        if not SavedToFile then
        begin
          try
            AttPropAccessor := Attachment.PropertyAccessor;
            try attachMethod := Integer(AttPropAccessor.GetProperty(PR_ATTACH_METHOD)); except attachMethod := 0; end;
            AttachBinaryVariant := Unassigned;
            try AttachBinaryVariant := AttPropAccessor.GetProperty(PR_ATTACH_DATA_BIN); except AttachBinaryVariant := Unassigned; end;
            if IsVariantAssigned(AttachBinaryVariant) then
              ContentRaw := VariantBinaryToByteStr(AttachBinaryVariant);
          except
            // ignore
          end;
        end;

        if Length(ContentRaw) > 0 then
        begin
          ShaHex := RawSHA256Hex(ContentRaw);
          if FOpts.AttachMode <> amBytes then
            ContentRaw := '';
        end;

        try
          if SavedToFile and TFile.Exists(TmpPath) then
            TFile.Delete(TmpPath);
        except end;
      end;

      SetLength(OutRows, Length(OutRows) + 1);
      OutRows[High(OutRows)].FileName  := FileName;
      OutRows[High(OutRows)].Mime      := Mime;
      OutRows[High(OutRows)].SizeBytes := SizeBytesVar;
      OutRows[High(OutRows)].ShaHex    := ShaHex;
      OutRows[High(OutRows)].Content   := ContentRaw;
    end;
  finally
    try
      if TDirectory.Exists(TmpDir) then
        TDirectory.Delete(TmpDir, True);
    except end;
  end;
end;

procedure TIngestor.SaveAttachments(const MessageId: Integer; const Rows: TArray<TAttachmentRow>);
var
  i: Integer;
begin
  for i := 0 to High(Rows) do
  begin
    try
      FDb.InsertAttachment(
        MessageId,
        Rows[i].FileName,
        Rows[i].Mime,
        Rows[i].SizeBytes,
        Rows[i].ShaHex,
        Rows[i].Content
      );
    except
      on E: Exception do
        FLogger.Warn(Format('Attachment insert error (%s): %s', [Rows[i].FileName, E.Message]));
    end;
  end;
end;

function TIngestor.LoadPst: Integer;
begin
  Assert(FOpts.PSTPath <> '', 'TIngestor.LoadPst: PSTPath is not set!');
  FOlk.AttachIfNeeded(FOpts.PSTPath, FPreAttachedPST, FTargetStore, FRootFolder, FRootDisplay);

  FPstID := FDb.EnsurePstRow(FOpts.PSTPath, RootDisplay);
  Result:= FPstId;
end;

procedure TIngestor.ClearPSTValues;
begin
  FTargetStore:= Unassigned;
  FRootFolder:= Unassigned;
  FRootDisplay:= Unassigned;
  FPreAttachedPST:= False;
  FPstId:= 0;
end;

function TIngestor.EnsureFolderPath(const OutlookPath: string; out FolderId: Integer;
  out FolderPath: string): Boolean;
var
  Segments: TArray<string>;
  segment, tSegment: string;
  ParentIdVar: Variant;
  CurrPath, CleanPath: string;
  Depth: Integer;
begin
  FolderId := 0;
  FolderPath := '';
  Result := False;
  CleanPath := Trim(OutlookPath);
  if CleanPath = '' then
    Exit(False);

  if CleanPath.StartsWith('\\') then
    Delete(CleanPath, 1, 2);

  Segments := CleanPath.Split(['\']);
  ParentIdVar := Null;
  CurrPath := '';
  Depth := 0;

  for tSegment in Segments do
  begin
    Segment := Trim(tSegment);
    if Segment = '' then
      Continue;

    if CurrPath <> '' then
      CurrPath := CurrPath + '/' + Segment
    else
      CurrPath := Segment;

    FolderId := FDb.EnsureFolderRow(FPstId, ParentIdVar, Segment, CurrPath, Depth);
    ParentIdVar := FolderId;
    Inc(Depth);
  end;

  FolderPath := CurrPath;
  Result := FolderId <> 0;
end;

function TIngestor.EnsureFolderFromVariant(const Folder: OleVariant; out FolderId: Integer;
  out FolderPath: string): Boolean;
var
  OutlookPath: string;
begin
  FolderId := 0;
  FolderPath := '';
  Result := False;
  if not IsVariantAssigned(Folder) then
    Exit(False);

  OutlookPath := TryGetFolderPath(Folder);
  if OutlookPath = '' then
    OutlookPath := VarToStrDefSafe(Folder.Name, '');

  Result := EnsureFolderPath(OutlookPath, FolderId, FolderPath);
end;

function TIngestor.EnsureFolderByEntryId(const FolderEntryId, StoreId: string;
  out FolderId: Integer; out FolderPath: string; out Folder: OleVariant): Boolean;
begin
  FolderId := 0;
  FolderPath := '';
  Folder := Unassigned;
  Result := False;
  if FolderEntryId = '' then
    Exit(False);

  try
    if StoreId <> '' then
      Folder := FOlk.Session.GetFolderFromID(FolderEntryId, StoreId)
    else
      Folder := FOlk.Session.GetFolderFromID(FolderEntryId);
  except
    Folder := Unassigned;
  end;

  if not IsVariantAssigned(Folder) then
    Exit(False);

  Result := EnsureFolderFromVariant(Folder, FolderId, FolderPath);
end;

procedure TIngestor.UpsertSingleMessage(const EntryId, StoreId, SourceFolderId: string);
var
  Item, Folder: OleVariant;
  InternetMessageId, OutlookEntryId, Subject, SenderName, SenderEmail: string;
  DisplayTo, DisplayCc, Headers, BodyText, BodyHtml, MsgClass: string;
  SentUtc, RecvUtc, CreateUtc, LastModNew: TDateTime;
  SearchKeyRead: TBytes;
  RecRows: TArray<TArray<Variant>>;
  AttachRows: TArray<TAttachmentRow>;
  FolderId: Integer;
  FolderFullPath: string;
  MessageId: Integer;
  SizeVariant: Variant;
begin
  if EntryId = '' then
    Exit;

  Item := Unassigned;
  try
    if StoreId <> '' then
      Item := FOlk.Session.GetItemFromID(EntryId, StoreId)
    else
      Item := FOlk.Session.GetItemFromID(EntryId);
  except
    on E: Exception do
    begin
      if Assigned(FLogger) then
        FLogger.Warn(Format('UpsertSingleMessage: Failed to get item with EntryId %s (StoreId: %s). Error: %s', [EntryId, StoreId, E.Message]));
      Item := Unassigned;
    end;
  end;
  if not IsVariantAssigned(Item) then
    Exit;

  FolderId := 0;
  FolderFullPath := '';
  try
    Folder := Item.Parent;
  except
    Folder := Unassigned;
  end;

  if not EnsureFolderFromVariant(Folder, FolderId, FolderFullPath) and (SourceFolderId <> '') then
    EnsureFolderByEntryId(SourceFolderId, StoreId, FolderId, FolderFullPath, Folder);

  if FolderId = 0 then
  begin
    if Assigned(FLogger) then
      FLogger.Warn(Format('Unable to resolve folder for entry %s', [EntryId]));
    Exit;
  end;

  ReadMessageCore(Item,
    InternetMessageId, OutlookEntryId, Subject, SenderName, SenderEmail,
    DisplayTo, DisplayCc, Headers, BodyText, BodyHtml, MsgClass,
    SentUtc, RecvUtc, CreateUtc, LastModNew, SearchKeyRead);

  if OutlookEntryId = '' then
    OutlookEntryId := EntryId;

  try
    GetRecipients(Item, Headers, RecRows);
  except
    SetLength(RecRows, 0);
  end;

  try
    GetAttachments(Item, AttachRows);
  except
    SetLength(AttachRows, 0);
  end;

  SizeVariant := Null;
  try
    SizeVariant := Item.Size;
  except
    SizeVariant := Null;
  end;

  MessageId := FDb.MessageExistsByEntryId(FPstId, OutlookEntryId);
  if MessageId = 0 then
  begin
    MessageId := FDb.InsertMessageReturnId(
      FPstId, FolderId,
      InternetMessageId, OutlookEntryId, Subject, SenderName, SenderEmail,
      DisplayTo, DisplayCc,
      SentUtc, RecvUtc, CreateUtc,
      Headers, BodyText, BodyHtml,
      SizeVariant, LastModNew, SearchKeyRead);

    if Length(AttachRows) > 0 then
      SaveAttachments(MessageId, AttachRows);
  end
  else
  begin
    FDb.UpdateMessageCoreFields(
      MessageId,
      InternetMessageId, OutlookEntryId, Subject, SenderName, SenderEmail,
      DisplayTo, DisplayCc,
      SentUtc, RecvUtc, CreateUtc,
      Headers, BodyText, BodyHtml,
      LastModNew, SearchKeyRead);
    FDb.UpdateMessageFolder(MessageId, FolderId);
  end;

  if Length(RecRows) > 0 then
    FDb.SaveRecipients(MessageId, RecRows);
end;

procedure TIngestor.HandleMoveEvent(const Event: TOutlookItemEvent);
var
  Folder: OleVariant;
  FolderId: Integer;
  FolderPath: string;
  StoreId: string;
  MessageId: Integer;
begin
  if Event.EntryId = '' then
    Exit;

  StoreId := Event.TargetStoreId;
  if StoreId = '' then
    StoreId := Event.StoreId;

  FolderId := 0;
  if Event.TargetFolderId <> '' then
    EnsureFolderByEntryId(Event.TargetFolderId, StoreId, FolderId, FolderPath, Folder);

  if FolderId <> 0 then
  begin
    MessageId := FDb.MessageExistsByEntryId(FPstId, Event.EntryId);
    if MessageId <> 0 then
      FDb.UpdateMessageFolder(MessageId, FolderId);
  end;

//  UpsertSingleMessage(Event.EntryId, StoreId, Event.TargetFolderId);
end;

procedure TIngestor.ProcessEvents(const Events: TArray<TOutlookItemEvent>);
var
  Ev: TOutlookItemEvent;
begin
  for Ev in Events do
  begin
    try
      case Ev.EventKind of
        oekAdded, oekChanged:
          UpsertSingleMessage(Ev.EntryId, Ev.StoreId, Ev.SourceFolderId);
        oekMoved:
          HandleMoveEvent(Ev);
      end;
    except
      on E: Exception do
        if Assigned(FLogger) then
          FLogger.Warn(Format('Event processing error for %s: %s',
            [Ev.EntryId, E.Message]));
    end;
  end;
end;

procedure TIngestor.UnloadPST;
begin
  FOlk.DetachStoreIfNeeded(FTargetStore, FPreAttachedPST);

  ClearPSTValues;
end;

function TryGetCurrentPstRoot(const AOutlookCom: TOutlookCom; const APSTPath: String; out RootFolder: OleVariant; out RootPath: string): Boolean;
var
  Stores, St: OleVariant;
  i, n: Integer;
  StorePath: string;
begin
  Result := False;
  RootFolder := Unassigned;
  RootPath := '';

  try
    Stores := AOutlookCom.Session.Stores;
    try n := Stores.Count; except n := 0; end;
    for i := 1 to n do
    begin
      St := Stores.Item(i);
      StorePath := VarToStrDefSafe(St.FilePath, '');
      if SameText(Trim(StorePath), Trim(APSTPath)) then
      begin
        RootFolder := St.GetRootFolder;
        RootPath := TryGetFolderPath(RootFolder); // e.g. \\Store Display
        Result := IsVariantAssigned(RootFolder) and (RootPath <> '');
        Exit;
      end;
    end;
  except
    // ignore and leave Result=False
  end;
end;

procedure TIngestor.PrintMessageInfo(const Item: OleVariant);
var
  subj, fromName, recvStr, folderPath: string;
  recv: TDateTime;
  parentFolder: OleVariant;
  searchKeyHex: string;
  pa, vSk: OleVariant;
  skRaw: TByteStr;

  function ByteStrToHexLower(const S: TByteStr): string;
  const
    Hex: array[0..15] of Char = '0123456789abcdef';
  var
    i: Integer;
  begin
    SetLength(Result, Length(S) * 2);
    for i := 0 to Length(S) - 1 do
    begin
      Result[i*2+1] := Hex[Byte(S[i+1]) shr 4];
      Result[i*2+2] := Hex[Byte(S[i+1]) and $0F];
    end;
  end;

begin
  subj := ''; fromName := ''; recvStr := ''; folderPath := ''; searchKeyHex := '';

  try subj := VarToStrDefSafe(Item.Subject, ''); except end;
  try fromName := VarToStrDefSafe(Item.SenderName, ''); except end;

  // Correct way to fetch PR_SEARCH_KEY (PT_BINARY) and make it printable
  try
    pa := Item.PropertyAccessor;
    vSk := pa.GetProperty(PR_SEARCH_KEY); // 'http://schemas.microsoft.com/mapi/proptag/0x300B0102'
    if IsVariantAssigned(vSk) then
    begin
      skRaw := VariantBinaryToByteStr(vSk);  // existing helper in your codebase
      searchKeyHex := ByteStrToHexLower(skRaw);
    end;
  except
    searchKeyHex := '';
  end;

  try
    recv := Item.ReceivedTime;
    if recv <> 0 then
      recvStr := FormatDateTime('yyyy-mm-dd hh:nn', recv);
  except end;

  parentFolder := Unassigned;
  try
    parentFolder := Item.Parent;
    if IsVariantAssigned(parentFolder) then
      folderPath := TryGetFolderPath(parentFolder);
  except end;

  try
    WriteLn(Format('  - [%s] %s - %s  (%s) sk: %s',
      [recvStr, subj, fromName, folderPath, searchKeyHex]));
  except
    on E: Exception do
      WriteLN(Format('Error on TIngestor.PrintHitInfo => %s: %s', [E.ClassName, E.Message]));
  end;
end;

procedure TIngestor.InteractiveSearch(AGetTable: Boolean = True);
var
  Root, Folder, Subs: OleVariant;
  Stack: TStack<OleVariant>;
  RootPath, JetFilter: string;
  Total, Shown: Integer;
  nSubs, i: Integer;

  procedure ProcessFolder(const Fld: OleVariant);
  var
    localItems, localView, it: OleVariant;
    countHere: Integer;

    procedure DoSearchGetTable;
    var
      Tbl, Row: OleVariant;
    begin
      Tbl:= Fld.GetTable(JetFilter, 0);
      try Tbl.MoveToStart; except end;
      try countHere := Integer(Tbl.GetRowCount); except countHere := 0; end;
      Inc(Total, countHere);
      while (not Tbl.EndOfTable) do
      begin
        Row := Tbl.GetNextRow;
        if not IsVariantAssigned(Row) then Break;
        it := Olk.Session.GetItemFromID(Row.Item('EntryId'));
        if Integer(it.Class) = OL_CLASS_MAIL then
        begin
          PrintMessageInfo(it);
          Inc(Shown);
        end;

//        if Shown >= 5 then
//          Exit;
      end;
    end;

    procedure DoSearchRestrict;
    var
      I: Integer;
    begin
       if JetFilter <> '' then
         localView := localItems.Restrict(JetFilter)
       else
         localView:= localItems;

      try countHere := Integer(localView.Count); except countHere := 0; end;
      Inc(Total, countHere);

      for I := 1 to countHere do
      begin
        it:= localView.Item(I);
        try
          if Integer(it.Class) = OL_CLASS_MAIL then
          begin
            PrintMessageInfo(it);
            Inc(Shown);
          end;
        except
          // ignore non-mail
        end;
//        if Shown >= 5 then
//          Exit;
      end;
    end;

  const
     // Mail-only constraint for Jet filters
    cJetMailClassFilter = '[MessageClass] Like ''IPM.Note%''';
  begin
    localItems := Unassigned;
    try localItems := Fld.Items; except localItems := Unassigned; end;
    if not IsVariantAssigned(localItems) then Exit;

  // Apply JET filter (local time semantics)
    try
     if AGetTable then
       DoSearchGetTable
    else
       DoSearchRestrict;


    except
      on E: Exception do
      begin
        Writeln(Format('WARN Restrict failed in "%s": %s',
          [TryGetFolderPath(Fld), E.Message]));
        Exit; // skip this folder, continue traversal
      end;
    end;
  end;

begin
  if not TryGetCurrentPstRoot(FOlk, FOpts.PSTPath, Root, RootPath) then
    raise Exception.Create('Unable to locate current PST root for JET search.');

  WriteLn('JET search on current PST. Type a JET filter. Blank = all. Type "exit" to quit.');
  WriteLn('Examples: [Subject] Like ''%invoice%''');
  WriteLn('          [LastModificationTime] > ''08/01/2025 00:00''');

  while True do
  try
    Write('JET> ');
    ReadLn(JetFilter);
    if SameText(Trim(JetFilter), 'exit') or SameText(Trim(JetFilter), 'quit') then Break;

    // Walk the PST (root + all subfolders)
    Total := 0;
    Shown := 0;

    Stack := TStack<OleVariant>.Create;
    try
      Stack.Push(Root);
      while (Stack.Count > 0) do
      begin
        Folder := Stack.Pop;
        ProcessFolder(Folder);

        // Recurse
        Subs := Unassigned;
        try Subs := Folder.Folders; except Subs := Unassigned; end;
        if IsVariantAssigned(Subs) then
        begin
          try nSubs := Subs.Count; except nSubs := 0; end;
          for i := 1 to nSubs do
            Stack.Push(Subs.Item(i));
        end;
      end;
    finally
      Stack.Free;
    end;

    WriteLn(Format('Total results: %d  (showing first %d)', [Total, Shown]));
  except
    on E: Exception do
    begin
      Writeln('ERROR ' + E.ClassName + ': ' + E.Message);
    end;
  end;
end;

{ ---------- Public Facade ---------- }

procedure RunIngest(const Opt: TCliOptions; out MessagesIngested: Int64);
var
  Logger: TLogger;
  Db: TDb;
  Olk: TOutlookCOM;
  Ingestor: TIngestor;
  EffectiveOpt: TCliOptions;
  DetPst, DetRoot, DetFolderPath: string;
  FScanCutoffModUtc: TDateTime;
  FIterMode: TIterMode;
begin
  MessagesIngested := 0;

  Logger := TLogger.Create(Opt.LogPath);
  try
    Db := TDb.Create(Opt.ConnOverride, Opt.EnvFile);
    try
      Olk := TOutlookCOM.Create;
      try
        Olk.Initialize;

        EffectiveOpt := Opt;

        if (Trim(EffectiveOpt.PSTPath) = '') or SameText(Trim(EffectiveOpt.PSTPath), 'active') then
        begin
          if not DetectCurrentPstFolder(Olk.Outlook, DetPst, DetRoot, DetFolderPath) then
            raise Exception.Create('Unable to detect the active Outlook folder/store via COM. ' +
                                   'Open Classic Outlook, select a folder in the target PST, or pass --pst=<path>.');

          if DetPst = '' then
            raise Exception.Create('Detected active store has no local file path (likely an online mailbox). ' +
                                   'Please pass a PST path explicitly with --pst=<path>.');

          Logger.Info(Format('Detected active store "%s" at: %s', [DetRoot, DetPst]));
          if DetFolderPath <> '' then
            Logger.Info('Active folder: ' + DetFolderPath);

          EffectiveOpt.PSTPath := DetPst;
        end;

        Ingestor := TIngestor.Create(Db, Olk, Logger, EffectiveOpt);
        try
          if Opt.QueryMode then
          begin
            Writeln('Running JET interactive filter. Write exit or quit to end.');
            Ingestor.InteractiveSearch(Opt.UseRestrict = False);
          end
          else begin
            Ingestor.LoadPst;
            if Opt.UseRestrict then
              FIterMode:= imRestrict
            else
              FIterMode:= imTable;
            with TMailFullSyncReconciler.Create(Ingestor, False, FIterMode) do
            begin
              if Opt.FullSync then
                Execute(Opt.MovesOnly, 0)
              else
              begin
                FScanCutoffModUtc := Db.GetScanCutoffLastModUtc(Ingestor.PstId);
                Execute(Opt.MovesOnly, UTCDateToLocal(FScanCutoffModUtc));
              end;
            end;
          end;
        finally
          try Ingestor.UnloadPst // Detach the PST if needed;
          except end;
          Ingestor.Free;
        end;

      finally
        Olk.Free;
      end;
    finally
      Db.Free;
    end;
  finally
    Logger.Free;
  end;
end;

end.

