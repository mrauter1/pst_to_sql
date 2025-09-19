unit mail.OutlookEvents;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, System.Variants,
  Winapi.ActiveX, System.Win.ComObj,
  mail.TypesUtils, mail.Logger, WinAPI.Windows;

type
  TOutlookItemEventKind = (oekAdded, oekChanged, oekRemoved, oekMoved);

  TOutlookItemEvent = record
    EntryId: string;
    StoreId: string;
    SourceFolderId: string;
    TargetFolderId: string;
    TargetStoreId: string;
    EventKind: TOutlookItemEventKind;
  end;

  TOutlookEventProc = procedure(const Event: TOutlookItemEvent) of object;

  TOutlookEventListener = class;

  TInvokeMethod = procedure(DispId: Integer; const Params: TArray<Variant>) of object;

  TComEventSink = class(TInterfacedObject, IDispatch)
  private
    FOnInvoke: TInvokeMethod;
    FLogger: TLogger;
  public
    constructor Create(const AOnInvoke: TInvokeMethod; ALogger: TLogger);

    { IDispatch }
    function GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; stdcall;
  end;

  TComConnection = class
  private
    FConnection: IConnectionPoint;
    FCookie: Integer;
    FSink: IDispatch;
  public
    constructor Create(const Source: IDispatch; const EventIID: TGUID; const Sink: IDispatch);
    destructor Destroy; override;

    procedure Disconnect;
  end;

  TFolderWatcher = class
  private
    FLogger: TLogger;
    FListener: TOutlookEventListener;
    FFolder: OleVariant;
    FItems: OleVariant;
    FFolderEntryId: string;
    FStoreId: string;
    FItemsSink: IDispatch;
    FFolderSink: IDispatch;
    FItemsConnection: TComConnection;
    FFolderConnection: TComConnection;
    procedure ItemsInvoke(DispId: Integer; const Params: TArray<Variant>);
    procedure FolderInvoke(DispId: Integer; const Params: TArray<Variant>);
    procedure NotifyAddChange(const Item: OleVariant; const Kind: TOutlookItemEventKind);
  public
    constructor Create(AListener: TOutlookEventListener; const Folder: OleVariant; ALogger: TLogger);
    destructor Destroy; override;
  end;

  TOutlookEventListener = class
  private
    FApp: OleVariant;
    FSession: OleVariant;
    FRootFolder: OleVariant;
    FOnEvent: TOutlookEventProc;
    FWatchers: TObjectList<TFolderWatcher>;
    FAppSink: IDispatch;
    FAppConnection: TComConnection;
    FLogger: TLogger;

    procedure ApplicationInvoke(DispId: Integer; const Params: TArray<Variant>);
    procedure BuildWatchers(const Folder: OleVariant);
    procedure Notify(const Event: TOutlookItemEvent);
  public
    constructor Create(const OutlookApp, Session, RootFolder: OleVariant; const Callback: TOutlookEventProc; ALogger: TLogger);
    destructor Destroy; override;

    procedure Refresh;
  end;

implementation

const
  DIID_ApplicationEvents_11: TGUID = '{0006304E-0000-0000-C000-000000000046}';
  DIID_ItemsEvents: TGUID           = '{00063079-0000-0000-C000-000000000046}';
  DIID_MAPIFolderEvents: TGUID      = '{000630F6-0000-0000-C000-000000000046}';

{ ------------------------------ Helpers ------------------------------ }

function VariantArgToVariant(const Arg: PVariantArg): Variant;
begin
  if Arg = nil then
    Result := Null
  else
    Result := OleVariant(Arg^);
end;

function VarToDispatch(const V: OleVariant): IDispatch;
begin
  Result := nil;
  if TVarData(V).VType = varDispatch then
    Result := IDispatch(TVarData(V).VDispatch)
  else if TVarData(V).VType = varUnknown then
    Result := IDispatch(TVarData(V).VUnknown);
end;

{ ------------------------------ TComEventSink ------------------------------ }

constructor TComEventSink.Create(const AOnInvoke: TInvokeMethod; ALogger: TLogger);
begin
  inherited Create;
  FLogger:= ALogger;
  FOnInvoke := AOnInvoke;
end;

function TComEventSink.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;

function TComEventSink.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult;
begin
  Pointer(TypeInfo) := nil;
  Result := E_NOTIMPL;
end;

function TComEventSink.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Count := 0;
  Result := E_NOTIMPL;
end;

function TComEventSink.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
var
  DispParams: PDispParams;
  Args: TArray<Variant>;
  i: Integer;
  P: PVariantArg;
begin
  Result := S_OK;
  if not Assigned(FOnInvoke) then
    Exit;

  try
    DispParams := @Params;
    if (DispParams <> nil) and (DispParams^.cArgs > 0) then
    begin
      SetLength(Args, DispParams^.cArgs);
      for i := 0 to DispParams^.cArgs - 1 do
      begin
        // COM passes arguments in reverse order
        P := @DispParams^.rgvarg[DispParams^.cArgs - 1 - i];
        Args[i] := VariantArgToVariant(P);
      end;
    end
    else
      SetLength(Args, 0);

    FOnInvoke(DispID, Args);
  except
    on E: Exception do
    begin
      if Assigned(FLogger) then
        FLogger.Error('TComEventSink.Invoke Error: ' + E.Message);

      Result := S_OK;
    end;
  end;
end;

{ ------------------------------ TComConnection ------------------------------ }

constructor TComConnection.Create(const Source: IDispatch; const EventIID: TGUID;
  const Sink: IDispatch);
var
  CPC: IConnectionPointContainer;
  CP: IConnectionPoint;
begin
  inherited Create;
  FSink := Sink;
  FCookie := 0;
  if (Source = nil) or (Sink = nil) then
    Exit;

  if Supports(Source, IConnectionPointContainer, CPC) then
  begin
    if Succeeded(CPC.FindConnectionPoint(EventIID, CP)) then
    begin
      FConnection := CP;
      if Failed(FConnection.Advise(Sink, FCookie)) then
      begin
        FCookie := 0;
        FConnection := nil;
      end;
    end;
  end;
end;

destructor TComConnection.Destroy;
begin
  Disconnect;
  inherited;
end;

procedure TComConnection.Disconnect;
begin
  if (FConnection <> nil) and (FCookie <> 0) then
    FConnection.Unadvise(FCookie);
  FCookie := 0;
  FConnection := nil;
  FSink := nil;
end;

{ ------------------------------ TFolderWatcher ------------------------------ }

constructor TFolderWatcher.Create(AListener: TOutlookEventListener; const Folder: OleVariant; ALogger: TLogger);
var
  ItemsDisp, FolderDisp: IDispatch;
begin
  inherited Create;
  FListener := AListener;
  FFolder := Folder;
  FLogger:= ALogger;
  FFolderEntryId := VarToStrDefSafe(Folder.EntryID, '');
  FStoreId := VarToStrDefSafe(Folder.StoreID, '');

  try
    FItems := Folder.Items;
  except
    FItems := Unassigned;
  end;

  ItemsDisp := VarToDispatch(FItems);
  if ItemsDisp <> nil then
  begin
    FItemsSink := TComEventSink.Create(ItemsInvoke, FLogger);
    FItemsConnection := TComConnection.Create(ItemsDisp, DIID_ItemsEvents, FItemsSink);
  end;

  FolderDisp := VarToDispatch(FFolder);
  if FolderDisp <> nil then
  begin
    FFolderSink := TComEventSink.Create(FolderInvoke, ALogger);
    FFolderConnection := TComConnection.Create(FolderDisp, DIID_MAPIFolderEvents, FFolderSink);
  end;
end;

destructor TFolderWatcher.Destroy;
begin
  FreeAndNil(FItemsConnection);
  FreeAndNil(FFolderConnection);
  FItemsSink := nil;
  FFolderSink := nil;
  FItems := Unassigned;
  FFolder := Unassigned;
  inherited;
end;

procedure TFolderWatcher.ItemsInvoke(DispId: Integer; const Params: TArray<Variant>);
var
  Item: OleVariant;
  Kind: TOutlookItemEventKind;
begin
  if (Length(Params) = 0) then
    Exit;

  Item := Variant(Params[0]);
  if not IsVariantAssigned(Item) then
    Exit;

  case DispId of
    61441: Kind := oekAdded;   // ItemAdd
    61442: Kind := oekChanged; // ItemChange
    61443: Kind := oekRemoved; // ItemRemove (no item reference in Outlook, but guard anyway)
  else
    Exit;
  end;

  if Kind = oekRemoved then
    Exit; // Outlook does not pass the item, so EntryID cannot be resolved reliably

  NotifyAddChange(Item, Kind);
end;

procedure TFolderWatcher.NotifyAddChange(const Item: OleVariant; const Kind: TOutlookItemEventKind);
var
  Ev: TOutlookItemEvent;
  Parent: OleVariant;
  StoreId: string;
begin
  if FListener = nil then
    Exit;

  FillChar(Ev, SizeOf(Ev), 0);
  Ev.EventKind := Kind;

  Ev.EntryId := VarToStrDefSafe(Item.EntryID, '');
  if Ev.EntryId = '' then
    Exit;

  StoreId := '';
  try
    Parent := Item.Parent;
  except
    Parent := Unassigned;
  end;

  if IsVariantAssigned(Parent) then
    StoreId := VarToStrDefSafe(Parent.StoreID, '');
  if StoreId = '' then
    StoreId := FStoreId;

  Ev.StoreId := StoreId;
  Ev.SourceFolderId := FFolderEntryId;

  FListener.Notify(Ev);
end;

procedure TFolderWatcher.FolderInvoke(DispId: Integer; const Params: TArray<Variant>);
var
  Item, MoveTo, Parent: OleVariant;
  Ev: TOutlookItemEvent;
  TargetStore: string;
begin
  if DispId <> 61444 then // BeforeItemMove
    Exit;

  if Length(Params) < 2 then
    Exit;

  Item := Variant(Params[0]);
  MoveTo := Variant(Params[1]);

  if not IsVariantAssigned(Item) then
    Exit;

  FillChar(Ev, SizeOf(Ev), 0);
  Ev.EventKind := oekMoved;
  Ev.EntryId := VarToStrDefSafe(Item.EntryID, '');
  if Ev.EntryId = '' then
    Exit;

  Parent := Unassigned;
  try
    Parent := Item.Parent;
  except
    Parent := Unassigned;
  end;
  if IsVariantAssigned(Parent) then
    Ev.StoreId := VarToStrDefSafe(Parent.StoreID, FStoreId)
  else
    Ev.StoreId := FStoreId;
  Ev.SourceFolderId := FFolderEntryId;

  if IsVariantAssigned(MoveTo) then
  begin
    Ev.TargetFolderId := VarToStrDefSafe(MoveTo.EntryID, '');
    TargetStore := VarToStrDefSafe(MoveTo.StoreID, '');
    if TargetStore = '' then
      TargetStore := Ev.StoreId;
    Ev.TargetStoreId := TargetStore;
  end;

  if FListener <> nil then
    FListener.Notify(Ev);
end;

{ ------------------------------ TOutlookEventListener ------------------------------ }

constructor TOutlookEventListener.Create(const OutlookApp, Session, RootFolder: OleVariant;
  const Callback: TOutlookEventProc; ALogger: TLogger);
var
  AppDisp: IDispatch;
begin
  inherited Create;
  FApp := OutlookApp;
  FSession := Session;
  FRootFolder := RootFolder;
  FOnEvent := Callback;
  FWatchers := TObjectList<TFolderWatcher>.Create(True);
  FLogger:= ALogger;

  AppDisp := VarToDispatch(FApp);
  if AppDisp <> nil then
  begin
    FAppSink := TComEventSink.Create(ApplicationInvoke, ALogger);
    FAppConnection := TComConnection.Create(AppDisp, DIID_ApplicationEvents_11, FAppSink);
  end;

  Refresh;
end;

destructor TOutlookEventListener.Destroy;
begin
  FreeAndNil(FWatchers);
  FreeAndNil(FAppConnection);
  FAppSink := nil;
  FApp := Unassigned;
  FSession := Unassigned;
  FRootFolder := Unassigned;
  inherited;
end;

procedure TOutlookEventListener.ApplicationInvoke(DispId: Integer; const Params: TArray<Variant>);
var
  EntryList: TArray<string>;
  RawIds, Id: string;
  Ev: TOutlookItemEvent;
  i: Integer;
begin
  if DispId <> 61446 then
    Exit;

  if Length(Params) = 0 then
    Exit;

  RawIds := VarToStrDefSafe(Params[0], '');
  if RawIds = '' then
    Exit;

  EntryList := RawIds.Split([';']);
  for i := 0 to High(EntryList) do
  begin
    Id := Trim(EntryList[i]);
    if Id = '' then
      Continue;

    FillChar(Ev, SizeOf(Ev), 0);
    Ev.EventKind := oekAdded;
    Ev.EntryId := Id;
    Notify(Ev);
  end;
end;

procedure TOutlookEventListener.BuildWatchers(const Folder: OleVariant);
var
  SubFolders: OleVariant;
  Count, i: Integer;
  NextFolder: OleVariant;
begin
  if not IsVariantAssigned(Folder) then
    Exit;

  FWatchers.Add(TFolderWatcher.Create(Self, Folder, FLogger));

  try
    SubFolders := Folder.Folders;
  except
    SubFolders := Unassigned;
  end;

  if not IsVariantAssigned(SubFolders) then
    Exit;

  try
    Count := SubFolders.Count;
  except
    Count := 0;
  end;

  for i := 1 to Count do
  begin
    try
      NextFolder := SubFolders.Item(i);
    except
      NextFolder := Unassigned;
    end;
    if IsVariantAssigned(NextFolder) then
      BuildWatchers(NextFolder);
  end;
end;

procedure TOutlookEventListener.Notify(const Event: TOutlookItemEvent);
begin
  if Assigned(FOnEvent) then
  begin
    try
      FOnEvent(Event);
    except
      // ignore callback errors (worker will handle/log)
    end;
  end;
end;

procedure TOutlookEventListener.Refresh;
begin
  if FWatchers <> nil then
    FWatchers.Clear;

  if IsVariantAssigned(FRootFolder) then
    BuildWatchers(FRootFolder);
end;

end.

