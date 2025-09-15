unit mail.TypesUtils;

interface

uses
  System.SysUtils, System.Classes, System.IOUtils, System.Hash, System.Variants,
  System.DateUtils, System.RegularExpressions, System.Generics.Collections,
  WinAPI.Windows, System.SyncObjs;

type
  // byte-string for FireDAC AsByteStr
  TByteStr = AnsiString;

  TAttMode = (amNone, amMeta, amMetaHash, amBytes);

  TCliOptions = record
    PSTPath: string;       // empty or 'active' -> detect via COM
    ConnOverride: string;
    EnvFile: string;
    LogPath: string;
    AttachMode: TAttMode;
    Batch: Integer;
    QueryMode: Boolean;
    FullSync: Boolean;
    Incremental: Boolean;
    MovesOnly: Boolean;
    UseRestrict: Boolean;
  end;

  // What ultimately controls the behavior
  TOutlookObjectModelGuardMode = (emPolicyControlled, emMachineObjectModelGuard, emNoExplicitSetting);

  TOutlookGuardResult = record
    Mode: TOutlookObjectModelGuardMode;
    // When Mode = emMachineObjectModelGuard:
    ObjectModelGuardValue: Integer;   // 0,1,2 or -1 if not applicable
    ObjectModelGuardMeaning: string;
    ObjectModelGuardPath: string;     // where it was found

    // When Mode = emPolicyControlled (AdminSecurityMode=3):
    AdminSecurityMode: Integer;       // 3 means "Use Outlook Security Group Policy"; -1 if not set
    PromptOOMSend: Integer;           // 0=Auto Deny, 1=Prompt, 2=Auto Approve, -1=not set
    PromptOOMAddressInformationAccess: Integer;  // as above
    PromptOOMAddressBookAccess: Integer;         // as above
    PromptOOMSaveAs: Integer;                    // as above

    // Helper: human-friendly description
    Description: string;
  end;

const
  // Outlook proptags (Unicode string type ...001F unless noted)
  PR_TRANSPORT_HEADERS    = 'http://schemas.microsoft.com/mapi/proptag/0x007D001F';
  PR_INTERNET_MESSAGE_ID  = 'http://schemas.microsoft.com/mapi/proptag/0x1035001F';
  PR_SMTP_ADDRESS         = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001F';
  PR_ATTACH_MIME_TAG      = 'http://schemas.microsoft.com/mapi/proptag/0x370E001F';
  PR_ATTACH_METHOD        = 'http://schemas.microsoft.com/mapi/proptag/0x37050003'; // PT_LONG
  PR_ATTACH_DATA_BIN      = 'http://schemas.microsoft.com/mapi/proptag/0x37010102'; // PT_BINARY

  PR_SEARCH_KEY                = 'http://schemas.microsoft.com/mapi/proptag/0x300B0102';
  PR_SUBJECT_W                 = 'http://schemas.microsoft.com/mapi/proptag/0x0037001F';
  PR_SENDER_NAME_W             = 'http://schemas.microsoft.com/mapi/proptag/0x0C1A001F';
  PR_DISPLAY_TO_W              = 'http://schemas.microsoft.com/mapi/proptag/0x0E04001F';
  PR_DISPLAY_CC_W              = 'http://schemas.microsoft.com/mapi/proptag/0x0E03001F';
  PR_CLIENT_SUBMIT_TIME        = 'http://schemas.microsoft.com/mapi/proptag/0x00390040'; // SentOn
  PR_MESSAGE_DELIVERY_TIME     = 'http://schemas.microsoft.com/mapi/proptag/0x0E060040'; // ReceivedTime
  PR_CREATION_TIME             = 'http://schemas.microsoft.com/mapi/proptag/0x30070040'; // CreationTime
  PR_LAST_MODIFICATION_TIME    = 'http://schemas.microsoft.com/mapi/proptag/0x30080040'; // LastModificationTime
  PR_SENDER_EMAIL_ADDRESS_W    = 'http://schemas.microsoft.com/mapi/proptag/0x0C1F001F'; // optional helper
  PR_MESSAGE_CLASS             = 'http://schemas.microsoft.com/mapi/proptag/0x001A001F';

  // FOLDER
  PR_LOCAL_COMMIT_TIME_MAX     = 'http://schemas.microsoft.com/mapi/proptag/0x670A0040'; //

  // Outlook constants
  OL_STORE_UNICODE = 2;  // AddStoreEx: olStoreUnicode
  OL_CLASS_MAIL    = 43; // OlObjectClass.olMail

  // MAPI attach methods (subset)
  ATTACH_BY_VALUE      = 1; // most common
  ATTACH_EMBEDDED_MSG  = 5;
  ATTACH_OLE           = 6;

  // Shared constants for the whole app
  cEmailRegex            = '^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$';
  cDedupWindowMinutes    = 5;
  cMaxSafeFileNameLen    = 200;
  cAttachTempPrefix      = 'pst_ingest_';

function ParseAttMode(const S: string): TAttMode;

procedure LoadEnvFile(const FileName: string);
function GetEnv(const Key, DefaultValue: string): string;

procedure SanitizeFileNameInPlace(var S: string);
function SanitizeFileName(const S: string): string;

function VarToStrDefSafe(const V: OleVariant; const Def: string = ''): string;
function VarToIntDefSafe(const V: OleVariant; const Def: Integer = 0): Integer;
function IsVariantAssigned(const V: OleVariant): Boolean;
function IsLikelyEmail(const S: string): Boolean;

function ToUTC(const V: OleVariant): TDateTime;

function UTCDateToLocal(const AUTC: TDateTime): TDateTime;

function DateToStrLocalNoSeconds(const DT: TDateTime): string;

// Build JET filter for LastModificationTime > ACutoffDate; empty when date = 0.
// This function expects a local date.
function BuildLastModJETFilter(const ACutoffDate: TDateTime): string;

function ParseHeadersFind(const Headers, FieldName: string): string;
function ParseMessageID(const Headers: string): string;

function ParseAddressList(const S: string): TArray<TPair<string,string>>;

function ReadAllBytesRaw(const FilePath: string): TByteStr;
function VariantBinaryToByteStr(const V: OleVariant): TByteStr;

function RawSHA256Hex(const Bytes: TByteStr): string;
function RawSHA256Bytes(const Bytes: TByteStr): TArray<System.Byte>;

// Updated: hex digest string for Message-ID (lowercased)
function MsgIdHashSql(const Mid: string): string;

// New: SHA256(lower(UTF-16(S))) -> raw 32 bytes (to match SQL NVARCHAR hashing)
function Sha256LowerUtf16ToBytes(const S: string): TBytes;

procedure CollectRecipientsFromCOM(const Recips: OleVariant; var Rows: TArray<TArray<Variant>>);

procedure GetRecipients(const Item: OleVariant; const Headers: string; out RecRows: TArray<TArray<Variant>>);

procedure ReadMessageCore(const Item: OleVariant; out InternetMessageId, OutlookEntryId, Subject, SenderName, SenderEmail,
      DisplayTo, DisplayCc, Headers, BodyText, BodyHtml, MessageClass: String;
      out SentUtc, RecvUtc, CreateUtc, LastModUtc: TDateTime;
      out SearchKey: TBytes);

function GetFolderCommitUtc(const Folder: OleVariant): TDateTime;

function GetLatestOutlookOfficeVersion: string;

function ReadEffectiveOutlookObjectModelGuard(
  const AOfficeVersion: string = ''): TOutlookGuardResult;

// 0,1,2 or -1 if not applicable
function GetOutlookObjectModelGuardValue(const AOfficeVersion: string = ''): Integer;

implementation

uses
  System.Win.Registry;

{ ---- Logging ---- }

function BuildLastModJETFilter(const ACutoffDate: TDateTime): string;
const
  Filter = '[LastModificationTime] >= "%s" ';
begin
  if ACutoffDate = 0 then
    Exit('');
  Result := Format(Filter, [DateToStrLocalNoSeconds(ACutoffDate)]);
end;

procedure CollectRecipientsFromCOM(const Recips: OleVariant; var Rows: TArray<TArray<Variant>>);
var
  n, i: Integer;
  r, PropAccessor, AddrEntry, ExchangeUser: OleVariant;
  Kind: Integer;
  DisplayName, Smtp, AddrType, AddressRaw: string;
  EmailRe: TRegEx;
  Row: TArray<Variant>;
begin
  SetLength(Rows, 0);
  if not IsVariantAssigned(Recips) then Exit;
  try
    n := Recips.Count;
  except
    Exit;
  end;

  EmailRe := TRegEx.Create(cEmailRegex);

  for i := 1 to n do
  begin
    try
      r := Recips.Item(i);
    except
      Continue;
    end;

    Kind := 1;
    try Kind := r.Type; except end;

    try
      DisplayName := VarToStrDef(r.Name, '');
    except
      DisplayName := '';
    end;

    Smtp := '';
    try
      PropAccessor := r.PropertyAccessor;
      Smtp := VarToStrDef(PropAccessor.GetProperty(PR_SMTP_ADDRESS), '');
    except
    end;

    if Smtp = '' then
    begin
      try
        AddrEntry := r.AddressEntry;
        try
          AddrType := VarToStrDef(AddrEntry.Type, '');
        except
          AddrType := '';
        end;
        if SameText(AddrType, 'EX') then
        begin
          try ExchangeUser := AddrEntry.GetExchangeUser; except ExchangeUser := Unassigned; end;
          if IsVariantAssigned(ExchangeUser) then
            try
              Smtp := VarToStrDef(ExchangeUser.PrimarySmtpAddress, '');
            except
              Smtp := '';
            end;
        end
        else
        begin
          try
            AddressRaw := VarToStrDef(AddrEntry.Address, '');
          except
            AddressRaw := '';
          end;
          if EmailRe.IsMatch(AddressRaw) then
            Smtp := AddressRaw;
        end;
      except
      end;
    end;

    if (DisplayName <> '') or (Smtp <> '') then
    begin
      SetLength(Rows, Length(Rows) + 1);
      Row := TArray<Variant>.Create(Kind, DisplayName, Smtp);
      Rows[High(Rows)] := Row;
    end;
  end;
end;

procedure GetRecipients(const Item: OleVariant; const Headers: string;
  out RecRows: TArray<TArray<Variant>>);
var
  ToList, CcList, BccList: TArray<TPair<string,string>>;
  k: Integer;
begin
  SetLength(RecRows, 0);
  try
    CollectRecipientsFromCOM(Item.Recipients, RecRows);
  except
    SetLength(RecRows, 0);
  end;
  if Length(RecRows) = 0 then
  begin
    ToList  := ParseAddressList(ParseHeadersFind(Headers, 'To'));
    CcList  := ParseAddressList(ParseHeadersFind(Headers, 'Cc'));
    BccList := ParseAddressList(ParseHeadersFind(Headers, 'Bcc'));
    for k := 0 to High(ToList) do
      RecRows := RecRows + [TArray<Variant>.Create(1, ToList[k].Key, ToList[k].Value)];
    for k := 0 to High(CcList) do
      RecRows := RecRows + [TArray<Variant>.Create(2, CcList[k].Key, CcList[k].Value)];
    for k := 0 to High(BccList) do
      RecRows := RecRows + [TArray<Variant>.Create(3, BccList[k].Key, BccList[k].Value)];
  end;
end;

procedure ReadMessageCore(const Item: OleVariant;
  out InternetMessageId, OutlookEntryId, Subject, SenderName, SenderEmail,
      DisplayTo, DisplayCc, Headers, BodyText, BodyHtml, MessageClass: String;
  out SentUtc, RecvUtc, CreateUtc, LastModUtc: TDateTime;
  out SearchKey: TBytes);
var
  MsgPropAccessor: OleVariant;
  Names, Values: OleVariant;
  vSk: OleVariant;
  lb, hb, i, vlb: Integer;

  // Local holders fetched via PropertyAccessor (to keep OOM fallbacks intact)
  paSubject, paSenderName, paDisplayTo, paDisplayCc: string;
  paSent, paRecv, paCreate, paLastMod: TDateTime;
  paSenderEmailAddr, paMessageClass: string;
  SenderAddrEntry, SenderExchangeUser: OleVariant;
  SenderAddressType, SenderAddressRaw: string;

  function SafeVarDate(const V: OleVariant): TDateTime;
  begin
    Result := 0;

    if VarIsNull(V) or VarIsEmpty(V) then
      Exit;
    try
      Result := VarToDateTime(V);
    except
      Result := 0;
    end;
  end;

begin
  // Defaults
  Headers := '';
  InternetMessageId := '';
  Subject := '';
  SenderName := '';
  OutlookEntryId := '';
  SenderEmail := '';
  DisplayTo := '';
  DisplayCc := '';
  BodyText := '';
  BodyHtml := '';
  SentUtc := 0; RecvUtc := 0; CreateUtc := 0; LastModUtc := 0;
  SetLength(SearchKey, 0);

  paSubject := ''; paSenderName := ''; paDisplayTo := ''; paDisplayCc := '';
  paSent := 0; paRecv := 0; paCreate := 0; paLastMod := 0;
  paSenderEmailAddr := '';
  paMessageClass:= '';

  // ---- Batch property fetch (Headers, Message-Id, SearchKey, plus other available props) ----
  MsgPropAccessor := Item.PropertyAccessor;
  try
    Names := VarArrayOf([
      PR_TRANSPORT_HEADERS,         // 0
      PR_INTERNET_MESSAGE_ID,       // 1
      PR_SEARCH_KEY,                // 2 (PT_BINARY)
      PR_SUBJECT_W,                 // 3
      PR_SENDER_NAME_W,             // 4
      PR_DISPLAY_TO_W,              // 5
      PR_DISPLAY_CC_W,              // 6
      PR_CLIENT_SUBMIT_TIME,        // 7 (date)
      PR_MESSAGE_DELIVERY_TIME,     // 8 (date)
      PR_CREATION_TIME,             // 9 (date)
      PR_LAST_MODIFICATION_TIME,    // 10 (date)
      PR_SENDER_EMAIL_ADDRESS_W,    // 11 (optional)
      PR_MESSAGE_CLASS              // 12 Message Class
    ]);
    Values := MsgPropAccessor.GetProperties(Names);

    if VarIsArray(Values) and (VarArrayDimCount(Values) = 1) then
    begin
      vlb := VarArrayLowBound(Values, 1);

      Headers           := VarToStrDefSafe(Values[vlb + 0], '');
      InternetMessageId := VarToStrDefSafe(Values[vlb + 1], '');

      // PR_SEARCH_KEY (PT_BINARY) -> TBytes
      vSk := Values[vlb + 2];
      if VarIsArray(vSk) and (VarArrayDimCount(vSk) = 1) then
      begin
        lb := VarArrayLowBound(vSk, 1);
        hb := VarArrayHighBound(vSk, 1);
        SetLength(SearchKey, hb - lb + 1);
        for i := 0 to Length(SearchKey) - 1 do
          SearchKey[i] := Byte(vSk[lb + i]);
      end
      else
        SetLength(SearchKey, 0);

      // Strings
      paSubject      := VarToStrDefSafe(Values[vlb + 3], '');
      paSenderName   := VarToStrDefSafe(Values[vlb + 4], '');
      paDisplayTo    := VarToStrDefSafe(Values[vlb + 5], '');
      paDisplayCc    := VarToStrDefSafe(Values[vlb + 6], '');

      // Date/time → UTC
      paSent    := SafeVarDate(Values[vlb + 7]);
      paRecv    := SafeVarDate(Values[vlb + 8]);
      paCreate  := SafeVarDate(Values[vlb + 9]);
      paLastMod := SafeVarDate(Values[vlb + 10]);

      // Optional sender email (may be empty / unreliable in some profiles)
      paSenderEmailAddr := VarToStrDefSafe(Values[vlb + 11], '');
      paMessageClass := VarToStrDefSafe(Values[vlb + 12], '');
    end
    else
    begin
      // Conservative fallback to single GetProperty calls for the essential trio
      Headers           := VarToStrDefSafe(MsgPropAccessor.GetProperty(PR_TRANSPORT_HEADERS), '');
      InternetMessageId := VarToStrDefSafe(MsgPropAccessor.GetProperty(PR_INTERNET_MESSAGE_ID), '');
      try
        vSk := MsgPropAccessor.GetProperty(PR_SEARCH_KEY);
        if VarIsArray(vSk) and (VarArrayDimCount(vSk) = 1) then
        begin
          lb := VarArrayLowBound(vSk, 1);
          hb := VarArrayHighBound(vSk, 1);
          SetLength(SearchKey, hb - lb + 1);
          for i := 0 to Length(SearchKey) - 1 do
            SearchKey[i] := Byte(vSk[lb + i]);
        end
        else
          SetLength(SearchKey, 0);
      except
        SetLength(SearchKey, 0);
      end;
      // Others (subject/sender/to/cc/dates) will fall back to OOM getters below
    end;
  except
    // Keep prior behavior: propagate unexpected COM errors (callers handle/log)
    raise;
  end;

  // ---- OOM fields & fallbacks (bodies excluded from batching by design) ----

  // Subject / SenderName / Display lists: prefer PA values; fall back to OOM if missing
  if paSubject <> ''    then Subject    := paSubject
  else begin try Subject    := VarToStrDef(Item.Subject,    ''); except Subject    := ''; end; end;

  if paSenderName <> '' then SenderName := paSenderName
  else begin try SenderName := VarToStrDef(Item.SenderName, ''); except SenderName := ''; end; end;

  if paDisplayTo <> ''  then DisplayTo  := paDisplayTo
  else begin try DisplayTo  := VarToStrDef(Item.&To,        ''); except DisplayTo  := ''; end; end;

  if paDisplayCc <> ''  then DisplayCc  := paDisplayCc
  else begin try DisplayCc  := VarToStrDef(Item.CC,         ''); except DisplayCc  := ''; end; end;

  if paMessageClass <> '' then MessageClass := paMessageClass
  else begin try MessageClass := VarToStrDef(Item.MessageClass, ''); except MessageClass:= ''; end; end;

  // EntryID: keep classic OOM semantics (string) — PR_ENTRYID is binary and not needed here
  try OutlookEntryId := VarToStrDef(Item.EntryID, ''); except OutlookEntryId := ''; end;
  // Bodies: explicitly use OOM (per requirement: "use ReadProperties for all available properties, but body")
  try BodyText := VarToStrDef(Item.Body, ''); except BodyText := ''; end;
  try BodyHtml := VarToStrDef(Item.HTMLBody, ''); except BodyHtml := ''; end;

  // Timestamps: prefer PA (already normalized), fall back to OOM → UTC
  if paSent <> 0    then SentUtc    := paSent    else begin try SentUtc    := ToUTC(Item.SentOn);            except SentUtc    := 0; end; end;
  if paRecv <> 0    then RecvUtc    := paRecv    else begin try RecvUtc    := ToUTC(Item.ReceivedTime);      except RecvUtc    := 0; end; end;
  if paCreate <> 0  then CreateUtc  := paCreate  else begin try CreateUtc  := ToUTC(Item.CreationTime);      except CreateUtc  := 0; end; end;
  if paLastMod <> 0 then LastModUtc := paLastMod else begin try LastModUtc := ToUTC(Item.LastModificationTime); except LastModUtc := 0; end; end;

  // Sender SMTP resolution (preserve robust logic; optionally seed with PA value if everything else fails)
  try
    if SameText(VarToStrDefSafe(Item.SenderEmailType, ''), 'SMTP') then
      try
        SenderEmail := VarToStrDef(Item.SenderEmailAddress, '');
      except
        SenderEmail := '';
      end
    else
    begin
      SenderEmail := '';
      try
        // Resolve through AddressEntry/Exchange path
        SenderAddrEntry := Item.Sender;
        try
          SenderAddressType := VarToStrDef(SenderAddrEntry.Type, '');
        except
          SenderAddressType := '';
        end;
        if SameText(SenderAddressType, 'EX') then
        begin
          try
            SenderExchangeUser := SenderAddrEntry.GetExchangeUser;
            if IsVariantAssigned(SenderExchangeUser) then
              try
                SenderEmail := VarToStrDef(SenderExchangeUser.PrimarySmtpAddress, '');
              except
                SenderEmail := '';
              end;
          except
            SenderEmail := '';
          end;
        end
        else
        begin
          try
            SenderAddressRaw := VarToStrDef(SenderAddrEntry.Address, '');
          except
            SenderAddressRaw := '';
          end;
          if IsLikelyEmail(SenderAddressRaw) then
            SenderEmail := SenderAddressRaw;
        end;
      except
        SenderEmail := '';
      end;

      if SenderEmail = '' then
        try SenderEmail := VarToStrDef(Item.SenderEmailAddress, ''); except SenderEmail := ''; end;

      if SenderEmail = '' then
      begin
        // Try well-known SMTP property (may or may not be present at item level)
        try
          SenderEmail := VarToStrDef(Item.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS), '');
        except
          SenderEmail := '';
        end;
      end;

      // Last resort: use PA sender email address if we got one from the batch call
      if (SenderEmail = '') and (paSenderEmailAddr <> '') then
        SenderEmail := paSenderEmailAddr;
    end;
  except
    try
      SenderEmail := VarToStrDef(Item.SenderEmailAddress, '');
    except
      SenderEmail := '';
    end;
  end;

  // If Message-Id was blank, try parsing from headers (unchanged behavior)
  if InternetMessageId = '' then
    InternetMessageId := ParseMessageID(Headers);
end;


procedure ReadMessageCoreOld(const Item: OleVariant;
  out InternetMessageId, OutlookEntryId, Subject, SenderName, SenderEmail,
      DisplayTo, DisplayCc, Headers, BodyText, BodyHtml: string;
  out SentUtc, RecvUtc, CreateUtc, LastModUtc: TDateTime;
  out SearchKey: TBytes);
var
  MsgPropAccessor, SenderAddrEntry, SenderExchangeUser: OleVariant;
  SenderAddressType, SenderAddressRaw: string;
  vSk: OleVariant;
  lb, hb, i: Integer;
begin
  Headers := '';
  InternetMessageId := '';
  SetLength(SearchKey, 0);
  try
    MsgPropAccessor := Item.PropertyAccessor;

    // Transport headers and Message-Id (Unicode)
    Headers           := VarToStrDefSafe(MsgPropAccessor.GetProperty(PR_TRANSPORT_HEADERS), '');
    InternetMessageId := VarToStrDefSafe(MsgPropAccessor.GetProperty(PR_INTERNET_MESSAGE_ID), '');

    // PR_SEARCH_KEY (PT_BINARY) -> TBytes, using the same SAFEARRAY approach as mail.FullSync
    try
      vSk := MsgPropAccessor.GetProperty(PR_SEARCH_KEY);
      if VarIsArray(vSk) and (VarArrayDimCount(vSk) = 1) then
      begin
        lb := VarArrayLowBound(vSk, 1);
        hb := VarArrayHighBound(vSk, 1);
        SetLength(SearchKey, hb - lb + 1);
        for i := 0 to Length(SearchKey) - 1 do
          SearchKey[i] := Byte(vSk[lb + i]);
      end
      else
        SetLength(SearchKey, 0);
    except
      SetLength(SearchKey, 0);
    end;
  except
    // leave defaults for headers, message-id, and search key
  end;

  // Standard OOM fields (guard each COM access)
  try Subject        := VarToStrDef(Item.Subject,        ''); except Subject        := ''; end;
  try SenderName     := VarToStrDef(Item.SenderName,     ''); except SenderName     := ''; end;
  try OutlookEntryId := VarToStrDef(Item.EntryID,        ''); except OutlookEntryId := ''; end;

  // Sender SMTP resolution with robust fallbacks
  try
    if SameText(VarToStrDefSafe(Item.SenderEmailType, ''), 'SMTP') then
      try
        SenderEmail := VarToStrDef(Item.SenderEmailAddress, '');
      except
        SenderEmail := '';
      end
    else
    begin
      SenderEmail := '';
      try
        SenderAddrEntry := Item.Sender;
        try
          SenderAddressType := VarToStrDef(SenderAddrEntry.Type, '');
        except
          SenderAddressType := '';
        end;
        if SameText(SenderAddressType, 'EX') then
        begin
          try
            SenderExchangeUser := SenderAddrEntry.GetExchangeUser;
            if IsVariantAssigned(SenderExchangeUser) then
              try
                SenderEmail := VarToStrDef(SenderExchangeUser.PrimarySmtpAddress, '');
              except
                SenderEmail := '';
              end;
          except
            SenderEmail := '';
          end;
        end
        else
        begin
          try
            SenderAddressRaw := VarToStrDef(SenderAddrEntry.Address, '');
          except
            SenderAddressRaw := '';
          end;
          if IsLikelyEmail(SenderAddressRaw) then
            SenderEmail := SenderAddressRaw;
        end;
      except
        SenderEmail := '';
      end;

      if SenderEmail = '' then
        try SenderEmail := VarToStrDef(Item.SenderEmailAddress, ''); except SenderEmail := ''; end;
      if SenderEmail = '' then
      begin
        try
          SenderEmail := VarToStrDef(Item.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS), '');
        except
          SenderEmail := '';
        end;
      end;
    end;
  except
    try
      SenderEmail := VarToStrDef(Item.SenderEmailAddress, '');
    except
      SenderEmail := '';
    end;
  end;

  try DisplayTo := VarToStrDef(Item.&To,     ''); except DisplayTo := ''; end;
  try DisplayCc := VarToStrDef(Item.CC,      ''); except DisplayCc := ''; end;
  try BodyText  := VarToStrDef(Item.Body,    ''); except BodyText  := ''; end;
  try BodyHtml  := VarToStrDef(Item.HTMLBody,''); except BodyHtml  := ''; end;

  // Normalize timestamps to UTC (existing behavior) with guards
  try SentUtc    := ToUTC(Item.SentOn);               except SentUtc    := 0; end;
  try RecvUtc    := ToUTC(Item.ReceivedTime);         except RecvUtc    := 0; end;
  try CreateUtc  := ToUTC(Item.CreationTime);         except CreateUtc  := 0; end;
  try LastModUtc := ToUTC(Item.LastModificationTime); except LastModUtc := 0; end;

  // If Message-Id property was blank, try parsing from headers (existing behavior)
  if InternetMessageId = '' then
    InternetMessageId := ParseMessageID(Headers);
end;

function DateToStrLocalNoSeconds(const DT: TDateTime): string;
  function StripSecondsFromTimeFormat(const Fmt: string): string;
  var
    i: Integer;
    inQuote: Boolean;
    outS: string;

    function IsSecondToken(ch: Char): Boolean;
    begin
      Result := (ch = 's') or (ch = 'S');
    end;

    procedure RemoveTrailingSep;
    begin
      if (outS <> '') and (outS[Length(outS)] in [':', '.', '-', '/', ' ']) then
        Delete(outS, Length(outS), 1);
    end;

  begin
    outS := '';
    inQuote := False;
    i := 1;
    while i <= Length(Fmt) do
    begin
      if Fmt[i] = '''' then
      begin
        inQuote := not inQuote;
        outS := outS + Fmt[i];
        Inc(i);
        Continue;
      end;

      if (not inQuote) and IsSecondToken(Fmt[i]) then
      begin
        // Skip the whole run of s/S
        while (i <= Length(Fmt)) and IsSecondToken(Fmt[i]) do
          Inc(i);
        // Remove the separator that would have preceded seconds
        RemoveTrailingSep;
        Continue;
      end;

      outS := outS + Fmt[i];
      Inc(i);
    end;

    // Trim any trailing separators left by removing seconds
    while (outS <> '') and (outS[Length(outS)] in [':', '.', '-', '/', ' ']) do
      Delete(outS, Length(outS), 1);

    Result := outS;
  end;

var
  fs: TFormatSettings;
  timeFmt: string;
begin
  fs := TFormatSettings.Create;                 // current computer locale
  timeFmt := StripSecondsFromTimeFormat(fs.ShortTimeFormat);

  Result :=
    FormatDateTime(fs.ShortDateFormat, DT, fs) + ' ' +
    FormatDateTime(timeFmt, DT, fs);
end;

{ ---- .env ---- }

procedure LoadEnvFile(const FileName: string);
var
  Lines: TStringList;
  i, p: Integer;
  L, Key, Val: string;
begin
  if not TFile.Exists(FileName) then Exit;
  Lines := TStringList.Create;
  try
    Lines.LoadFromFile(FileName, TEncoding.UTF8);
    for i := 0 to Lines.Count - 1 do
    begin
      L := Trim(Lines[i]);
      if (L = '') or (L.StartsWith('#')) then
        Continue;
      p := L.IndexOf('=');
      if p <= 0 then
        Continue;
      Key := Trim(Copy(L, 1, p));
      Val := Trim(Copy(L, p + 2, MaxInt));
      if (Length(Val) >= 2) and
         (((Val.StartsWith('"')) and (Val.EndsWith('"'))) or ((Val.StartsWith('''')) and (Val.EndsWith('''')))) then
        Val := Val.Substring(1, Val.Length - 2);
      SetEnvironmentVariable(PChar(Key), PChar(Val));
    end;
  finally
    Lines.Free;
  end;
end;

function GetEnv(const Key, DefaultValue: string): string;
var
  buf: array[0..32767] of Char;
  n: DWORD;
begin
  n := GetEnvironmentVariable(PChar(Key), @buf[0], Length(buf));
  if n = 0 then
    Exit(DefaultValue);
  SetString(Result, buf, n);
end;

{ ---- Misc ---- }

function ParseAttMode(const S: string): TAttMode;
begin
  if SameText(S, 'meta') then Exit(amMeta);
  if SameText(S, 'meta-hash') then Exit(amMetaHash);
  if SameText(S, 'bytes') then Exit(amBytes);
  Result := amNone;
end;

procedure SanitizeFileNameInPlace(var S: string);
const
  BadChars: array[0..8] of Char = ('<','>',':','"','/','\','|','?','*');
var
  c: Char;
  Base, Ext, BaseUpper: string;

  function IsReserved(const U: string): Boolean;
  var
    n: Integer;
  begin
    Result := (U = 'CON') or (U = 'PRN') or (U = 'AUX') or (U = 'NUL');
    if Result then Exit;
    for n := 1 to 9 do
      if (U = 'COM' + IntToStr(n)) or (U = 'LPT' + IntToStr(n)) then
        Exit(True);
  end;

begin
  for c in BadChars do
    S := S.Replace(c, '_');
  S := S.Trim;
  if S = '' then S := 'attachment';
  Base := TPath.GetFileNameWithoutExtension(S);
  Ext  := TPath.GetExtension(S);
  BaseUpper := UpperCase(Base);
  if IsReserved(BaseUpper) then
    Base := '_' + Base;
  S := Base + Ext;
  if S.Length > cMaxSafeFileNameLen then
    S := Copy(S, 1, cMaxSafeFileNameLen);
end;

function SanitizeFileName(const S: string): string;
begin
  Result := S;
  SanitizeFileNameInPlace(Result);
end;

function VarToStrDefSafe(const V: OleVariant; const Def: string): string;
begin
  try
    if VarIsNull(V) or VarIsEmpty(V) then
      Exit(Def);
    Result := VarToStr(V);
  except
    Result := Def;
  end;
end;

function VarToIntDefSafe(const V: OleVariant; const Def: Integer): Integer;
var
  S: string;
begin
  try
    if VarIsNull(V) or VarIsEmpty(V) then
      Exit(Def);
    S := VarToStr(V);
    if not TryStrToInt(S, Result) then
      Result := Def;
  except
    Result := Def;
  end;
end;

function IsVariantAssigned(const V: OleVariant): Boolean;
begin
  Result := (not VarIsNull(V)) and (not VarIsEmpty(V));
end;

function IsLikelyEmail(const S: string): Boolean;
begin
  Result := TRegEx.IsMatch(S, cEmailRegex);
end;

function ToUTC(const V: OleVariant): TDateTime;
var
  dt: TDateTime;
begin
  Result := 0; // sentinel "unknown"
  if VarIsNull(V) or VarIsEmpty(V) then Exit;
  try
    dt := VarToDateTime(V);
  except
    Exit;
  end;
  Result := TTimeZone.Local.ToUniversalTime(dt);
end;

function UTCDateToLocal(const AUTC: TDateTime): TDateTime;
begin
  Result := TTimeZone.Local.ToLocalTime(AUTC);  // -> local time for this computer
end;

{ ---- Header parsing ---- }

function ParseHeadersFind(const Headers, FieldName: string): string;
var
  RE: TRegEx;
  M: TMatch;
  Lines: TArray<string>;
  i: Integer;
  Acc: string;
  H: string;
begin
  Result := '';
  if Headers = '' then Exit;
  H := Headers.Replace(#13#10, #10);
  Lines := H.Split([#10]);
  Acc := '';
  for i := 0 to High(Lines) do
  begin
    if (i = 0) or not ((Length(Lines[i]) > 0) and (Lines[i][1] in [' ', #9])) then
      Acc := Acc + #10 + Lines[i]
    else
      Acc := Acc + ' ' + TrimLeft(Lines[i]);
  end;
  RE := TRegEx.Create('(?im)^\s*' + TRegEx.Escape(FieldName) + '\s*:\s*(.+)$');
  M := RE.Match(Acc);
  if M.Success then
    Result := Trim(M.Groups[1].Value);
end;

function ParseMessageID(const Headers: string): string;
begin
  Result := ParseHeadersFind(Headers, 'Message-ID');
  if Result = '' then
    Result := ParseHeadersFind(Headers, 'Message-Id');
  Result := Result.Trim(['<','>',' ','"','''']);
end;

function ParseAddressList(const S: string): TArray<TPair<string,string>>;
var
  Parts: TArray<string>;
  i: Integer;
  P, NamePart, EmailPart: string;
  M: TMatch;
  RE: TRegEx;
  Pair: TPair<string,string>;
begin
  SetLength(Result, 0);
  if S = '' then Exit;
  Parts := TRegEx.Split(S, '\s*[;,]\s*');
  RE := TRegEx.Create('([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})');
  for i := 0 to High(Parts) do
  begin
    P := Trim(Parts[i]);
    if P = '' then Continue;
    M := RE.Match(P);
    if M.Success then
    begin
      EmailPart := LowerCase(M.Groups[1].Value);
      NamePart := Trim(StringReplace(P, M.Groups[1].Value, '', []));
      NamePart := NamePart.Trim(['<','>','"','''',' ']);
    end
    else
    begin
      EmailPart := '';
      NamePart := P;
    end;
    Pair := TPair<string,string>.Create(NamePart, EmailPart);
    Result := Result + [Pair];
  end;
end;

{ ---- Binary utilities ---- }

function ReadAllBytesRaw(const FilePath: string): TByteStr;
var
  fs: TFileStream;
  L: Integer;
begin
  Result := '';
  fs := TFileStream.Create(FilePath, fmOpenRead or fmShareDenyWrite);
  try
    L := fs.Size;
    SetLength(Result, L);
    if L > 0 then
      fs.ReadBuffer(PAnsiChar(Result)^, L);
  finally
    fs.Free;
  end;
end;

function VariantBinaryToByteStr(const V: OleVariant): TByteStr;
var
  L, i, lb, hb: Integer;
begin
  Result := '';
  if VarIsArray(V) and (VarArrayDimCount(V) = 1) then
  begin
    lb := VarArrayLowBound(V, 1);
    hb := VarArrayHighBound(V, 1);
    L  := hb - lb + 1;
    SetLength(Result, L);
    if L > 0 then
      for i := 0 to L - 1 do
        PByte(PAnsiChar(Result) + i)^ := Byte(V[lb + i]);
  end;
end;

function RawSHA256Hex(const Bytes: TByteStr): string;
begin
  Result := THashSHA2.GetHashString(Bytes, THashSHA2.TSHA2Version.SHA256);
end;

function RawSHA256Bytes(const Bytes: TByteStr): TArray<System.Byte>;
begin
  Result := THashSHA2.GetHashBytes(Bytes, THashSHA2.TSHA2Version.SHA256);
end;

function MsgIdHashSql(const Mid: string): string;
var
  S: string;
begin
  if Mid = '' then
  begin
    SetLength(Result, 0);
    Exit;
  end;
  S := LowerCase(Mid);
  Result := THashSHA2.GetHashString(S, THashSHA2.TSHA2Version.SHA256);
end;

// Compute SHA256 over LOWER(UTF-16LE(S)) and return raw 32 bytes
function Sha256LowerUtf16ToBytes(const S: string): TBytes;
var
  LowerS: string;
begin
  Result := THashSHA2.GetHashBytes(LowerCase(S), THashSHA2.TSHA2Version.SHA256);
end;

function GetFolderCommitUtc(const Folder: OleVariant): TDateTime;
var
  pa, v: OleVariant;
begin
  Result := 0;

  pa := Folder.PropertyAccessor;
  v  := pa.GetProperty(PR_LOCAL_COMMIT_TIME_MAX);
  // PR_* 0040 props are MAPI SYSTIME (UTC). Treat as UTC to match the rest of your pipeline.
  Result := VarToDateTime(v);
end;

function DwordMeaningOMG(const V: Integer): string;
begin
  case V of
    0: Result := 'Warn only if antivirus is inactive/out-of-date';
    1: Result := 'Always warn';
    2: Result := 'Never warn (not recommended)';
  else
    Result := 'Unknown';
  end;
end;

function PromptMeaning(const V: Integer): string;
begin
  case V of
    0: Result := 'Automatically deny';
    1: Result := 'Prompt user';
    2: Result := 'Automatically approve';
  else
    Result := '<not set>';
  end;
end;

function TryReadHKLM_DWORD(const SubKey, ValueName: string; out Value: Cardinal; out AccessUsed: Cardinal): Boolean;
var
  Reg: TRegistry;
begin
  Result := False;

  // 1) Try the 64-bit view
  Reg := TRegistry.Create(KEY_READ or KEY_WOW64_64KEY);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKeyReadOnly(SubKey) and Reg.ValueExists(ValueName) then
    begin
      Value := Reg.ReadInteger(ValueName);
      AccessUsed := KEY_WOW64_64KEY;
      Exit(True);
    end;
  finally
    Reg.Free;
  end;

  // 2) Try the 32-bit view
  Reg := TRegistry.Create(KEY_READ or KEY_WOW64_32KEY);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKeyReadOnly(SubKey) and Reg.ValueExists(ValueName) then
    begin
      Value := Reg.ReadInteger(ValueName);
      AccessUsed := KEY_WOW64_32KEY;
      Exit(True);
    end;
  finally
    Reg.Free;
  end;
end;

function TryReadHKCU_DWORD(const SubKey, ValueName: string; out Value: Cardinal): Boolean;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    Result := Reg.OpenKeyReadOnly(SubKey) and Reg.ValueExists(ValueName);
    if Result then
      Value := Reg.ReadInteger(ValueName);
  finally
    Reg.Free;
  end;
end;

function ReadEffectiveOutlookObjectModelGuard(
  const AOfficeVersion: string): TOutlookGuardResult;
var
  // policy
  polBase: string;
  v: Cardinal;
  found: Boolean;

  // machine ObjectModelGuard candidates (normal + ClickToRun virtualized)
  paths: TArray<string>;
  i: Integer;
  accessUsed: Cardinal;
  pathHit: string;
  OfficeVersion: String;
begin
  if AOfficeVersion='' then
    OfficeVersion:= GetLatestOutlookOfficeVersion
  else
    OfficeVersion:= AOfficeVersion;
  // defaults
  Result.Mode := emNoExplicitSetting;
  Result.ObjectModelGuardValue := -1;
  Result.ObjectModelGuardMeaning := '';
  Result.ObjectModelGuardPath := '';
  Result.AdminSecurityMode := -1;
  Result.PromptOOMSend := -1;
  Result.PromptOOMAddressInformationAccess := -1;
  Result.PromptOOMAddressBookAccess := -1;
  Result.PromptOOMSaveAs := -1;
  Result.Description := 'No explicit policy or ObjectModelGuard value found';

  // 1) Check policy: HKCU\Software\Policies\Microsoft\Office\{ver}\Outlook\Security
  polBase := Format('Software\Policies\Microsoft\Office\%s\Outlook\Security', [OfficeVersion]);
  if TryReadHKCU_DWORD(polBase, 'AdminSecurityMode', v) and (v = 3) then
  begin
    Result.Mode := emPolicyControlled;
    Result.AdminSecurityMode := v;

    if TryReadHKCU_DWORD(polBase, 'PromptOOMSend', v) then
      Result.PromptOOMSend := Integer(v);
    if TryReadHKCU_DWORD(polBase, 'PromptOOMAddressInformationAccess', v) then
      Result.PromptOOMAddressInformationAccess := Integer(v);
    if TryReadHKCU_DWORD(polBase, 'PromptOOMAddressBookAccess', v) then
      Result.PromptOOMAddressBookAccess := Integer(v);
    if TryReadHKCU_DWORD(polBase, 'PromptOOMSaveAs', v) then
      Result.PromptOOMSaveAs := Integer(v);

    Result.Description :=
      Format('Policy-controlled (AdminSecurityMode=3): Send=%s; AddressInfo=%s; AddressBook=%s; SaveAs=%s',
        [ PromptMeaning(Result.PromptOOMSend),
          PromptMeaning(Result.PromptOOMAddressInformationAccess),
          PromptMeaning(Result.PromptOOMAddressBookAccess),
          PromptMeaning(Result.PromptOOMSaveAs) ]);
    Exit; // policy takes precedence; ObjectModelGuard is ignored
  end;

  // 2) No policy -> try machine ObjectModelGuard
  paths := TArray<string>.Create(
    // native Office hive (x64/x86)
    Format('SOFTWARE\Microsoft\Office\%s\Outlook\Security', [OfficeVersion]),
    Format('SOFTWARE\Wow6432Node\Microsoft\Office\%s\Outlook\Security', [OfficeVersion]),
    // ClickToRun virtualized machine hives (common on RDS / C2R)
    Format('SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\%s\Outlook\Security', [OfficeVersion]),
    Format('SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\%s\Outlook\Security', [OfficeVersion])
  );

  found := False;
  for i := 0 to High(paths) do
  begin
    if TryReadHKLM_DWORD(paths[i], 'ObjectModelGuard', v, accessUsed) then
    begin
      found := True;
      pathHit := paths[i];
      Break;
    end;
  end;

  if found then
  begin
    Result.Mode := emMachineObjectModelGuard;
    Result.ObjectModelGuardValue := Integer(v);
    Result.ObjectModelGuardMeaning := DwordMeaningOMG(Result.ObjectModelGuardValue);
    Result.ObjectModelGuardPath := 'HKLM:\' + pathHit;
    Result.Description := Format('Machine ObjectModelGuard: %d (%s) at %s',
      [Result.ObjectModelGuardValue, Result.ObjectModelGuardMeaning, Result.ObjectModelGuardPath]);
  end
  else
  begin
    // No policy and no OMG -> default/AV-driven on client OS; on Server OS Outlook typically prompts.
    Result.Description :=
      'No policy and no ObjectModelGuard found. On Server OS, Outlook will usually prompt unless you set policy or ObjectModelGuard.';
  end;
end;

function GetOutlookObjectModelGuardValue(const AOfficeVersion: string = ''): Integer;
var
  OfficeVersion: String;
begin

  if AOfficeVersion='' then
    OfficeVersion:= GetLatestOutlookOfficeVersion
  else
    OfficeVersion:= AOfficeVersion;
  Result:= ReadEffectiveOutlookObjectModelGuard(OfficeVersion).ObjectModelGuardValue;
end;

function GetLatestOutlookOfficeVersion: string;

  function TryReadHKLMString(const SubKey, ValueName: string; Access: Cardinal; out Value: string): Boolean;
  var
    Reg: TRegistry;
  begin
    Result := False;
    Value := '';
    Reg := TRegistry.Create(KEY_READ or Access);
    try
      Reg.RootKey := HKEY_LOCAL_MACHINE;
      if Reg.OpenKeyReadOnly(SubKey) and Reg.ValueExists(ValueName) then
      begin
        Value := Reg.ReadString(ValueName);
        Result := Value <> '';
      end;
    finally
      Reg.Free;
    end;
  end;

  function ParseMajorMinor(const S: string; out Major, Minor: Integer): Boolean;
  var
    p: Integer;
    a, b: string;
  begin
    Result := False;
    Major := 0; Minor := 0;
    p := S.IndexOf('.');
    if p > 0 then
    begin
      a := S.Substring(0, p);
      b := S.Substring(p + 1);
      // keep only leading digits of the minor
      p := 1;
      while (p <= b.Length) and CharInSet(b[p], ['0'..'9']) do Inc(p);
      b := b.Substring(0, p-1);
      Result := TryStrToInt(a, Major) and TryStrToInt(b, Minor);
    end;
  end;

  function OutlookKeyExistsForVersion(const Version: string; Access: Cardinal): Boolean;
  var
    Reg: TRegistry;
    function KeyExists(const K: string): Boolean;
    begin
      Result := Reg.OpenKeyReadOnly(K);
      if Result then Reg.CloseKey;
    end;
  begin
    Result := False;
    Reg := TRegistry.Create(KEY_READ or Access);
    try
      Reg.RootKey := HKEY_LOCAL_MACHINE;

      // Native Office hives
      if KeyExists(Format('SOFTWARE\Microsoft\Office\%s\Outlook', [Version])) then Exit(True);
      if KeyExists(Format('SOFTWARE\Wow6432Node\Microsoft\Office\%s\Outlook', [Version])) then Exit(True);

      // Click-to-Run virtualized machine hives (commonly used on RDS)
      if KeyExists(Format('SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\%s\Outlook', [Version])) then Exit(True);
      if KeyExists(Format('SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\%s\Outlook', [Version])) then Exit(True);
    finally
      Reg.Free;
    end;
  end;

  procedure ConsiderVersion(const Version: string; var BestMajor, BestMinor: Integer);
  var
    Mj, Mn: Integer;
  begin
    if ParseMajorMinor(Version, Mj, Mn) then
      if (Mj > BestMajor) or ((Mj = BestMajor) and (Mn > BestMinor)) then
      begin
        // accept only if Outlook subkey exists somewhere
        if OutlookKeyExistsForVersion(Version, KEY_WOW64_64KEY) or
           OutlookKeyExistsForVersion(Version, KEY_WOW64_32KEY) then
        begin
          BestMajor := Mj;
          BestMinor := Mn;
        end;
      end;
  end;

  procedure EnumerateOfficeVersions(const BaseKey: string; Access: Cardinal; var BestMajor, BestMinor: Integer);
  var
    Reg: TRegistry;
    Names: TStrings;
    i: Integer;
  begin
    Names:= TStrings.Create;
    try
      Reg := TRegistry.Create(KEY_READ or Access);
      try
        Reg.RootKey := HKEY_LOCAL_MACHINE;
        if Reg.OpenKeyReadOnly(BaseKey) then
        begin
          Reg.GetKeyNames(Names);
          Reg.CloseKey;
          for i := 0 to Names.Count-1 do
            // Only consider keys that look like "N.N"
            if (Pos('.', Names[i]) > 0) then
              ConsiderVersion(Names[i], BestMajor, BestMinor);
        end;
      finally
        Reg.Free;
      end;
    finally
      Names.Free;
    end;
  end;

var
  s: string;
  mj, mn: Integer;

begin
  Result := '';
  mj := -1; mn := -1;

  // 1) Prefer Click-to-Run’s reported version (most reliable on modern Office)
  if not TryReadHKLMString('SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'VersionToReport', KEY_WOW64_64KEY, s) then
    TryReadHKLMString('SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'ClientVersionToReport', KEY_WOW64_64KEY, s);

  if (s = '') then
  begin
    // try 32-bit view as well
    if not TryReadHKLMString('SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'VersionToReport', KEY_WOW64_32KEY, s) then
      TryReadHKLMString('SOFTWARE\Microsoft\Office\ClickToRun\Configuration', 'ClientVersionToReport', KEY_WOW64_32KEY, s);
  end;

  if (s <> '') and ParseMajorMinor(s, mj, mn) then
  begin
    // ensure Outlook actually exists for that major.minor
    if OutlookKeyExistsForVersion(Format('%d.%d', [mj, mn]), KEY_WOW64_64KEY) or
       OutlookKeyExistsForVersion(Format('%d.%d', [mj, mn]), KEY_WOW64_32KEY) then
    begin
      Exit(Format('%d.%d', [mj, mn]));
    end;
    // otherwise fall through to enumeration
    mj := -1; mn := -1;
  end;

  // 2) Enumerate Office\* version keys and pick the highest that has an Outlook subkey
  EnumerateOfficeVersions('SOFTWARE\Microsoft\Office', KEY_WOW64_64KEY, mj, mn);
  EnumerateOfficeVersions('SOFTWARE\Microsoft\Office', KEY_WOW64_32KEY, mj, mn);

  if (mj >= 0) and (mn >= 0) then
    Result := Format('%d.%d', [mj, mn])
  else
    Result := ''; // not found
end;


end.

