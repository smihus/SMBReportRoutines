unit SMBExcel;

interface
uses
   Excel_TLB, System.Generics.Collections;

const
  ExcelApp = 'Excel.Application';

type
  TSMBExcel = class
  private
    FExcel: ExcelApplication;
    FLCID: Integer;
    FWorkbook: ExcelWorkbook;
    function GetWorksheet(WSName: String): ExcelWorksheet;
    function GetWorkwheets: TList<ExcelWorksheet>;
    function GetActiveWorksheet: ExcelWorksheet;
    procedure SetActiveWorksheet(const Value: ExcelWorksheet);
    function GetField(const vWorksheet: String; const Text: string): ExcelRange;
    function GetRange(const vWorksheet, Name: String): ExcelRange;
  public
    constructor Create(const Visible: Boolean = False); overload;
    constructor Create(FileName: String; const Visible: Boolean = False); overload;
    destructor Destroy; override;
    class function OpenWorkbook(const ExcelApp: ExcelApplication; FileName: String; const ALCID: Integer = 0): ExcelWorkbook;
    class function CreateExcelObject(const Visible: Boolean = False; const ALCID: Integer = 0): ExcelApplication;
    class function FreeExcelObject(var ExcelApp: ExcelApplication; const ALCID: Integer = 0): Boolean;
    class function GetLCID: Integer;
    class function CheckExcelInstall: Boolean;
    class procedure Show(var ExcelApp: ExcelApplication; const ALCID: Integer = 0); overload;
    class procedure Hide(var ExcelApp: ExcelApplication; const ALCID: Integer = 0); overload;
    procedure Show; overload;
    procedure Hide; overload;
    procedure Copy(SourceRange: ExcelRange; DestinationRange: ExcelRange);
    property Worksheet[WSName: String]: ExcelWorksheet read GetWorksheet;
    property Worksheets: TList<ExcelWorksheet> read GetWorkwheets;
    property ActiveWorksheet: ExcelWorksheet read GetActiveWorksheet write SetActiveWorksheet;
    property Field[const vWorksheet: String; const Text: String]: ExcelRange read GetField;
    property Range[const vWorksheet: String; const Name: String]: ExcelRange read GetRange;
  end;
implementation
uses
  System.Variants, ActiveX, Windows, System.SysUtils;

{ TSMBReport }

class function TSMBExcel.CheckExcelInstall: Boolean;
var
  ClassID: TCLSID;
  Rez : HRESULT;
begin
  Rez     := CLSIDFromProgID(PWideChar(WideString(ExcelApp)), ClassID);
  Result  := (Rez = S_OK);
end;

constructor TSMBExcel.Create(const Visible: Boolean = False);
begin
  inherited Create;
  if CheckExcelInstall then
  begin
    FLCID  := GetLCID;
    FExcel := CreateExcelObject(False, FLCID);
  end;
end;

procedure TSMBExcel.Copy(SourceRange, DestinationRange: ExcelRange);
begin
  SourceRange.Copy(DestinationRange);
end;

constructor TSMBExcel.Create(FileName: String; const Visible: Boolean);
begin
  Create(Visible);
  FWorkbook := TSMBExcel.OpenWorkbook(FExcel, FileName, FLCID);
end;

class function TSMBExcel.CreateExcelObject(const Visible: Boolean = False; const ALCID: Integer = 0): ExcelApplication;
var
  vLCID: Integer;
begin
  vLCID := ALCID;
  if ALCID = 0 then
    vLCID := GetLCID;
  if CheckExcelInstall then
  begin
    Result := CoExcelApplication.Create;
    Result.Visible[vLCID] := Visible;
  end
  else
    Result := nil;
end;

destructor TSMBExcel.Destroy;
begin
  if Assigned(FExcel) then FreeExcelObject(FExcel, FLCID);
  inherited;
end;

class function TSMBExcel.FreeExcelObject(var ExcelApp: ExcelApplication; const ALCID: Integer = 0): Boolean;
var
  vLCID: Integer;
begin
  vLCID := ALCID;
  if ALCID = 0 then
    vLCID := GetLCID;
  try
    if ExcelApp.Visible[vLCID] then
      ExcelApp.Visible[vLCID] := False;
    ExcelApp.Quit;
    ExcelApp  := nil;
    Result    := True;
  except
    Result := False;
  end;
end;

function TSMBExcel.GetActiveWorksheet: ExcelWorksheet;
begin
  Result := FWorkbook.ActiveSheet as ExcelWorksheet;
end;

function TSMBExcel.GetField(const vWorksheet: String; const Text: string): ExcelRange;
begin
  Result := Worksheet[vWorksheet].Cells.Find(
    Text,        // What
    EmptyParam,  // After
    EmptyParam,  // LookIn
    xlWhole,     // LookAt
    EmptyParam,  // SearchOrder
    xlByRows,    // SearchDirection
    True,        // MatchCase
    False,       // MatchByte
    EmptyParam); // SearchFormat
end;

class function TSMBExcel.GetLCID: Integer;
begin
  Result := GetUserDefaultLCID;
end;

function TSMBExcel.GetRange(const vWorksheet, Name: String): ExcelRange;
begin
  Result := Worksheet[vWorksheet].Range[Name, EmptyParam];
end;

function TSMBExcel.GetWorksheet(WSName: String): ExcelWorksheet;
begin
  Result := FWorkbook.Sheets[WSName] as ExcelWorksheet;
end;

function TSMBExcel.GetWorkwheets: TList<ExcelWorksheet>;
var
  vCount: Integer;
  i: Integer;
begin
  Result := TList<ExcelWorksheet>.Create;
  vCount := FWorkbook.Sheets.Count;
  for i := 1 to vCount do
    Result.Add(FWorkbook.Sheets[i] as ExcelWorksheet);
end;

procedure TSMBExcel.Hide;
begin
 FExcel.Visible[FLCID] := False;
end;

class procedure TSMBExcel.Hide(var ExcelApp: ExcelApplication;
  const ALCID: Integer);
var
  vLCID: Integer;
begin
  vLCID := ALCID;
  if ALCID = 0 then
    vLCID := GetLCID;
  if ExcelApp.Visible[vLCID] then
      ExcelApp.Visible[vLCID] := False;
end;

class function TSMBExcel.OpenWorkbook(const ExcelApp: ExcelApplication; FileName: String; const ALCID: Integer = 0): ExcelWorkbook;
var
  vLCID: Integer;
begin
  vLCID := ALCID;
  if ALCID = 0 then
    vLCID := GetLCID;
  Result := ExcelApp.Workbooks.Open(
    FileName, // Filename: WideString;
    2, // UpdateLinks: OleVariant; 2 - never update
    False, // ReadOnly: OleVariant;
    EmptyParam, // Format: OleVariant;
    EmptyParam, // Password: OleVariant;
    EmptyParam, // WriteResPassword: OleVariant;
    EmptyParam, // IgnoreReadOnlyRecommended: OleVariant;
    EmptyParam, // Origin: OleVariant;
    EmptyParam, // Delimiter: OleVariant;
    EmptyParam, // Editable: OleVariant;
    EmptyParam, // Notify: OleVariant;
    EmptyParam, // Converter: OleVariant;
    False, // AddToMru: OleVariant;
    EmptyParam, // Local: OleVariant;
    EmptyParam, // CorruptLoad: OleVariant;
    vLCID);
end;

procedure TSMBExcel.SetActiveWorksheet(const Value: ExcelWorksheet);
begin
  if Assigned(Value) then Value.Activate(FLCID);
end;

procedure TSMBExcel.Show;
begin
  FExcel.Visible[FLCID] := True;
end;

class procedure TSMBExcel.Show(var ExcelApp: ExcelApplication;
  const ALCID: Integer);
var
  vLCID: Integer;
begin
  vLCID := ALCID;
  if ALCID = 0 then
    vLCID := GetLCID;
  if not ExcelApp.Visible[vLCID] then
      ExcelApp.Visible[vLCID] := True;
end;

end.
