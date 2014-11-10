unit SMBExcel;

interface
uses
   Excel_TLB;

const
  ExcelApp = 'Excel.Application';

   type
  TSMBExcel = class
  private
    FExcel: ExcelApplication;
    FLCID: Integer;
    FWorkbook: ExcelWorkbook;
    function GetWorksheet(WSName: String): ExcelWorksheet;
  public
    constructor Create(const Visible: Boolean = False); overload;
    constructor Create(FileName: String; const Visible: Boolean = False); overload;
    destructor Destroy; override;
    class function OpenWorkbook(const ExcelApp: ExcelApplication; FileName: String; const ALCID: Integer = 0): ExcelWorkbook;
    class function CreateExcelObject(const Visible: Boolean = False; const ALCID: Integer = 0): ExcelApplication;
    class function FreeExcelObject(var ExcelApp: ExcelApplication; const ALCID: Integer = 0): Boolean;
    class function GetLCID: Integer;
    class function CheckExcelInstall: Boolean;
    class procedure Show(var ExcelApp: ExcelApplication; const ALCID: Integer = 0);
    class procedure Hide(var ExcelApp: ExcelApplication; const ALCID: Integer = 0);

    property Worksheet[WSName: String]: ExcelWorksheet read GetWorksheet;
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

constructor TSMBExcel.Create(FileName: String; const Visible: Boolean);
begin
  Create(Visible);
  FWorkbook := TSMBExcel.OpenWorkbook(FExcel, FileName, FLCID);
end;

class function TSMBExcel.CreateExcelObject(const Visible: Boolean = False; const ALCID: Integer = 0): ExcelApplication;
var
  _LCID: Integer;
begin
  _LCID := ALCID;
  if ALCID = 0 then
    _LCID := GetLCID;
  if CheckExcelInstall then
  begin
    Result := CoExcelApplication.Create;
    Result.Visible[_LCID] := Visible;
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
  _LCID: Integer;
begin
  _LCID := ALCID;
  if ALCID = 0 then
    _LCID := GetLCID;
  try
    if ExcelApp.Visible[_LCID] then
      ExcelApp.Visible[_LCID] := False;
    ExcelApp.Quit;
    ExcelApp  := nil;
    Result    := True;
  except
    Result := False;
  end;
end;

class function TSMBExcel.GetLCID: Integer;
begin
  Result := GetUserDefaultLCID;
end;

function TSMBExcel.GetWorksheet(WSName: String): ExcelWorksheet;
begin
  Result := nil;
end;

class procedure TSMBExcel.Hide(var ExcelApp: ExcelApplication;
  const ALCID: Integer);
var
  _LCID: Integer;
begin
  _LCID := ALCID;
  if ALCID = 0 then
    _LCID := GetLCID;
  if ExcelApp.Visible[_LCID] then
      ExcelApp.Visible[_LCID] := False;
end;

class function TSMBExcel.OpenWorkbook(const ExcelApp: ExcelApplication; FileName: String; const ALCID: Integer = 0): ExcelWorkbook;
var
  _LCID: Integer;
begin
  _LCID := ALCID;
  if ALCID = 0 then
    _LCID := GetLCID;
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
    _LCID);
end;

class procedure TSMBExcel.Show(var ExcelApp: ExcelApplication;
  const ALCID: Integer);
var
  _LCID: Integer;
begin
  _LCID := ALCID;
  if ALCID = 0 then
    _LCID := GetLCID;
  if not ExcelApp.Visible[_LCID] then
      ExcelApp.Visible[_LCID] := True;
end;

end.
