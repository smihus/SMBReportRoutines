unit TestSMBExcel;
{

  Delphi DUnit Test Case
  ----------------------
  This unit contains a skeleton test case class generated by the Test Case Wizard.
  Modify the generated code to correctly setup and call the methods from the unit 
  being tested.

}

interface

uses
  TestFramework, SMBExcel, System.Variants, Excel_TLB;

type
  // Test methods for class TSMBReport

  TestTSMBExcel = class(TTestCase)
  strict private
    FSMBExcel: TSMBExcel;
  public
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure OpenWorkbook;
    procedure CreateExcelObj;
    procedure FreeExcelObj;
    procedure CheckExcelInstall;
    procedure Show;
  end;

implementation

procedure TestTSMBExcel.CheckExcelInstall;
begin
  CheckTrue(FSMBExcel.CheckExcelInstall);
end;

procedure TestTSMBExcel.CreateExcelObj;
var
  FExcel: ExcelApplication;
begin
  FExcel := FSMBExcel.CreateExcelObject(True);
  CheckTrue(Assigned(FExcel));
  FSMBExcel.FreeExcelObject(FExcel);
end;

procedure TestTSMBExcel.FreeExcelObj;
var
  FExcel: ExcelApplication;
begin
  FExcel := FSMBExcel.CreateExcelObject();
  CheckTrue(FSMBExcel.FreeExcelObject(FExcel));
  CheckFalse(Assigned(FExcel));
end;

procedure TestTSMBExcel.SetUp;
begin
  FSMBExcel := TSMBExcel.Create;
end;

procedure TestTSMBExcel.Show;
var
  FileName: string;
  FExcel: ExcelApplication;
  WB: ExcelWorkbook;
begin
  FExcel    := FSMBExcel.CreateExcelObject();
  FileName  := 'C:\Users\1\Google ����\RAD Studio Projects\SMBComponents\SMBReport\Patterns\pattern1.xlsx';
  WB        := FSMBExcel.OpenWorkbook(FExcel, FileName);
  FSMBExcel.Show(FExcel);
  FSMBExcel.FreeExcelObject(FExcel);
end;

procedure TestTSMBExcel.TearDown;
begin
  FSMBExcel.Free;
  FSMBExcel := nil;
end;

procedure TestTSMBExcel.OpenWorkbook;
var
  FileName: string;
  FExcel: ExcelApplication;
begin
  FExcel := FSMBExcel.CreateExcelObject();
  FileName := 'C:\Users\1\Google ����\RAD Studio Projects\SMBComponents\SMBReport\Patterns\pattern1.xlsx';
  CheckNotNull(FSMBExcel.OpenWorkbook(FExcel, FileName));
  FSMBExcel.FreeExcelObject(FExcel);
end;

initialization
  // Register any test cases with the test runner
  RegisterTest(TestTSMBExcel.Suite);
end.

