unit SMBReport;

interface

uses
  SMBExcel;
type
  TSMBReport = class
  private
    FPattern: TSMBExcel;
    FDestination: TSMBExcel;
    FCurrentCellAddress: String;
  public
    constructor Create(Pattern, Destination: TSMBExcel);
    procedure StartFrom(CellAddress: String);
    property CurrentCellAddress: String read FCurrentCellAddress;
  end;
implementation

{ TSMBReport }

constructor TSMBReport.Create(Pattern, Destination: TSMBExcel);
begin
  FPattern      := Pattern;
  FDestination  := Destination;
  FCurrentCellAddress := 'A1';
end;

procedure TSMBReport.StartFrom(CellAddress: String);
begin
  FCurrentCellAddress := CellAddress;
end;

end.
