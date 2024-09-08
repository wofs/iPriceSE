unit FmFormatsImportU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, BufDataset, db, dbf, Forms, Controls, Graphics, Dialogs, ExtCtrls, Buttons, Grids, DBGrids, StdCtrls,
  LazUTF8,
  fpspreadsheetctrls, fpspreadsheetgrid, fpsexport;

type

  { TFmFormatsImport }

  TFmFormatsImport = class(TForm)
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    Panel2: TPanel;
    StringGrid1: TStringGrid;
    sWorkbookSource1: TsWorkbookSource;
  private

  public

  end;

var
  FmFormatsImport: TFmFormatsImport;

implementation

{$R *.lfm}

end.

