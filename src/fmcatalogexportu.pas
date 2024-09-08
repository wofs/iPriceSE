unit FmCatalogExportU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Buttons, Classes, ExtCtrls, StdCtrls, SysUtils, Forms, Controls, Graphics, Dialogs, wFuncU, wTypesU;

type

  { TFmCatalogExport }

  TFmCatalogExport = class(TForm)
    btnCancel: TBitBtn;
    btnReport: TBitBtn;
    CheckBox1: TCheckBox;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    lstPrices: TListBox;
    lstStocks: TListBox;
    Panel2: TPanel;
    pButtom: TPanel;
    Splitter1: TSplitter;
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    function GetSelectedStocks: ArrayOfInteger;
    function GetSelectedPrices: ArrayOfInteger;
    function GetStockOnly: boolean;

  public
    property ListPrices: TListBox read lstPrices write lstPrices;
    property SelectedPrices: ArrayOfInteger read GetSelectedPrices;
    property SelectedStocks: ArrayOfInteger read GetSelectedStocks;
    property StockOnly: boolean read GetStockOnly;
  end;

var
  FmCatalogExport: TFmCatalogExport;

implementation

{$R *.lfm}

{ TFmCatalogExport }

procedure TFmCatalogExport.FormDestroy(Sender: TObject);
begin
  lbxClearData(lstPrices);
  lbxClearData(lstStocks);
end;

procedure TFmCatalogExport.FormShow(Sender: TObject);
begin
  lstPrices.Selected[0]:= true;
  lstStocks.Selected[0]:= true;
end;

function TFmCatalogExport.GetSelectedStocks: ArrayOfInteger;
begin
  Result:= lbxSelectIDs(lstStocks);
end;

function TFmCatalogExport.GetSelectedPrices: ArrayOfInteger;
begin
  Result:= lbxSelectIDs(lstPrices);
end;

function TFmCatalogExport.GetStockOnly: boolean;
begin
  Result:= CheckBox1.Checked;
end;

end.

