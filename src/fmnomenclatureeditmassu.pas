unit FmNomenclatureEditMassU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  wBaseU,
  StdCtrls, Buttons, FmNomenclatureEditU;

type

  { TFmNomenclatureEditMass }

  TFmNomenclatureEditMass = class(TForm)
    btnCancel: TBitBtn;
    btnOK: TBitBtn;
    cbAll: TCheckBox;
    cbUnit: TCheckBox;
    cbGroup: TCheckBox;
    cbPrice: TCheckBox;
    cbMain: TGroupBox;
    cbP1: TCheckBox;
    cbN: TCheckBox;
    cbM: TCheckBox;
    cbD: TCheckBox;
    cbC: TCheckBox;
    cbK: TCheckBox;
    cbUnselect: TCheckBox;
    gbUnit: TGroupBox;
    gbGroup: TGroupBox;
    gbPrice: TGroupBox;
    gbChange: TGroupBox;
    l_edGroupText: TLabel;
    pMain: TPanel;
    pBtn: TPanel;
    pUnit: TPanel;
    pGroup: TPanel;
    pPrice: TPanel;
    procedure btnCancelClick(Sender: TObject);
    procedure cbGroupChange(Sender: TObject);
    procedure cbAllChange(Sender: TObject);
    procedure cbPriceChange(Sender: TObject);
    procedure cbUnitChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure gbPriceClick(Sender: TObject);
  private
    fBase: TwBase;
    procedure SeTwBase(AValue: TwBase);
    { private declarations }
  public
    { public declarations }
    property Base: TwBase read fBase write SeTwBase;
  end;

var
  FmNomenclatureEditMass: TFmNomenclatureEditMass;
  _Form: TFmNomenclatureEdit;

implementation

uses
  wDBTreeU, pkgCatalogU;

{$R *.lfm}

{ TFmNomenclatureEditMass }

procedure TFmNomenclatureEditMass.cbUnitChange(Sender: TObject);
begin
  if cbUnit.Checked then
     begin
       _Form.kEdini.Parent:= gbUnit;
       _Form.kEdini.Top:= -2;
       _Form.kEdini.Left:= 154;
     end else
     begin
       _Form.kEdini.Parent:= _Form.gbGroup;
     end;
end;

procedure TFmNomenclatureEditMass.FormCloseQuery(Sender: TObject;
  var CanClose: boolean);
begin
  if ModalResult = mrCancel then
    if MessageDlg('Закрыть без сохранения?',mtConfirmation, mbOKCancel, 0) = mrCancel
     then
       CanClose:= false
     else
       ModalResult:= mrCancel;
end;

procedure TFmNomenclatureEditMass.FormCreate(Sender: TObject);
begin
  _Form:= TFmNomenclatureEdit.Create(Self);

end;

procedure TFmNomenclatureEditMass.cbGroupChange(Sender: TObject);
begin
  if cbGroup.Checked then
     begin
       _Form.edGroup.Parent:= gbGroup;
       _Form.edGroup.Top:= -2;
       _Form.edGroup.Left:= 130;
       _Form.btnChangeGroup.Parent:= gbGroup;
       _Form.btnChangeGroup.Top:= -2;
       _Form.btnChangeGroup.Left:= 544;
       _Form.edGroup.Tag:= gbGroup.Tag; //_DBTreeGroupPrice
       _Form.edGroup.Text:=l_edGroupText.Caption;
     end else
     begin
       _Form.edGroup.Parent:= _Form.gbGroup;
       _Form.btnChangeGroup.Parent:= _Form.gbGroup;

     end;
end;

procedure TFmNomenclatureEditMass.cbAllChange(Sender: TObject);
begin
   cbP1.Checked:= cbAll.Checked;
   cbN.Checked:= cbAll.Checked;
   cbM.Checked:= cbAll.Checked;
   cbD.Checked:= cbAll.Checked;
   cbC.Checked:= cbAll.Checked;
   cbK.Checked:= cbAll.Checked;
end;

procedure TFmNomenclatureEditMass.btnCancelClick(Sender: TObject);
begin

end;

procedure TFmNomenclatureEditMass.cbPriceChange(Sender: TObject);
begin
  if cbPrice.Checked then
     begin
       _Form.pPrices.Parent:= pPrice;
       pPrice.Height:=pPrice.Height+300;
       _Form.pPrices.Align:= alClient;
       gbChange.Visible:=true;
       _Form.GridPriceFieldsRefrash();
     end else
     begin
       _Form.pPrices.Parent:= _Form.pcTbPrice;
       gbChange.Visible:=false;
       pPrice.Height:=pPrice.Height-300;
     end;
end;

procedure TFmNomenclatureEditMass.FormShow(Sender: TObject);
begin

end;

procedure TFmNomenclatureEditMass.gbPriceClick(Sender: TObject);
begin

end;

procedure TFmNomenclatureEditMass.SeTwBase(AValue: TwBase);
begin
  fBase:= AValue;
  _Form.Base:= AValue;
end;

end.

