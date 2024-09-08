unit FmInvoceU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Buttons, Classes, ComCtrls, ExtCtrls, Spin, StdCtrls, SysUtils, Forms, Controls, Graphics, Dialogs, wBaseU, wFuncU, wTypesU;

type

  { TFmInvoce }

  TFmInvoce = class(TForm)
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    cbInvoced: TComboBox;
    gbIdent: TGroupBox;
    GroupBox1: TGroupBox;
    gbVipiska: TGroupBox;
    GroupBox2: TGroupBox;
    edLabel: TLabeledEdit;
    edScod: TLabeledEdit;
    edVendorCode: TLabeledEdit;
    edUnit: TLabeledEdit;
    edPrice: TLabeledEdit;
    edSum: TLabeledEdit;
    bgInvoced: TGroupBox;
    Label1: TLabel;
    edName: TMemo;
    edRemark: TMemo;
    pcInvoce: TPageControl;
    Panel1: TPanel;
    Panel2: TPanel;
    edQuantity: TSpinEdit;
    tsAddEdit: TTabSheet;
    procedure cbInvocedChange(Sender: TObject);
    procedure edQuantityChange(Sender: TObject);
    procedure edQuantityKeyPress(Sender: TObject; var Key: char);
    procedure edSumChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    fBase: TwBase;
    fMResult: Integer;
    fonCbInvoiceChange: TIntValueNotify;

  public
    property Base:TwBase read fBase write fBase;
    property onCbInvoiceChange: TIntValueNotify read fonCbInvoiceChange write fonCbInvoiceChange;
  end;

var
  FmInvoce: TFmInvoce;

implementation

{$R *.lfm}

{ TFmInvoce }

procedure TFmInvoce.edQuantityChange(Sender: TObject);
begin
  edSum.Text:= FloatToStr(GetValue(edPrice)*edQuantity.Value);
end;

procedure TFmInvoce.edQuantityKeyPress(Sender: TObject; var Key: char);
begin
  if Key = #13 then
    begin
      fMResult:= mrOK;
      Close;
    end;
end;

procedure TFmInvoce.cbInvocedChange(Sender: TObject);
begin
  if Assigned(onCbInvoiceChange) then onCbInvoiceChange(self,cmbxSelectID(cbInvoced));
end;

procedure TFmInvoce.edSumChange(Sender: TObject);
begin
  FormatValue(TLabeledEdit(Sender));
end;

procedure TFmInvoce.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
   if ModalResult<>mrOK then ModalResult:= fMResult;
end;

procedure TFmInvoce.FormCreate(Sender: TObject);
begin
  fMResult:= mrCancel;
end;

procedure TFmInvoce.FormShow(Sender: TObject);
begin
   edQuantity.SetFocus;
end;

end.

