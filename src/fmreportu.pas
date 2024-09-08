unit FmReportU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Buttons, Classes, ComCtrls, ExtCtrls, StdCtrls, SysUtils, Forms, Controls, Graphics, Dialogs, wBaseU, wDBGridU, wDBTreeU, wTypesU;

type

  { TFmReport }

  TFmReport = class(TForm)
    btnCancel: TBitBtn;
    btnReport: TBitBtn;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox4: TGroupBox;
    ImageListTree: TImageList;
    ListPriceBase: TListBox;
    ListPriceCompare: TListBox;
    PageControl1: TPageControl;
    pPriceSelectPos: TPanel;
    pGroup: TPanel;
    Panel2: TPanel;
    pButtom: TPanel;
    Panel4: TPanel;
    pPriceAnalogPos: TPanel;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    TabSheet1: TTabSheet;
    TreeOwner: TTreeView;
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TreeOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
  private
    fBase: TwBase;
    fIdMainOwner: Integer;
    fPriceBase: TPriceType;
    fPriceCompare: TPriceType;
    fSelectedItems: ArrayOfInteger;
    fTreeOwner: TwDBTree;
    function GetPrice(aIndex: integer): TPriceType;

  public
    property Base: TwBase read fBase write fBase;
    property SelectedItems: ArrayOfInteger read fSelectedItems;
    property PriceBase: TPriceType read fPriceBase;
    property PriceCompare: TPriceType read fPriceCompare;

  end;

var
  FmReport: TFmReport;

implementation

{$R *.lfm}

{ TFmReport }

function TFmReport.GetPrice(aIndex: integer): TPriceType;
begin
case aIndex of
    0: Result:= ptBase;
    1: Result:= ptPrice2;
    2: Result:= ptPrice3;
    3: Result:= ptPrice4;
    4: Result:= ptPrice5;
  end;
end;

procedure TFmReport.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  fSelectedItems:= fTreeOwner.SelectedItems;
  fPriceBase:= GetPrice(ListPriceBase.ItemIndex);
  fPriceCompare:= GetPrice(ListPriceCompare.ItemIndex);
end;

procedure TFmReport.FormDestroy(Sender: TObject);
begin
  fTreeOwner.Destroy();
end;

procedure TFmReport.FormShow(Sender: TObject);
begin

    fIdMainOwner:= fBase.ReadSettingByName('setDefaultOwner');

    fTreeOwner:= TwDBTree.Create(fBase,TreeOwner,'OWNER','IDPARENT, NAME',[]);
    fTreeOwner.MultiSelect:= true;
    fTreeOwner.Expanded:= true;
    fTreeOwner.ShowChildrenItems:= true;
    fTreeOwner.Fill();
end;

procedure TFmReport.TreeOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
  if TTreeData(Node.Data).Value = fIdMainOwner then
  begin
    Node.ImageIndex:=2;
    exit;
  end;

  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmReport.TreeOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

end.

