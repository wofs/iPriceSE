unit FmTreeU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Buttons, Classes, ExtCtrls, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  wLogU,
  wBaseU, wDBTreeU, wTypesU,
  ComCtrls;

type

  { TFmTree }

  TFmTree = class(TForm)
    BitBtn1: TBitBtn;
    ImageList16: TImageList;
    ImageListTreeGroup: TImageList;
    ImageListTreeOwner: TImageList;
    PanelBtn: TPanel;
    TreeGroup: TTreeView;
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TreeGroupDblClick(Sender: TObject);
    procedure TreeGroupGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupSelectionChanged(Sender: TObject);
  private
    fFormName: string;
    fBase: TwBase;
    fIdGroup: Int64;
    fTree: TwDBTree;
    fIdMainOwner: Int64; // ID основного контрагента (к которому привязан каталог)
    fChanged: boolean;
  private
    FMode: integer;
    { private declarations }
    procedure SetStatus(_Text:string);
  public
    property Base: TwBase read fBase write fBase;
    property Tree: TwDBTree read fTree write fTree;
    property IdGroup: Int64 read fIdGroup write fIdGroup;
    property IdMainOwner: Int64 read fIdMainOwner write fIdMainOwner;
    property Mode: integer read FMode write FMode;
    { public declarations }
  end;

var
  FmTree: TFmTree;

implementation

{$R *.lfm}

{ TFmTree }

procedure TFmTree.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if fChanged then
  begin
    if MessageDlg('Закрыть без сохранения?',mtConfirmation, mbOKCancel, 0) = mrCancel
     then
       CanClose:= false
     else
       ModalResult:= mrCancel;
  end else
  if FMode<>2 then
    ModalResult:= mrOk;

end;

procedure TFmTree.FormCreate(Sender: TObject);
begin
  fFormName:= Self.Name;
  FMode:=0;
  fIdGroup:=0;
  fChanged:= false;
end;

procedure TFmTree.FormDestroy(Sender: TObject);
begin
  try
   wLog('FmTree','Выгрузка...');

      // выгружаем подгруженные DBTree
   fTree.Destroy();


  wLog('FmTree','Выгрузка успешно завершена.');

  except
    on E: Exception do
    begin
        SetStatus('Сбой выгрузки FmTree');
        wLog('FmTree','Ошибка [FmDestroy]: "' + E.Message + '"');
        wLog('FmTree','Сбой выгрузки FmTree.');
        ShowMessage('Ошибка [FmDestroy]: "' + E.Message + '"');
     end;
  end;
end;

procedure TFmTree.FormShow(Sender: TObject);
begin
  //DBTree
  fIdMainOwner:= fBase.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист

case FMode of
    0:
      begin
          fTree:= TwDBTree.Create(fBase,TreeGroup,'CATALOG_GROUP','IDPARENT,NAME',['IDOWNER',fIdMainOwner]);
          if Assigned(TreeGroup.PopupMenu) then
          begin
            with TreeGroup.PopupMenu do
            begin
              Images:= ImageList16;
              Items[0].ImageIndex:= 0;
              Items[1].ImageIndex:= 1;
              Items[2].ImageIndex:= 2;
            end;
          end;
      end;
    1:
      begin
        fTree:= TwDBTree.Create(fBase,TreeGroup,'PL_GROUP','IDPARENT,NAME',['IDOWNER',fIdMainOwner]);
        TreeGroup.PopupMenu:= nil;
      end;
    2:
      begin
        fTree:= TwDBTree.Create(fBase,TreeGroup,'OWNER','IDPARENT,NAME',nil);
        fTree.MultiSelect:= true;
        TreeGroup.PopupMenu:= nil;
        TreeGroup.Images:= ImageListTreeOwner;
      end;
    3:
      begin
        fTree:= TwDBTree.Create(fBase,TreeGroup,'OWNER','IDPARENT,NAME',nil);
        TreeGroup.PopupMenu:= nil;
        TreeGroup.Images:= ImageListTreeOwner;
      end;
  end;

   fTree.Fill();

   fTree.FindNodeWithDataInt(fIdGroup);
   fTree.FirstFillTree:= false;
end;

procedure TFmTree.TreeGroupDblClick(Sender: TObject);
begin
  if FMode = 2  then exit;

  fChanged:= false;
  Close();
end;

procedure TFmTree.TreeGroupGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmTree.TreeGroupGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmTree.TreeGroupSelectionChanged(Sender: TObject);
var
  _arr: ArrayOfInteger;
begin
if fTree.FirstFillTree then exit;

if fTree.Tree.Selected.Level <> 0 then
begin
  _arr:= fTree.SelectedItems;
  if Assigned(_arr) then
    fIdGroup:= _arr[0];
end else
    fIdGroup:= 0;

  if FMode <>2 then fChanged:= true;

  _arr:=nil;
end;

procedure TFmTree.SetStatus(_Text: string);
begin
  wStatus(fFormName,_Text,true);
end;

end.

