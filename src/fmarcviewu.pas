unit FmArcViewU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  wFuncU, wTypesU, wZipperU
  ;

type

  { TFmArcView }

  TFmArcView = class(TForm)
    cbDecodeFileName: TCheckBox;
    ListFiles: TListBox;
    Panel1: TPanel;
    procedure cbDecodeFileNameChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ListFilesDblClick(Sender: TObject);
  private
    FArcFileName: string;
    FDecodeFileName: boolean;
    fSelectPriceConvert: boolean;
    fFormatID: integer;
    FZipper: TwZipper;

  public
   SelectedFileName: string;
   property Zipper: TwZipper read FZipper write FZipper;
   property ArcFileName: string read FArcFileName write FArcFileName;
   property DecodeFileName: boolean read FDecodeFileName write FDecodeFileName default false;
   property FormatID: integer read fFormatID;
  end;

var
  FmArcView: TFmArcView;

implementation

{$R *.lfm}

{ TFmArcView }

procedure TFmArcView.ListFilesDblClick(Sender: TObject);
begin
 SelectedFileName:= ListFiles.Items[ListFiles.ItemIndex];
 if Assigned(TwData(ListFiles.Items.Objects[ListFiles.ItemIndex])) then
   fFormatID:= TwData(ListFiles.Items.Objects[ListFiles.ItemIndex]).Value;
 Close;
end;

procedure TFmArcView.FormCreate(Sender: TObject);
begin
  SelectedFileName:='';
end;

procedure TFmArcView.FormShow(Sender: TObject);
begin
  cbDecodeFileName.Checked:= DecodeFileName;
end;

procedure TFmArcView.cbDecodeFileNameChange(Sender: TObject);
begin
  DecodeFileName:= cbDecodeFileName.Checked;
  if Assigned(Zipper) then Zipper.DecodeFileName:= DecodeFileName;
  ListFiles.Items:= Zipper.ReadFileList(FArcFileName);
end;

procedure TFmArcView.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  i: Integer;
begin
  for i:=0 to ListFiles.Count-1 do
    TwData(ListFiles.Items.Objects[i]).Free;
end;

end.

