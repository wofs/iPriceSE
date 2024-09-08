unit FmAboutU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  wBaseU, wFuncU, wLogU, LCLintf,
  StdCtrls, Buttons;

type

  { TFmAbout }

  TFmAbout = class(TForm)
    btnClose: TBitBtn;
    lbMail: TLabel;
    lbMail1: TLabel;
    p1: TPanel;
    Panel1: TPanel;
    t_ProgramName: TStaticText;
    StaticText2: TStaticText;
    procedure FormCreate(Sender: TObject);
    procedure lbLICENSEClick(Sender: TObject);
    procedure lbMailClick(Sender: TObject);
    procedure StaticText2Click(Sender: TObject);
  private
    fBase: TwBase;
    FormIDent: string;
    property wFormID: string read FormIDent write FormIDent;
  public

  end;

var
  FmAbout: TFmAbout;

implementation
uses
  FmMainU;

{$R *.lfm}

{ TFmAbout }


procedure TFmAbout.FormCreate(Sender: TObject);
begin

  wFormID:= Self.Name;

end;

procedure TFmAbout.lbLICENSEClick(Sender: TObject);
begin
  OpenDocument(PathApplication_Unsafe+'LICENSE.txt');
end;

procedure TFmAbout.lbMailClick(Sender: TObject);
begin
  OpenURL(lbMail.Caption);
end;

procedure TFmAbout.StaticText2Click(Sender: TObject);
begin

end;

end.

