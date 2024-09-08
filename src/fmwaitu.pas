unit FmWaitU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  ComCtrls, StdCtrls;

type

  { TFmWait }

  TFmWait = class(TForm)
    mStatus: TMemo;
    Panel1: TPanel;
    pbStatus: TProgressBar;
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
  private
    fNoClose: boolean;
     StepIt: boolean;
  public
    procedure SetStatus(const _Text: string = ''; const ProcessMessages: boolean = false);
    procedure InitBar(_Max, _Step: integer);
    procedure SetBar(const _Position: integer);
    property Memo:TMemo read mStatus write mStatus;
    property NoClose:boolean read fNoClose write fNoClose;
  end;

var
  FmWait: TFmWait;

implementation

{$R *.lfm}

{ TFmWait }

procedure TFmWait.FormCreate(Sender: TObject);
begin
  fNoClose:= false;
end;

procedure TFmWait.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if NoClose then
      CanClose:= false;
end;

procedure TFmWait.SetStatus(const _Text: string; const ProcessMessages: boolean);
begin
  mStatus.Lines.Add(_Text);
  if ProcessMessages then Application.ProcessMessages;
end;

procedure TFmWait.InitBar(_Max, _Step: integer);
begin
    pbStatus.Max:=_Max;
    if _Step = 0 then
        StepIt:= false else
        begin
          StepIt:= true;
          pbStatus.Step:=_Step;
        end;
end;

procedure TFmWait.SetBar(const _Position: integer);
begin
      if StepIt then pbStatus.StepIt else
                 pbStatus.Position:=_Position;
end;

end.

