unit wTProgressU;

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

  TProgressBarValues = record
    MaxValue: integer;
    CurValue: integer;
  end;

  TProgressBarName = (pbTop, pbBottom);

  TProgressEventInit = procedure(const aProgressBarName:TProgressBarName; aValue: integer) of object;
  TProgressEventUpdate = procedure(const aProgressBarName:TProgressBarName; aValue: integer) of object;
  TProgressEventStatus = procedure(const aValue: string) of object;

  { TwBar }

  TwBar = class (TProgressBar)
    private
      fValues: TProgressBarValues;

    public
      constructor Create(AOwner: TComponent); override;

      property Values: TProgressBarValues read fValues write fValues;

      procedure Init(aMax:integer; const aStep: integer = 1);
      procedure SetBar(const aPosition: integer= -1);
      procedure Clear();
  end;

  { TProgress }

  TProgress = class(TForm)
    fStatus: TMemo;
    fpbStatusTop: TwBar;
    fpbStatusBottom: TwBar; // Fruction
    procedure fonCloseQuery(Sender: TObject; var CanClose: boolean);
  private
    fHeight: Integer;
    fNoClose: boolean;
    fonStopForce: TNotifyEvent;
    fPanel: TPanel;
    fWidth: Integer;
     procedure fOnClose(Sender: TObject; var CloseAction: TCloseAction);
     procedure SetShowBottom(aValue: boolean);
     procedure SetShowMaxMinButtons(aValue: boolean);
     procedure SetShowLog(aValue: boolean);
     procedure SetShowTop(aValue: boolean);
  public
    constructor Create(TheOwner: TComponent); override;

    procedure SetStatus(const aText: string);

    procedure InitBar(aProgressBar: TProgressBarName; aMax: integer; const aStep: integer = 1); // Top
    procedure InitBarTop(aMax,aStep: integer);
    procedure InitBarBottom(aMax,aStep: integer);

    procedure SetBar(aProgressBar: TProgressBarName; const aPosition: integer = -1); // Top
    procedure SetLog(const aText: string); // Top
    procedure SetBarTop(aPosition: integer);
    procedure SetBarBottom(aPosition: integer);

    procedure SetSize(aHeight, aWidth: integer);
    procedure ForceClose;
    procedure onSetSize();

    property ShowMaxMinButtons: boolean write SetShowMaxMinButtons;
    property ShowLog: boolean write SetShowLog;
    property ShowTop: boolean write SetShowTop;
    property ShowBottom: boolean write SetShowBottom;
    property onStopForce:TNotifyEvent read fonStopForce write fonStopForce;

    property NoClose:boolean read fNoClose write fNoClose;
  end;

implementation

{ TwBar }

constructor TwBar.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  self.Align:= alBottom;
  self.Height:= 20;
  self.Smooth:= true;
end;

procedure TwBar.Init(aMax: integer; const aStep: integer);
begin
  self.Max:= aMax;
  self.Step:= aStep;
end;

procedure TwBar.SetBar(const aPosition: integer);
begin

  if aPosition = -1 then
    self.StepIt
  else
    self.Position:=aPosition;

end;

procedure TwBar.Clear();
begin
   SetBar(0);
end;

{ TProgress }

procedure TProgress.fonCloseQuery(Sender: TObject; var CanClose: boolean);
begin

  if fNoClose then
     if MessageDlg('Отменить текущую операцию?',mtConfirmation, mbOKCancel, 0) = mrOk then
       onStopForce(self);

  if fNoClose then
    CanClose:= false;
end;

procedure TProgress.SetShowBottom(aValue: boolean);
begin
  fpbStatusBottom.Visible:= aValue;
  onSetSize();
end;

procedure TProgress.fOnClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  //CloseAction:= caFree;
end;

procedure TProgress.SetShowMaxMinButtons(aValue: boolean);
begin
  if aValue then
    self.BorderIcons:= [biMaximize, biMinimize, biSystemMenu]
  else
    self.BorderIcons:= [biSystemMenu];
end;

procedure TProgress.SetShowLog(aValue: boolean);
begin
  fStatus.Visible:= aValue;
  onSetSize();
end;

procedure TProgress.SetShowTop(aValue: boolean);
begin
  fpbStatusTop.Visible:= aValue;
  onSetSize();
end;

constructor TProgress.Create(TheOwner: TComponent);
begin
    inherited CreateNew(TheOwner);
    fNoClose:= true;
    self.Position:=poScreenCenter;
    self.FormStyle:= fsStayOnTop;
    self.OnCloseQuery:=@fonCloseQuery;
    self.OnClose:= @fOnClose;
    self.ShowInTaskBar:= stAlways;

    fPanel:= TPanel.Create(self);
    fPanel.Parent:= self;
    fPanel.Align:= alBottom;
    fPanel.AutoSize:= true;
    fPanel.BorderStyle:= bsNone;
    fPanel.BevelOuter:= bvNone;

    fpbStatusTop:= TwBar.Create(self);
    fpbStatusTop.Parent:= fPanel;


    fpbStatusBottom:= TwBar.Create(self);
    fpbStatusBottom.Parent:= fPanel;

    fStatus:= TMemo.Create(self);
    fStatus.ScrollBars:= ssAutoBoth;
    fStatus.Parent:= self;
    fStatus.Align:= alClient;
    fStatus.Color:= clMoneyGreen;
    fStatus.ReadOnly:= True;

    SetSize(100,400);
end;

procedure TProgress.SetStatus(const aText: string);
begin
  fStatus.Lines.Add(aText);
end;

procedure TProgress.InitBar(aProgressBar:TProgressBarName; aMax: integer; const aStep: integer);
begin
  fNoClose:= true;

  case aProgressBar of
    pbBottom: InitBarBottom(aMax, aStep);
    pbTop: InitBarTop(aMax, aStep);
  end;
end;

procedure TProgress.InitBarTop(aMax, aStep: integer);
begin
  fpbStatusTop.Visible:= true;
  fpbStatusTop.Height:=fpbStatusTop.Height;

  fpbStatusTop.Max:= aMax;
  fpbStatusTop.Step:= aStep;
end;

procedure TProgress.InitBarBottom(aMax, aStep: integer);
begin
  fpbStatusBottom.Max:= aMax;
  fpbStatusBottom.Step:= aStep;
end;

procedure TProgress.SetBar(aProgressBar: TProgressBarName; const aPosition: integer);
begin
  case aProgressBar of
    pbBottom: SetBarBottom(aPosition);
    pbTop: SetBarTop(aPosition);
  end;
end;

procedure TProgress.SetLog(const aText: string);
begin
  fStatus.Lines.Add(aText);
end;

procedure TProgress.SetBarTop(aPosition: integer);
begin
  fpbStatusTop.SetBar(aPosition);
end;

procedure TProgress.SetBarBottom(aPosition: integer);
begin
  fpbStatusBottom.SetBar(aPosition);
end;

procedure TProgress.SetSize(aHeight, aWidth: integer);
begin
  fHeight:= aHeight;
  fWidth:= aWidth;
  onSetSize();
end;

procedure TProgress.ForceClose;
begin
  fNoClose:= false;
  self.Close();
end;

procedure TProgress.onSetSize();
begin
  self.Height:= 0;
  self.Width:= fWidth;

  if fpbStatusBottom.Visible then self.Height:= self.Height + fpbStatusBottom.Height;
  if fpbStatusTop.Visible then self.Height:= self.Height + fpbStatusTop.Height;
  if fStatus.Visible then self.Height:= fHeight;

  self.MoveToDefaultPosition;
end;


end.

