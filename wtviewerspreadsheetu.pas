unit wTViewerSpreadsheetU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Buttons, Classes, Controls, db, DBGrids, Dialogs, ExtCtrls, fpspreadsheet, fpspreadsheetctrls, fpsallformats, fpspreadsheetgrid, fpsTypes, Graphics, Grids,
  LCLIntf, LCLType, SysUtils, Forms, wTypesU

  ;

  type

    { TwViewer }

    TwViewer = class(TForm)

      procedure fonCloseQuery(Sender: TObject; var CanClose: boolean);
    private
     const
      fDefaultFilterIndex = 2;
     var
      fBtnSave: TBitBtn;
      fGrid: TsWorksheetGrid;
      fHeight: Integer;
      fPanelCenter: TPanel;
      fPanelTop: TPanel;
      fWidth: Integer;
      fNoClose: boolean;
      fonStopForce: TNotifyEvent;
      fPanelBottom: TPanel;
      fWorkBook: TsWorkbook;
      procedure BtnSaveOnClick(Sender: TObject);
      function GetWorkBook: TsWorkbook;
      function GetWorkbookSource: TsWorkbookSource;
      function GetWorkSheet: TsWorksheet;
      procedure SetWorkbookSource(aValue: TsWorkbookSource);
      procedure sWorksheetGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

    public
      constructor Create(TheOwner: TComponent); override;
      destructor Destroy; override;

      procedure SetStatus(const aText: string);

      procedure SetSize(aHeight, aWidth: integer);
      procedure ForceClose;
      procedure onSetSize();

      property onStopForce:TNotifyEvent read fonStopForce write fonStopForce;

      property NoClose:boolean read fNoClose write fNoClose;

      property WorkBook: TsWorkbook read GetWorkBook write fWorkBook;
      property WorkSheet: TsWorksheet read GetWorkSheet;
      property WorkbookSource: TsWorkbookSource read GetWorkbookSource write SetWorkbookSource;
    end;

implementation

{ TwViewer }

procedure TwViewer.fonCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if fNoClose then
     if MessageDlg('Закрыть?',mtConfirmation, mbOKCancel, 0) = mrOk then
                                                             fNoClose:= false;
  if fNoClose then
    CanClose:= false;
end;

function TwViewer.GetWorkbookSource: TsWorkbookSource;
begin
  Result:= fGrid.WorkbookSource;
end;

function TwViewer.GetWorkBook: TsWorkbook;
begin
  Result:= fGrid.Workbook;
end;

procedure TwViewer.BtnSaveOnClick(Sender: TObject);
var
  _SaveDialog: TSaveDialog;
begin
  _SaveDialog:= TSaveDialog.Create(fBtnSave);

  try
    try
      _SaveDialog.FileName:= Caption;
      _SaveDialog.Filter:='OpenDocument (*.ods)|*.ods| Excel (*.xls)|*.xls| Excel (*.xlsx)|*.xlsx|Comma Text (*.csv)|*.csv';
      _SaveDialog.FilterIndex:= fDefaultFilterIndex;
      _SaveDialog.Options:= [ofOverwritePrompt];
      screen.Cursor:= crHourGlass;
      Application.ProcessMessages;

      if _SaveDialog.Execute then
           fGrid.WorkbookSource.SaveToSpreadsheetFile(_SaveDialog.FileName,true);

      if MessageDlg('Файл успешно сохранен! '+LineEnding+'Открыть файл в программе просмотра?',
            mtInformation, mbOKCancel, 0) = mrOK then
            OpenDocument(_SaveDialog.FileName);

    finally
      screen.Cursor:= crDefault;
      _SaveDialog.Free;
    end;
  except
    ShowMessage('Сохранение завершено с ошибкой!');
    raise;
  end;
end;

function TwViewer.GetWorkSheet: TsWorksheet;
begin
  Result:= fGrid.Workbook.ActiveWorksheet;
end;

procedure TwViewer.SetWorkbookSource(aValue: TsWorkbookSource);
begin
  fGrid.WorkbookSource:= aValue;
end;

constructor TwViewer.Create(TheOwner: TComponent);
begin
  inherited CreateNew(TheOwner);

   fNoClose:= true;
   self.Position:=poScreenCenter;

   //self.FormStyle:= fsStayOnTop;
   self.OnCloseQuery:=@fonCloseQuery;
   self.ShowInTaskBar:= stAlways;

   fPanelTop:= TPanel.Create(self);
   fPanelTop.Parent:= self;
   fPanelTop.Align:= alTop;
   //fPanelTop.AutoSize:= true;
   fPanelTop.BorderStyle:= bsNone;
   fPanelTop.BevelOuter:= bvNone;
   fPanelTop.Height:=0;
   fPanelTop.Visible:= false;

   fPanelCenter:= TPanel.Create(self);
   fPanelCenter.Parent:= self;
   fPanelCenter.Align:= alClient;
   fPanelCenter.BorderStyle:= bsNone;
   fPanelCenter.BevelOuter:= bvNone;

   fPanelBottom:= TPanel.Create(self);
   fPanelBottom.Parent:= self;
   fPanelBottom.Align:= alBottom;
   //fPanelBottom.AutoSize:= true;
   fPanelBottom.BorderStyle:= bsNone;
   fPanelBottom.BevelOuter:= bvNone;
   fPanelBottom.Height:=35;

   fBtnSave:= TBitBtn.Create(fPanelBottom);
   fBtnSave.Parent:= fPanelBottom;
   fBtnSave.Caption:= ' Сохранить в файл ';
   fBtnSave.AutoSize:= true;
   fBtnSave.Visible:= true;

   fBtnSave.OnClick:= @BtnSaveOnClick;

   fBtnSave.AnchorSideTop.Control:= fPanelBottom;
   fBtnSave.AnchorSideTop.Side:= asrCenter;
   //fBtnSave.AnchorSideRight.Control:= fPanelBottom;
   fBtnSave.Left:= fPanelBottom.Width-fBtnSave.Width-15;
   fBtnSave.Anchors:= [akTop, akRight];

   fGrid:= TsWorksheetGrid.Create(self);
   fGrid.Parent:= fPanelCenter;
   fGrid.Align:= alClient;
   fGrid.OnKeyDown:= @sWorksheetGridKeyDown;
   //fGrid.FixedColor:= fGrid.Color;
   //fGrid.Color:= clCream;
   //fGrid.ShowGridLines:= false;
   fGrid.Options:= fGrid.Options+ [goColMoving, goColSizing, goRowMoving, goRowSizing, goDblClickAutoSize];
   //fGrid.GridLineColor:= clBlack;

   //fWorkSheet.WriteText(1,1,'текст текст текст');
   //fWorkSheet.WriteBorders(1,1,[cbNorth, cbWest, cbEast, cbSouth]);

   self.WindowState:= wsNormal;

   SetSize(600,1024);
end;

destructor TwViewer.Destroy;
begin
  inherited Destroy;
end;

procedure TwViewer.SetStatus(const aText: string);
begin

end;

procedure TwViewer.SetSize(aHeight, aWidth: integer);
begin
  self.Height:= aHeight;
  self.Width:= aWidth;

  onSetSize();
end;

procedure TwViewer.ForceClose;
begin
  fNoClose:= false;
  Close();
end;

procedure TwViewer.onSetSize();
begin
  self.MoveToDefaultPosition;
end;

procedure TwViewer.sWorksheetGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

var
  _Zoom: Double;
begin

  if (ssCtrl in Shift) and (Key in [VK_LCL_EQUAL, VK_LCL_MINUS, VK_0..VK_9])then
  begin
    _Zoom:= TsWorksheetGrid(Sender).ZoomFactor;

    if Key = VK_LCL_EQUAL then
      _Zoom:= _Zoom+0.1;

    if Key = VK_LCL_MINUS then
      _Zoom:= _Zoom-0.1;

    if Key = VK_0 then
      _Zoom:= 1;

    if Key in [VK_1..VK_9] then
    case Key of
      VK_1: _Zoom:= 0.1;
      VK_2: _Zoom:= 0.2;
      VK_3: _Zoom:= 0.3;
      VK_4: _Zoom:= 0.4;
      VK_5: _Zoom:= 0.5;
      VK_6: _Zoom:= 0.6;
      VK_7: _Zoom:= 0.7;
      VK_8: _Zoom:= 0.8;
      VK_9: _Zoom:= 0.9;
    end;

    TsWorksheetGrid(Sender).ZoomFactor:= _Zoom;
    TsWorksheetGrid(Sender).Repaint;
  end;
end;

end.

