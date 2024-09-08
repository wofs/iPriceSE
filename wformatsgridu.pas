unit wFormatsGridU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

// Назначение:
// Работа с StringGrid. Добавляет возможность автоматического заполнения из БД, добавлять к строкам чекбоксы.

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Controls, ComCtrls, StdCtrls, fgl, Grids, Dialogs, wBaseU,
  Menus, Buttons, LazUTF8, Forms, LCLIntf, LCLType, wGetU, wTypesU, db,
  fpsutils,

  IBCustomDataSet, IBDatabase
  ;

type

//TObjectList
 TFileFormatArr = array of TFileFormat;

 TFormatTypeArr = array of TFormatType;

 { TwEditor }

 TwEditor = class

   private
     FStringGrid: TStringGrid;
     mEdit: TMemo;
     eForm: TForm;
     BtnClose: TButton;
     Col,Row: integer;
     fParse: boolean;
     procedure BtnCloseClick(Sender: TObject);

   public
     constructor Create(aOwner: TStringGrid);
     destructor Destroy;

     function Load(aStringGrid: TStringGrid; aCol, aRow: integer; aParse: boolean = false): boolean;
     procedure Save();

     property StringGrid: TStringGrid read FStringGrid write FStringGrid;
 end;

{ TStringGridData }

TStringGridData = class
private

  fCategory: integer;
  fFileType: TFileFormatArr;
  fFormatType: TFormatTypeArr;
  fFieldName: string;

  fComboBox: TComboBox;
  fCheckBox: TCheckBox;
  fBitBtn: TBitBtn;
  fBitBtn1: TBitBtn;
  fBitBtnEditor: TBitBtn;
  fMenuItem: TMenuItem;
  fMenuSpliter: TMenuItem;

public
  constructor Create();
  property FieldName: string read fFieldName write fFieldName;
  property Category: integer read fCategory write fCategory default 0;
  property FileType: TFileFormatArr read fFileType write fFileType;
  property FormatType: TFormatTypeArr read fFormatType write fFormatType;
  property ComboBox: TComboBox read fComboBox write fComboBox;
  property CheckBox: TCheckBox read fCheckBox write fCheckBox;
  property BitBtn: TBitBtn read fBitBtn write fBitBtn;
  property BitBtn1: TBitBtn read fBitBtn1 write fBitBtn1;
  property BitBtnEditor: TBitBtn read fBitBtnEditor write fBitBtnEditor;
  property MenuItem: TMenuItem read fMenuItem write fMenuItem;
  property MenuSpliter: TMenuItem read fMenuSpliter write fMenuSpliter;

  function ComboBoxCreate(aParent: TComponent; const aHint: string = ''): TComboBox;
  function CheckBoxCreate(aParent: TComponent; const aHint: string = ''): TCheckBox;
  function MenuItemCreate(aParent: TPopupMenu; const aFieldName:string; const aCaption: string = ''; const aImageIndex:integer = 0): TMenuItem;
  function MenuSpliterCreate(aParent: TPopupMenu): TMenuItem;
  function BitBtnCreate(aParent: TComponent; const aCaption: string = ''; const aHint: string = ''): TBitBtn;

end;

{ TwFormatsGrid }

TwFormatsGrid = class
private
  fFormatID: Integer;
  FFormatType: TFormatType;
  fFileType: TFileFormat;
  FonChangedFormat: TNotifyEvent;
  FonFillGrid: TNotifyEvent;
  FonSavedFormat: TNotifyEvent;
  FxGridCurrentColRow: ArrayOfInteger;
  fxGridCurrentColRowAddr: string;
  //FxGridWorksheet: integer;
  wFormatCategory: integer;
  wFillGridOn: boolean;

  fBase: TwBase;
  fdbDataSet: TIBDataSet;
  fdbDataSource: TDataSource;
  fdbTransaction: TIBTransaction;

  wRow, wCol: integer;
  // PopupMenu
  sgPopupMenu:TPopupMenu;
  fxGridPopupMenu:TPopupMenu;
  sgPopupMenuItem :TMenuItem;
  fForm: TObject;
  fFormName: string;

  _StringGrid : TStringGrid;

  cbx_FormatType: TComboBox;
  fTabCategory: TTabControl;
  fMasterMode: boolean;

  wEditor: TwEditor;

  function FileFormatInArr(aFileFormat: TFileFormat; aFileFormatArr: TFileFormatArr): boolean;
  function FormatTypeInArr(aFormatType: TFormatType; aFormatTypeArr: TFormatTypeArr): boolean;
  function FindMenuItemByName(aName: string; aPopupMenu: TPopupMenu): TMenuItem;
  function GetComboBoxCodePage: TComboBox;
  function GetComboBoxCSVDelimiter: TComboBox;
  function GetComboBoxFileFormat: TComboBox;
  function GetxGridWorksheet: string;
  procedure LoadFile();
  procedure Log(_Text:string);
  procedure SelectCellByName(aText: string; aCol, aRow: integer);
  procedure SetStatus(_Text:string); // вывод статуса
  procedure SetTabCategory(AValue: TTabControl);
  procedure SetxGridWorksheet(AValue: string);
  procedure StringGridClear();

  procedure _onDrawCell(Sender: TObject; aCol, aRow: Integer; aRect: TRect; aState: TGridDrawState);
  procedure _onTopLeftChanged(Sender: TObject);
  procedure _on_BtnEditorClick(Sender: TObject);
  procedure _on_sgValidateEntry(sender: TObject; aCol, aRow: Integer; const OldValue: string; var NewValue: String);
  procedure _on_sgGetEditText(Sender: TObject; ACol, ARow: Integer; var Value: string);
  procedure _on_sgPopupMenuClearClick(Sender: TObject);
  procedure _on_sgPopupMenuSelectColumnClick(Sender: TObject);
  procedure _on_sgPopupMenuClearAllClick(Sender: TObject);
  procedure _on_BtnURLMasterClick(Sender: TObject);
  procedure _on_BtnURLClear(Sender: TObject);
  procedure _on_CheckBoxChange(Sender: TObject);
  procedure _on_ComboBoxChange(Sender: TObject);
  procedure _on_cbx_FormatTypeChange(Sender: TObject);
  procedure _on_xGridMenuItemClick(Sender: TObject);
  procedure _onMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
  procedure cbx_IdFileFormat_onChange(Sender: TObject);
  procedure on_TabCategotyChange(Sender: TObject);

public

  constructor Create(Sender:TObject; _StGrid: TStringGrid; aFormatType: TComboBox; const _MenuON: boolean = false; const aBase: TwBase = nil);
  destructor Destroy();

  procedure FillGrid(const aFormatID: integer; const Sender: TObject = nil); // заполнение грида
  procedure ClearCell(aCol, aRow: integer); // очистка ячеек
  procedure ClearAll; // очистка всего грида

  procedure ChangeView(const aFormat: string='xls');
  procedure Save(Sender: TObject);
  procedure SetGridCellValue(aFieldName: string; aValue: boolean);
  procedure SetGridCellValue(aFieldName: string; aValue: integer);
  procedure SetGridCellValue(aFieldName: string; aValue: string);
  function FormatCopy(aID: integer): integer;

  function GetCellCheckBox(aFieldName: string):boolean;
  function GetCellComboBox(aFieldName: string):integer;
  function GetCellVal(aFieldName: string):string;

  property Row: integer read wRow write wRow;
  property Col: integer read wCol write wCol;

  property xGridPopupMenu: TPopupMenu read fxGridPopupMenu write fxGridPopupMenu;
  property StringGrid: TStringGrid read _StringGrid write _StringGrid;
  property FormatCategory: integer read wFormatCategory write wFormatCategory;
  property TabCategory: TTabControl read fTabCategory write SetTabCategory;
  property FormatType: TFormatType read FFormatType write FFormatType;
  property xGridCurrentColRow: ArrayOfInteger read FxGridCurrentColRow write FxGridCurrentColRow;
  property xGridCurrentColRowAddr: string read fxGridCurrentColRowAddr write fxGridCurrentColRowAddr;
  property FormatID: integer read fFormatID;
  property onChangedFormat: TNotifyEvent read FonChangedFormat write FonChangedFormat;
  property onFillGrid: TNotifyEvent read FonFillGrid write FonFillGrid;
  property onSavedFormat: TNotifyEvent read FonSavedFormat write FonSavedFormat;
  property MasterMode: boolean read fMasterMode write fMasterMode default false;
  property ComboBoxFileFormat: TComboBox read GetComboBoxFileFormat;
  property ComboBoxCodePage: TComboBox read GetComboBoxCodePage;
  property ComboBoxCSVDelimiter: TComboBox read GetComboBoxCSVDelimiter;
  property FillGridOn: boolean read wFillGridOn write wFillGridOn;


end;


implementation
uses
  wLogU, wFuncU, FmURLsU, FmMasterU, FmWaitU;

{ TwEditor }

procedure TwEditor.BtnCloseClick(Sender: TObject);
begin
  Save();
  eForm.ModalResult:= mrOK;
  eForm.Close;
end;

constructor TwEditor.Create(aOwner: TStringGrid);
begin
    //if (aOwner is TEdit) then Edit:= TEdit(aOwner);

    eForm:= TForm.Create(aOwner);
    eForm.Visible:= false;

    fParse:= true;

    mEdit:= TMemo.Create(eForm);
    mEdit.Parent:= eForm;
    mEdit.ScrollBars:= ssAutoBoth;
    mEdit.Width:= 250;
    mEdit.Clear;
    mEdit.Lines.Delimiter := ',';

    BtnClose:= TButton.Create(eForm);
    BtnClose.Parent:= eForm;
    BtnClose.Width:=mEdit.Width;
    BtnClose.OnClick:=@BtnCloseClick;
    BtnClose.ModalResult:= mrOK;

    BtnClose.Align:= alBottom;
    BtnClose.Caption:= 'Сохранить и закрыть';
    mEdit.Align:= alClient;
end;

destructor TwEditor.Destroy;
begin
  eForm.Free;
end;

function TwEditor.Load(aStringGrid: TStringGrid; aCol, aRow: integer; aParse: boolean): boolean;
begin
  Col:= aCol;
  Row:= aRow;
  StringGrid:= aStringGrid;

  mEdit.Clear;
  mEdit.Lines.StrictDelimiter := True;
  mEdit.Lines.Delimiter:=',';
  fParse:= aParse;

  if fParse then
    mEdit.Lines.DelimitedText:= StringGrid.Cells[Col,Row]
  else
    mEdit.Lines.Text:= StringGrid.Cells[Col,Row];

  eForm.Position:= poScreenCenter;
  eForm.Caption:= 'Редактирование';

  eForm.ShowModal;

  //if eForm.ModalResult = mrOK then
  Result:= true;
  //else
  //  Result:= false;
end;

procedure TwEditor.Save();
var
  i: Integer;
  _Result: string;
begin
  _Result:= '';

  if fParse then
    begin
      for i:=0 to mEdit.Lines.Count-1 do
        if i>0 then
           _Result:= _Result+','+mEdit.Lines[i] else
             _Result:= mEdit.Lines[i];
    end else
      _Result:= mEdit.Lines.Text;

  StringGrid.Cells[Col,Row]:= _Result;

end;

{ TStringGridData }

constructor TStringGridData.Create();
begin
   fComboBox:=nil;
   fCheckBox:= nil;
   fBitBtn:= nil;
end;

function TStringGridData.ComboBoxCreate(aParent: TComponent; const aHint: string): TComboBox;
begin
   Result:= TComboBox.Create(aParent);

   with Result do begin
     Parent:= TWinControl(aParent);
     Style:= csDropDownList;
     Hint:= aHint;
     if Length(aHint)>0 then
         ShowHint:= true;
     Visible:= false;
   end;
end;

function TStringGridData.CheckBoxCreate(aParent: TComponent; const aHint: string): TCheckBox;
begin
  Result:= TCheckBox.Create(aParent);

  with Result do begin
   Parent:= TWinControl(aParent);
   Hint:= aHint;
     if Length(aHint)>0 then
         ShowHint:= true;
     //ParentColor:= false;
     Visible:= false;
  end;

end;

function TStringGridData.MenuItemCreate(aParent: TPopupMenu; const aFieldName: string; const aCaption: string; const aImageIndex: integer): TMenuItem;
begin
  Result:= TMenuItem.Create(aParent);
  Result.Caption:= aCaption;
  Result.ImageIndex:= aImageIndex;
  Result.Name:= aFieldName;
end;

function TStringGridData.MenuSpliterCreate(aParent: TPopupMenu): TMenuItem;
begin
  Result:= TMenuItem.Create(aParent);
  Result.Caption:= '-';
  aParent.Items.Add(Result);
end;

function TStringGridData.BitBtnCreate(aParent: TComponent; const aCaption: string; const aHint: string): TBitBtn;
begin
  Result:= TBitBtn.Create(aParent);
  Result.Parent:= TWinControl(aParent);
  Result.Caption:= aCaption;
  Result.Hint:= aHint;
  if Length(aHint)>0 then
      Result.ShowHint:= true;

  Result.Visible:= false;
end;

{ TStringGridData }

{ TwFormatsGrid }

procedure TwFormatsGrid.Log(_Text: string);
begin
  // здесь напишите процедуру ведения лог-файла
  // если вы не ведете лог-файл, то оставьте тело функции пустым
 wLog('['+fFormName+']'+'['+StringGrid.Name+'] '+'[wTab] ',_Text);
end;

procedure TwFormatsGrid.SetStatus(_Text: string);
begin
 wStatus(fFormName,_Text,true);
end;

procedure TwFormatsGrid.SetTabCategory(AValue: TTabControl);
begin
  if fTabCategory=AValue then Exit;
  fTabCategory:=AValue;
  fTabCategory.OnChange:=@on_TabCategotyChange;
end;

procedure TwFormatsGrid.SetxGridWorksheet(AValue: string);
var
  i: Integer;
  _Obj: TStringGridData;
begin

  for i:=0 to StringGrid.RowCount-1 do
  begin
    _Obj:= TStringGridData(StringGrid.Objects[1,i]);
    if Assigned(_Obj) then
      if _Obj.FieldName = 'SPREADSHEET' then
        begin
            StringGrid.Cells[1,i]:= AValue;
            Break;
        end;

  end;

end;

procedure TwFormatsGrid._onDrawCell(Sender: TObject; aCol, aRow: Integer;
  aRect: TRect; aState: TGridDrawState);
var
  _obj: TStringGridData;
  VertScroolWidth: Integer;
  // определение статуса и размера скрулбаров
  function ScrollIsVisible(Handle : HWnd; Style : Longint) : Boolean;
  begin
     Result := (GetWindowLong(Handle, GWL_STYLE) and Style) <> 0;
  end;

  function GetSizeVertScrool(Grid_Handle: THandle): Integer;
  begin
     if ScrollIsVisible(Grid_Handle, WS_VSCROLL) then
        Result := GetSystemMetrics(SM_CXVSCROLL)
     else
        Result := 0;
  end;

  function GetSizeHorizScrool(Grid_Handle: THandle): Integer;
  begin
    if ScrollIsVisible(Grid_Handle, WS_HSCROLL) then
       Result := GetSystemMetrics(SM_CXHSCROLL)
    else
       Result := 0;
   end;

begin

_obj:= TStringGridData(StringGrid.Objects[aCol,aRow]);

VertScroolWidth:= GetSizeVertScrool(StringGrid.Handle);

  if Assigned(_obj) then
  begin
     if  Assigned(_obj.ComboBox) then
     begin
        _obj.ComboBox.Width:=Arect.Width-5;
        _obj.ComboBox.Top:=arect.Top+2;
        _obj.ComboBox.Left:=arect.Left+2;
        _obj.ComboBox.Visible:= true;
     end;

     if  Assigned(_obj.CheckBox) then
     begin
       _obj.CheckBox.Top:=arect.Top+2;
       _obj.CheckBox.Left:=arect.Left+50;
       _obj.CheckBox.Visible:= true;
     end;

     if  Assigned(_obj.BitBtn) then
     begin
       if Assigned(_obj.BitBtn1) then
       begin
         _obj.BitBtn.Height:=25;
         _obj.BitBtn.Width:=Arect.Width-30;
         _obj.BitBtn.Top:=arect.Top+1;
         _obj.BitBtn.Left:=arect.Left+2;
         _obj.BitBtn.Visible:=True;

         _obj.BitBtn1.Height:=25;
         _obj.BitBtn1.Width:=25;
         _obj.BitBtn1.Top:=arect.Top+1;
         _obj.BitBtn1.Left:=arect.Left+_obj.BitBtn.Width+5;
         _obj.BitBtn1.Visible:=True;
       end else
       begin
         _obj.BitBtn.Height:=25;
         _obj.BitBtn.Width:=Arect.Width-4;
         _obj.BitBtn.Top:=arect.Top+1;
         _obj.BitBtn.Left:=arect.Left+2;
         _obj.BitBtn.Visible:=True;
       end;
     end;

    if Assigned(_obj.BitBtnEditor) then
    begin
      _obj.BitBtnEditor.Height:=25;
      _obj.BitBtnEditor.Width:=25;
      _obj.BitBtnEditor.Top:=arect.Top+1;
      _obj.BitBtnEditor.Left:=StringGrid.Left+StringGrid.Width-_obj.BitBtnEditor.Width-VertScroolWidth-5;
      _obj.BitBtnEditor.Visible:= true;
    end;
  end;

StringGrid.DefaultDrawCell(aCol,aRow,aRect,aState);
end;

procedure TwFormatsGrid._onTopLeftChanged(Sender: TObject);
var
    i:integer;
    _obj: TStringGridData;
begin

   for i:=TStringGrid(Sender).VisibleRowCount to TStringGrid(Sender).rowcount-1 do
   begin
     _obj:= TStringGridData(TStringGrid(Sender).Objects[1,i]);

       if Assigned(_obj) then
       begin
          if  Assigned(_obj.ComboBox) then
              _obj.ComboBox.Visible:= false;

          if  Assigned(_obj.CheckBox) then
              _obj.CheckBox.Visible:= false;

          if  Assigned(_obj.BitBtn) then
              _obj.BitBtn.Visible:= false;

          if  Assigned(_obj.BitBtn1) then
              _obj.BitBtn1.Visible:= false;

          if  Assigned(_obj.BitBtnEditor) then
              _obj.BitBtnEditor.Visible:= false;

       end;
   end;

end;

procedure TwFormatsGrid._on_BtnEditorClick(Sender: TObject);
var
  _Obj: TStringGridData;
  pnt: TPoint;
  fCol, fRow: Integer;
  fResult: boolean;
begin
  pnt:= Mouse.CursorPos;
  pnt:= StringGrid.ScreenToClient(pnt);
  StringGrid.MouseToCell(pnt.X, pnt.Y, fCol, fRow);


  _Obj:= TStringGridData(StringGrid.Objects[fCol,fRow]);
  if Assigned(_Obj) then
  begin
     if (_Obj.FieldName = 'SPREADSHEET') or (_Obj.FieldName = 'STOCKSYMBOLS') then
         fResult:= wEditor.Load(StringGrid,fCol,fRow,true);

     if _Obj.FieldName = 'REMARK' then
         fResult:= wEditor.Load(StringGrid,fCol,fRow, false);
  end;

  if fResult and Assigned(onChangedFormat) then onChangedFormat(Sender);

end;

procedure TwFormatsGrid._on_sgValidateEntry(sender: TObject; aCol, aRow: Integer; const OldValue: string; var NewValue: String);
var
  _Obj: TStringGridData;
  i: Longint;
begin
   if OldValue <> NewValue then
   begin
     _Obj:= TStringGridData(TStringGrid(Sender).Objects[aCol,aRow]);
     if Assigned(_Obj) then
     begin
       if  Assigned(_obj.ComboBox) or Assigned(_obj.CheckBox) or Assigned(_obj.BitBtn) or Assigned(_obj.BitBtn1) then
           begin
             NewValue:='';
             exit;
           end;

         if (_Obj.FieldName <> 'FILE') and
            (_Obj.FieldName <> 'SPREADSHEET') and
            (_Obj.FieldName <> 'STOCKSYMBOLS') and
            (_Obj.FieldName <> 'REMARK') and
            (_Obj.FieldName <> 'OUTCELLTEXT') and
            (_Obj.FieldName <> 'ADDRCELLTEXT')  then
           if not TryStrToInt(NewValue,i) then NewValue:= OldValue;
     end;

     onChangedFormat(self);
   end;

end;

procedure TwFormatsGrid._on_sgGetEditText(Sender: TObject; ACol, ARow: Integer; var Value: string);
begin

end;


procedure TwFormatsGrid._on_sgPopupMenuClearClick(Sender: TObject);
begin
  ClearCell(Col, Row);
end;

procedure TwFormatsGrid._on_sgPopupMenuSelectColumnClick(Sender: TObject);
begin
  SelectCellByName(ReplaceStr(TMenuItem(Sender).Name,'m',''),Col,Row);
end;

procedure TwFormatsGrid._on_sgPopupMenuClearAllClick(Sender: TObject);
begin
  ClearAll();
end;

procedure TwFormatsGrid.LoadFile();
var
  _XMLText: string;
  wGet: TwGet;
  _arr: ArrayArrayOfString;
  i: Integer;
  _Pos: PtrInt;
  _Obj: TStringGridData;
begin
_XMLText:= fBase.SQLReadArr('FORMATS',['URL'],'ID='+IntToStr(StringGrid.Tag),'')[0,0];

if Length(_XMLText)=0 then
   begin
     ShowMessage('Для загрузки файла сперва укажите ссылку в мастере!');
     exit;
   end;

wGet:= TwGet.Create(StringGrid);

try
  _arr:= wGet.ExecuteXML(_XMLText);
finally
  wGet.Destroy();
end;

for i:=0 to High(_arr) do
begin
  _Pos:= UTF8Pos('File loaded in ',_arr[i,1]);
  if _Pos>0 then
     begin
       _arr[i,1]:= StringReplace(_arr[i,1],'File loaded in ','',[]);
        SetGridCellValue('FILE',UnsafePath(PathApplication_Unsafe,_arr[i,1]));
        TFmMaster(fForm).OpenFile(_arr[i,1]);
       Break;
     end;
end;
end;

procedure TwFormatsGrid._on_BtnURLMasterClick(Sender: TObject);
var
  _Form: TFmURLs;
begin
  try
    _Form:= nil;
    if MessageDlg('Перед запуском мастера формат будет автоматически сохранен - продолжить?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

    fBase.SQLTransactionEnd(true);
    if Assigned(onSavedFormat) then onSavedFormat(self);

    _Form:= TFmURLs.Create(StringGrid);

    _Form.Base:= fBase;
    _Form.FormatID:= StringGrid.Tag;
    _Form.Caption:='Мастер загрузки прайс-листа из сети';
    try
      _Form.XMLText:= fBase.SQLReadArr('FORMATS',['URL'],'ID='+IntToStr(StringGrid.Tag),'')[0,0];
      _Form.ShowModal;
    finally
      if _Form.ModalResult = mrOK then
         begin
           //TFmMaster(fForm).Repaint;
           if Length(fBase.SQLReadArr('FORMATS',['URL'],'ID='+IntToStr(StringGrid.Tag),'')[0,0])>0 then
              begin
                StringGrid.TitleImageList.GetBitmap(0,TBitBtn(Sender).Glyph);
                if fMasterMode then
                   begin
                      ShowMessage('Сейчас будет выполнена загрузка файла и отображение его в мастере формата.');
                      screen.Cursor:= crHourGlass;
                      LoadFile();
                      screen.Cursor:= crDefault;
                   end;
              end
               else
                 StringGrid.TitleImageList.GetBitmap(1,TBitBtn(Sender).Glyph);

         //onChangedFormat(self);
         end;
       _Form.Free;
    end;
  except
    on E: Exception do
     begin
       screen.Cursor:= crDefault;
       __Log.SaveLogError(E);
       wLog('wStringGrid','Ошибка [_on_BtnURLMasterClick]: "' + E.Message + '"');
       raise;
     end;
  end;
end;

procedure TwFormatsGrid._on_BtnURLClear(Sender: TObject);
var
  _obj: TStringGridData;
  i: Integer;
begin
  if MessageDlg('Очистить ссылку на прайс-лист?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

  if Assigned(onChangedFormat) then onChangedFormat(self);

  fBase.SQLUpdate('FORMATS',['URL'],[''],'ID='+IntTOStr(StringGrid.Tag),false);

  for i:=0 to StringGrid.RowCount-1 do
  begin
    _obj:= TStringGridData(StringGrid.Objects[1,i]);
    if Assigned(_Obj) then
      if _Obj.FieldName = 'URL' then StringGrid.TitleImageList.GetBitmap(1,_Obj.BitBtn.Glyph);
  end;
end;

procedure TwFormatsGrid._on_CheckBoxChange(Sender: TObject);
begin
  if Assigned(onChangedFormat) then onChangedFormat(self);
  on_TabCategotyChange(self);
end;

procedure TwFormatsGrid._on_ComboBoxChange(Sender: TObject);
begin
  if (TComboBox(Sender).Name = 'IDCURRENCY') and (fFormatType = ftPRICE) then
    begin
        if MessageDlg('Изменить валюту прайс-листа? Это приведет к очистке цен и архива ранее импортированного прайс-листа! Выставленные соответствия останутся без изменений. Для занесения цен по новому курсу валют - просто обновите прайс-лист.',mtWarning, mbOKCancel, 0) = mrCancel then
        begin
          TComboBox(Sender).ItemIndex:=cmbxItemIndexByID(TComboBox(Sender),fdbDataSet.FieldByName('IDCURRENCY').AsInteger);
          exit;
        end else
        begin
          try
            if Assigned(onChangedFormat) then onChangedFormat(self);
            Screen.Cursor:= crSQLWait;
            try
              fBase.SQLUpdate('PL_ITEMS',['PRICE','PRICE2','PRICE3','PRICE4','PRICE5','PRICE6','PRICE7','PRICE8','PRICE9','PRICE10','PRICECALC','PRICECALC2','PRICECALC3','PRICECALC4','PRICECALC5','PRICECALC6','PRICECALC7','PRICECALC8','PRICECALC9','PRICECALC10'],[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],'IDFORMATS='+IntToStr(FormatID),false);
              fBase.SQLDelete('PRICELISTS_TIMESTAMPS','IDFORMATS='+IntToStr(FormatID),false);
              fBase.SQLDelete('PL_VERSIONS','IDFORMATS='+IntToStr(FormatID),false);
            finally
              Screen.Cursor:= crDefault;
            end;

            ShowMessage('Прайс-лист успешно изменен! Для принятия изменений сохраните формат.');
          except
            on E: Exception do
             begin
               __Log.SaveLogError(E);
               wLog('wStringGrid','Ошибка [Create]: "' + E.Message + '"');
               raise;
             end;
          end;
        end;

    end else
      if Assigned(onChangedFormat) then onChangedFormat(self);

  on_TabCategotyChange(self);
end;

procedure TwFormatsGrid._on_cbx_FormatTypeChange(Sender: TObject);
begin
  FillGrid(fFormatID,Sender);
  if Assigned(onChangedFormat) then onChangedFormat(self);
end;

procedure TwFormatsGrid._on_xGridMenuItemClick(Sender: TObject);
var
  _Obj: TStringGridData;
  i: Integer;
begin
  for i:=0 to StringGrid.RowCount-1 do
  begin
    _Obj:= TStringGridData(StringGrid.Objects[1,i]);
    if Assigned(_Obj) then
      if _Obj.FieldName = TMenuItem(Sender).Name then
        begin

          case _Obj.FieldName of
            'FIRSTLINE'   :
                          begin
                            StringGrid.Cells[1,i]:= IntToStr(xGridCurrentColRow[1]);
                            TMenuItem(Sender).Enabled:= false;
                          end;
            'SPREADSHEET' :
                          begin
                           if Length(StringGrid.Cells[1,i])>0 then
                             StringGrid.Cells[1,i]:= StringGrid.Cells[1,i]+',';
                             StringGrid.Cells[1,i]:= StringGrid.Cells[1,i]+IntToStr(xGridCurrentColRow[2]+1);
                             TMenuItem(Sender).Caption:= 'Выбраные листы: '+StringGrid.Cells[1,i];
                          end;
            'ADDRCELLTEXT'   :
                          begin
                            StringGrid.Cells[1,i]:= xGridCurrentColRowAddr;
                            TMenuItem(Sender).Enabled:= false;
                          end;
            else
             begin
               StringGrid.Cells[1,i]:= IntToStr(xGridCurrentColRow[0]);
               TMenuItem(Sender).Enabled:= false;
             end;
          end;
            Break;
        end;

  end;

  onChangedFormat(self);

end;

procedure TwFormatsGrid._onMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  _Col, _Row: Longint;
begin
  //if Button = mbRight then
  //begin
    (Sender as TStringGrid).MouseToCell(X, Y, _Col, _Row);
    Col:= _Col;
    Row:= _Row;
    (Sender as TStringGrid).Selection := TGridRect(rect(Col, Row, Col, Row));
//  end;
end;

procedure TwFormatsGrid.cbx_IdFileFormat_onChange(Sender: TObject);
begin
   ChangeView(fBase.SQLReadArr('Select Lower(CODE) from "FILEFORMAT" WHERE ID='+IntTOStr(cmbxSelectID(TComboBox(Sender))))[0,0]);
   if Assigned(onChangedFormat) then onChangedFormat(self);
end;

function TwFormatsGrid.FileFormatInArr(aFileFormat: TFileFormat; aFileFormatArr: TFileFormatArr): boolean;
var
  i: Integer;
begin
 Result:= false;
 for i:=0 to High(aFileFormatArr) do
     if aFileFormat = aFileFormatArr[i] then
        begin
          Result:= true;
          Break;
        end;
end;

function TwFormatsGrid.FormatTypeInArr(aFormatType: TFormatType; aFormatTypeArr: TFormatTypeArr): boolean;
var
  i: Integer;
begin
 Result:= false;
 for i:=0 to High(aFormatTypeArr) do
     if aFormatType = aFormatTypeArr[i] then
        begin
          Result:= true;
          Break;
        end;

end;

function TwFormatsGrid.FindMenuItemByName(aName: string; aPopupMenu: TPopupMenu):TMenuItem;
var
  i: Integer;
begin
  Result:= nil;
  if Assigned(aPopupMenu) then
    for i:=0 to aPopupMenu.Items.Count-1 do
        if aPopupMenu.Items[i].Name = aName then
          begin
            Result:= aPopupMenu.Items[i];
            break;
          end;
end;

function TwFormatsGrid.GetComboBoxCodePage: TComboBox;
var
  i: Integer;
  _Obj: TStringGridData;
begin
for i:=0 to StringGrid.RowCount-1 do
begin
  _Obj:= TStringGridData(StringGrid.Objects[1,i]);
  if Assigned(_Obj) then
    if _Obj.FieldName = 'IDCODEPAGETEXT' then
      begin
          Result:=_Obj.ComboBox;
          Break;
      end;
end;
end;

function TwFormatsGrid.GetComboBoxCSVDelimiter: TComboBox;
var
  i: Integer;
  _Obj: TStringGridData;
begin
for i:=0 to StringGrid.RowCount-1 do
begin
  _Obj:= TStringGridData(StringGrid.Objects[1,i]);
  if Assigned(_Obj) then
    if _Obj.FieldName = 'IDCSVDELIMITER' then
      begin
          Result:=_Obj.ComboBox;
          Break;
      end;
end;
end;

function TwFormatsGrid.GetComboBoxFileFormat: TComboBox;
var
  i: Integer;
  _Obj: TStringGridData;
begin
for i:=0 to StringGrid.RowCount-1 do
begin
  _Obj:= TStringGridData(StringGrid.Objects[1,i]);
  if Assigned(_Obj) then
    if _Obj.FieldName = 'IDFILEFORMAT' then
      begin
          Result:=_Obj.ComboBox;
          Break;
      end;
end;
end;

function TwFormatsGrid.GetxGridWorksheet: string;
var
  i: Integer;
  _Obj: TStringGridData;
begin
for i:=0 to StringGrid.RowCount-1 do
begin
  _Obj:= TStringGridData(StringGrid.Objects[1,i]);
  if Assigned(_Obj) then
    if _Obj.FieldName = 'SPREADSHEET' then
      begin
          Result:= StringGrid.Cells[1,i];
          Break;
      end;

end;
end;

procedure TwFormatsGrid.on_TabCategotyChange(Sender: TObject);
var
  _Category, i: Integer;
  _obj: TStringGridData;
  _GroupInRows, _CurrencyRUR: boolean;
  _MenuItem: TMenuItem;
begin
  _Category:= fTabCategory.TabIndex;
  _GroupInRows:= false;
  _CurrencyRUR:= false;

  StringGrid.BeginUpdate;

  for i:=0 to StringGrid.RowCount-1 do
  begin
    _obj:= TStringGridData(StringGrid.Objects[1,i]);

    //FormatTypeInArr сравнение форматов и корректировка списка видимых строк

    if Assigned(_obj) then
    begin

      if (_obj.Category=_Category) and FileFormatInArr(fFileType,_Obj.FileType) and FormatTypeInArr(fFormatType,_Obj.FormatType) then
      begin

        if Assigned(_Obj.BitBtn) then
           StringGrid.RowHeights[i]:= _Obj.BitBtn.Height+2
        else
          if Assigned(_Obj.BitBtn1) then
             StringGrid.RowHeights[i]:= _Obj.BitBtn1.Height+2
          else
            if Assigned(_Obj.BitBtnEditor) then
               StringGrid.RowHeights[i]:= _Obj.BitBtnEditor.Height+2
            else
              if Assigned(_Obj.ComboBox) then
                 StringGrid.RowHeights[i]:= _Obj.ComboBox.Height+6
              else
                if Assigned(_Obj.CheckBox) then
                   StringGrid.RowHeights[i]:= _Obj.CheckBox.Height+6
                else
                   StringGrid.RowHeights[i]:= 22;

        //if Assigned(_Obj.BitBtn) or Assigned(_Obj.BitBtn1) or Assigned(_Obj.BitBtnEditor) then
        // StringGrid.RowHeights[i]:= 28 else
        // StringGrid.RowHeights[i]:= 22;


        if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
        begin
           _MenuItem:= FindMenuItemByName(_Obj.FieldName, xGridPopupMenu);
           if Assigned(_MenuItem) then
             _MenuItem.Visible:= true;
        end;

        if _Obj.FieldName = '' then
           if not (fFileType in [ffXLS,ffXLSX]) then
                 _Obj.CheckBox.Checked:= false;

        if _Obj.FieldName = 'GROUPSINROWS' then
           _GroupInRows:= _Obj.CheckBox.Checked;

        if  (_Obj.FieldName = 'GROUPALGORITHM') and not _GroupInRows then
        begin
           StringGrid.RowHeights[i]:= 0;
           _Obj.ComboBox.Visible:= false;
        end;

        if  (_Obj.FieldName = 'IDCURRENCY') then
                if (cmbxSelectID(_Obj.ComboBox)=1) then
                    _CurrencyRUR:= true else _CurrencyRUR:= false;

        if  (_Obj.FieldName = 'CURRENCYPERCENT') and _CurrencyRUR then
            begin
              StringGrid.Cells[1,i]:='0';
              StringGrid.RowHeights[i]:= 0;
            end;

        if  (_Obj.FieldName = 'SUBGROUPS1') and _GroupInRows then
          begin
             StringGrid.RowHeights[i]:= 0;
             StringGrid.Cells[1,i]:='0';
             if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
               FindMenuItemByName(_Obj.FieldName, xGridPopupMenu).Visible:= false;
          end;


        if  (_Obj.FieldName = 'SUBGROUPS2') and _GroupInRows then
          begin
             StringGrid.RowHeights[i]:= 0;
             StringGrid.Cells[1,i]:='0';
            if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
               FindMenuItemByName(_Obj.FieldName, xGridPopupMenu).Visible:= false;
          end;


        if  (_Obj.FieldName = 'SUBGROUPS3') and _GroupInRows then
          begin
             StringGrid.RowHeights[i]:= 0;
             StringGrid.Cells[1,i]:='0';
            if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
               FindMenuItemByName(_Obj.FieldName, xGridPopupMenu).Visible:= false;
          end;

      end
      else
         begin
           StringGrid.RowHeights[i]:= 0;
           if Assigned(_Obj.ComboBox) then
              _Obj.ComboBox.Visible:= false;

           if Assigned(_Obj.CheckBox) then
              _Obj.CheckBox.Visible:= false;

           if Assigned(_Obj.BitBtn) then
              _Obj.BitBtn.Visible:= false;

           if Assigned(_Obj.BitBtn1) then
              _Obj.BitBtn1.Visible:= false;

           if Assigned(_Obj.BitBtnEditor) then
              _Obj.BitBtnEditor.Visible:= false;
         end;

    end;
  end;

  StringGrid.EndUpdate();
end;

constructor TwFormatsGrid.Create(Sender: TObject; _StGrid: TStringGrid; aFormatType: TComboBox; const _MenuON: boolean; const aBase: TwBase);
var
  i: Integer;
  _NewItem: TMenuItem;
begin
       fBase:=aBase;
       fdbDataSet:= TIBDataSet.Create(_StGrid);
       fdbDataSource:= TDataSource.Create(_StGrid);
       fdbTransaction:= TIBTransaction.Create(_StGrid);
       fxGridPopupMenu:= nil;
       with fdbTransaction do
       begin
         DefaultDatabase := fBase.DataBase;
         Params.Add('read');
         Params.Add('read_committed');
         Params.Add('rec_version');
         Params.Add('nowait');
       end;

       with fdbDataSet do
       begin
         Database := fBase.DataBase;
         AutoCalcFields := False;
       end;
       fdbDataSource.DataSet:= fdbDataSet;


       fForm:= Sender;
     try
       fFormName:=TComponent(Sender).Name;
       StringGrid:= _StGrid;
       cbx_FormatType:= aFormatType;
       cbx_FormatType.OnChange:=@_on_cbx_FormatTypeChange;
       fTabCategory:= nil;

       FillGridOn:= false;

       wEditor:= TwEditor.Create(StringGrid);

     if (StringGrid.PopupMenu = nil) then
     begin
        sgPopupMenu:= TPopupMenu.Create(StringGrid);

        if _MenuON then
        begin
          if StringGrid.TitleImageList<> nil then
             sgPopupMenu.Images:= StringGrid.TitleImageList;

          sgPopupMenuItem:= TMenuItem.Create(sgPopupMenu);
          sgPopupMenuItem.Caption:= 'Очистить';
          sgPopupMenuItem.OnClick:=@_on_sgPopupMenuClearClick;
          sgPopupMenu.Items.Add(sgPopupMenuItem);

          sgPopupMenuItem:= TMenuItem.Create(sgPopupMenu);
          sgPopupMenuItem.Caption:= 'Очистить все';
          sgPopupMenuItem.OnClick:=@_on_sgPopupMenuClearAllClick;
          sgPopupMenu.Items.Add(sgPopupMenuItem);


          sgPopupMenuItem:= TMenuItem.Create(sgPopupMenu);
          sgPopupMenuItem.Caption:= '-';
          sgPopupMenu.Items.Add(sgPopupMenuItem);


          sgPopupMenuItem:= TMenuItem.Create(sgPopupMenu);
          sgPopupMenuItem.Caption:= 'Выбрать колонку';
          sgPopupMenu.Items.Add(sgPopupMenuItem);

          for i:=0 to 255 do
          begin
            _NewItem:= NewItem(GetColString(i)+' ('+IntToStr(i+1)+')', 0, False, True, @_on_sgPopupMenuSelectColumnClick, 0, 'm'+IntToStr(i+1));
            _NewItem.ImageIndex:=9;
              sgPopupMenuItem.Add(
                  _NewItem
              );
          end;

{
function GetColString(AColIndex: Integer): String;
function ParseCellString(const AStr: String; out ACellRow, ACellCol: Cardinal;
  out AFlags: TsRelFlags): Boolean;
}
          if sgPopupMenu.Images<> nil then
          begin
            sgPopupMenu.Items[0].ImageIndex:=1;
            sgPopupMenu.Items[1].ImageIndex:=2;
          end;
        end;

        StringGrid.PopupMenu:=sgPopupMenu;
     end;

     if StringGrid.OnDrawCell = nil then
        StringGrid.OnDrawCell:=@_onDrawCell;
     if StringGrid.OnMouseDown = nil then
        StringGrid.OnMouseDown:=@_onMouseDown;

     StringGrid.OnTopLeftChanged:=@_onTopLeftChanged;
     StringGrid.OnValidateEntry:=@_on_sgValidateEntry;

     except
       on E: Exception do
        begin
          __Log.SaveLogError(E);
          wLog('wStringGrid','Ошибка [Create]: "' + E.Message + '"');
          raise;
        end;
     end;
end;

procedure TwFormatsGrid.StringGridClear();
var
  i: integer;
  _Obj: TObject;
begin
  if Assigned(StringGrid) then
  begin
    if Assigned(xGridPopupMenu) then xGridPopupMenu.Items.Clear;

    //StringGrid.BeginUpdate;

    for i:=0 to StringGrid.RowCount-1 do
    begin
      _Obj:= StringGrid.Objects[1,i];
      if Assigned(_Obj) then
         begin
           if Assigned(TStringGridData(_Obj).ComboBox)
              then cmbxClearData(TStringGridData(_Obj).ComboBox);

           if Assigned(TStringGridData(_Obj).ComboBox) then
              TStringGridData(_Obj).ComboBox.Free;

           if Assigned(TStringGridData(_Obj).CheckBox) then
              TStringGridData(_Obj).CheckBox.Free;

           if Assigned(TStringGridData(_Obj).BitBtn) then
              TStringGridData(_Obj).BitBtn.Free;

           if Assigned(TStringGridData(_Obj).BitBtn1) then
              TStringGridData(_Obj).BitBtn1.Free;

           if Assigned(TStringGridData(_Obj).BitBtnEditor) then
              TStringGridData(_Obj).BitBtnEditor.Free;

           TStringGridData(_Obj).Free;
         end;
    end;
    StringGrid.Clear;
    //StringGrid.EndUpdate();
  end;
end;

destructor TwFormatsGrid.Destroy();
var
  i: Integer;
begin
  StringGridClear();
  wEditor.Free;
  FreeAndNil(fdbDataSet);
  FreeAndNil(fdbDataSource);
  FreeAndNil(fdbTransaction);
end;

procedure TwFormatsGrid.FillGrid(const aFormatID: integer; const Sender: TObject);
var
  i:integer;
  _SQLText: String;
  _Obj: TStringGridData;
  _arr: ArrayOfArrayVariant;
begin
   if (fBase = nil) then exit;
   FillGridOn:= true;

   fFormatID:= aFormatID;

   _arr:= fBase.SQLReadArr('FORMATS',['IDFMTS_CATEGORY'],'ID='+IntToStr(aFormatID),'');

   if not Assigned(_arr) then exit;

  try
      StringGrid.BeginUpdate;

      if Assigned(Sender) and not (Sender is TComboBox) then
        begin
          cbx_FormatType.OnChange:=nil;
          cbx_FormatType.ItemIndex:= _arr[0,0]-1;
          cbx_FormatType.OnChange:=@_on_cbx_FormatTypeChange;
        end;

      case (cbx_FormatType.ItemIndex+1) of
        1: fFormatType:= ftPRICE;
        2: fFormatType:= ftNAKL;
      end;

      StringGridClear();

      //case fFormatComboBox.ItemIndex+1 of
      //    1: // Прайс-лист
      //      begin

              _SQLText:=' SELECT '
              +'FILE, '
              +'FILEZIPNAMEDECODE, '
              +'URL, '
              +'IDFILEFORMAT, '
              +'FCONVERTLIBRE, '
              +'IDCODEPAGETEXT, '
              +'IDCSVDELIMITER, '
              +'IDVENDORCODEVARIANT, '
              +'IDSTOCKVARIANT, '
              +'IDPRICEVARIANT, '
              +'IDCURRENCY, '
              +'CURRENCYPERCENT, '
              +'STORAGEDAYS, '
              +'STOCKSYMBOLS, '
              +'STOCKONLY, '
              +'YMLID,'
              +'YMLPRICE,'
              +'YMLQUANTITY,'
              +'SPREADSHEET, '
              +'GROUPSINROWS, '
              +'GROUPALGORITHM, '
              +'GROUPS, '
              +'SUBGROUPS1, '
              +'SUBGROUPS2, '
              +'SUBGROUPS3, '
              +'FIRSTLINE, '
              +'VENDORCODE, '
              +'FNAME, '
              +'UNIT, '
              +'QUANTITY, '
              +'FSUM, '
              +'CUSTOMSDECLARATION, '
              +'COUNTRY, '
              +'STOCK2, '
              +'STOCK3, '
              +'STOCK4, '
              +'STOCK5, '
              +'TRANSIT, '
              +'PRICE, '
              +'PRICE2, '
              +'PRICE3, '
              +'PRICE4, '
              +'PRICE5, '
              +'PRICE6, '
              +'PRICE7, '
              +'PRICE8, '
              +'PRICE9, '
              +'PRICE10, '
              +'LABEL, '
              +'SCOD, '
              +'FURL, '
              +'FURLPICTURE, '
              +'FREMARK, '
              +'FCOLOR, '
              +'OUTCELLTEXT, '
              +'ADDRCELLTEXT, '
              +'ADDRCELLFORINVOCE, '
              +'INVOCEDAYS, '
              +'ACTUALDAYS, '
              +'NOMINPRICE, '
              +'REMARK, '
              +'FCLOSE '
              + 'FROM "FORMATS" WHERE ID='+IntToStr(aFormatID);

              fBase.SQLReadDS(fdbDataSet, fdbTransaction, fdbDataSource, _SQLText);
              StringGrid.RowCount:= fdbDataSet.Fields.Count;
              StringGrid.ColWidths[0]:= 200;

              //StringGrid.Columns.Add;
              StringGrid.FixedCols:=1;
              //ffXLS, ffXLSX, ffODS, ffCSV, ffYML, mfNONE
              //категории 0 - основные, 1 - группировка, 3 - столбцы с данными
              with StringGrid do
              begin
                BeginUpdate;
                for i:=0 to RowCount-1 do
                begin
                   case fdbDataSet.Fields[i].FieldName of
                     'FILE'            :
                                       begin
                                         Cells[0,i]:='Файл';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'FILEZIPNAMEDECODE'         :
                                       begin
                                         Cells[0,i]:='Декодировать имя файла';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.CheckBox:= _Obj.CheckBoxCreate(StringGrid,'Установите, если есть проблемы при распаковке из архивов файлов с кирилицей в наименовании.');
                                         _Obj.CheckBox.OnChange:= @_on_CheckBoxChange;


                                         if fdbDataSet.FieldByName(_Obj.FieldName).AsInteger = 1 then
                                            _obj.CheckBox.Checked:= true else
                                            _obj.CheckBox.Checked:= false;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'URL'             :
                                       begin
                                         Cells[0,i]:='Ссылка на прайс-лист';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.BitBtn:= _Obj.BitBtnCreate(StringGrid,'Мастер','Запустить мастер ссылок');
                                         _Obj.BitBtn.OnClick:=@_on_BtnURLMasterClick;
                                         _Obj.BitBtn1:= _Obj.BitBtnCreate(StringGrid,'','Нажмите для очистки');
                                         _Obj.BitBtn1.OnClick:= @_on_BtnURLClear;

                                         if Assigned(StringGrid.TitleImageList) then
                                            begin
                                              if Length(fdbDataSet.FieldByName(_Obj.FieldName).AsString)>0 then
                                                 StringGrid.TitleImageList.GetBitmap(0,_Obj.BitBtn.Glyph) else
                                                 StringGrid.TitleImageList.GetBitmap(1,_Obj.BitBtn.Glyph);
                                                 StringGrid.TitleImageList.GetBitmap(2,_Obj.BitBtn1.Glyph)
                                            end;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'IDFILEFORMAT'    :
                                       begin
                                         Cells[0,i]:='Формат файла';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;
                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Выберите формат файла');
                                         cmbxFill(_Obj.ComboBox,fBase.SQLReadDS('SELECT NAME,ID FROM "FILEFORMAT" ORDER BY ID'),['NAME','ID']);
                                         _Obj.ComboBox.ItemIndex:= cmbxItemIndexByID(_Obj.ComboBox,fdbDataSet.FieldByName(_Obj.FieldName).AsInteger);
                                         _Obj.ComboBox.OnChange:= @cbx_IdFileFormat_onChange;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'FCONVERTLIBRE'    :
                                       begin
                                         Cells[0,i]:='Конвертир. с LibreOffice';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX];
                                         _Obj.Category:= 0;
                                         _Obj.CheckBox:= _Obj.CheckBoxCreate(StringGrid,'Установите, если файл читается некорректно. Необходим установленный LibreOffice.');
                                         _Obj.CheckBox.OnChange:= @_on_CheckBoxChange;

                                         if fdbDataSet.FieldByName(_Obj.FieldName).AsInteger = 1 then
                                            _obj.CheckBox.Checked:= true else
                                            _obj.CheckBox.Checked:= false;

                                         Objects[1,i]:= _Obj;
                                       end;

                     'IDCODEPAGETEXT'  :
                                       begin
                                         Cells[0,i]:='Кодировка';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffCSV];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Выберите кодировку, в случае проблем.');

                                         cmbxFill(_Obj.ComboBox,fBase.SQLReadDS('SELECT NAME,ID FROM "CODEPAGETEXT" ORDER BY ID'),['NAME','ID']);
                                         _Obj.ComboBox.ItemIndex:= cmbxItemIndexByID(_Obj.ComboBox,fdbDataSet.FieldByName(_Obj.FieldName).AsInteger);
                                         _Obj.ComboBox.OnChange:= @_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;

                     'IDCSVDELIMITER':
                                       begin
                                         Cells[0,i]:='Разделитель';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffCSV];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Если не знаете, что выбрать - оставьте по-умолчанию');

                                         _Obj.ComboBox.Items.Add('По-умолчанию');
                                         _Obj.ComboBox.Items.Add(';');
                                         _Obj.ComboBox.Items.Add(',');
                                         _Obj.ComboBox.Items.Add('$');

                                         _Obj.ComboBox.ItemIndex:= fdbDataSet.FieldByName(_Obj.FieldName).AsInteger;
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;

                     'IDVENDORCODEVARIANT':
                                       begin
                                         Cells[0,i]:='Тип Кода Контр.';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Если не знаете, что выбрать - оставьте по-умолчанию');

                                         _Obj.ComboBox.Items.Add('По-умолчанию');
                                         _Obj.ComboBox.Items.Add('Число');
                                         _Obj.ComboBox.Items.Add('Текст');

                                         _Obj.ComboBox.ItemIndex:= fdbDataSet.FieldByName(_Obj.FieldName).AsInteger;
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;

                     'IDSTOCKVARIANT':
                                       begin
                                         Cells[0,i]:='Тип остатка';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Если не знаете, что выбрать - оставьте по-умолчанию');

                                         _Obj.ComboBox.Items.Add('По-умолчанию');
                                         _Obj.ComboBox.Items.Add('Число');

                                         _Obj.ComboBox.ItemIndex:= fdbDataSet.FieldByName(_Obj.FieldName).AsInteger;
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;

                     'IDPRICEVARIANT':
                                       begin
                                         Cells[0,i]:='Тип цен';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Если не знаете, что выбрать - оставьте по-умолчанию');

                                         _Obj.ComboBox.Items.Add('По-умолчанию');
                                         _Obj.ComboBox.Items.Add('Число');

                                         _Obj.ComboBox.ItemIndex:= fdbDataSet.FieldByName(_Obj.FieldName).AsInteger;
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;

                     'IDCURRENCY'      :
                                       begin
                                         Cells[0,i]:='Валюта';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Выберите валюту прайс-листа');
                                         _Obj.ComboBox.Name:='IDCURRENCY';
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         cmbxFill(_Obj.ComboBox,fBase.SQLReadDS('SELECT CODE,ID FROM "CURRENCY" ORDER BY ID'),['CODE','ID']);
                                         _Obj.ComboBox.ItemIndex:= cmbxItemIndexByID(_Obj.ComboBox,fdbDataSet.FieldByName(_Obj.FieldName).AsInteger);

                                         Objects[1,i]:= _Obj;
                                       end;
                     'CURRENCYPERCENT' :
                                       begin
                                         Cells[0,i]:='Валюта. % конвертации';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'STORAGEDAYS'     :
                                       begin
                                         Cells[0,i]:='Хранить дней';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'STOCKONLY'  :
                                       begin
                                         Cells[0,i]:='Только в наличии';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.CheckBox:= _Obj.CheckBoxCreate(StringGrid,'Загружать только в наличии');
                                         _Obj.CheckBox.OnChange:= @_on_CheckBoxChange;

                                         if fdbDataSet.FieldByName(_Obj.FieldName).AsInteger = 1 then
                                            _obj.CheckBox.Checked:= true else
                                            _obj.CheckBox.Checked:= false;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'STOCKSYMBOLS'  :
                                       begin
                                         Cells[0,i]:='Заменить наличие';
                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.BitBtnEditor := _Obj.BitBtnCreate(StringGrid,'','Открыть редактор. В случае, если остаток не числовой - можете составить список замен. Формат: СЛОВО=ЧИСЛО. Значения перечисляются через пятятую.');
                                         _Obj.BitBtnEditor.OnClick:=@_on_BtnEditorClick;

                                         if Assigned(StringGrid.TitleImageList) then
                                              StringGrid.TitleImageList.GetBitmap(3,_Obj.BitBtnEditor.Glyph);

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;

                     'YMLID':
                                       begin
                                         Cells[0,i]:='ID позиции';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Если не знаете, что выбрать - оставьте по-умолчанию');

                                         _Obj.ComboBox.Items.Add('По-умолчанию');
                                         _Obj.ComboBox.Items.Add('<product_id_1c>');
                                         _Obj.ComboBox.Items.Add('<vendorcode>');

                                         _Obj.ComboBox.ItemIndex:= fdbDataSet.FieldByName(_Obj.FieldName).AsInteger;
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'YMLPRICE':
                                       begin
                                         Cells[0,i]:='Цена. Базовая.';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Если не знаете, что выбрать - оставьте по-умолчанию');

                                         _Obj.ComboBox.Items.Add('По-умолчанию');
                                         _Obj.ComboBox.Items.Add('<key_partner>');
                                         _Obj.ComboBox.Items.Add('<price>');
                                         _Obj.ComboBox.Items.Add('<oldprice>');

                                         _Obj.ComboBox.ItemIndex:= fdbDataSet.FieldByName(_Obj.FieldName).AsInteger;
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'YMLPRICE2':
                                       begin
                                         Cells[0,i]:='Цена. Колонка 2';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Если не знаете, что выбрать - оставьте по-умолчанию');

                                         _Obj.ComboBox.Items.Add('По-умолчанию');
                                         _Obj.ComboBox.Items.Add('<key_partner>');
                                         _Obj.ComboBox.Items.Add('<price>');
                                         _Obj.ComboBox.Items.Add('<oldprice>');

                                         _Obj.ComboBox.ItemIndex:= fdbDataSet.FieldByName(_Obj.FieldName).AsInteger;
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'YMLQUANTITY':
                                       begin
                                         Cells[0,i]:='Наличие';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Если не знаете, что выбрать - оставьте по-умолчанию.'+LineEnding+' Будет использовано до 3х складов (при наличии) согласно стандарта.');

                                         _Obj.ComboBox.Items.Add('По-умолчанию');
                                         _Obj.ComboBox.Items.Add('<quantity>');

                                         _Obj.ComboBox.ItemIndex:= fdbDataSet.FieldByName(_Obj.FieldName).AsInteger;
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'FCLOSE'         :
                                       begin
                                         Cells[0,i]:='-= Закрыт =-';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.CheckBox:= _Obj.CheckBoxCreate(StringGrid,'Закрытый формат будет пропущен при импорте.');
                                         _Obj.CheckBox.OnChange:= @_on_CheckBoxChange;


                                         if fdbDataSet.FieldByName(_Obj.FieldName).AsInteger = 1 then
                                            _obj.CheckBox.Checked:= true else
                                            _obj.CheckBox.Checked:= false;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'GROUPSINROWS'  :
                                       begin
                                         Cells[0,i]:='Группы в строках';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS];
                                         _Obj.Category:= 1;

                                         _Obj.CheckBox:= _Obj.CheckBoxCreate(StringGrid,'Если в прайс-листе под группы товаров не выделено'+LineEnding+' отдельного столбца, то установите этот флажок.');
                                         _Obj.CheckBox.OnChange:= @_on_CheckBoxChange;

                                         if fdbDataSet.FieldByName(_Obj.FieldName).AsInteger = 1 then
                                            _obj.CheckBox.Checked:= true else
                                            _obj.CheckBox.Checked:= false;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'GROUPALGORITHM':
                                       begin
                                         Cells[0,i]:='Алгоритм поиска группы';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS];
                                         _Obj.Category:= 1;

                                         _Obj.ComboBox:= _Obj.ComboBoxCreate(StringGrid,'Выберите алгоритм поиска группы.'+LineEnding+' В большинстве случаев работает "Фон+Цена"');

                                         _Obj.ComboBox.Items.Add('Фон + Цена');
                                         _Obj.ComboBox.Items.Add('Фон + Идент.');

                                         _Obj.ComboBox.ItemIndex:= fdbDataSet.FieldByName(_Obj.FieldName).AsInteger;
                                         _Obj.ComboBox.OnChange:=@_on_ComboBoxChange;

                                         Objects[1,i]:= _Obj;
                                       end;
                     'GROUPS'          :
                                       begin
                                         Cells[0,i]:='Группа товаров';
                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 1;


                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],0);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'SUBGROUPS1'      :
                                       begin
                                         Cells[0,i]:='Подгруппа товаров [1]';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 1;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],0);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'SUBGROUPS2'      :
                                       begin
                                         Cells[0,i]:='Подгруппа товаров [2]';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 1;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],0);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'SUBGROUPS3'      :
                                       begin
                                         Cells[0,i]:='Подгруппа товаров [3]';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 1;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],0);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             _Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);

                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'SPREADSHEET'     :
                                       begin
                                         Cells[0,i]:='Листы';
                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 0;

                                         _Obj.BitBtnEditor := _Obj.BitBtnCreate(StringGrid,'','Открыть редактор. Каждый лист вводится с новой строки.');
                                         _Obj.BitBtnEditor.OnClick:=@_on_BtnEditorClick;

                                         if Assigned(StringGrid.TitleImageList) then
                                              StringGrid.TitleImageList.GetBitmap(3,_Obj.BitBtnEditor.Glyph);


                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,'Выбранные листы: '+fdbDataSet.FieldByName(_Obj.FieldName).AsString,13);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             _Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'FIRSTLINE'       :
                                       begin
                                         Cells[0,i]:='Первая строка';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],1);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'VENDORCODE'      :
                                       begin
                                         Cells[0,i]:='Идентификатор';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],2);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'FNAME'           :
                                       begin
                                         Cells[0,i]:='Наименование';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],3);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'UNIT'            :
                                       begin
                                         Cells[0,i]:='Единица измерения';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],4);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             _Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'QUANTITY'        :
                                       begin
                                        case FFormatType of
                                          ftPRICE: Cells[0,i]:='Остаток [1]';
                                          ftNAKL : Cells[0,i]:='Количество';
                                        end;

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],5);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'FSUM'        :
                                       begin
                                         Cells[0,i]:='Сумма';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],5);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                'CUSTOMSDECLARATION':
                                       begin
                                         Cells[0,i]:='ГТД';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],5);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'COUNTRY'        :
                                       begin
                                         Cells[0,i]:='Страна';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],5);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'STOCK2'        :
                                       begin
                                         Cells[0,i]:='Остаток [2]';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],5);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'STOCK3'        :
                                       begin
                                         Cells[0,i]:='Остаток [3]';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],5);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             //_Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'STOCK4'        :
                                       begin
                                         Cells[0,i]:='Остаток [4]';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],5);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             //_Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'STOCK5'        :
                                       begin
                                         Cells[0,i]:='Остаток [5]';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],5);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             _Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'TRANSIT'         :
                                       begin
                                         Cells[0,i]:='Транзит';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],6);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             _Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE'           :
                                       begin
                                         Cells[0,i]:='Цена. Базовая. (P)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE2'           :
                                       begin
                                         Cells[0,i]:='Цена [2] (P2)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE3'           :
                                       begin
                                         Cells[0,i]:='Цена [3] (P3)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             //_Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE4'           :
                                       begin
                                         Cells[0,i]:='Цена [4] (P4)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             //_Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE5'           :
                                       begin
                                         Cells[0,i]:='Цена [5] (P5)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             //_Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE6'           :
                                       begin
                                         Cells[0,i]:='Цена [6] (P6)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             //_Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE7'           :
                                       begin
                                         Cells[0,i]:='Цена [7] (P7)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             //_Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE8'           :
                                       begin
                                         Cells[0,i]:='Цена [8] (P8)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             //_Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE9'           :
                                       begin
                                         Cells[0,i]:='Цена [9] (P9)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             //_Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'PRICE10'           :
                                       begin
                                         Cells[0,i]:='Цена [10] (P10)';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],7);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             _Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'LABEL'           :
                                       begin
                                         Cells[0,i]:='Артикул';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],8);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'SCOD'            :
                                       begin
                                         Cells[0,i]:='Штрих-код';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],9);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                             _Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'FURL'         :
                                       begin
                                         Cells[0,i]:='Ссылка';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],10);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'FURLPICTURE'    :
                                       begin
                                         Cells[0,i]:='Ссылка на изображение';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],11);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'FREMARK'         :
                                       begin
                                         Cells[0,i]:='Примечание';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],12);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;

                     'FCOLOR'         :
                                       begin
                                         Cells[0,i]:='Цветовой маркер';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 2;

                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],14);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'OUTCELLTEXT'     :
                                       begin
                                         Cells[0,i]:='Текст, вывод. в накл.';
                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 0;

                                         //_Obj.BitBtnEditor := _Obj.BitBtnCreate(StringGrid,'','Открыть редактор');
                                         //_Obj.BitBtnEditor.OnClick:=@_on_BtnEditorClick;
                                         //
                                         //if Assigned(StringGrid.TitleImageList) then
                                         //     StringGrid.TitleImageList.GetBitmap(3,_Obj.BitBtnEditor.Glyph);


                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'ADDRCELLTEXT'     :
                                       begin
                                         Cells[0,i]:='Яч. с текст. [напр. А5]';
                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV];
                                         _Obj.Category:= 0;


                                         if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         begin
                                             _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,Cells[0,i],12);
                                             _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                             xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'ADDRCELLFORINVOCE'     :
                                       begin
                                         Cells[0,i]:='Ячейка для заказа';
                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS];
                                         _Obj.Category:= 0;
                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;

                     'INVOCEDAYS'     :
                                       begin
                                         Cells[0,i]:='Доставка дней';
                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;
                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;

                     'REMARK'     :
                                       begin
                                         Cells[0,i]:='Примечание к формату';
                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE,ftNAKL];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.BitBtnEditor := _Obj.BitBtnCreate(StringGrid,'','Открыть редактор');
                                         _Obj.BitBtnEditor.OnClick:=@_on_BtnEditorClick;

                                         if Assigned(StringGrid.TitleImageList) then
                                              StringGrid.TitleImageList.GetBitmap(3,_Obj.BitBtnEditor.Glyph);


                                         //if Assigned(xGridPopupMenu) and FormatTypeInArr(FFormatType,_Obj.fFormatType) then
                                         //begin
                                         //    _Obj.MenuItem:= _Obj.MenuItemCreate(xGridPopupMenu,_Obj.FieldName,'Выбранные листы: '+fdbDataSet.FieldByName(_Obj.FieldName).AsString,13);
                                         //    _Obj.MenuItem.OnClick:=@_on_xGridMenuItemClick;
                                         //    xGridPopupMenu.Items.Add(_Obj.MenuItem);
                                         //    _Obj.MenuSpliter:= _Obj.MenuSpliterCreate(xGridPopupMenu);
                                         //end;

                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'ACTUALDAYS'     :
                                       begin
                                         Cells[0,i]:='Актуальность, дни. 0 - всегда';
                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;
                                         Objects[1,i]:= _Obj;
                                         Cells[1,i]:= fdbDataSet.FieldByName(_Obj.FieldName).AsString;
                                       end;
                     'NOMINPRICE'         :
                                       begin
                                         Cells[0,i]:='Не использовать в формулах';

                                         _Obj:= TStringGridData.Create();
                                         _Obj.fFormatType:= [ftPRICE];
                                         _Obj.FieldName:= fdbDataSet.Fields[i].FieldName;
                                         _Obj.FileType:= [ffXLS, ffXLSX, ffODS, ffCSV, ffYML];
                                         _Obj.Category:= 0;

                                         _Obj.CheckBox:= _Obj.CheckBoxCreate(StringGrid,'Не учитывать при расчёте минимальной цены');
                                         _Obj.CheckBox.OnChange:= @_on_CheckBoxChange;


                                         if fdbDataSet.FieldByName(_Obj.FieldName).AsInteger = 1 then
                                            _obj.CheckBox.Checked:= true else
                                            _obj.CheckBox.Checked:= false;

                                         Objects[1,i]:= _Obj;
                                       end;
                   end;
                end;
                EndUpdate();
              end;

              ChangeView(fBase.SQLReadArr('Select Lower(CODE) from "FILEFORMAT" WHERE ID='+fdbDataSet.FieldByName('IDFILEFORMAT').AsString)[0,0]);
              if Assigned(onFillGrid) then onFillGrid(self);
            //end;
      //end;

      StringGrid.EndUpdate();
    FillGridOn:= false;
  except
    on E: Exception do
    begin
        StringGrid.EndUpdate();
        SetStatus('Сбой при чтении данных из БД.');
        Log('Ошибка [FillGrid]: "' + E.Message + '"');

     end;
  end;
end;

procedure TwFormatsGrid.SelectCellByName(aText:string; aCol, aRow: integer);
var
  _Obj: TStringGridData;
begin
  _Obj:= TStringGridData(StringGrid.Objects[aCol,aRow]);
  if Assigned(_Obj) then
  begin

      if not Assigned(_obj.ComboBox) and not Assigned(_obj.CheckBox) and not Assigned(_obj.BitBtn) and not Assigned(_obj.BitBtn1) then
      if (_Obj.FieldName<>'FILE') or (_Obj.FieldName<>'SPREADSHEET') then
          StringGrid.Cells[aCol,aRow]:= aText;

      if Assigned(onChangedFormat) then onChangedFormat(self);
  end;
end;

procedure TwFormatsGrid.ClearCell(aCol, aRow: integer);
var
  _Obj: TStringGridData;
begin
_Obj:= TStringGridData(StringGrid.Objects[aCol,aRow]);
if Assigned(_Obj) then
begin
    if Assigned(_obj.CheckBox) then _obj.CheckBox.Checked:= false;

    if  Assigned(_obj.ComboBox) then
    begin

      case _Obj.FieldName of
       'IDFILEFORMAT':
                       begin
                         _Obj.ComboBox.ItemIndex:= 2;
                       end;
     'IDCODEPAGETEXT':
                       begin
                         _Obj.ComboBox.ItemIndex:= 0;
                       end;
         'IDCURRENCY':
                       begin
                         _Obj.ComboBox.ItemIndex:= 0;
                       end;
              'YMLID':
                       begin
                         _Obj.ComboBox.ItemIndex:= 0;
                       end;
           'YMLPRICE':
                       begin
                         _Obj.ComboBox.ItemIndex:= 0;
                       end;
          'YMLPRICE2':
                       begin
                         _Obj.ComboBox.ItemIndex:= 0;
                       end;
        'YMLQUANTITY':
                       begin
                         _Obj.ComboBox.ItemIndex:= 0;
                       end;
     'GROUPALGORITHM':
                       begin
                         _Obj.ComboBox.ItemIndex:= 0;
                       end;

      end;

    end;

    if not Assigned(_obj.ComboBox) and not Assigned(_obj.CheckBox) and not Assigned(_obj.BitBtn) and not Assigned(_obj.BitBtn1) then
    if (_Obj.FieldName<>'FILE') or (_Obj.FieldName<>'SPREADSHEET') or (_Obj.FieldName<>'STOCKSYMBOLS') or (_Obj.FieldName<>'REMARK') then
        StringGrid.Cells[aCol,aRow]:='0';

    if _Obj.FieldName = 'SPREADSHEET' then StringGrid.Cells[aCol,aRow]:='1';
end;
    if Assigned(onChangedFormat) then onChangedFormat(self);
end;

procedure TwFormatsGrid.ClearAll;
var
  i: integer;
begin
     for i:=0 to StringGrid.RowCount-1 do
     begin
       ClearCell(1,i);
     end;

end;

procedure TwFormatsGrid.ChangeView(const aFormat: string);
var
  i: Integer;
begin

case aFormat of
  'xls'      : fFileType:= ffXLS;
  'xlsx'     : fFileType:= ffXLSX;
  'ods'      : fFileType:= ffODS;
  //'txt'      : fFileType:= mfTXT;
  'csv'      : fFileType:= ffCSV;
  'yml'      : fFileType:= ffYML;
  else
   fFileType:= ffNONE;
end;

  on_TabCategotyChange(self);
end;

procedure TwFormatsGrid.Save(Sender: TObject);
var
  _obj: TStringGridData;
  i, _IDFILEFORMAT, _FCONVERTLIBRE, _IDCODEPAGETEXT, _IDCURRENCY, _STOCKONLY, _YMLID, _YMLPRICE, _YMLQUANTITY, _FCLOSE, _GROUPSINROWS, _GROUPALGORITHM,
    _IDFMTS_CATEGORY, _STORAGEDAYS, _GROUPS, _SUBGROUPS1, _SUBGROUPS2, _SUBGROUPS3, _FIRSTLINE, _VENDORCODE, _FNAME, _UNIT, _QUANTITY,
    _STOCK2, _STOCK3, _STOCK4, _STOCK5, _TRANSIT, _PRICE, _PRICE2, _PRICE3, _PRICE4, _PRICE5, _PRICE6, _PRICE7, _PRICE8, _PRICE9, _PRICE10, _LABEL,
    _SCOD, _FREMARK, _FURL, _FURLPICTURE, _FILEZIPNAMEDECODE, _STOCKONLYINFO, _FCOLOR, _CUSTOMSDECLARATION, _COUNTRY, _FSUM, _IDCSVDELIMITER,
    _IDVENDORCODEVARIANT, _IDSTOCKVARIANT, _IDPRICEVARIANT, _ACTUALDAYS, _NOMINPRICE: Integer;

  _FILE, _SPREADSHEET, _STOCKSYMBOLS, _REMARK, _OUTCELLTEXT, _ADDRCELLTEXT, _ADDRCELLFORINVOCE, _INVOCEDAYS: String;
  _CURRENCYPERCENT: Double;
begin
try
  _IDFMTS_CATEGORY:= cbx_FormatType.ItemIndex+1;

  for i:=0 to StringGrid.RowCount-1 do
    begin
      _obj:= TStringGridData(StringGrid.Objects[1,i]);
      if Assigned(_Obj) then
      begin

                case _Obj.FieldName of
                  'FILE'            : _FILE:= StringGrid.Cells[1,i];
                  'FILEZIPNAMEDECODE': if _Obj.CheckBox.Checked then _FILEZIPNAMEDECODE:=1 else _FILEZIPNAMEDECODE:= 0;
                  'IDFILEFORMAT'    : _IDFILEFORMAT:= cmbxSelectID(_Obj.ComboBox);
                  'FCONVERTLIBRE'   : if _Obj.CheckBox.Checked then _FCONVERTLIBRE:=1 else _FCONVERTLIBRE:= 0;
                  'IDCODEPAGETEXT'  : _IDCODEPAGETEXT:= cmbxSelectID(_Obj.ComboBox);
                  'IDCSVDELIMITER'  : _IDCSVDELIMITER:= _Obj.ComboBox.ItemIndex;
                  'IDVENDORCODEVARIANT'  : _IDVENDORCODEVARIANT:= _Obj.ComboBox.ItemIndex;
                  'IDSTOCKVARIANT'  : _IDSTOCKVARIANT:= _Obj.ComboBox.ItemIndex;
                  'IDPRICEVARIANT'  : _IDPRICEVARIANT:= _Obj.ComboBox.ItemIndex;
                  'IDCURRENCY'      : _IDCURRENCY:= cmbxSelectID(_Obj.ComboBox);
                  'CURRENCYPERCENT' : TryStrToFloat(StringGrid.Cells[1,i],_CURRENCYPERCENT);
                  'STORAGEDAYS'     : TryStrToInt(StringGrid.Cells[1,i],_STORAGEDAYS);
                  'STOCKONLY'       : if _Obj.CheckBox.Checked then _STOCKONLY:=1 else _STOCKONLY:= 0;
                  'STOCKSYMBOLS'    :
                                    begin
                                      _STOCKSYMBOLS:= StringGrid.Cells[1,i];
                                      if Length(_STOCKSYMBOLS)>0 then _STOCKONLYINFO:=1 else _STOCKONLYINFO:= 0;
                                    end;
                  'YMLID'           : _YMLID:= _Obj.ComboBox.ItemIndex;
                  'YMLPRICE'        : _YMLPRICE:= _Obj.ComboBox.ItemIndex;
                  'YMLQUANTITY'     : _YMLQUANTITY:= _Obj.ComboBox.ItemIndex;
                  'SPREADSHEET'     : _SPREADSHEET:= StringGrid.Cells[1,i];
                  'FCLOSE'          : if _Obj.CheckBox.Checked then _FCLOSE:=1 else _FCLOSE:= 0;
                  'GROUPSINROWS'    : if _Obj.CheckBox.Checked then _GROUPSINROWS:=1 else _GROUPSINROWS:= 0;
                  'GROUPALGORITHM'  : _GROUPALGORITHM:= _Obj.ComboBox.ItemIndex;
                  'GROUPS'          : TryStrToInt(StringGrid.Cells[1,i],_GROUPS);
                  'SUBGROUPS1'      : TryStrToInt(StringGrid.Cells[1,i],_SUBGROUPS1);
                  'SUBGROUPS2'      : TryStrToInt(StringGrid.Cells[1,i],_SUBGROUPS2);
                  'SUBGROUPS3'      : TryStrToInt(StringGrid.Cells[1,i],_SUBGROUPS3);
                  'FIRSTLINE'       : TryStrToInt(StringGrid.Cells[1,i],_FIRSTLINE);
                  'VENDORCODE'      : TryStrToInt(StringGrid.Cells[1,i],_VENDORCODE);
                  'FNAME'           : TryStrToInt(StringGrid.Cells[1,i],_FNAME);
                  'UNIT'            : TryStrToInt(StringGrid.Cells[1,i],_UNIT);
                  'QUANTITY'        : TryStrToInt(StringGrid.Cells[1,i],_QUANTITY);
                  'STOCK2'          : TryStrToInt(StringGrid.Cells[1,i],_STOCK2);
                  'STOCK3'          : TryStrToInt(StringGrid.Cells[1,i],_STOCK3);
                  'STOCK4'          : TryStrToInt(StringGrid.Cells[1,i],_STOCK4);
                  'STOCK5'          : TryStrToInt(StringGrid.Cells[1,i],_STOCK5);
                  'TRANSIT'         : TryStrToInt(StringGrid.Cells[1,i],_TRANSIT);
                  'PRICE'           : TryStrToInt(StringGrid.Cells[1,i],_PRICE);
                  'PRICE2'          : TryStrToInt(StringGrid.Cells[1,i],_PRICE2);
                  'PRICE3'          : TryStrToInt(StringGrid.Cells[1,i],_PRICE3);
                  'PRICE4'          : TryStrToInt(StringGrid.Cells[1,i],_PRICE4);
                  'PRICE5'          : TryStrToInt(StringGrid.Cells[1,i],_PRICE5);
                  'PRICE6'          : TryStrToInt(StringGrid.Cells[1,i],_PRICE6);
                  'PRICE7'          : TryStrToInt(StringGrid.Cells[1,i],_PRICE7);
                  'PRICE8'          : TryStrToInt(StringGrid.Cells[1,i],_PRICE8);
                  'PRICE9'          : TryStrToInt(StringGrid.Cells[1,i],_PRICE9);
                  'PRICE10'          : TryStrToInt(StringGrid.Cells[1,i],_PRICE10);
                  'LABEL'           : TryStrToInt(StringGrid.Cells[1,i],_LABEL);
                  'SCOD'            : TryStrToInt(StringGrid.Cells[1,i],_SCOD);
                  'FURL'           : TryStrToInt(StringGrid.Cells[1,i],_FURL);
                  'FURLPICTURE'    : TryStrToInt(StringGrid.Cells[1,i],_FURLPICTURE);
                  'FREMARK'        : TryStrToInt(StringGrid.Cells[1,i],_FREMARK);
                  'FCOLOR'        : TryStrToInt(StringGrid.Cells[1,i],_FCOLOR);
                  'CUSTOMSDECLARATION': TryStrToInt(StringGrid.Cells[1,i],_CUSTOMSDECLARATION);
                  'COUNTRY'        : TryStrToInt(StringGrid.Cells[1,i],_COUNTRY);
                  'FSUM'        : TryStrToInt(StringGrid.Cells[1,i],_FSUM);
                  'OUTCELLTEXT'        : _OUTCELLTEXT:= StringGrid.Cells[1,i];
                  'ADDRCELLTEXT'        : _ADDRCELLTEXT:= StringGrid.Cells[1,i];
                  'ADDRCELLFORINVOCE'   : _ADDRCELLFORINVOCE:= StringGrid.Cells[1,i];
                  'INVOCEDAYS'          : _INVOCEDAYS:= StringGrid.Cells[1,i];
                  'REMARK'        : _REMARK:= StringGrid.Cells[1,i];
                  'ACTUALDAYS'          :  TryStrToInt(StringGrid.Cells[1,i],_ACTUALDAYS);
                  'NOMINPRICE'          : if _Obj.CheckBox.Checked then _NOMINPRICE:=1 else _NOMINPRICE:= 0;
                end;

      end;
    end;

        fBase.SQLUpdate('FORMATS',[
                            'FILE',
                            'FILEZIPNAMEDECODE',
                            'IDFILEFORMAT',
                            'FCONVERTLIBRE',
                            'IDCODEPAGETEXT',
                            'IDCSVDELIMITER',
                            'IDVENDORCODEVARIANT',
                            'IDSTOCKVARIANT',
                            'IDPRICEVARIANT',
                            'IDCURRENCY',
                            'CURRENCYPERCENT',
                            'STORAGEDAYS',
                            'STOCKONLY',
                            'STOCKSYMBOLS',
                            'STOCKONLYINFO',
                            'YMLID',
                            'YMLPRICE',
                            'YMLQUANTITY',
                            'SPREADSHEET',
                            'FCLOSE',
                            'GROUPSINROWS',
                            'GROUPALGORITHM',
                            'GROUPS',
                            'SUBGROUPS1',
                            'SUBGROUPS2',
                            'SUBGROUPS3',
                            'FIRSTLINE',
                            'VENDORCODE',
                            'FNAME',
                            'UNIT',
                            'QUANTITY',
                            'STOCK2',
                            'STOCK3',
                            'STOCK4',
                            'STOCK5',
                            'TRANSIT',
                            'PRICE',
                            'PRICE2',
                            'PRICE3',
                            'PRICE4',
                            'PRICE5',
                            'PRICE6',
                            'PRICE7',
                            'PRICE8',
                            'PRICE9',
                            'PRICE10',
                            'LABEL',
                            'SCOD',
                            'FURL',
                            'FURLPICTURE',
                            'FREMARK',
                            'FCOLOR',
                            'IDFMTS_CATEGORY',
                            'CUSTOMSDECLARATION',
                            'COUNTRY',
                            'FSUM',
                            'OUTCELLTEXT',
                            'ADDRCELLTEXT',
                            'ADDRCELLFORINVOCE',
                            'INVOCEDAYS',
                            'REMARK',
                            'ACTUALDAYS',
                            'NOMINPRICE'
                            ],[
                            _FILE,
                            _FILEZIPNAMEDECODE,
                            _IDFILEFORMAT,
                            _FCONVERTLIBRE,
                            _IDCODEPAGETEXT,
                            _IDCSVDELIMITER,
                            _IDVENDORCODEVARIANT,
                            _IDSTOCKVARIANT,
                            _IDPRICEVARIANT,
                            _IDCURRENCY,
                            _CURRENCYPERCENT,
                            _STORAGEDAYS,
                            _STOCKONLY,
                            _STOCKSYMBOLS,
                            _STOCKONLYINFO,
                            _YMLID,
                            _YMLPRICE,
                            _YMLQUANTITY,
                            _SPREADSHEET,
                            _FCLOSE,
                            _GROUPSINROWS,
                            _GROUPALGORITHM,
                            _GROUPS,
                            _SUBGROUPS1,
                            _SUBGROUPS2,
                            _SUBGROUPS3,
                            _FIRSTLINE,
                            _VENDORCODE,
                            _FNAME,
                            _UNIT,
                            _QUANTITY,
                            _STOCK2,
                            _STOCK3,
                            _STOCK4,
                            _STOCK5,
                            _TRANSIT,
                            _PRICE,
                            _PRICE2,
                            _PRICE3,
                            _PRICE4,
                            _PRICE5,
                            _PRICE6,
                            _PRICE7,
                            _PRICE8,
                            _PRICE9,
                            _PRICE10,
                            _LABEL,
                            _SCOD,
                            _FURL,
                            _FURLPICTURE,
                            _FREMARK,
                            _FCOLOR,
                            _IDFMTS_CATEGORY,
                            _CUSTOMSDECLARATION,
                            _COUNTRY,
                            _FSUM,
                            _OUTCELLTEXT,
                            _ADDRCELLTEXT,
                            _ADDRCELLFORINVOCE,
                            _INVOCEDAYS,
                            _REMARK,
                            _ACTUALDAYS,
                            _NOMINPRICE
                            ],'ID='+IntToStr(fFormatID),false);
             if Assigned(onSavedFormat) then onSavedFormat(self);


except
  on E: Exception do
   begin
     screen.Cursor:= crDefault;
     __Log.SaveLogError(E);
     wLog('wStringGrid','Ошибка [_on_BtnURLMasterClick]: "' + E.Message + '"');
     raise;
   end;
end;
end;

procedure TwFormatsGrid.SetGridCellValue(aFieldName: string; aValue: boolean);
var
  _obj: TStringGridData;
  i: Integer;
begin
for i:=0 to StringGrid.RowCount-1 do
  begin
    _obj:= TStringGridData(StringGrid.Objects[1,i]);
    if Assigned(_Obj) then
    begin
     if _Obj.FieldName = aFieldName then
       if Assigned(_Obj.CheckBox) then
       begin
          _Obj.CheckBox.Checked:= aValue;
          Exit;
       end;
    end;
  end;

end;

procedure TwFormatsGrid.SetGridCellValue(aFieldName: string; aValue: integer);
var
  _obj: TStringGridData;
  i: Integer;
begin
for i:=0 to StringGrid.RowCount-1 do
  begin
    _obj:= TStringGridData(StringGrid.Objects[1,i]);
    if Assigned(_Obj) then
    begin
     if _Obj.FieldName = aFieldName then
       if Assigned(_Obj.ComboBox) then
       begin
          _Obj.ComboBox.ItemIndex:= aValue;
          Exit;
       end;
    end;
  end;
end;

procedure TwFormatsGrid.SetGridCellValue(aFieldName: string; aValue: string);
var
  _obj: TStringGridData;
  i: Integer;
begin
for i:=0 to StringGrid.RowCount-1 do
  begin
    _obj:= TStringGridData(StringGrid.Objects[1,i]);
    if Assigned(_Obj) then
    begin
     if _Obj.FieldName = aFieldName then
     begin
       StringGrid.Cells[1,i]:= aValue;
       Exit;
     end;
    end;
  end;
end;

function TwFormatsGrid.GetCellCheckBox(aFieldName: string): boolean;
var
  _obj: TStringGridData;
  i: Integer;
begin
Result:= false;
for i:=0 to StringGrid.RowCount-1 do
  begin
    _obj:= TStringGridData(StringGrid.Objects[1,i]);
    if Assigned(_Obj) then
    begin
     if _Obj.FieldName = aFieldName then
     begin
       if Assigned(_Obj.CheckBox) then
          Result:= _Obj.CheckBox.Checked;
       Exit;
     end;
    end;
  end;

end;

function TwFormatsGrid.GetCellComboBox(aFieldName: string): integer;
var
  _obj: TStringGridData;
  i: Integer;
begin
Result:=-1;
for i:=0 to StringGrid.RowCount-1 do
  begin
    _obj:= TStringGridData(StringGrid.Objects[1,i]);
    if Assigned(_Obj) then
    begin
     if _Obj.FieldName = aFieldName then
     begin
       if Assigned(_Obj.ComboBox) then
          Result:= cmbxSelectID(_Obj.ComboBox);
       Exit;
     end;
    end;
  end;

end;

function TwFormatsGrid.GetCellVal(aFieldName: string): string;
var
  _obj: TStringGridData;
  i: Integer;
begin
Result:='';
for i:=0 to StringGrid.RowCount-1 do
  begin
    _obj:= TStringGridData(StringGrid.Objects[1,i]);
    if Assigned(_Obj) then
    begin
     if _Obj.FieldName = aFieldName then
     begin
       if not Assigned(_Obj.CheckBox) or not Assigned(_Obj.ComboBox) then Result:= StringGrid.Cells[1,i];
       Exit;
     end;
    end;
  end;

end;

function TwFormatsGrid.FormatCopy(aID: integer): integer;
var
  _arr: ArrayOfArrayVariant;
begin
    Result:= -1;

       try
         Result:= fBase.SQLInsert(
            'INSERT INTO FORMATS '
            +' (IDOWNER, '
            +' PRIORITY, '
            +' NAME, '
            +' FIRSTLINE, '
            +' VENDORCODE, '
            +' FNAME, '
            +' UNIT, '
            +' QUANTITY, '
            +' PRICE, '
            +' FSUM, '
            +' LABEL, '
            +' SCOD, '
            +' CUSTOMSDECLARATION, '
            +' COUNTRY, '
            +' "FILE", '
            +' URL, '
            +' SPREADSHEET, '
            +' GROUPSINROWS, '
            +' GROUPS, '
            +' SUBGROUPS1, '
            +' SUBGROUPS2, '
            +' SUBGROUPS3, '
            +' STOCKONLY, '
            +' IDUSER, '
            +' GROUPALGORITHM, '
            +' FTIMESTAMP, '
            +' FILEHASH, '
            +' STORAGEDAYS, '
            +' IDFILEFORMAT, '
            +' IDCURRENCY, '
            +' CURRENCYPERCENT, '
            +' IDCODEPAGETEXT, '
            +' REMARK, '
            +' FREMARK, '
            +' TRANSIT, '
            +' FTIMESTAMPLASTIMPORT, '
            +' FCLOSE, '
            +' YMLID, '
            +' YMLPRICE, '
            +' YMLPRICE2, '
            +' YMLQUANTITY, '
            +' FURLPICTURE, '
            +' FURL, '
            +' PRICE2, '
            +' PRICE3, '
            +' PRICE4, '
            +' PRICE5, '
            +' PRICE6, '
            +' PRICE7, '
            +' PRICE8, '
            +' PRICE9, '
            +' PRICE10, '
            +' STOCK2, '
            +' STOCK3, '
            +' STOCK4, '
            +' STOCK5, '
            +' FILEZIPNAMEDECODE, '
            +' STOCKONLYINFO, '
            +' FCOLOR, '
            +' IDFMTS_CATEGORY, '
            +' FCONVERTLIBRE, '
            +' STOCKSYMBOLS, '
            +' IDCSVDELIMITER, '
            +' IDVENDORCODEVARIANT, '
            +' ADDRCELLTEXT, '
            +' OUTCELLTEXT, '
            +' ADDRCELLFORINVOCE,'
            +' INVOCEDAYS,'
            +' IDSTOCKVARIANT, '
            +' IDPRICEVARIANT '
            +') '
            +' SELECT '
            +' IDOWNER, '
            +' PRIORITY+1, '
            +' NAME||'' Копия'', '
            +' FIRSTLINE, '
            +' VENDORCODE, '
            +' FNAME, '
            +' UNIT, '
            +' QUANTITY, '
            +' PRICE, '
            +' FSUM, '
            +' LABEL, '
            +' SCOD, '
            +' CUSTOMSDECLARATION, '
            +' COUNTRY, '
            +' "FILE", '
            +' URL, '
            +' SPREADSHEET, '
            +' GROUPSINROWS, '
            +' GROUPS, '
            +' SUBGROUPS1, '
            +' SUBGROUPS2, '
            +' SUBGROUPS3, '
            +' STOCKONLY, '
            +' IDUSER, '
            +' GROUPALGORITHM, '
            +' FTIMESTAMP, '
            +' FILEHASH, '
            +' STORAGEDAYS, '
            +' IDFILEFORMAT, '
            +' IDCURRENCY, '
            +' CURRENCYPERCENT, '
            +' IDCODEPAGETEXT, '
            +' REMARK, '
            +' FREMARK, '
            +' TRANSIT, '
            +' FTIMESTAMPLASTIMPORT, '
            +' FCLOSE, '
            +' YMLID, '
            +' YMLPRICE, '
            +' YMLPRICE2, '
            +' YMLQUANTITY, '
            +' FURLPICTURE, '
            +' FURL, '
            +' PRICE2, '
            +' PRICE3, '
            +' PRICE4, '
            +' PRICE5, '
            +' PRICE6, '
            +' PRICE7, '
            +' PRICE8, '
            +' PRICE9, '
            +' PRICE10, '
            +' STOCK2, '
            +' STOCK3, '
            +' STOCK4, '
            +' STOCK5, '
            +' FILEZIPNAMEDECODE, '
            +' STOCKONLYINFO, '
            +' FCOLOR, '
            +' IDFMTS_CATEGORY, '
            +' FCONVERTLIBRE, '
            +' STOCKSYMBOLS, '
            +' IDCSVDELIMITER, '
            +' IDVENDORCODEVARIANT, '
            +' ADDRCELLTEXT, '
            +' OUTCELLTEXT,  '
            +' ADDRCELLFORINVOCE,'
            +' INVOCEDAYS,'
            +' IDSTOCKVARIANT, '
            +' IDPRICEVARIANT '
            +' FROM FORMATS WHERE ID='+IntToStr(aID),true);

         _arr:= fBase.SQLReadArr('FORMATS',['PRIORITY'],'ID='+IntToStr(Result),'');
         if Assigned(_arr) then
         Result:= _arr[0,0];

       except
         raise;
       end;

end;

end.

