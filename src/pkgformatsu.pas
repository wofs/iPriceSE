unit pkgFormatsU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  SysUtils, Dialogs, Forms, Controls, ComCtrls, ExtCtrls, StdCtrls, DBGrids,
  Classes, Graphics, Grids, Buttons, LCLProc, Menus, LazUTF8, fpsTypes, fpspreadsheet, db,
  fpdbfexport, fpstdexports,
  DateUtils,
  DOM, xmlread,
  wLogU, wBaseU, wDBTreeU,
  wTabU, wFormatsGridU, wFuncU, wTypesU,
  FmMasterU,FmFormatsImportU, fpsexport;

type

  { TFmFormats }

  TFmFormats = class(TForm)
    btnSave: TBitBtn;
    cb_FormatType: TComboBox;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    ImageListGrid: TImageList;
    ImageListTree: TImageList;
    Images16: TImageList;
    Label1: TLabel;
    MenuItem1: TMenuItem;
    OpenDialog: TOpenDialog;
    pmFormatsDelete: TMenuItem;
    pmFormatsRename: TMenuItem;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    pmFormats: TPopupMenu;
    SaveDialog: TSaveDialog;
    Splitter1: TSplitter;
    sgFormat: TStringGrid;
    TabCategory: TTabControl;
    TabFormats: TTabControl;
    tbTree: TToolBar;
    tbTreeBtnAdd: TToolButton;
    tbTreeBtnRename: TToolButton;
    tbTreeBtnDelete: TToolButton;
    ToolBar2: TToolBar;
    btnFormatAdd: TToolButton;
    btnFormatDelete: TToolButton;
    btnFormatWisard: TToolButton;
    btnFormatRename: TToolButton;
    ToolButton1: TToolButton;
    tbTreeBtnSetMain: TToolButton;
    ToolButton2: TToolButton;
    tbExport: TToolButton;
    tbImport: TToolButton;
    btnFormatCopy: TToolButton;
    TreeGroupOwner: TTreeView;
    procedure btnFormatAddClick(Sender: TObject);
    procedure btnFormatCopyClick(Sender: TObject);
    procedure btnFormatRenameClick(Sender: TObject);
    procedure btnFormatWisardClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnFormatDeleteClick(Sender: TObject);
    procedure sgFormatResize(Sender: TObject);
    procedure TabFormatsChange(Sender: TObject);
    procedure tbExportClick(Sender: TObject);
    procedure tbImportClick(Sender: TObject);
    procedure tbTreeBtnAddClick(Sender: TObject);
    procedure tbTreeBtnDeleteClick(Sender: TObject);
    procedure tbTreeBtnRenameClick(Sender: TObject);
    procedure tbTreeBtnSetMainClick(Sender: TObject);
    procedure TreeGroupOwnerClick(Sender: TObject);
    procedure TreeGroupOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
  private
    { private declarations }
    FormIDent: string;
    OwnerID: integer;     // ID выбранного контрагента
    IdMainOwner: integer; // ID основного контрагента (к которому привязан каталог)
    fTabIndexCurrent: integer;

    fTabFormat: TwTab;
    fBase: TwBase;
    _DBTreeTreeGroupOwner: TwDBTree;
    fFormatsGrid: TwFormatsGrid;

    _FieldsArr:ArrayOfString;

    function GetDataString(_ACell: PCell): string;
    procedure RefrashTabControl(Tab: TwTab; _Index: integer);
    procedure _onChangeFormat(Sender: TObject);
    procedure _onFillGrid(Sender: TObject);
    procedure _onSavedFormat(Sender: TObject);
    procedure _onTreeMenuSetMainClick(Sender: TObject);
    property wFormID: string read FormIDent write FormIDent;
    property FieldsArr: ArrayOfString read  _FieldsArr write _FieldsArr;

    procedure SetStatus(_Text:string);
  public
    { public declarations }
  end;

var
  FmFormats: TFmFormats;

const

  _Fields = 'ID,NAME,FIRSTLINE,VENDORCODE,FNAME,UNIT,QUANTITY,PRICE,FSUM,LABEL,SCOD,CUSTOMSDECLARATION,COUNTRY,FILE,URL,SPREADSHEET,GROUPSINROWS,GROUPS,SUBGROUPS1,SUBGROUPS2,SUBGROUPS3,STOCKONLY,IDFMTS_CATEGORY';
  // поля, выбираемые для формирования списка полей формата

implementation

{$R *.lfm}

{ TFmFormats }

procedure TFmFormats.SetStatus(_Text: string);
begin
  wStatus(wFormID,_Text,true);
  Application.ProcessMessages;
end;

procedure TFmFormats._onChangeFormat(Sender: TObject);
begin
  if fFormatsGrid.FillGridOn then exit;

  if not fBase.LongTransaction then
     fBase.LongTransaction:= true;

  btnSave.Enabled:= true;
end;

procedure TFmFormats._onFillGrid(Sender: TObject);
begin
  btnSave.Enabled:= false;
end;

procedure TFmFormats._onSavedFormat(Sender: TObject);
begin
  fBase.SQLTransactionEnd(true);
  ShowMessage('Формат успешно сохранен!');
  btnSave.Enabled:=false;
end;

procedure TFmFormats.FormCreate(Sender: TObject);
var
  _TreeMenu, _TabMenu: TPopupMenu;
  TreeMenuSetMain, TreeMenuSpliter: TMenuItem;
begin
   wFormID:=Self.Name;
   OwnerID:=0; // инициализируем.

   screen.Cursor:= crSQLWait;

   //  := TwTabList.Create(wFormID,TabFormats);

  wLog('Formats','Инициализация плагина... ['+wFormID+']');

  try
    //wTab

    fTabFormat:= TwTab.Create(Sender,TabFormats);

    _TabMenu:=pmFormats;
    with _TabMenu do
    begin
      Images:= Images16;
      Items[0].ImageIndex:= 1;
      Items[1].ImageIndex:= 2;
    end;

    //DBase

    fBase:= TwBase.Create(Sender);

    TryStrToInt(fBase.ReadSettingByName('setDefaultOwner'),IdMainOwner); // считываем настройки - текущий основной прайс-лист

    cmbxFill(cb_FormatType,fBase.SQLReadDS('FORMATS_CATEGORY',['NAME','ID'],'FCLOSE=0','ID'),['NAME','ID']);

    FieldsArr:= fBase.MakeArrayFromString(_Fields); // поля, выбираемые для формирования списка полей формата


    //wFormatsGrid
    fFormatsGrid:= TwFormatsGrid.Create(Sender,sgFormat,cb_FormatType,true,fBase);
    fFormatsGrid.TabCategory:= TabCategory;
    fFormatsGrid.onChangedFormat:= @_onChangeFormat;
    fFormatsGrid.onFillGrid:= @_onFillGrid;
    fFormatsGrid.onSavedFormat:=@_onSavedFormat;
  // заполняем StringGrid в FormShow

     //DBTree
     _DBTreeTreeGroupOwner:= TwDBTree.Create(fBase,TreeGroupOwner,'OWNER','IDPARENT,ID',[]);
     _DBTreeTreeGroupOwner.Expanded:= true;

     _TreeMenu:=TreeGroupOwner.PopupMenu;

     TreeMenuSpliter:= TMenuItem.Create(_TreeMenu);
     TreeMenuSpliter.Caption:= '-';
     TreeMenuSpliter.Enabled:=false;;
     //TreeMenuSpliter.OnClick:=@_onTreePopupMenuDelete;
     _TreeMenu.Items.Add(TreeMenuSpliter);

     TreeMenuSetMain:= TMenuItem.Create(_TreeMenu);
     TreeMenuSetMain.Caption:= 'Установить основным';
     TreeMenuSetMain.OnClick:=@_onTreeMenuSetMainClick;
     _TreeMenu.Items.Add(TreeMenuSetMain);

     with _TreeMenu do
     begin
       Images:= Images16;
       Items[0].ImageIndex:= 0;
       Items[1].ImageIndex:= 1;
       Items[2].ImageIndex:= 2;
       Items[4].ImageIndex:= 4;
     end;

     _DBTreeTreeGroupOwner.Fill(); // заполнение дерева

    wLog('Formats','Инициализация плагина успешно завершена.');

    screen.Cursor:= crDefault;
  except
    on E: Exception do
    begin
        screen.Cursor:= crDefault;
        SetStatus('Сбой инициализации плагина.');
        wLog('Formats','Ошибка [FmCreate]: "' + E.Message + '"');
        wLog('Formats','Сбой инициализации плагина.');
        ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

     end;
  end;
end;

procedure TFmFormats.btnFormatAddClick(Sender: TObject);
var
  _Tree: TTreeView;
  _Priority: integer;
  _ID: integer;
begin
  try
   _Tree:=TreeGroupOwner;

   if _Tree.Selected.Level = 0 then exit;

   OwnerID:= _DBTreeTreeGroupOwner.SelectedItems[0];

  // _TabControl:= TabFormats;
   if (_Tree.Selected.Count>0) then
      begin
        ShowMessage('Формат можно добавить только для контрагента!');
        exit;
      end;

   _Priority:= fTabFormat.Count;

   fBase.SQLInsert('FORMATS',['IDOWNER','PRIORITY','NAME','FILE','URL','STORAGEDAYS','IDFMTS_CATEGORY','FTIMESTAMP','FTIMESTAMPLASTIMPORT'],[OwnerID,_Priority,'Формат '+IntToStr(_Priority),'','',integer(5),integer(1),now,now],true);  //SQLReadArr('FORMATS',FieldsArr,'IDOWNER='+_Where,'PRIORITY, Name');

   // если формат добавлен впервые, то создаем корень группы товаров для этого контрагента
   if fBase.SQLReadDS('FORMATS',['ID'],'IDOWNER='+IntToStr(OwnerID),'').DataSet.RecordCount=1 then
      if fBase.SQLReadDS('PL_GROUP',['ID'],'IDOWNER='+IntToStr(OwnerID),'').DataSet.RecordCount=0 then
            fBase.SQLInsert('PL_GROUP',['IDOWNER','IDPARENT','NAME'],[OwnerID,integer(0),'Номенклатура'],true);


   RefrashTabControl(fTabFormat,_Priority);

   SetStatus('Формат успешно добавлен.');
 except
       on E: Exception do
   begin
      wLog('Formats','Ошибка [AddFormat]: "' + E.Message + '"');
      ShowMessage('Ошибка [AddFormat]: "' + E.Message + '"');
   end;
 end;
end;

procedure TFmFormats.btnFormatCopyClick(Sender: TObject);
var
  _Priority, _Index: Integer;
begin

 _Index:= fTabFormat.TabIndex;

  if MessageDlg('Копировать формат '+fTabFormat.Text[_Index]+'? Все несохраненные изменения формата будут отменены!',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

 try
    btnSave.Enabled:= false;

    if fBase.LongTransaction then fBase.SQLTransactionEnd(false);

    _Priority:= fFormatsGrid.FormatCopy(fTabFormat.Value[_Index]);

                               //_Priority
    RefrashTabControl(fTabFormat,_Priority);
 except
   raise;
 end;

end;

procedure TFmFormats.btnFormatRenameClick(Sender: TObject);
var
    _Index: integer;
    _Name:string;
    _Where: string;
begin
 try

   _Index:= fTabFormat.TabIndex;

  if fTabFormat.Count=0 then exit;

    _Name:=InputBox('Введите новое имя формата','Наименование:',fTabFormat.Text[_Index]);

    if _Name = fTabFormat.Text[_Index] then exit;
    _Where:= fTabFormat.ValueToString[_Index];
    fBase.SQLUpdate('FORMATS',['NAME'],[_Name],'ID='+_Where,true);


   RefrashTabControl(fTabFormat,_Index);

    //fTabFormat.TabIndex:=_Index;

    SetStatus('Формат успешно переименован.');
  except
        on E: Exception do
    begin
       wLog('Formats','Ошибка [AddFormat]: "' + E.Message + '"');
       ShowMessage('Ошибка [AddFormat]: "' + E.Message + '"');
    end;
  end;

end;

procedure TFmFormats.btnFormatWisardClick(Sender: TObject);
var
   _Form: TFmMaster;
   _Tree: TTreeVIew;
   _Index: integer;

begin

  if fTabFormat.Count=0 then exit;

  _Tree:= TreeGroupOwner;
  _Form := TFmMaster.Create(Self);

  _Index:= fTabFormat.TabIndex;
  _Form.sgFormat.Tag:=fTabFormat.Value[_Index];
  _Form.e_Owner.Text:=_Tree.Selected.Text;
  _Form.OwnerID:= OwnerID;
  _Form.FormatName:= fTabFormat.Text[_Index];
  _Form.Caption:= 'Мастер формата ['+fTabFormat.Text[_Index]+']';

  try
  _Form.ShowModal;

  finally
    _Form.Free;

    fFormatsGrid.FillGrid(sgFormat.Tag, self);

  end;
end;

procedure TFmFormats.btnSaveClick(Sender: TObject);
begin
  fFormatsGrid.Save(self);
end;

procedure TFmFormats.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if btnSave.Enabled then
  if MessageDlg('Закрыть без сохранения?',mtConfirmation, mbOKCancel, 0) = mrOK
   then
     ModalResult:= mrCancel
   else
     CanClose:= false
end;

procedure TFmFormats.FormDestroy(Sender: TObject);
begin
      try
       wLog('Formats','Выгрузка плагина...');

        // выгружаем подгруженные TabControls
    fTabFormat.Destroy();

        // выгружаем подгруженные DBTree
    _DBTreeTreeGroupOwner.Destroy();


    fBase.SQLTransactionEnd(false);
        // выгружаем DBase
    fBase.Destroy();

    // выгружаем wFormatsGrid;
   fFormatsGrid.Destroy();

   cmbxClearData(cb_FormatType);

      wLog('Formats','Выгрузка плагина успешно завершена.');

      except
        on E: Exception do
        begin
            SetStatus('Сбой выгрузки плагина: Каталог.');
            wLog('Formats','Ошибка [FmDestroy]: "' + E.Message + '"');
            wLog('Formats','Сбой выгрузки плагина.');
            ShowMessage('Ошибка [FmDestroy]: "' + E.Message + '"');
         end;
      end;
end;

procedure TFmFormats._onTreeMenuSetMainClick(Sender: TObject);
var
  _Tree: TTreeVIew;
  _Text: string;
  _TimeStamp, _TimeStampWherePriceList: string;
  _DataSet: TDataSet;

  //_TimeStampMaxArrPriceList: ArrayOfDateTime;

begin
  _Tree:= TreeGroupOwner;
  //_Index:= _Tree.Selected.Index;
  if (_Tree.Selected.Level = 0) or (_Tree.SelectionCount>1) then exit;
  _Text:= _Tree.Selected.Text;

  _TimeStamp := DateTimeToStr(now());

  if _Tree.Selected.Count>0 then exit;

  if MessageDlg('Сделать контрагента "'+_Text+'" основным ? Это приведет к очистке каталога товаров!',mtWarning, mbOKCancel, 0) = mrCancel then exit;

  if MessageDlg('Вы уверены? Очистка каталога необратима!',mtWarning, mbOKCancel, 0) = mrCancel then exit;

  SetStatus('Смена основного контрагента...');

  try
    if fBase.SetSettingByName('setDefaultOwner',TTreeData(_Tree.Selected.Data).Value) then
       begin
           // подчищаем хвосты в каталоге
           fBase.SQLDelete('CATALOG','IDOWNER='+IntToStr(IdMainOwner));
           fBase.SQLDelete('CATALOG_GROUP','IDOWNER='+IntToStr(IdMainOwner));

           IdMainOwner:= fBase.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист

           //fBase.SQLInsert('CATALOG_GROUP',['NAME','IDPARENT','IDOWNER','FTIMESTAMP'],['Номенклатура',0,IdMainOwner,_TimeStamp],true);

           _DBTreeTreeGroupOwner.Fill(); // заполнение дерева
           _DBTreeTreeGroupOwner.FindNodeWithDataInt(IdMainOwner);

       end else
       begin
         ShowMessage('Произошла ошибка при записи настрек!');
         exit;
       end;
//////

// заполнение каталога на основании прайс-листа

           try
              screen.Cursor:= crHourGlass;
              SetStatus('Заполнение каталога... Ждите...');

              //fBase.SQLDelete('CATALOG_GROUP','IDOWNER='+IntToStr(IdMainOwner));

              fBase.SQLUpdate('EXECUTE PROCEDURE CTG_UPDT_PL('+IntToStr(IdMainOwner)+',true);'); // заполняю каталог согласно прайс-листа

              Screen.Cursor := crDefault;

              SetStatus('Обновление индекса...');
              fBase.SQLUpdate('SET STATISTICS INDEX CTG_VENDORCODE;');

              SetStatus('Каталог успешно заполнен.');
              ShowMessage('Каталог успешно заполнен на основании прайс-листа контрагента.');
           except
             on E: Exception do
             begin
                ShowMessage('Ошибка заполнения каталога : "' + E.Message + '"');
                raise;
             end;
           end;



/////

     SetStatus('"'+_Text+'" успешно установлен основным.');
     ShowMessage('"'+_Text+'" успешно установлен основным. Если каталог открыт - закройте и заново откройте его.');

  except
        on E: Exception do
    begin
      __LOg.SaveLogError(E);
       wLog('Formats','Ошибка [_onTreeMenuSetMainClick]: "' + E.Message + '"');
       ShowMessage('Ошибка [_onTreeMenuSetMainClick]: "' + E.Message + '"');
    end;
  end;
end;

procedure TFmFormats.btnFormatDeleteClick(Sender: TObject);
var
  _Where: string;
  i, _Index:integer;
  _arr: ArrayOfArrayVariant;
begin

  _Index:= fTabFormat.TabIndex;

  OwnerID:= _DBTreeTreeGroupOwner.SelectedItems[0];

  if TabFormats.Tabs.Count=0 then exit;

  if MessageDlg('Удалить формат "'+fTabFormat.Text[_Index]+'" ?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;


  _Where:= IntToStr(OwnerID);

     // если формат единственный, то удаляем записи группы товаров для этого контрагента
   if fBase.SQLReadDS('FORMATS',['ID'],'IDOWNER='+IntToStr(OwnerID),'').DataSet.RecordCount =1 then
         fBase.SQLDelete('PL_GROUP','IDOWNER='+IntToStr(OwnerID),true);

  try
    _Where:= fTabFormat.ValueToString[_Index];
    fBase.SQLDelete('FORMATS','ID='+_Where);


    _arr:=fBase.SQLReadArr('FORMATS',FieldsArr,'IDOWNER='+_Where,'PRIORITY, Name');

    for i:=0 to Length(_arr)-1 do
                    fBase.SQLUpdate('FORMATS',['PRIORITY'],[integer(i)],'ID='+string(_arr[i,0]),true);


    RefrashTabControl(fTabFormat,0);

    SetStatus('Формат успешно удален.');
  except
        on E: Exception do
    begin
       wLog('Formats','Ошибка [DeleteFormat]: "' + E.Message + '"');
       ShowMessage('Ошибка [DeleteFormat]: "' + E.Message + '"');
    end;
  end;

end;

procedure TFmFormats.sgFormatResize(Sender: TObject);
begin
  self.Repaint;
end;

procedure TFmFormats.TabFormatsChange(Sender: TObject);
begin
   if btnSave.Enabled then
      if MessageDlg('Отменить все изменения формата?',mtConfirmation, mbOKCancel, 0) = mrCancel then
         begin
           fTabFormat.TabIndex:= fTabIndexCurrent;
           exit;
         end else
         begin
             fBase.SQLTransactionEnd(false);
         end;

  sgFormat.Tag:= fTabFormat.Value[fTabFormat.TabIndex];

  fFormatsGrid.FillGrid(sgFormat.Tag, self); // заполняем стринг грид;

end;

procedure TFmFormats.tbExportClick(Sender: TObject);
var
  _DataSourse: TDataSource;
  _FPSExport: TFPSExport;
  i: Integer;
  _DateTimeStr, _OwnersWhere, _FileName, _VersionProgram: string;
  _XMLFile: TStringList;
begin

   _OwnersWhere:= fBase.PrepareWhereString('OWN.ID',_DBTreeTreeGroupOwner.SelectedItems);

   if MessageDlg('Экспортировать форматы выбранных контрагентов в файл?',mtWarning, mbOKCancel, 0) = mrCancel then exit;

   DateTimeToString(_DateTimeStr, 'dd_mm_yy_hh-mm-ss', now);

   _VersionProgram:= GetVersion;

   SaveDialog.InitialDir:=PathExport_Unsafe;
   SaveDialog.FileName:= _DateTimeStr+'.fmts';

   if SaveDialog.Execute then
      _FileName:= SaveDialog.FileName else
      exit;

   try
     _FPSExport:= TFPSExport.Create(self);

     try
       _DataSourse:= fBase.SQLReadDS('SELECT '
           +'    '+QuotedStr(_VersionProgram)+' AS VERSION,'
           +'    OWN.NAME OWNERNAME, '
           +'    FMTS.ID, '
           +'    FMTS.IDOWNER, '
           +'    FMTS.PRIORITY, '
           +'    FMTS.NAME, '
           +'    FMTS.FIRSTLINE, '
           +'    FMTS.VENDORCODE, '
           +'    FMTS.FNAME, '
           +'    FMTS.UNIT, '
           +'    FMTS.QUANTITY, '
           +'    FMTS.PRICE, '
           +'    FMTS.FSUM, '
           +'    FMTS.LABEL, '
           +'    FMTS.SCOD, '
           +'    FMTS.CUSTOMSDECLARATION, '
           +'    FMTS.COUNTRY, '
           +'    FMTS.FILE, '
           +'    FMTS.URL, '
           +'    FMTS.SPREADSHEET, '
           +'    FMTS.GROUPSINROWS, '
           +'    FMTS.GROUPS, '
           +'    FMTS.SUBGROUPS1, '
           +'    FMTS.SUBGROUPS2, '
           +'    FMTS.SUBGROUPS3, '
           +'    FMTS.STOCKONLY, '
           +'    FMTS.IDFMTS_CATEGORY, '
           +'    FMTS.IDUSER, '
           +'    FMTS.GROUPALGORITHM, '
           +'    FMTS.FTIMESTAMP, '
           +'    FMTS.FILEHASH, '
           +'    FMTS.STORAGEDAYS, '
           +'    FMTS.IDFILEFORMAT, '
           +'    FMTS.IDCURRENCY, '
           +'    FMTS.CURRENCYPERCENT, '
           +'    FMTS.IDCODEPAGETEXT, '
           +'    FMTS.REMARK, '
           +'    FMTS.FREMARK, '
           +'    FMTS.TRANSIT, '
           +'    FMTS.FTIMESTAMPLASTIMPORT, '
           +'    FMTS.FCLOSE, '
           +'    FMTS.STOCKSYMBOLS, '
           +'    FMTS.YMLID, '
           +'    FMTS.YMLPRICE, '
           +'    FMTS.YMLPRICE2, '
           +'    FMTS.YMLQUANTITY, '
           +'    FMTS.FURLPICTURE, '
           +'    FMTS.FURL, '
           +'    FMTS.PRICE2, '
           +'    FMTS.PRICE3, '
           +'    FMTS.PRICE4, '
           +'    FMTS.PRICE5, '
           +'    FMTS.STOCK2, '
           +'    FMTS.STOCK3, '
           +'    FMTS.STOCK4, '
           +'    FMTS.STOCK5, '
           +'    FMTS.FILEZIPNAMEDECODE, '
           +'    FMTS.STOCKONLYINFO, '
           +'    FMTS.FCOLOR, '
           +'    FMTS.FCONVERTLIBRE, '
           +'    FMTS.IDCSVDELIMITER, '
           +'    FMTS.IDVENDORCODEVARIANT, '
           +'    FMTS.ADDRCELLTEXT, '
           +'    FMTS.OUTCELLTEXT, '
           +'    FMTS.PRICE6, '
           +'    FMTS.PRICE7, '
           +'    FMTS.PRICE8, '
           +'    FMTS.PRICE9, '
           +'    FMTS.PRICE10, '
           +'    FMTS.ADDRCELLFORINVOCE, '
           +'    FMTS.INVOCEDAYS, '
           +'    FMTS.ACTUALDAYS, '
           +'    FMTS.NOMINPRICE '
            //+' FROM "FORMATS" FMTS '
           //+' LEFT JOIN OWNER OWN ON (OWN.ID=FMTS.IDOWNER) '
           +' FROM  OWNER OWN '
           +' RIGHT JOIN "FORMATS" FMTS ON (OWN.ID=FMTS.IDOWNER) '
           +' WHERE ('+_OwnersWhere+')'
           +' ORDER BY OWN.IDPARENT, OWN.NAME, FMTS.NAME ');

       _FPSExport.FileName:= _FileName;
       _FPSExport.ExportFields.Clear;
       _FPSExport.Dataset:= _DataSourse.DataSet;
       _FPSExport.FormatSettings.HeaderRow:= true;
       _FPSExport.FormatSettings.ExportFormat:= efODS;
       //_FPSExport.

       for i:=0 to _DataSourse.DataSet.Fields.Count-1 do
          _FPSExport.ExportFields.AddField(_DataSourse.DataSet.Fields[i].FieldName);

        _FPSExport.Execute;
     finally
       _FPSExport.Free;
     end;

     ShowMessage('Форматы успешно экспортированы');
   except
     on E: Exception do
       begin
          wLog('Formats','Ошибка [Export]: "' + E.Message + '"');
          ShowMessage('Ошибка [Export]: "' + E.Message + '"');
       end;
   end;

end;

function TFmFormats.GetDataString(_ACell: PCell): string;
begin
   if Assigned(_Acell) then

   case _ACell^.ContentType of
     cctNumber: Result := FloatToStr(_ACell^.NumberValue);
     cctDateTime: Result := DateToStr(_ACell^.DateTimeValue);
     else
       Result := _ACell^.UTF8StringValue;
   end
   else
   Result:='';
end;

procedure TFmFormats.tbImportClick(Sender: TObject);
var
  _Form: TFmFormatsImport;
  _FileName, _NAME, _FIRSTLINE, _VENDORCODE, _FNAME, _UNIT, _QUANTITY, _PRICE, _FSUM, _LABEL, _SCOD, _CUSTOMSDECLARATION, _COUNTRY, _FILE, _URL,
    _SPREADSHEET, _GROUPSINROWS, _GROUPS, _SUBGROUPS1, _SUBGROUPS2, _SUBGROUPS3, _STOCKONLY, _IDFMTS_CATEGORY, _IDUSER, _GROUPALGORITHM,
    _FILEHASH, _STORAGEDAYS, _IDFILEFORMAT, _IDCURRENCY, _CURRENCYPERCENT, _IDCODEPAGETEXT, _REMARK, _FREMARK, _TRANSIT, _FCLOSE,
    _STOCKSET, _YMLID, _YMLPRICE, _YMLPRICE2, _YMLQUANTITY, _FURLPICTURE, _FURL, _PRICE2, _PRICE3, _PRICE4, _PRICE5, _PRICE6, _PRICE7, _PRICE8, _PRICE9, _PRICE10,
    _STOCK2, _STOCK3, _STOCK4, _STOCK5, _FILEZIPNAMEDECODE, _STOCKONLYINFO: String;
  _FTIMESTAMP,_FTIMESTAMPLASTIMPORT, _FCOLOR, _FCONVERTLIBRE, _STOCKSYMBOLS, _VERSIONCREATEPROG, _IDCSVDELIMITER, _IDVENDORCODEVARIANT, _ADDRCELLTEXT,
    _OUTCELLTEXT, _ADDRCELLFORINVOCE, _INVOCEDAYS, _ACTUALDAYS,
    _NOMINPRICE: string;
  i, _PRIORITY, iRows, _RootOwner: Integer;
  _arr, _arrOwner: ArrayOfArrayVariant;
  _Worksheet: TsWorksheet;
  _IDOWNER: Longint;
  _IncSpreadSheet: boolean;

  function FromNull(aValue: string):string;
  begin
    if aValue = null then Result:= '0' else Result:= aValue;
  end;

begin

  if OpenDialog.Execute then
     _FileName:= OpenDialog.FileName else
        exit;

 _IncSpreadSheet:= false;
  _Form:= TFmFormatsImport.Create(self);

  _Form.sWorkbookSource1.LoadFromSpreadsheetFile(_FileName,sfOpenDocument,0);

  try
    with _Form do
    begin
      _Worksheet:= sWorkbookSource1.Worksheet;
      StringGrid1.RowCount:=  _Worksheet.GetCellCountInCol(1);
      StringGrid1.Columns[1].PickList.Clear;
      fBase.LongTransaction:= true;
      _arr:=nil;
      _arr:= fBase.SQLReadArr('OWNER',['ID','NAME'],'','IDPARENT,NAME');

      if Assigned(_arr) then
      begin
        StringGrid1.Columns[1].PickList.Add('Создать нового контрагента...');
        StringGrid1.Columns[1].PickList.Add('Добавить к предыдущему');

        for i:=0 to High(_arr) do
           StringGrid1.Columns[1].PickList.Add(string(_arr[i,0])+'|'+string(_arr[i,1]));

           //StringGrid1.Columns[1].
        _VERSIONCREATEPROG:= GetDataString(_Worksheet.GetCell(i,0));

        for i:=1 to StringGrid1.RowCount-1 do
           begin
            StringGrid1.Cells[1,i]:='1';
            _arrOwner:= nil;

            _RootOwner:= fBase.SQLReadArr('OWNER',['ID'],'IDPARENT=0','')[0,0];

            _arrOwner:= fBase.SQLReadArr('OWNER',['ID','NAME'],'NAME='+QuotedStr(GetDataString(_Worksheet.GetCell(i,1))),'');
            if Assigned(_arrOwner) then
                   StringGrid1.Cells[2,i]:= string(_arrOwner[0,0])+'|'+string(_arrOwner[0,1]) else
                   StringGrid1.Cells[2,i]:='Создать нового контрагента...';

            StringGrid1.Cells[3,i]:=GetDataString(_Worksheet.GetCell(i,1));
            StringGrid1.Cells[4,i]:=GetDataString(_Worksheet.GetCell(i,5));
           end;
      end else
      begin
        ShowMessage('Список контрагентов пуст!');
      end;
    end;

    _Form.ShowModal;

    if _Form.ModalResult = mrOK then
    begin
      try
        iRows:=0;

        _IDOWNER:=0;

        for i:=1 to _Form.StringGrid1.RowCount-1 do
           begin

            if _Form.StringGrid1.Cells[1,i] = '1' then
            begin
                Inc(iRows);
                if _Form.StringGrid1.Cells[2,i]<>'Добавить к предыдущему' then
                      TryStrToInt(UTF8Copy(_Form.StringGrid1.Cells[2,i],1,UTF8Pos('|',_Form.StringGrid1.Cells[2,i])-1),_IDOWNER);

                if _IDOWNER = 0 then
                begin
                  _IDOWNER:= fBase.SQLInsert('OWNER',['NAME','IDPARENT'],[GetDataString(_Worksheet.GetCell(i,1)),_RootOwner],false);
                end;

                _PRIORITY:= 0;
                _NAME:= _Form.StringGrid1.Cells[4,i];
                _FIRSTLINE:= GetDataString(_Worksheet.GetCell(i,6));
                _VENDORCODE:= GetDataString(_Worksheet.GetCell(i,7));
                _FNAME:= GetDataString(_Worksheet.GetCell(i,8));
                _UNIT:= GetDataString(_Worksheet.GetCell(i,9));
                _QUANTITY:= GetDataString(_Worksheet.GetCell(i,10));
                _PRICE:= GetDataString(_Worksheet.GetCell(i,11));
                _FSUM:= GetDataString(_Worksheet.GetCell(i,12));
                _LABEL:= GetDataString(_Worksheet.GetCell(i,13));
                _SCOD:= GetDataString(_Worksheet.GetCell(i,14));
                _CUSTOMSDECLARATION:= GetDataString(_Worksheet.GetCell(i,15));
                _COUNTRY:= GetDataString(_Worksheet.GetCell(i,16));
                _FILE:= GetDataString(_Worksheet.GetCell(i,17));
                _URL:= StringReplace(GetDataString(_Worksheet.GetCell(i,18)),'&','&amp;',[rfReplaceAll]);
                _SPREADSHEET:= GetDataString(_Worksheet.GetCell(i,19));
                _GROUPSINROWS:= GetDataString(_Worksheet.GetCell(i,20));
                _GROUPS:= GetDataString(_Worksheet.GetCell(i,21));
                _SUBGROUPS1:= GetDataString(_Worksheet.GetCell(i,22));
                _SUBGROUPS2:= GetDataString(_Worksheet.GetCell(i,23));
                _SUBGROUPS3:= GetDataString(_Worksheet.GetCell(i,24));
                _STOCKONLY:= GetDataString(_Worksheet.GetCell(i,25));
                _IDFMTS_CATEGORY:= GetDataString(_Worksheet.GetCell(i,26));
                _IDUSER:= GetDataString(_Worksheet.GetCell(i,27));
                _GROUPALGORITHM:= GetDataString(_Worksheet.GetCell(i,28));
                _FTIMESTAMP:= GetDataString(_Worksheet.GetCell(i,29));
                _FILEHASH:=  '';
                _STORAGEDAYS:= GetDataString(_Worksheet.GetCell(i,31));
                _IDFILEFORMAT:= GetDataString(_Worksheet.GetCell(i,32));
                _IDCURRENCY:= GetDataString(_Worksheet.GetCell(i,33));
                _CURRENCYPERCENT:= GetDataString(_Worksheet.GetCell(i,34));
                _IDCODEPAGETEXT:= GetDataString(_Worksheet.GetCell(i,35));
                _REMARK:= GetDataString(_Worksheet.GetCell(i,36));
                _FREMARK:= GetDataString(_Worksheet.GetCell(i,37));
                _TRANSIT:= GetDataString(_Worksheet.GetCell(i,38));
                _FTIMESTAMPLASTIMPORT:= GetDataString(_Worksheet.GetCell(i,39));
                _FCLOSE:= GetDataString(_Worksheet.GetCell(i,40));
                _STOCKSYMBOLS:= GetDataString(_Worksheet.GetCell(i,41));
                _YMLID:= GetDataString(_Worksheet.GetCell(i,42));
                _YMLPRICE:= GetDataString(_Worksheet.GetCell(i,43));
                _YMLPRICE2:= GetDataString(_Worksheet.GetCell(i,44));
                _YMLQUANTITY:= GetDataString(_Worksheet.GetCell(i,45));
                _FURLPICTURE:= GetDataString(_Worksheet.GetCell(i,46));
                _FURL:= GetDataString(_Worksheet.GetCell(i,47));
                _PRICE2:= GetDataString(_Worksheet.GetCell(i,48));
                _PRICE3:= GetDataString(_Worksheet.GetCell(i,49));
                _PRICE4:= GetDataString(_Worksheet.GetCell(i,50));
                _PRICE5:= GetDataString(_Worksheet.GetCell(i,51));
                _STOCK2:= GetDataString(_Worksheet.GetCell(i,52));
                _STOCK3:= GetDataString(_Worksheet.GetCell(i,53));
                _STOCK4:= GetDataString(_Worksheet.GetCell(i,54));
                _STOCK5:= GetDataString(_Worksheet.GetCell(i,55));
                _FILEZIPNAMEDECODE:= GetDataString(_Worksheet.GetCell(i,56));
                _STOCKONLYINFO:= GetDataString(_Worksheet.GetCell(i,57));
                _FCOLOR:= FromNull(GetDataString(_Worksheet.GetCell(i,58)));
                _FCONVERTLIBRE:= FromNull(GetDataString(_Worksheet.GetCell(i,59)));

                _IDCSVDELIMITER:= FromNull(GetDataString(_Worksheet.GetCell(i,60)));
                _IDVENDORCODEVARIANT:= FromNull(GetDataString(_Worksheet.GetCell(i,61)));
                _ADDRCELLTEXT:= GetDataString(_Worksheet.GetCell(i,62));
                _OUTCELLTEXT:= GetDataString(_Worksheet.GetCell(i,63));

                _PRICE6:= GetDataString(_Worksheet.GetCell(i,64));
                _PRICE7:= GetDataString(_Worksheet.GetCell(i,65));
                _PRICE8:= GetDataString(_Worksheet.GetCell(i,66));
                _PRICE9:= GetDataString(_Worksheet.GetCell(i,67));
                _PRICE10:= GetDataString(_Worksheet.GetCell(i,68));
                _ADDRCELLFORINVOCE:= GetDataString(_Worksheet.GetCell(i,69));
                _INVOCEDAYS:= GetDataString(_Worksheet.GetCell(i,70));
                _ACTUALDAYS:= GetDataString(_Worksheet.GetCell(i,71));
                _NOMINPRICE:= GetDataString(_Worksheet.GetCell(i,72));

                if Length(_FCOLOR)=0 then _FCOLOR:= '0';

                  fBase.SQLInsert('FORMATS',[
                  'IDOWNER',
                  'PRIORITY',
                  'NAME',
                  'FIRSTLINE',
                  'VENDORCODE',
                  'FNAME',
                  'UNIT',
                  'QUANTITY',
                  'PRICE',
                  'FSUM',
                  'LABEL',
                  'SCOD',
                  'CUSTOMSDECLARATION',
                  'COUNTRY',
                  'FILE',
                  'URL',
                  'SPREADSHEET',
                  'GROUPSINROWS',
                  'GROUPS',
                  'SUBGROUPS1',
                  'SUBGROUPS2',
                  'SUBGROUPS3',
                  'STOCKONLY',
                  'IDFMTS_CATEGORY',
                  'IDUSER',
                  'GROUPALGORITHM',
                  'FTIMESTAMP',
                  'FILEHASH',
                  'STORAGEDAYS',
                  'IDFILEFORMAT',
                  'IDCURRENCY',
                  'CURRENCYPERCENT',
                  'IDCODEPAGETEXT',
                  'REMARK',
                  'FREMARK',
                  'TRANSIT',
                  'FTIMESTAMPLASTIMPORT',
                  'FCLOSE',
                  'STOCKSYMBOLS',
                  'YMLID',
                  'YMLPRICE',
                  'YMLPRICE2',
                  'YMLQUANTITY',
                  'FURLPICTURE',
                  'FURL',
                  'PRICE2',
                  'PRICE3',
                  'PRICE4',
                  'PRICE5',
                  'PRICE6',
                  'PRICE7',
                  'PRICE8',
                  'PRICE9',
                  'PRICE10',
                  'STOCK2',
                  'STOCK3',
                  'STOCK4',
                  'STOCK5',
                  'FILEZIPNAMEDECODE',
                  'STOCKONLYINFO',
                  'FCOLOR',
                  'FCONVERTLIBRE',
                  'IDCSVDELIMITER',
                  'IDVENDORCODEVARIANT',
                  'ADDRCELLTEXT',
                  'OUTCELLTEXT',
                  'ADDRCELLFORINVOCE',
                  'INVOCEDAYS',
                  'ACTUALDAYS',
                  'NOMINPRICE'
                                ],[
                                _IDOWNER,
                                _PRIORITY,
                                _NAME,
                                _FIRSTLINE,
                                _VENDORCODE,
                                _FNAME,
                                _UNIT,
                                _QUANTITY,
                                _PRICE,
                                _FSUM,
                                _LABEL,
                                _SCOD,
                                _CUSTOMSDECLARATION,
                                _COUNTRY,
                                _FILE,
                                _URL,
                                _SPREADSHEET,
                                _GROUPSINROWS,
                                _GROUPS,
                                _SUBGROUPS1,
                                _SUBGROUPS2,
                                _SUBGROUPS3,
                                _STOCKONLY,
                                _IDFMTS_CATEGORY,
                                _IDUSER,
                                _GROUPALGORITHM,
                                _FTIMESTAMP,
                                _FILEHASH,
                                _STORAGEDAYS,
                                _IDFILEFORMAT,
                                _IDCURRENCY,
                                _CURRENCYPERCENT,
                                _IDCODEPAGETEXT,
                                _REMARK,
                                _FREMARK,
                                _TRANSIT,
                                _FTIMESTAMPLASTIMPORT,
                                _FCLOSE,
                                _STOCKSYMBOLS,
                                _YMLID,
                                _YMLPRICE,
                                _YMLPRICE2,
                                _YMLQUANTITY,
                                _FURLPICTURE,
                                _FURL,
                                _PRICE2,
                                _PRICE3,
                                _PRICE4,
                                _PRICE5,
                                _PRICE6,
                                _PRICE7,
                                _PRICE8,
                                _PRICE9,
                                _PRICE10,
                                _STOCK2,
                                _STOCK3,
                                _STOCK4,
                                _STOCK5,
                                _FILEZIPNAMEDECODE,
                                _STOCKONLYINFO,
                                _FCOLOR,
                                _FCONVERTLIBRE,
                                _IDCSVDELIMITER,
                                _IDVENDORCODEVARIANT,
                                _ADDRCELLTEXT,
                                _OUTCELLTEXT,
                                _ADDRCELLFORINVOCE,
                                _INVOCEDAYS,
                                _ACTUALDAYS,
                                _NOMINPRICE
                                ],false);
            end;

            // если формат добавлен впервые, то создаем корень группы товаров для этого контрагента
            if fBase.SQLReadDS('FORMATS',['ID'],'IDOWNER='+IntToStr(_IDOWNER),'').DataSet.RecordCount=1 then
               if fBase.SQLReadDS('PL_GROUP',['ID'],'IDOWNER='+IntToStr(_IDOWNER),'').DataSet.RecordCount=0 then
                     fBase.SQLInsert('PL_GROUP',['IDOWNER','IDPARENT','NAME'],[_IDOWNER,integer(0),'Номенклатура'],false);

           end;
        fBase.SQLTransactionEnd(true);


        ShowMessage('Успешно испортировано форматов: '+IntToStr(iRows));

        _DBTreeTreeGroupOwner.Fill();

      except
        on E: Exception do
         begin
           fBase.SQLTransactionEnd(false);
           __Log.SaveLogError(E);
           wLog('Formats','Ошибка [Import]: "' + E.Message + '"');
         end;
      end;
    end;

  finally
    _Form.Free;
  end;
end;

procedure TFmFormats.tbTreeBtnAddClick(Sender: TObject);
begin
  _DBTreeTreeGroupOwner.onTreePopupMenuAdd(Self);
end;

procedure TFmFormats.tbTreeBtnDeleteClick(Sender: TObject);
begin
  _DBTreeTreeGroupOwner.onTreePopupMenuDelete(Self);
end;

procedure TFmFormats.tbTreeBtnRenameClick(Sender: TObject);
begin
  _DBTreeTreeGroupOwner.onTreePopupMenuRename(Self);
end;

procedure TFmFormats.tbTreeBtnSetMainClick(Sender: TObject);
begin
  _onTreeMenuSetMainClick(Self);
end;

procedure TFmFormats.RefrashTabControl(Tab: TwTab; _Index: integer);
var
  _arr: ArrayOfArrayVariant;
  i: integer;
  _Where: string;
begin
  // обновление таб контрол после изменений

  fTabIndexCurrent:= _Index;

  _Where:= IntToStr(OwnerID);

  screen.Cursor:= crSQLWait;

  try
    _arr:=fBase.SQLReadArr('FORMATS',FieldsArr,'IDOWNER='+_Where,'PRIORITY, Name');

    if Length(_arr)<1 then
    begin
      Tab.Visible:=false;
      Tab.Clear;
      exit;
    end else
    begin
      Tab.Clear;
      Tab.Visible:=true;
    end;

    for i:=0 to Length(_arr)-1 do
    begin
         Tab.Add(string(_arr[i,1]),integer(_arr[i,0]));
    end;
    if _Index >-1 then
      begin
        Tab.TabIndex:=_Index;
      end;

    sgFormat.Tag:= fTabFormat.Value[fTabFormat.TabIndex];

    fFormatsGrid.FillGrid(sgFormat.Tag, self); // заполняем стринг грид;


    _arr:=nil;
  finally
    screen.Cursor:= crDefault;
  end;

end;

procedure TFmFormats.TreeGroupOwnerClick(Sender: TObject);
var
  _Tree: TTreeView;
begin
  _Tree:=(Sender as TTreeView);

   if btnSave.Enabled then
      if MessageDlg('Отменить все изменения формата?',mtConfirmation, mbOKCancel, 0) = mrCancel then
         begin
           _DBTreeTreeGroupOwner.FindNodeWithDataInt(OwnerID);
           exit;
         end else
         begin
             fBase.SQLTransactionEnd(false);
         end;

  if _Tree.Selected.Count>0 then
     begin
       fTabFormat.Clear;
       fTabFormat.Visible:=false;
       exit;
     end;

  OwnerID:= TTreeData(_Tree.Selected.Data).Value;

  RefrashTabControl(fTabFormat,0);

end;

procedure TFmFormats.TreeGroupOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
  if TTreeData(Node.Data).Value = IdMainOwner then
  begin
    Node.ImageIndex:=2;
    exit;
  end;

  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmFormats.TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

end.

