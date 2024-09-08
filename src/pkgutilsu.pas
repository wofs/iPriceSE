unit pkgUtilsU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Buttons, CheckLst, Classes, ComCtrls, Controls, db, Dialogs, ExtCtrls, LazFileUtils,
  FileUtil, Forms, fpcsvexport, fpdbfexport,
  fpexprpars, fpspreadsheet
  , fpspreadsheetctrls, fpsTypes, Graphics, Grids, LazUTF8,
  mUtilsU, StdCtrls, SysUtils, wBaseU, wFuncU, wLogU, wPlugin, wZipperU, wTypesU;

type

  { TFmUtils }

  TFmUtils = class(TForm)
    btnDataExport: TBitBtn;
    btnDBRestore: TBitBtn;
    btnDataImport: TBitBtn;
    btnDBBackup: TBitBtn;
    btnImport: TBitBtn;
    btnOpenFile: TBitBtn;
    CheckBox2: TCheckBox;
    CheckBox4: TCheckBox;
    CheckBox0: TCheckBox;
    CheckBox3: TCheckBox;
    CheckBox1: TCheckBox;
    CheckBox5: TCheckBox;
    eOurCode: TLabeledEdit;
    eQIP1:    TLabeledEdit;
    eQIP2:    TLabeledEdit;
    eQIP3:    TLabeledEdit;
    GroupBox1: TGroupBox;
    ImageList16: TImageList;
    eVendorCode: TLabeledEdit;
    Label1:   TLabel;
    mFormula: TMemo;
    mLog:     TMemo;
    od1:      TOpenDialog;
    PageControl1: TPageControl;
    pcReservVariants: TPageControl;
    Panel5:   TPanel;
    Panel6:   TPanel;
    Panel7:   TPanel;
    Panel8:   TPanel;
    pcImport: TPageControl;
    Panel2:   TPanel;
    Panel3:   TPanel;
    Panel4:   TPanel;
    pcUtils:  TPageControl;
    Panel1:   TPanel;
    pbAll:    TProgressBar;
    pbOne:    TProgressBar;
    SaveDialog: TSaveDialog;
    StringGridImport: TStringGrid;
    sWorkbookSource1: TsWorkbookSource;
    tsDBService: TTabSheet;
    tsDBase:  TTabSheet;
    tsData:   TTabSheet;
    tsImportMatching: TTabSheet;
    tsBackupRestore: TTabSheet;
    tsImportData: TTabSheet;
    procedure btnDataExportClick(Sender: TObject);
    procedure btnDataImportClick(Sender: TObject);
    procedure btnDBBackupClick(Sender: TObject);
    procedure btnDBRestoreClick(Sender: TObject);
    procedure btnImportClick(Sender: TObject);
    procedure btnOpenFileClick(Sender: TObject);
    procedure CheckBox4Change(Sender: TObject);
    procedure CheckBox3Change(Sender: TObject);
    procedure CheckBox5Change(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    fBase:   TwBase;
    Utils:   TUtils;
    _FormID: string;
    IdMainOwner: longint;
    procedure fIF(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    function GetDataString(_ACell: PCell): string;
    procedure ImportData(aIdOwner: integer; aFileName: string;
      aParser: TFpExpressionParser; aClear: boolean);
    procedure LogBackup(AText: string);
    procedure SetStatus(_Text: string);
    property wFormID: string read _FormID write _FormID;
  public

  end;

var
  FmUtils: TFmUtils;

const
  cChechBoxCount = 6;

implementation

uses
  FmMainU;

{$R *.lfm}

{ TFmUtils }

procedure TFmUtils.LogBackup(AText: string);
begin
  mLog.Lines.Add(DateTimeToStr(now)+' | '+AText);
  wLog('[console.log]', AText);
end;

procedure TFmUtils.SetStatus(_Text: string);
begin
  wStatus(self.Name, _Text, True);
end;

procedure TFmUtils.btnDBBackupClick(Sender: TObject);
var
  i: integer;
  LinesCount: integer;
  _DateTimeStr, _Patch, _BackupFileName: string;
begin
  try
    SetStatus('Запущена операция резервного копирования...');
    LogBackup('Соединяюсь с БД...');

    DateTimeToString(_DateTimeStr, 'dd-mm-yy_hh-mm-ss', now);

    _Patch := ExtractFileDir(Application.ExeName);

    if not DirectoryExistsUTF8(_Patch+'\'+'BackupDB') then
      ForceDirectoriesUTF8(_Patch+'\'+'BackupDB');

    _BackupFileName := _Patch+'\BackupDB\'+_DateTimeStr+'_db.fbk';

    fBase.BackupBase(_BackupFileName);
    SetStatus('Резервное копирование успешно завершено.');
  except

  end;
end;

//экспорт данных в файл
procedure TFmUtils.btnDataExportClick(Sender: TObject);
var
  i: integer;
  _CheckBox: TControl;
  _ExportObjects: TExportObjects;
begin
 _ExportObjects:=[];

  for i := 0 to cChechBoxCount-1 do
  begin
    _CheckBox := nil;
    _CheckBox := TsData.FindChildControl('CheckBox'+IntToStr(i));

    if Assigned(_CheckBox) then
      if TCheckBox(_CheckBox).Checked then
        case i of
          0: _ExportObjects:= _ExportObjects+ [eoOwner];
          1: _ExportObjects:= _ExportObjects+ [eoFormats];
          2: _ExportObjects:= _ExportObjects+ [eoCatalog];
          3: _ExportObjects:= _ExportObjects+ [eoPrice];
          4: _ExportObjects:= _ExportObjects+ [eoMatchings];
          5: _ExportObjects:= _ExportObjects+ [eoPriceVersions];
        end;

  end;

  Utils.ExportData(_ExportObjects,[emData, emSaveFile]);

end;

//импорт данных из файла
procedure TFmUtils.btnDataImportClick(Sender: TObject);
var
  _FileName: String;

  _Zipper: TwZipper;
  _FileList: TStringList;
begin
  if MessageDlg(
    'Импортировать выбранные данные в БД из внешнего файла?',
    mtConfirmation, mbOKCancel, 0) = mrCancel then
    exit;

  if (fBase.SQLReadDS('SELECT ID FROM PL_ITEMS ROWS 1').DataSet.RecordCount>0)
  or (fBase.SQLReadDS('SELECT ID FROM CATALOG ROWS 1').DataSet.RecordCount>0)
  or (fBase.SQLReadDS('SELECT ID FROM OWNER ROWS 3').DataSet.RecordCount>2)
  then
    begin
      ShowMessage('Импорт данных возможен только в пустую БД! Очистите БД (Операции->Очистка БД) и повторите попытку импорта.');
      exit;
    end;


  od1.FileName:='';
  od1.Filter:='Backup Data Package|*.bdpkg';

  if not od1.Execute then
    exit;

  _FileName := od1.FileName;

  _Zipper:= TwZipper.Create();
  try
    _FileList:= _Zipper.ReadFileList(_FileName);

    if MessageDlg(
      'Найденные таблицы:'+LineEnding+_FileList.Text+LineEnding+'Импортировать их в БД?',
      mtConfirmation, mbOKCancel, 0) = mrCancel then
      exit;
  finally
    _Zipper.Destroy();
  end;

  Utils.ImportData(_FileName);

end;

procedure TFmUtils.btnDBRestoreClick(Sender: TObject);
var
  i:      integer;
  _UtilsIndex, _PluginPageIndex: integer;
  _Patch: string;
  _FileName: string;
begin
  if db_portable then
    begin
    if MessageDlg(
      'Восстановить БД из резервной копии? ВНИМАНИЕ! Вся существующая информация в БД будет утеряна!!!', mtWarning, mbOKCancel, 0) = mrCancel then
      exit;
    end else
    begin
      if MessageDlg(
        'Восстановить БД из резервной копии? Резервная копия будет восстановлена в файл на сервере базы данных с именем Restored_'+ExtractFileName(db_dbName), mtWarning, mbOKCancel, 0) = mrCancel then
        exit;
    end;

  if MessageDlg(
    'Все открытые окна будут закрыты без сохранения. Продолжить?',
    mtConfirmation, mbOKCancel, 0) = mrCancel then
    exit;
  _FileName:= EmptyStr;
  try
    _UtilsIndex := __wPluginGetForm(wFormID).Index;
    // выгружаем подгруженные плагины
    if Plugin<>nil then
    begin
      for  i := Plugin.Count-1 downto 0 do
        if i<>_UtilsIndex then
        begin
          _PluginPageIndex := Plugin[i].PageIndex;
          try
            Plugin[i].Unload();
          finally
            FmMain.pcPlugins.Pages[_PluginPageIndex].Free;
          end;
        end;
      Plugin[0].PageIndex := 0;
    end;

    if db_portable then
    begin
    _Patch := includeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
    if not DirectoryExistsUTF8(_Patch+'BackupDB') then
      ForceDirectoriesUTF8(_Patch+'BackupDB');

    od1.InitialDir := _Patch+'BackupDB';
    od1.Filter:='Резервные копии|*.fbk';

    if not od1.Execute then exit;
      _FileName := od1.FileName;
    end;

      SetStatus(
        'Запущена операция восстановления из резервной копии...');
      LogBackup('Соединяюсь с БД...');
      fBase.DataBase.Connected := False;

      fBase.RestoreBase(_FileName);

      fBase.DataBase.Connected := True;
      LogBackup('Синхронизация версии БД с версией программы...');

      try
        //
        FmMain.CheckDBVersion(fBase);

      finally
        LogBackup(
          'Синхронизация версии БД с версией программы...[ОК]');
        if not db_portable then
          LogBackup('-= ВНИМАНИЕ: чтобы закончить процесс восстановления, обязательно выполните инструкции выше. =-')
        else
          LogBackup('-= Все операции восстановления завершены =-');

      end;

      SetStatus(
        'Операция восстановления БД из резервной копии успешно завершена.');

      if not db_portable then
        SetStatus(
          'ВНИМАНИЕ: чтобы закончить процесс восстановления, обязательно выполните инструкции выше.');
  except

  end;

end;

procedure TFmUtils.fIF(var Result: TFPExpressionResult;
  const Args: TExprParameterArray);
begin
  if Args[0].resBoolean then
    Result.resfloat := Args[1].resfloat
  else
    Result.resfloat := Args[2].resfloat;
end;

function TFmUtils.GetDataString(_ACell: PCell): string;
begin
  if Assigned(_Acell) then

    case _ACell^.ContentType of
      cctNumber: Result   := FloatToStr(_ACell^.NumberValue);
      cctDateTime: Result := DateToStr(_ACell^.DateTimeValue);
      else
        Result := _ACell^.UTF8StringValue;
    end
  else
    Result := '';
end;

procedure TFmUtils.ImportData(aIdOwner: integer; aFileName: string;
  aParser: TFpExpressionParser; aClear: boolean);
var
  _Worksheet:      TsWorksheet;
  _WorksheetCount: cardinal;
  _QUANTITYINPACKING: double;
  _VENDORCODE, _OURCODE: string;
  _eOurCode, _eVendorCode, _eQIP1, _eQIP2, _eQIP3: longint;
  i, _IDCATALOG, _PriceRoot, _IdPrice: integer;
  _arr, _PIRCEarr: ArrayOfArrayVariant;
begin
  try
    fBase.LongTransaction := True;

    if aClear then
      fBase.SQLDelete('CATALOG_MATCHING', 'IDOWNER='+IntToStr(aIdOwner), False);

    _PriceRoot := 0;
    _IdPrice   := 0;

    _PriceRoot := fBase.SQLReadArr('PL_GROUP', ['ID'], 'IDOWNER='+
      IntToStr(aIdOwner), '')[0, 0];

    //FWorkBook. := true;
    sWorkbookSource1.AutoDetectFormat := True;
    sWorkbookSource1.FileName := aFileName;
    //sWorkbookSource1.LoadFromSpreadsheetFile(aFileName,sfCSV,0);
    _Worksheet      := sWorkbookSource1.Worksheet;
    _WorksheetCount := _Worksheet.GetCellCountInCol(0);

    TryStrToInt(eOurCode.Text, _eOurCode);
    TryStrToInt(eVendorCode.Text, _eVendorCode);
    TryStrToInt(eQIP1.Text, _eQIP1);
    TryStrToInt(eQIP2.Text, _eQIP2);
    TryStrToInt(eQIP3.Text, _eQIP3);

    //if (_eOurCode>0) and (_eVendorCode>0) then
    pbOne.Position := 0;
    pbOne.Max      := _WorksheetCount;
    pbOne.Step     := 1;
    SetStatus('Импортирую файл: '+aFileName);

    for i := 0 to _WorksheetCount do
    begin
      //wLog('[debug]',IntTOStr(i)+'|'+GetDataString(_Worksheet.GetCell(i,_eOurCode-1))+'|'+GetDataString(_Worksheet.GetCell(i,_eVendorCode-1))+'|'+GetDataString(_Worksheet.GetCell(i,_eQIP1-1)));

      _OURCODE := GetDataString(_Worksheet.GetCell(i, _eOurCode-1));

      _arr := fBase.SQLReadArr('SELECT ID, NAME FROM CATALOG WHERE VENDORCODE='+
        QuotedStr(_OURCODE)+' AND IDOWNER='+IntToStr(IdMainOwner));
      if Assigned(_arr) then
      begin
        _IDCATALOG := integer(_arr[0, 0]);

        if _eQIP1>0 then
          aParser.Identifiers[0].AsFloat := _Worksheet.GetCell(i, _eQIP1-1)^.Numbervalue
        else
          aParser.Identifiers[0].AsFloat := 0;

        if _eQIP2>0 then
          aParser.Identifiers[1].AsFloat :=
            _Worksheet.GetCell(i, _eQIP2-1)^.Numbervalue
        else
          aParser.Identifiers[1].AsFloat := 0;

        if _eQIP3>0 then
          aParser.Identifiers[2].AsFloat := _Worksheet.GetCell(i, _eQIP3-1)^.Numbervalue
        else
          aParser.Identifiers[2].AsFloat := 0;

        _QUANTITYINPACKING := aParser.Evaluate.ResFloat;
        _VENDORCODE := GetDataString(_Worksheet.GetCell(i, _eVendorCode-1));

        if Length(_VENDORCODE)>0 then
          begin
            _PIRCEarr := nil;
          _PIRCEarr   := fBase.SQLReadArr('PL_ITEMS', ['ID'], 'VENDORCODE='+
            QuotedStr(_VENDORCODE)+' AND IDOWNER='+IntToStr(aIdOwner), '');
          if Assigned(_PIRCEarr) then
            fBase.SQLInsert('CATALOG_MATCHING',
              ['IDCATALOG', 'IDOWNER', 'IDPL_ITEMS', 'QUANTITYINPACKING',
              'IDUSER', 'FTIMESTAMP'],
              [_IDCATALOG, aIdOwner, _PIRCEarr[0, 0], _QUANTITYINPACKING,
              integer(1), now],
              'IDPL_ITEMS',
              False
              )
          else
          begin
            _IdPrice := fBase.SQLInsert('PL_ITEMS',
              ['IDPL_GROUP', 'IDOWNER', 'IDFORMATS', 'NAME', 'UNIT',
              'LABEL', 'REMARK', 'FURL', 'FURLPICTURE', 'VENDORCODE', 'FTIMESTAMP'],
              [_PriceRoot, aIdOwner, integer(0),
              'Создано процедурой импорта соответствий',
              '', '', '', '', '', string(_VENDORCODE), now], False);
            fBase.SQLInsert('CATALOG_MATCHING',
              ['IDCATALOG', 'IDOWNER', 'IDPL_ITEMS', 'QUANTITYINPACKING',
              'IDUSER', 'FTIMESTAMP'],
              [_IDCATALOG, aIdOwner, _IdPrice, _QUANTITYINPACKING,
              integer(1), now],


              'IDPL_ITEMS',
              False
              );
          end;
          end;
      end;
      if i mod 20 = 0 then
      begin
        pbOne.StepBy(i-pbOne.Position);
        Application.ProcessMessages;
      end;

    end;

    pbOne.Position := pbOne.Max;
    fBase.SQLTransactionEnd(True);
  except
    on E: Exception do
    begin
      fBase.SQLTransactionEnd(False);
      __Log.SaveLogError(E);
      wLog('Utils', 'Ошибка [ImportData]: "'+E.Message+'"');
    end;
  end;

end;

procedure TFmUtils.btnImportClick(Sender: TObject);
var
  _Parser:  TFpExpressionParser;
  i, iRows: integer;
  _IDOWNER: longint;
  _CLEAR:   boolean;
begin
  try
    _Parser := TFpExpressionParser.Create(TComponent(self));
    btnImport.Enabled := False;
    btnOpenFile.Enabled := False;
    Screen.Cursor := crSQLWait;
    try
      _Parser.Builtins := [bcMath]+[bcBoolean];
      with _Parser.Identifiers do
      begin
        AddFloatVariable('QIP1', 0);
        AddFloatVariable('QIP2', 0);
        AddFloatVariable('QIP3', 0);
        AddFunction('IF', 'F', 'BFF', @fIF);  // если (полный аналог IFF)
      end;
      _Parser.Expression := mFormula.Text;

      iRows      := 0;
      pbAll.Position := 0;
      pbAll.Max  := StringGridImport.RowCount-1;
      pbAll.Step := 1;
      for i := 1 to StringGridImport.RowCount-1 do
      begin

        TryStrToInt(UTF8Copy(StringGridImport.Cells[2, i], 1,
          UTF8Pos('|', StringGridImport.Cells[2, i])-1), _IDOWNER);

        if StringGridImport.Cells[1, i] = '1' then
          _CLEAR := True
        else
          _CLEAR := False;

        if _IDOWNER>0 then
        begin
          Inc(iRows);
          ImportData(_IDOWNER, StringGridImport.Cells[3, i], _Parser, _CLEAR);
          pbAll.StepIt;
          SetStatus('Импортировано '+IntToStr(
            iRows)+' из '+IntToStr(pbAll.Max));
          Application.ProcessMessages;
        end;

      end;

      ShowMessage('Успешно импортировано файлов: '+
        IntToStr(iRows));

    finally
      Screen.Cursor     := crDefault;
      btnImport.Enabled := True;
      btnOpenFile.Enabled := True;
      _Parser.Free;
    end;
  except
    on E: Exception do
    begin
      __Log.SaveLogError(E);
      wLog('Utils', 'Ошибка [Import]: "'+E.Message+'"');
    end;
  end;

end;

procedure TFmUtils.btnOpenFileClick(Sender: TObject);
var
  i:    integer;
  _arr: ArrayOfArrayVariant;
begin
  od1.Filter:='All supported files|*.xls;*.xlsx;*.xlsm;*.ods;*.csv;|All spreadsheet files|*.xls;*.xlsx;*.ods;*.csv|All Excel files (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm|LibreOffice/OpenOffice spreadsheets (*.ods)|*.ods|Comma-separated text files (*.csv)|*.csv;';
  if not od1.Execute then
    exit;

  StringGridImport.RowCount := od1.Files.Count+1;
  for i := 0 to od1.Files.Count-1 do
  begin
    StringGridImport.Cells[1, i+1] := '0';
    StringGridImport.Cells[3, i+1] := od1.Files[i];
  end;

  _arr := nil;
  _arr := fBase.SQLReadArr(
    'SELECT DISTINCT OWN.ID, OWN.NAME FROM OWNER OWN INNER JOIN FORMATS FMTS ON (OWN.ID= FMTS.IDOWNER) ORDER BY OWN.IDPARENT, OWN.NAME');

  if Assigned(_arr) then
  begin
    StringGridImport.Columns[1].PickList.Clear;
    for i := 0 to High(_arr) do
      StringGridImport.Columns[1].PickList.Add(
        string(_arr[i, 0])+'|'+string(_arr[i, 1]));
  end;
end;

procedure TFmUtils.CheckBox4Change(Sender: TObject);
begin
  if TCheckBox(Sender).Checked then
      if MessageDlg(
      'Для экспорта/импорта соответствий необходим экспорт/импорт каталога и прайс-листов. Отметить каталог и прайс-листы на экспорт/импорт?', mtWarning, mbOKCancel, 0) = mrCancel then
          TCheckBox(Sender).Checked := False
      else
      begin
        CheckBox2.Checked:= true;
        CheckBox3.Checked:= true;
      end;

end;

procedure TFmUtils.CheckBox3Change(Sender: TObject);
begin
  if TCheckBox(Sender).Checked then
    if MessageDlg(
      'Таблица прайс-листов обычно содержит несколько сотен тысяч строк. Экспорт/импорт такого количества данных займет несколько минут времени. Вы уверены, что ходите включить прайс-листы в экспорт/импорт?', mtWarning, mbOKCancel, 0) = mrCancel then
          TCheckBox(Sender).Checked := False;

end;

procedure TFmUtils.CheckBox5Change(Sender: TObject);
begin
  if TCheckBox(Sender).Checked then
    if MessageDlg(
      'Архив прайс-листов обычно содержит сотни тысяч (иногда миллионы) строк. Экспорт/импорт такого количества данных займет много времени. Вы уверены, что ходите включить архив в экспорт/импорт?', mtWarning, mbOKCancel, 0) = mrCancel then
          TCheckBox(Sender).Checked := False;
end;

procedure TFmUtils.FormCreate(Sender: TObject);
begin
  try
    wFormID := Self.Name;
    fBase   := TwBase.Create(self);
    fBase.Memo:= mLog;

    Utils      := TUtils.Create(self, fBase);
    Utils.Memo := mLog;

    TryStrToInt(fBase.ReadSettingByName('setDefaultOwner'), IdMainOwner);
    // считываем настройки - текущий основной прайс-лист
  except
    on E: Exception do
    begin
      SetStatus('Сбой инициализации плагина.');
      wLog('Utils', 'Ошибка [FmCreate]: "'+E.Message+'"');
      wLog('Utils', 'Сбой инициализации плагина.');
      // ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

    end;
  end;
end;

procedure TFmUtils.FormDestroy(Sender: TObject);
begin
  //inherited Destroy();
  Utils.Destroy();
  fBase.Destroy();
end;

end.
