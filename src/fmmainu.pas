unit FmMainU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  LCLIntf, mCatalogU, mInvoceU, StdCtrls, SysUtils, Forms, Controls, Dialogs, ComCtrls,
  ExtCtrls, Menus, wReportU, wTProgressU, XMLPropStorage, FmPriceFieldsU, INIFiles,
  wPlugin, pkgCatalogU, pkgFormatsU, pkgPricesU, pkgAnalisisU, pkgOrdersU,
  pkgUtilsU, FmAboutU, FmWaitU, Clipbrd,
  wLogU, DOM, xmlread,
  wBaseU, wDBImportU, wGetU, wFuncU, wZipperU, wTypesU,
  DB, Graphics, ColorBox, LazUTF8, LazFileUtils, FileUtil,
  md5, IBDatabase, Classes;

type

  { TFmMain }

  TFmMain = class(TForm)
    ILpcPlugins: TImageList;
    ImageListTray: TImageList;
    mClearProp: TMenuItem;
    mmChangeLog: TMenuItem;
    mmBag: TMenuItem;
    mmHelp: TMenuItem;
    mExportCatalogInCSV: TMenuItem;
    mExportCatalogInSpreadsheet: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem7: TMenuItem;
    mmRestoreFormSize: TMenuItem;
    mShowMainForm: TMenuItem;
    MenuItem9: TMenuItem;
    mViewInTray: TMenuItem;
    mmViews: TMenuItem;
    mmBooksPriceField: TMenuItem;
    mmBooks: TMenuItem;
    MenuItem6: TMenuItem;
    mm_ClearDB: TMenuItem;
    MenuItem8: TMenuItem;
    mmAbout: TMenuItem;
    pics24: TImageList;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    mClose: TMenuItem;
    MainMenu: TMainMenu;
    pcPlugins: TPageControl;
    Panel2: TPanel;
    PopupMenu1: TPopupMenu;
    mTray: TPopupMenu;
    Status: TStatusBar;
    tbPluginBtn: TToolBar;
    tbPluginBtnCatalog: TToolButton;
    tbPluginBtnFormats: TToolButton;
    tbPluginBtnPrices: TToolButton;
    tbPluginBtnControl: TToolButton;
    ToolButton1: TToolButton;
    ToolButton15: TToolButton;
    ToolButton2: TToolButton;
    tbPluginBtnOrders: TToolButton;
    ToolButton3: TToolButton;
    TrayIcon: TTrayIcon;
    XMLPropStorage1: TXMLPropStorage;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure mClearPropClick(Sender: TObject);
    procedure mExportCatalogInCSVClick(Sender: TObject);
    procedure mExportCatalogInSpreadsheetClick(Sender: TObject);
    procedure mmBagClick(Sender: TObject);
    procedure mmChangeLogClick(Sender: TObject);
    procedure mmHelpClick(Sender: TObject);
    procedure mmRestoreFormSizeClick(Sender: TObject);
    procedure mShowMainFormClick(Sender: TObject);
    procedure mCloseClick(Sender: TObject);
    procedure MenuItem9Click(Sender: TObject);
    procedure mmAboutClick(Sender: TObject);
    procedure mm_ClearDBClick(Sender: TObject);
    procedure mmBooksPriceFieldClick(Sender: TObject);
    procedure tbPluginBtnCatalogClick(Sender: TObject);
    procedure tbPluginBtnControlClick(Sender: TObject);
    procedure tbPluginBtnFormatsClick(Sender: TObject);
    procedure tbPluginBtnOrdersClick(Sender: TObject);
    procedure tbPluginBtnPricesClick(Sender: TObject);
    procedure ToolButton15Click(Sender: TObject);
  private
    fCatalog: TCatalog;
    fDataClearThread: TDataClearThread;
    fIdMainOwner: integer;
    { private declarations }
    FormIDent: string;
    fBase: TwBase;
    fReport: TwReport;
    fSilentMode: boolean;
    _Version, _VersionRaw: string;
    procedure DBUpdate(aSelectedPrices: ArrayOfInteger);
    procedure ExportCatalog(aPatch: string; aStocks: string; aPrices: string;
      aType: TwfExportFormat);

    procedure DBBackup();
    procedure EnabledFuncional(AValue: boolean);
    procedure Init(Sender: TObject);
    procedure mKursClick(Sender: TObject);
    procedure mm_ClearDB_onEnd(Sender: TObject);
    procedure mm_ClearDB_onStatusUpdate(Sender: TObject);
    procedure mTrayClick(Sender: TObject);
    procedure mTrayMenuFill(Sender: TObject);
    procedure UpdateKurs(const aSilent: boolean = False);
    procedure ClearProps();

    property wFormID: string read FormIDent write FormIDent;
  public
    { public declarations }
    procedure SetStatus(_Text: string; _Log: boolean);
    function GetStatus(_Log: boolean): string;
    procedure CheckDBVersion(ADBase: TwBase);
    property LicenseOK: boolean write EnabledFuncional;

  end;

var
  FmMain: TFmMain;
  dbArr: array [0..7, 0..2] of variant;

const
  wProgName = 'iPriceSE — работа с прайс-листами';

  {$IFDEF WINDOWS}

  {$IFDEF CPU32}
    wTargetOS = 'Win32';
  {$ELSE}
    wTargetOS = 'Win64';
  {$ENDIF}
  {$ENDIF}

  {$IFDEF UNIX}// mac & linux

  {$IFDEF CPU32}
    wTargetOS = 'Unix32';
  {$ELSE}
    wTargetOS = 'Unix64';
  {$ENDIF}
  {$ENDIF}

implementation

{$R *.lfm}

{ TFmMain }


procedure TFmMain.EnabledFuncional(AValue: boolean);
var
  i: integer;
begin
  for i := 0 to tbPluginBtn.ButtonCount - 1 do
    tbPluginBtn.Buttons[i].Enabled := AValue;

  tbPluginBtnControl.Enabled := True;
  MainMenu.Items[1].Enabled := AValue;
  MainMenu.Items[2].Enabled := AValue;
end;

procedure TFmMain.CheckDBVersion(ADBase: TwBase);
var
  _Invoce: TInvoce;
  _Invoce_CreateTable: ArrayOfString;
  i: integer;
  _CheckVersion, aSQL, aNewVersion: string;
begin
  try
    if ADBase.ReadSettingByName('dbVersion') <> _VersionRaw then
    begin
      wLog('Main', 'Рефакторинг метаданных БД...');

      aNewVersion := '0.0.6.0';
      if CheckLessVersion(ADBase.ReadSettingByName('dbVersion'), aNewVersion) then
      begin

        FmWait := TFmWait.Create(self);

        FmWait.Show;
        FMWait.InitBar(4, 0);
        wLog('Main', 'Обновление метаданных БД... до версии '
          +
          aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus(
          'Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate('ALTER TABLE FORMATS ADD ACTUALDAYS INTEGER DEFAULT 3 ',
            True, False);
          ADBase.SQLUpdate('ALTER TABLE FORMATS ADD NOMINPRICE INTEGER DEFAULT 0 ',
            True, False);
          FMWait.SetBar(2);
          ADBase.SQLUpdate('UPDATE FORMATS SET ACTUALDAYS=3, NOMINPRICE=0', True, False);
          FMWait.SetBar(3);
          ADBase.SQLUpdate('CREATE OR ALTER procedure CATALOG_PL_ITEMS_PRICE (  ' +
            wfLE + '     IDCATALOG bigint)  ' + wfLE + ' returns (  ' +
            wfLE + '     ID bigint, ' + wfLE + '     PRICEPL numeric(15,10))  ' +
            wfLE + ' as ' + wfLE + ' BEGIN    ' +
            wfLE + '      FOR     ' + wfLE + '        SELECT    ' +
            wfLE + '        PL.ID,    ' +
            wfLE + '        (PL.PRICECALC/MTH.QUANTITYINPACKING) AS PRICEPL     ' +
            wfLE + '        FROM "PL_ITEMS" PL      ' +
            wfLE +
            '        INNER JOIN "CATALOG_MATCHING" MTH ON (PL.ID = MTH.IDPL_ITEMS AND MTH.IDCATALOG=:IDCATALOG)  '
            +
            wfLE + '        LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PL.IDFORMATS) ' +
            wfLE +
            '        WHERE (PL.STOCK>0 OR PL.STOCK2>0 OR PL.STOCK3>0 OR PL.STOCK4>0 OR PL.STOCK5>0) AND (PL.PRICECALC>0) AND FMTS.FCLOSE = 0  AND FMTS.NOMINPRICE = 0 AND IIF(FMTS.ACTUALDAYS>0,(DATEDIFF(day, PL.FTIMESTAMP, CURRENT_TIMESTAMP)<=FMTS.ACTUALDAYS),true)' + wfLE + '        INTO :ID,     ' + wfLE + '             :PRICEPL     ' + wfLE + '      DO    ' + wfLE + '      BEGIN     ' + wfLE + '        SUSPEND;    ' + wfLE + '      END     ' + wfLE + '    END ', True, False);

          wLog('Main', 'Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion', aNewVersion);

          FMWait.Free;
        except
          wLog('Main', 'Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;
    end;

    ADBase.SetSettingByName('dbVersion', _VersionRaw);
  except

  end;
end;

procedure TFmMain.Init(Sender: TObject);
var
  i, _progCountStart: integer;
  _DbIsClear: boolean;
  _LogUpdate: TStringList;
  _LogUpdateFile: string;
begin
  try

    fBase := TwBase.Create(Sender);

    if fBase.SQLReadDS('OWNER', ['ID'], '', '').DataSet.RecordCount = 1 then
    begin
      ShowMessage(
        'Список контрагентов пуст! Для работы приложения добавьте хотя бы одного контрагента!');
      _DbIsClear := True;
    end
    else
    begin
      _DbIsClear := False;
      SetStatus(
        'Сервер БД доступен и ожидает соединения', True);
      SetStatus('Отсоединен от БД', False);
    end;

    if CheckLessVersion(_VersionRaw, fBase.ReadSettingByName(
      'dbVersion')) then
      // проверка версии не ниже
    begin
      ShowMessage(
        'Версия БД новее используемой версии программы! Для продолжения работы используйте версию приложения не ниже v. ' + fBase.ReadSettingByName('dbVersion') + '!');
      fBase.Free;
      exit;
    end;

    if CheckLessVersion(fBase.ReadSettingByName('dbVersion'), '0.0.1.27') then
      // был сильный рефакторинг БД ниже не поддерживается
    begin
      ShowMessage(
        'Версия БД несовместима с текущей версией приложения!');
      fBase.Free;
      exit;
    end;

    CheckDBVersion(fBase);

    //End Проверка версии БД

    // обслуживание индексов CATALOG каждые 100 запусков

    TryStrToInt(fBase.ReadSettingByName('progCountStart'), _progCountStart);

    if _progCountStart > 300 then
    begin
      FmWait := TFmWait.Create(self);

      FmWait.Show;
      FMWait.InitBar(5, 0);
      wLog('Main', 'Обслуживание БД...');
      FmWait.SetStatus('--=== iPriceSE ===--');
      FmWait.SetStatus('Обновление индексов...');
      FmWait.SetStatus(
        'Это может занять некоторое время...');
      FMWait.SetBar(1);

      try

        fBase.SQLUpdate('ALTER INDEX PL_VERSIONS_FTIMESTAMP INACTIVE;');
        fBase.SQLUpdate('ALTER INDEX PL_VERSIONS_FTIMESTAMP ACTIVE;');
        FMWait.SetBar(2);
        fBase.SQLUpdate('ALTER INDEX PL_ITEMS_VENDORCODE INACTIVE;');
        fBase.SQLUpdate('ALTER INDEX PL_ITEMS_VENDORCODE ACTIVE;');
        FMWait.SetBar(3);
        fBase.SQLUpdate('ALTER INDEX CTG_VENDORCODE INACTIVE;');
        fBase.SQLUpdate('ALTER INDEX CTG_VENDORCODE ACTIVE;');
        FMWait.SetBar(4);
        fBase.SQLUpdate('ALTER INDEX CTG_NAME INACTIVE;');
        fBase.SQLUpdate('ALTER INDEX CTG_NAME ACTIVE;');

      finally
        fBase.SetSettingByName('progCountStart', 0);

        FmWait.SetStatus('Обновление индексов... [ОК]');
        FMWait.SetBar(5);
        FMWait.Free;
      end;
    end
    else
    begin
      // счетчик запусков программы
      fBase.SetSettingByName('progCountStart', _progCountStart + 1);
    end;

    //End обслуживание индексов CATALOG каждые 100 запусков

    wLog('Main', 'Формирование списка плагинов...');

    if _DbIsClear then
      __wPluginSettings([TFmCatalog, TFmPrices, TFmFormats,
        TFmUtils, TFmAnalisis, TFmOrders], 0)
    else
      __wPluginSettings(
        [TFmCatalog, TFmPrices, TFmFormats, TFmUtils, TFmAnalisis, TFmOrders], 0);

    __wPluginInit(FmMain, pcPlugins, self);

    if _DbIsClear then
    begin
      if Plugin <> nil then
        Plugin.Add(TwPlugin.Create(2)); // Formats
    end;

    wLog('Main',
      'Формирование списка плагинов успешно завершено.');

    fBase.Free;
  except
    SetStatus('Отсоединен от БД', False);

    EnabledFuncional(False); // ограничиваем функционал

    //В случае отсутствия БД запускаем плагин Утилиты
    __wPluginSettings([TFmUtils], 0);
    __wPluginInit(FmMain, pcPlugins, self);
    if Plugin <> nil then
      Plugin.Add(TwPlugin.Create(0));
    if Assigned(fBase) then
      fBase.Free;

    raise Exception.Create(
      'Ошибка соединения с БД. Проверьте настройки соединения');
  end;

  _LogUpdateFile := PathLogFiles_Unsafe + 'log-update.txt';
  if FileExistsUTF8(_LogUpdateFile) then
  begin
    _LogUpdate := TStringList.Create;
    try
      _LogUpdate.LoadFromFile(_LogUpdateFile);

      if (UTF8Pos('error', UTF8LowerCase(_LogUpdate.Text)) > 0) then
        if MessageDlg(
          'Внимание! При последнем автоматическом обновлении были обнаружены ошибки!' +
          LineEnding + 'Открыть файл в программе просмотра?',
          mtWarning, mbOKCancel, 0) = mrOk then
          OpenDocument(_LogUpdateFile);
    finally
      _LogUpdate.Free;
    end;
  end;
end;

procedure TFmMain.DBUpdate(aSelectedPrices: ArrayOfInteger);
var
  fDBImport: TwDBImport;
  _Form: TFmWait;
  i: integer;
  _DataSet: TDataSet;
  _FieldsArray: ArrayOfString;
  _Where, _LogUpdateFile: string;
begin
  fBase := TwBase.Create(self);
  _Form := TFmWait.Create(self);

  fDBImport := TwDBImport.Create(self, _Form.Memo);

  _FieldsArray := nil;

  _FieldsArray := fBase.MakeArrayFromString(FormatImportFields);

  if Assigned(aSelectedPrices) then
    _Where := ' IDFMTS_CATEGORY=1 AND FCLOSE=0 AND IDOWNER IN(' +
      fBase.MakeStringFromArray(aSelectedPrices) + ')'
  else// только прайс-листы
    _Where := ' IDFMTS_CATEGORY=1 AND FCLOSE=0 ';// только прайс-листы

  _DataSet := fBase.SQLReadDS('FORMATS', _FieldsArray, _Where,
    'IDOWNER, PRIORITY, NAME').DataSet;
  _DataSet.Last;
  _DataSet.First;
  fDBImport.FormatsPrice.Clear;
  for i := 0 to _DataSet.RecordCount - 1 do
  begin

    fDBImport.FormatsPrice.PushBack(
      fDBImport.CreateFormatPrice(
      _DataSet.FieldByName('IDOWNER').AsInteger,
      _DataSet.FieldByName('ID').AsInteger,
      _DataSet.FieldByName('NAME').AsString,
      _DataSet.FieldByName('FILE').AsString,
      _DataSet.FieldByName('FILEZIPNAMEDECODE').AsInteger,
      _DataSet.FieldByName('FILEHASH').AsString,
      _DataSet.FieldByName('URL').AsString,
      _DataSet.FieldByName('IDFILEFORMAT').AsInteger,
      _DataSet.FieldByName('FCONVERTLIBRE').AsInteger,
      _DataSet.FieldByName('IDCODEPAGETEXT').AsInteger,
      _DataSet.FieldByName('IDCURRENCY').AsInteger,
      _DataSet.FieldByName('CURRENCYPERCENT').AsInteger,
      _DataSet.FieldByName('STORAGEDAYS').AsInteger,
      _DataSet.FieldByName('STOCKONLY').AsInteger,
      fBase.MakeArrayArrayVariantFromString(_DataSet.FieldByName(
      'STOCKSYMBOLS').AsString), _DataSet.FieldByName('STOCKONLYINFO').AsInteger,
      _DataSet.FieldByName('YMLID').AsInteger,
      _DataSet.FieldByName('YMLPRICE').AsInteger,
      _DataSet.FieldByName('YMLQUANTITY').AsInteger,
      _DataSet.FieldByName('FCLOSE').AsInteger,
      _DataSet.FieldByName('GROUPSINROWS').AsInteger,
      _DataSet.FieldByName('GROUPALGORITHM').AsInteger,
      _DataSet.FieldByName('GROUPS').AsInteger,
      _DataSet.FieldByName('SUBGROUPS1').AsInteger,
      _DataSet.FieldByName('SUBGROUPS2').AsInteger,
      _DataSet.FieldByName('SUBGROUPS3').AsInteger,
      _DataSet.FieldByName('FIRSTLINE').AsInteger,
      _DataSet.FieldByName('VENDORCODE').AsInteger,
      _DataSet.FieldByName('FNAME').AsInteger,
      _DataSet.FieldByName('UNIT').AsInteger,
      _DataSet.FieldByName('QUANTITY').AsInteger,
      _DataSet.FieldByName('STOCK2').AsInteger,
      _DataSet.FieldByName('STOCK3').AsInteger,
      _DataSet.FieldByName('STOCK4').AsInteger,
      _DataSet.FieldByName('STOCK5').AsInteger,
      _DataSet.FieldByName('TRANSIT').AsInteger,
      _DataSet.FieldByName('PRICE').AsInteger,
      _DataSet.FieldByName('PRICE2').AsInteger,
      _DataSet.FieldByName('PRICE3').AsInteger,
      _DataSet.FieldByName('PRICE4').AsInteger,
      _DataSet.FieldByName('PRICE5').AsInteger,
      _DataSet.FieldByName('PRICE6').AsInteger,
      _DataSet.FieldByName('PRICE7').AsInteger,
      _DataSet.FieldByName('PRICE8').AsInteger,
      _DataSet.FieldByName('PRICE9').AsInteger,
      _DataSet.FieldByName('PRICE10').AsInteger,
      _DataSet.FieldByName('LABEL').AsInteger,
      _DataSet.FieldByName('SCOD').AsInteger,
      _DataSet.FieldByName('FURL').AsInteger,
      _DataSet.FieldByName('FURLPICTURE').AsInteger,
      _DataSet.FieldByName('FREMARK').AsInteger,
      _DataSet.FieldByName('FCOLOR').AsInteger,
      _DataSet.FieldByName('IDFMTS_CATEGORY').AsInteger,
      fBase.MakeArrayArrayIntegerFromString(_DataSet.FieldByName(
      'SPREADSHEET').AsString, _DataSet.FieldByName('FIRSTLINE').AsInteger),
      _DataSet.FieldByName('IDVENDORCODEVARIANT').AsInteger,
      _DataSet.FieldByName('IDCSVDELIMITER').AsInteger,
      _DataSet.FieldByName('IDSTOCKVARIANT').AsInteger,
      _DataSet.FieldByName('IDPRICEVARIANT').AsInteger)
      );

    _DataSet.Next;
  end;


  fDBImport.IgnoreVersion := False;
  _Form.BorderStyle := bsSizeable;
  _Form.Height := 500;
  _Form.Width := 600;
  _Form.Memo.Alignment := taLeftJustify;
  _Form.Memo.Font.Style := _Form.Memo.Font.Style - [fsBold];
  _Form.Show;
  _Form.NoClose := True;
  try
    fDBImport.Import();
    while not fDBImport.EndThread do
    begin
      Application.ProcessMessages;
    end;
    _Form.NoClose := False;
  finally
    if Assigned(aSelectedPrices) and (aSelectedPrices[0] = fIdMainOwner) then
      _LogUpdateFile := PathLogFiles_Unsafe + 'log-update_ourprice.txt'
    else
      _LogUpdateFile := PathLogFiles_Unsafe + 'log-update.txt';
    if FileExistsUTF8(_LogUpdateFile) then DeleteFileUTF8(_LogUpdateFile);

    if FileExistsUTF8(PathLogFiles_Unsafe + 'log-crash.txt') then
    begin
      _Form.Memo.Lines.Append('');
      _Form.Memo.Lines.Append('======= log-crash.txt =======');
      _Form.Memo.Lines.Append(__Log.Text);
    end;

    _Form.Memo.Lines.SaveToFile(_LogUpdateFile);
    _Form.Free;
    _DataSet := nil;
    _FieldsArray := nil;

    fBase.Destroy();
    fDBImport.Destroy();
    SetStatus('Импорт завершен.', False);
  end;

end;

procedure TFmMain.ExportCatalog(aPatch: string; aStocks: string;
  aPrices: string; aType: TwfExportFormat);
var
  aCatalog: TCatalog;
  aLog: TStringList;
  aLogExportFile: string;
  aPriceArr, aStockArr: ArrayOfInteger;
  aArr: ArrayOfArrayVariant;
  i: integer;
  //aLogExportFileSpreadSheet= PathLogFiles_Unsafe+'log-exportcatalogSpreadSheet.txt';

  procedure WriteLog(aText: string);
  begin
    aLog.Append(DateTimeToStr(now()) + ' | ' + aText);
  end;

begin
  fBase := TwBase.Create(self);
  aCatalog := TCatalog.Create(self, fBase, True);
  aLog := TStringList.Create;

  try

    case aType of
      eftCSV:
      begin
        aLogExportFile := PathLogFiles_Unsafe + 'log-exportcatalogCSV.txt';
        WriteLog('ExportCatalog in CSV...');
        aCatalog.ExportCatalogInCSV(aPatch, True);
        WriteLog('ExportCatalog in CSV... [OK]');
      end;
      eftSpreadSheet:
      begin
        aLogExportFile := PathLogFiles_Unsafe + 'log-exportcatalogSpreadSheet.txt';
        WriteLog('ExportCatalog in SpreadSheet...');

        aArr := nil;
        aArr := fBase.SQLReadArr('SELECT ID FROM PRICEFIELD WHERE PRIORITY IN (' +
          aPrices + ')');
        SetLength(aPriceArr, Length(aArr));
        for i := 0 to High(aArr) do
          aPriceArr[i] := aArr[i, 0];
        aArr := nil;

        aStockArr := fBase.MakeArrayIntegerFromString(aStocks);
        aCatalog.ExportCatalogInSpreadsheet(aPatch, aStockArr, aPriceArr, True);
        WriteLog('ExportCatalog in SpreadSheet... [OK]');
      end;
    end;


  finally
    if FileExistsUTF8(aLogExportFile) then DeleteFileUTF8(aLogExportFile);
    aLog.SaveToFile(aLogExportFile);

    FreeAndNil(aLog);
    FreeAndNil(aCatalog);
    FreeAndNil(fBase);
  end;
end;

procedure TFmMain.DBBackup();
var
  _Form: TFmWait;
  _DateTimeStr, _BackupFileName, _LogBackupFile: string;
begin
  try
    fBase := TwBase.Create(self);
    _Form := TFmWait.Create(self);
    fBase.Memo := _Form.Memo;

    try
      _Form.Memo.Lines.Add(DateTimeToStr(now) +
        '|Запущена операция резервного копирования...');
      _Form.Memo.Lines.Add(DateTimeToStr(now) + '|Соединяюсь с БД...');

      DateTimeToString(_DateTimeStr, 'dd_mm_yy_hh-mm-ss', now);

      if not DirectoryExistsUTF8(PathApplication_Unsafe + 'BackupDB') then
        ForceDirectoriesUTF8(PathApplication_Unsafe + 'BackupDB');

      _BackupFileName := PathApplication_Unsafe + includeTrailingPathDelimiter(
        'BackupDB') + _DateTimeStr + '_db.fbk';

      _Form.BorderStyle := bsSizeable;
      _Form.Height := 500;
      _Form.Width := 600;
      _Form.Memo.Alignment := taLeftJustify;
      _Form.Memo.Font.Style := _Form.Memo.Font.Style - [fsBold];
      _Form.Show;
      _Form.NoClose := True;

      try
        fBase.BackupBase(_BackupFileName);
        _Form.Memo.Lines.Add(DateTimeToStr(now) +
          '|Резервное копирование успешно завершено.');
      except
        raise;
      end;

    finally
      _LogBackupFile := PathLogFiles_Unsafe + 'log-backup.txt';
      if FileExists(_LogBackupFile) then DeleteFile(_LogBackupFile);
      _Form.Memo.Lines.SaveToFile(_LogBackupFile);
      fBase.Free;
      _Form.Free;
    end;

  except

  end;
end;

procedure TFmMain.mTrayMenuFill(Sender: TObject);
var
  _arr: ArrayOfArrayVariant;
  _PopupMenu: TPopupMenu;
  _MenuItem: TMenuItem;
  i: integer;
begin
  fBase := TwBase.Create(self);

  _arr := fBase.SQLReadArr('CURRENCY', ['ID', 'KURS', 'FTIMESTAMP'], 'ID<>1', 'ID');
  try

    if Assigned(_arr) and Assigned(TrayIcon.PopupMenu) then
    begin
      _PopupMenu := TrayIcon.PopUpMenu;
      _PopupMenu.Items.Clear;

      _MenuItem := TMenuItem.Create(_PopupMenu);
      _MenuItem.Name := 'ShowMainForm';
      _MenuItem.Caption := 'Восстановить/Спрятать';
      _MenuItem.ImageIndex := 0;
      _MenuItem.OnClick := @mShowMainFormClick;
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem := TMenuItem.Create(_PopupMenu);
      _MenuItem.Name := 'RestoreFormSize';
      _MenuItem.Caption :=
        'Восстановить размер окна по-умолчанию';
      _MenuItem.ImageIndex := 5;
      _MenuItem.OnClick := @mTrayClick;
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem := TMenuItem.Create(_PopupMenu);
      _MenuItem.Caption := '-';
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem := TMenuItem.Create(_PopupMenu);
      _MenuItem.Name := 'Data';
      _MenuItem.Caption := _arr[0, 2];
      _MenuItem.ImageIndex := 1;
      _MenuItem.OnClick := @mKursClick;
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem := TMenuItem.Create(_PopupMenu);
      _MenuItem.Caption := '-';
      _PopupMenu.Items.Add(_MenuItem);

      for i := 0 to High(_arr) do
      begin
        case i of
          0:
          begin
            _MenuItem := TMenuItem.Create(_PopupMenu);
            _MenuItem.Name := 'USD';
            _MenuItem.Caption := _arr[i, 1];
            _MenuItem.ImageIndex := 2;
            _MenuItem.OnClick := @mKursClick;
            _PopupMenu.Items.Add(_MenuItem);
          end;
          1:
          begin
            _MenuItem := TMenuItem.Create(_PopupMenu);
            _MenuItem.Name := 'EUR';
            _MenuItem.Caption := _arr[i, 1];
            _MenuItem.ImageIndex := 3;
            _MenuItem.OnClick := @mKursClick;
            _PopupMenu.Items.Add(_MenuItem);
          end;
        end;
      end;

      _MenuItem := TMenuItem.Create(_PopupMenu);
      _MenuItem.Caption := '-';
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem := TMenuItem.Create(_PopupMenu);
      _MenuItem.Name := 'CloseMainForm';
      _MenuItem.Caption := 'Выйти';
      //_MenuItem.Sty
      _MenuItem.ImageIndex := 4;
      _MenuItem.OnClick := @mCloseClick;
      _PopupMenu.Items.Add(_MenuItem);
    end;

    //mClose
  finally
    fBase.Destroy();
  end;
end;

procedure TFmMain.FormCreate(Sender: TObject);
var
  _FileSettings: TINIFile;
  i: integer;
  _res: string;
  aTmpArray: ArrayOfInteger;
  fClearProps: boolean;
begin
  try
    TrayIcon.Show;
    wFormID := Self.Name;
    fSilentMode := False;
    DefaultFormatSettings.ThousandSeparator := ' ';
    DefaultFormatSettings.DecimalSeparator := ',';

    PathApplication_Unsafe := SysToUTF8(ExtractFilePath(ParamStr(0)));
    PathLogFiles_Unsafe := PathApplication_Unsafe + 'logs' + DirectorySeparator;
    if not DirectoryExistsUTF8(PathLogFiles_Unsafe) then
      ForceDirectoriesUTF8(PathLogFiles_Unsafe);

    PathExport_Unsafe := PathApplication_Unsafe + DirectorySeparator +
      'export' + DirectorySeparator;
    if not DirectoryExistsUTF8(PathExport_Unsafe) then
      ForceDirectoriesUTF8(PathExport_Unsafe);

    PathTmp_Unsafe := PathApplication_Unsafe + 'tmp' + DirectorySeparator;

    PathTemplates_Unsafe := PathApplication_Unsafe + 'templates' + DirectorySeparator;

    DeleteDirectory(PathTmp_Unsafe, True);

    if not DirectoryExistsUTF8(PathTmp_Unsafe) then
      ForceDirectoriesUTF8(PathTmp_Unsafe);

    //FileExistsUTF8
    if FileExists(PathLogFiles_Unsafe + 'log-crash.txt') then
      DeleteFile(PathLogFiles_Unsafe + 'log-crash.txt');

    if FileExists(PathLogFiles_Unsafe + 'log.txt') then
      DeleteFile(PathLogFiles_Unsafe + 'log.txt');

    FileSettingsPath := '';
    // проверка пути к конфигурационному файлу
    for i := 1 to ParamCount do
    begin
      if UTF8Pos('-settings=', ParamStr(i)) > 0 then
      begin
        FileSettingsPath := StringReplace(ParamStr(i), '-settings=',
          '', [rfReplaceAll, rfIgnoreCase]);
        FileSettingsPath :=
          SafePath(StringReplace(FileSettingsPath, #34, '', [rfReplaceAll, rfIgnoreCase]));
      end;
    end;

    if Length(FileSettingsPath) = 0 then
      FileSettingsPath := PathApplication_Unsafe + 'dbconfig.ini';


    try
      _FileSettings := TINIFile.Create(FileSettingsPath, True);
      __onLog := _FileSettings.ReadBool('Others', 'LogFileON', False);

      CatalogVendorCodeAsNumber :=
        _FileSettings.ReadBool('Others', 'CatalogVendorCodeAsNumber', False);

      fClearProps := _FileSettings.ReadBool('Others', 'ClearProps', False);
      if fClearProps then
      begin
        DeleteFileUTF8(PathApplication_Unsafe + 'ipricese.xml');
        _FileSettings.WriteString('Others', 'ClearProps', '0');
      end;

        {$IFDEF WINDOWS}
          PathLibreOffice:= _FileSettings.ReadString('Others','LibreOffice','');
          if Length(PathLibreOffice)=0 then PathLibreOffice:= GetLibreOfficeInstallation;
        {$ENDIF}


      db_portable := _FileSettings.ReadBool('Others', 'Portable', True);
      // узнаем - портабле ли мы
      ReportHeaderColor := RGBtoBGR(_FileSettings.ReadString(
        'Others', 'ReportHeaderColor', '#FFFBF0'));//$F0FBFF_
      FreeAndNil(_FileSettings);
    finally

      __Log := TwLog.Create;

      _VersionRaw := GetVersion;
      _Version := _VersionRaw + ' [' + wTargetOS + '] ';

      __Log.Add('Main', '-= ' + wProgName + ' | версия: ' + _Version + ' =-');
      FmMain.Caption := wProgName + ' | версия: ' + _Version;

      if FileExists(PathLibreOffice) then
        __Log.Add('Main', 'Найден LibreOffice ' + PathLibreOffice)
      else
        __Log.Add('Main', 'LibreOffice ' + PathLibreOffice + ' НЕ НАЙДЕН!');

      __Log.Add('Main', 'Используемый файл настроек: ' +
        FileSettingsPath);

      if __onLog then
        __Log.Add('Main', 'Ведение лог-файла [ON]')
      else
        __Log.Add('Main', 'Ведение лог-файла [OFF]');
      if db_portable then
        __Log.Add('Main',
          'Приложение запущено в режиме [Portable]')
      else
        __Log.Add('Main',
          'Приложение запущено в режиме [Network]')
    end;

    __Log.Add('Main', 'Инициализация приложения...');

    //__dbReadSettings();

    //DBase
    __wDBaseReadSettings();


    fBase := TwBase.Create(self);
    try
      fIdMainOwner := fBase.ReadSettingByName('setDefaultOwner');
      // считываем настройки - текущий основной прайс-лист
    finally
      fBase.Free;
    end;

    // проверка остальных параметров
    i := 1;
    while i < ParamCount + 1 do
    begin
      case ParamStr(i) of
        '-update':
        begin
          fSilentMode := True;
          DBUpdate(nil);
        end;
        '-updateour':
          { TODO : выбор прайс-листов для загрузки }
        begin
          fSilentMode := True;
          DBUpdate([fIdMainOwner]);
        end;
        '-updateis':
        begin
          fSilentMode := True;
          Inc(i);
          fBase := TwBase.Create(self);
          try
            aTmpArray := fBase.MakeArrayIntegerFromString(ParamStr(i));
          finally
            fBase.Free;
          end;
          DBUpdate(aTmpArray);
        end;
        '-updatekurs':
        begin
          fSilentMode := True;
          UpdateKurs(True);
        end;
        '-exportcatalogcsv':
        begin
          fSilentMode := True;
          Inc(i);
          ExportCatalog(ParamStr(i), '', '', eftCSV);
        end;
        '-exportcatalogxls':
        begin
          fSilentMode := True;
          Inc(i);
          ExportCatalog(ParamStr(i), ParamStr(i + 1),
            ParamStr(i + 2), eftSpreadSheet);
        end;
        '-backup':
        begin
          fSilentMode := True;
          DBBackup();
        end;
      end;
      Inc(i);
    end;

    if not fSilentMode then
      Init(Sender);// инициализация

    mTrayMenuFill(self); // заполнение трей меню

    __Log.Add('Main', 'Инициализация приложения завершена.');

  except
    on E: Exception do
    begin
      if FmWait <> nil then  FmWait.Free;

      __Log.SaveLogError(E);
      wLog('Main', 'Ошибка [FmMC]: "' + E.Message + '"');
      wLog('Main', 'Сбой инициализации приложения.');
      ShowMessage('Ошибка [FmMC]: "' + E.Message + '"');
      SetStatus(E.Message, True);
      exit;
    end;
  end;
end;

procedure TFmMain.FormDestroy(Sender: TObject);
var
  i: integer;
begin
  try
    TrayIcon.Free;
    try
      //  if MainData.dbConnect <> nil then MainData.dbConnect.Connected:=false;

    except
      on E: Exception do
      begin
        ShowMessage('Ошибка [FmMClose]: "' + E.Message + '"');
        wLog('Main', 'Ошибка [FmMC]: "' + E.Message + '"');
        wLog('Main', 'Сбой завершения приложения.');
        __Log.SaveLogError(E);
        ShowMessage(__Log.Text);
      end;
    end;
  finally

    // очищаем память
    FreeAndNil(Plugin);
    FreeAndNil(PluginList);

    __Log.Add('Main', 'Завершение приложения.');

    if __onLog then
    begin
      if __Log.SaveLog() then
      begin  // если лог-файл удалось сохранить, то удаляем краш
        if FileExists(PathLogFiles_Unsafe + 'log-crash.txt') then
          DeleteFile(PathLogFiles_Unsafe + 'log-crash.txt');
      end;

    end;

    if __Log <> nil then  FreeAndNil(__Log);
  end;
end;

procedure TFmMain.FormResize(Sender: TObject);
begin
  Status.Panels[0].Width := Status.Width - 200;
end;

procedure TFmMain.FormShow(Sender: TObject);
begin
  if fSilentMode then
  begin
    FmMain.Close;
    exit;
  end;
end;

procedure TFmMain.ClearProps();
var
  _FileSettings: TIniFile;
begin
  if MessageDlg(
    'Вы уверены, что хотите сбросить настройки внешнего вида: размер окна, ширину колонок и т.п.?', mtWarning, mbOKCancel, 0) = mrCancel then exit;
  _FileSettings := TINIFile.Create(FileSettingsPath, True);

  try
    _FileSettings.WriteString('Others', 'ClearProps', '1');
    ShowMessage(
      'Настройки внешнего вида будут сброшены при следующем запуске, перезапустите программу.');
  finally
    FreeAndNil(_FileSettings);
  end;
end;

procedure TFmMain.mClearPropClick(Sender: TObject);
begin
  ClearProps();
end;

procedure TFmMain.mExportCatalogInCSVClick(Sender: TObject);
begin
  fBase := TwBase.Create(self);
  fCatalog := TCatalog.Create(self, fBase, True);

  try
    fCatalog.ExportCatalogInCSV;
  finally
    screen.Cursor := crDefault;
    fCatalog.Free;
    fBase.Free;
  end;
end;

procedure TFmMain.mExportCatalogInSpreadsheetClick(Sender: TObject);
begin
  fBase := TwBase.Create(self);
  fCatalog := TCatalog.Create(self, fBase, True);

  try
    fCatalog.ExportCatalogInSpreadsheet;
  finally
    screen.Cursor := crDefault;
    fCatalog.Free;
    fBase.Free;
  end;
end;

procedure TFmMain.mmBagClick(Sender: TObject);
begin
  OpenURL('https://bitbucket.org/wofs/ipricese/issues?status=new&status=open');
end;

procedure TFmMain.mmChangeLogClick(Sender: TObject);
begin
  OpenURL('https://bitbucket.org/wofs/ipricese/wiki/%D0%96%D1%83%D1%80%D0%BD%D0%B0%D0%BB%20%D0%B8%D0%B7%D0%BC%D0%B5%D0%BD%D0%B5%D0%BD%D0%B8%D0%B9%20%D0%B2%20%D0%B2%D0%B5%D1%80%D1%81%D0%B8%D1%8F%D1%85');
end;

procedure TFmMain.mmHelpClick(Sender: TObject);
begin
  OpenURL('https://bitbucket.org/wofs/ipricese/wiki/Home.md');
end;

procedure TFmMain.mmRestoreFormSizeClick(Sender: TObject);
begin
  case TMenuItem(Sender).Name of
    'mmRestoreFormSize':
    begin
      if FmMain.WindowState <> wsNormal then
        FmMain.WindowState := wsNormal;
      FmMain.Height := 500;
      FmMain.Width := 900;
      FmMain.MoveToDefaultPosition;
    end;
  end;
end;

procedure TFmMain.mShowMainFormClick(Sender: TObject);
begin
  if FmMain.Showing then
  begin
    Application.Minimize;
    FmMain.Hide;
  end
  else
  begin
    FmMain.WindowState := wsNormal;
    FmMain.Show;
    Application.Restore;
  end;
end;

procedure TFmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  i: longint;
begin
  // выгружаем подгруженные плагины
  if Plugin <> nil then
  begin
    for  i := Plugin.Count - 1 downto 0 do
    begin
      Plugin[i].Unload();
    end;
  end;
  if FmMain.WindowState in [wsMaximized, wsMinimized, wsFullScreen] then
    FmMain.WindowState := wsNormal;
end;

procedure TFmMain.Button1Click(Sender: TObject);
begin

end;

procedure TFmMain.mCloseClick(Sender: TObject);
begin
  Close();
end;

procedure TFmMain.UpdateKurs(const aSilent: boolean = False);
var
  _DBImport: TwDBImport;
  aForm: TProgress;
begin
  _DBImport := TwDBImport.Create(self);

  try
    _DBImport.ImportKursValut(aSilent);

    aForm := TProgress.Create(self);
    aForm.BorderStyle := bsNone;
    aForm.SetStatus('Идет обновление прайс-листа...');
    aForm.SetStatus('Дождитесь окончания операции');
    aForm.SetStatus('Это может занять несколько минут...');
    aForm.InitBar(pbTop, 2);
    aForm.SetBar(pbTop, 1);
    aForm.ShowBottom := False;

    aForm.Show;
    try
      while not _DBImport.EndThread do
      begin
        Application.ProcessMessages;
      end;

    finally
      wLog('debug', 'Освобождаю форму...');
      FreeAndNil(aForm);
      wLog('debug', 'Освобождаю форму... OK');
    end;

    if aSilent then exit;

    if Length(_DBImport.ErrorMessage) > 0 then ShowMessage(_DBImport.ErrorMessage);
    if Length(_DBImport.OKMessage) > 0 then ShowMessage(_DBImport.OKMessage);

  finally
    wLog('debug', 'Освобождаю объект DBImport..');
    _DBImport.Destroy();
    wLog('debug', 'Освобождаю объект DBImport..');

    wLog('debug', 'Заполняю меню в трее...');
    mTrayMenuFill(self);
    wLog('debug', 'Заполняю меню в трее... ОК');
  end;

end;

procedure TFmMain.MenuItem9Click(Sender: TObject);
begin
  UpdateKurs;
end;

procedure TFmMain.mmAboutClick(Sender: TObject);
var
  _Form: TFmAbout;
begin

  _Form := TFmAbout.Create(Self);

  _Form.t_ProgramName.Caption := wProgName + ' | версия: ' + GetVersion;
  try
    _Form.ShowModal;

  finally
    _Form.Free;
  end;

end;


procedure TFmMain.mm_ClearDB_onStatusUpdate(Sender: TObject);
begin
  if Assigned(FmWait) then
  begin
    FmWait.SetStatus(fDataClearThread.Status);
    FMWait.SetBar(fDataClearThread.ProgressPosition);
  end;
end;

procedure TFmMain.mTrayClick(Sender: TObject);
begin
  case TMenuItem(Sender).Name of
    'RestoreFormSize':
    begin
      if FmMain.WindowState <> wsNormal then
        FmMain.WindowState := wsNormal;
      FmMain.Height := 500;
      FmMain.Width := 900;
      FmMain.MoveToDefaultPosition;
    end;
  end;
end;

procedure TFmMain.mm_ClearDB_onEnd(Sender: TObject);
var
  _Result: boolean;
  _Status: string;
begin
  try
    //FmMain.Enabled:= true;
    FmWait.Close;
    _Result := fDataClearThread.Result;
    _Status := fDataClearThread.Status;

    fDataClearThread.Terminate;
    fDataClearThread := nil;

    if _Result then
    begin
      SetStatus('Очистка БД успешно произведена.', True);
      ShowMessage(
        '[DBClean] Очистка БД успешно произведена. После очистки рекомендуется сделать Backup/Restore через [ Утилиты->Обслуживание БД ].');
    end
    else
    begin
      SetStatus('Произошла ошибка при очистке БД', True);
      ShowMessage(_Status);
    end;
  finally
    fBase.Destroy();

    //if Plugin <> nil then
    //    Plugin.Add(TwPlugin.Create(2));
  end;
end;

procedure TFmMain.mm_ClearDBClick(Sender: TObject);
var
  i, _IdOwner: integer;
  _PluginPageIndex: longint;
begin

  if MessageDlg(
    'Очистить Базу данных? Это приведет к потере всех введенных данных!',
    mtWarning, mbOKCancel, 0) = mrCancel then exit;

  if MessageDlg('Вы уверены? Очистка БД необратима!',
    mtWarning, mbOKCancel, 0) = mrCancel then exit;

  fBase := TwBase.Create(self);

  fDataClearThread := nil;
  fDataClearThread := TDataClearThread.Create(True);
  fDataClearThread.Base := fBase;

  if Plugin <> nil then
  begin
    for  i := Plugin.Count - 1 downto 0 do
    begin
      _PluginPageIndex := Plugin[i].PageIndex;
      try
        Plugin[i].Unload();
      finally
        FmMain.pcPlugins.Pages[_PluginPageIndex].Free;
      end;
    end;
  end;

  //DBase

  FmWait := TFmWait.Create(self);

  FmWait.Height := 250;
  FmWait.Width := 540;

  FmWait.mStatus.Alignment := taLeftJustify;
  FMWait.InitBar(8, 0);
  wLog('Main', 'Очистка БД...');
  FmWait.SetStatus('--=== iPriceSE ===--');
  FmWait.SetStatus('Очистка БД...');
  FMWait.SetBar(2);

  fDataClearThread.ProgressPosition := 2;
  fDataClearThread.onEndThread := @mm_ClearDB_onEnd;
  fDataClearThread.onStatusUpdate := @mm_ClearDB_onStatusUpdate;
  fDataClearThread.Start;

  try
    FmWait.ShowModal;
  finally
    FmWait.Free;
  end;
end;


procedure TFmMain.mmBooksPriceFieldClick(Sender: TObject);
var
  _Form: TFmPriceFields;

begin
  _Form := TFmPriceFields.Create(Self);

  try
    _Form.ShowModal;
  finally
    _Form.Free;
  end;
end;

procedure TFmMain.mKursClick(Sender: TObject);
begin
  Clipboard.AsText := TMenuItem(Sender).Name + ' = ' + TMenuItem(Sender).Caption;
  ShowMessage('Строка скопирована в буфер обмена.');
end;

procedure TFmMain.tbPluginBtnCatalogClick(Sender: TObject);
begin
  wLog('Main', 'Запуск плагина: Каталог...');

  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(0))
  else
  begin
    SetStatus('Нет соединения с БД.', False);
    SetStatus('Сбой инициализации плагина: Каталог.',
      True);
    wLog('Main', 'Сбой инициализации плагина: Каталог.');
    ShowMessage('Сбой инициализации плагина: Каталог.');
  end;
end;

procedure TFmMain.tbPluginBtnControlClick(Sender: TObject);
begin
  wLog('Main', 'Запуск плагина: Утилиты...');

  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(3))
  else
  begin
    SetStatus('Нет соединения с БД.', False);
    SetStatus('Сбой инициализации плагина: Утилиты.',
      True);
    wLog('Main', 'Сбой инициализации плагина: Утилиты.');
    ShowMessage('Сбой инициализации плагина: Утилиты.');
  end;
end;

procedure TFmMain.tbPluginBtnFormatsClick(Sender: TObject);
begin
  wLog('Main', 'Запуск плагина: Форматы...');

  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(2))
  else
  begin
    SetStatus('Нет соединения с БД.', False);
    SetStatus('Сбой инициализации плагина: Форматы.',
      True);
    wLog('Main', 'Сбой инициализации плагина: Форматы.');
    ShowMessage('Сбой инициализации плагина: Форматы.');
  end;
end;

procedure TFmMain.tbPluginBtnOrdersClick(Sender: TObject);
begin
  wLog('Main', 'Запуск плагина: Накладные...');

  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(5))
  else
  begin
    SetStatus('Нет соединения с БД.', False);
    SetStatus('Сбой инициализации плагина: Накладные.',
      True);
    wLog('Main', 'Сбой инициализации плагина: Накладные.');
    ShowMessage('Сбой инициализации плагина: Накладные.');
  end;
end;

procedure TFmMain.tbPluginBtnPricesClick(Sender: TObject);
begin
  wLog('Main', 'Запуск плагина: Прайс-листы...');

  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(1))
  else
  begin
    SetStatus('Нет соединения с БД.', False);
    SetStatus('Сбой инициализации плагина: Прайс-листы.', True);
    wLog('Main', 'Сбой инициализации плагина: Прайс-листы.');
    ShowMessage('Сбой инициализации плагина: Прайс-листы.');
  end;
end;

procedure TFmMain.ToolButton15Click(Sender: TObject);
begin
  wLog('Main', 'Запуск плагина: Аналитика...');
  SetStatus('Запуск плагина "Аналитика"... Это может занять немного времени...', True);
  Application.ProcessMessages;
  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(4))
  else
  begin
    SetStatus('Нет соединения с БД.', False);
    SetStatus('Сбой инициализации плагина: Аналитика.',
      True);
    wLog('Main', 'Сбой инициализации плагина: Аналитика.');
    ShowMessage('Сбой инициализации плагина: Аналитика.');
  end;
end;

procedure TFmMain.SetStatus(_Text: string; _Log: boolean);
begin
  if _Log then
    Status.Panels.Items[0].Text := _Text
  else
    Status.Panels.Items[1].Text := _Text;

  wLog(wFormID, _Text);
end;

function TFmMain.GetStatus(_Log: boolean): string;
begin
  if _Log then
    Result := Status.Panels.Items[0].Text
  else
    Result := Status.Panels.Items[1].Text;
end;

end.

