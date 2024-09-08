unit wBaseU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

 // Назначение:
 // Простая работа с базой

{$mode objfpc}{$H+}

interface

uses
  Classes, Controls,
  db, Dialogs, Forms, IBCustomDataSet, IBDatabase, IBEvents, IBQuery, IBServices,
  IBSQL, INIFiles, csvdocument, fpspreadsheet, fpsTypes, fpsutils,
  LazUTF8, LCLProc, messages, StrUtils, StdCtrls, SysUtils,
  wTProgressU, wTypesU, wCustomClassThreadU,
   wLogU;

type

  { TwBase }

  TwBase = class
  private
    fDataSet: TDataSet;
    fdbConnect: TIBDatabase; // подключение к БД
    fdbTransaction: TIBTransaction;   // транзакция
    fdbTransactionUpdate: TIBTransaction;   // транзакция
    fdbQuery:  TIBQuery;     // запрос на чтение данных
    fdbQueryUpdate: TIBSQL;
    fdbEvent: TIBEVents;
    // запрос на изменение данных данных
    fdbDataSource: TDataSource;
    fMemo:     TMemo;
    fonProgressInit: TProgressEventInit;
    fonProgressUpdate: TProgressEventUpdate;
    fonStatusUpdate: TNotifyEvent;
    fOutStringArr: ArrayOfString;
    fParent:   TComponent;
    fStatus: string;
    fStopForce: Boolean;
    IBBackup:  TIBBackupService;
    IBRestore: TIBRestoreService;

    fFormName: string;
    fLongTransaction: boolean;

    function ExportTableToSpreadSheetFile(aSQL, aFileName: string; aTableHeads: ArrayOfString; aFileFormat: TFileFormat; const aStep: integer): boolean;
    procedure Log(aText: string);
    function PrintQueryParams(adbQueryUpdate: TIBSQL): string;
    procedure ProgressBarInit(aValue: integer);
    procedure ProgressBarPosition(aValue: integer);
    procedure SetStatus(aText: string; aLog: boolean = False);
    // вывод статуса

    function GetTransaction: TIBTransaction;
    procedure SetLongTransaction(AValue: boolean);
    function DataToStr(aField: TField; aCSVComma: boolean= false; aBr: boolean = false): string;
    procedure onWriteCellData(Sender: TsWorksheet; ARow, ACol: cardinal; var AValue: variant; var AStyleCell: PCell);

  public
    property Transaction: TIBTransaction read GetTransaction;
    property LongTransaction: boolean read fLongTransaction write SetLongTransaction;
    property Query: TIBQuery read fdbQuery write fdbQuery;
    property QueryUpdate: TIBSQL read fdbQueryUpdate write fdbQueryUpdate;
    property EventDB: TIBEvents read fdbEvent write fdbEvent;
    property DataBase: TIBDatabase read fdbConnect;
    property Memo: TMemo read fMemo write fMemo;
    property Status: string read fStatus;
    property StopForce: boolean write fStopForce;

    property onStatusUpdate: TNotifyEvent read fonStatusUpdate write fonStatusUpdate;
    property onProgressUpdate: TProgressEventUpdate read fonProgressUpdate write fonProgressUpdate;
    property onProgressInit: TProgressEventInit read fonProgressInit write fonProgressInit;
    property OutStringArr: ArrayOfString read fOutStringArr write fOutStringArr;

    //FORMULA
    //property PriceField: boolean read fPriceField write fPriceField;
    //property PriceFieldValue: ArrayOfDouble read fPriceFieldValue write fPriceFieldValue;
    //property BaseFormula: TFormula read fFormula write fFormula;
    property FormName: string read fFormName write fFormName;

    constructor Create(Sender: TObject);
    destructor Destroy(); override;
    procedure BackupBase(aBackupFileName: string);
    procedure RestoreBase(aBackupFileName: string);

    function ReadSettingByName(aName: string): variant;
    // функция чтения настройки
    function SetSettingByName(_Name: string; _Value: variant): boolean;

    function AddExt(const aFileName, aExt: string): string;
    function ExportTableToCSVFile(aSQL, aFileName: string; aTableHeads: ArrayOfString; aTableFields: string; const aStep: integer): boolean;

    function ExportTableToFile(aSQL, aFileName: string; aExportFormat: TFileFormat; aTableHeads: ArrayOfString; const aStep: integer=10000): boolean;
    function ImportTableFromFile(aFileName: string; aCommit: boolean): boolean;

    function SQLItemGet(aSQL: string; aID: integer): ArrayOfArrayVariant;
    function SQLItemGetDS(aSQL: string; aParamValues:ArrayOfVariant): TDataSource;

    function SQLItemUpdate(aSQL: string; aParamValues:ArrayOfVariant; const aReturnID:boolean = false; const aCommit: boolean = true):integer;

    function SQLReadArr(aSQL: string): ArrayOfArrayVariant;
    function SQLReadArr(aTableName: string; aFieldsArray: array of string;
      aWhere, aOrderBy: string): ArrayOfArrayVariant;

    function SQLReadDS(aSQL: string; aFetchAll: boolean=false): TDataSource;
    procedure SQLReadDS(aDataSet: TIBDataSet; aTransaction: TIBTransaction;
      aDataSource: TDataSource; aSQL: string);
    function SQLReadDS(aTableName: string; aFieldsArray: array of string;
      aWhere, aOrderBy: string): TDataSource;

    function SQLInsert(aTableName: string; aFieldsArray: array of string;
      aValuesArray: array of variant; const aCommit: boolean = True): integer;
    function SQLInsert(aTableName: string; aFieldsArray: array of string;
      aValuesArray: array of variant; aWhereFields: string;
      const aCommit: boolean = True): integer;
    function SQLInsert(aSQLText: string; const aCommit: boolean = True): integer;
    function SQLExecStringList(aStringList: TStringList;
      const aCommit: boolean = True): boolean;
    function SQLUpdate(aTableName: string; aFieldsArray: array of string;
      aValuesArray: array of variant; aWhereFields: string;
      const aCommit: boolean = True): boolean;
    function SQLUpdate(aSQLText: string; const aCommit: boolean = True;
      const aParamCheck: boolean = True): boolean;
    function SQLDelete(aTableName: string; aWhereFields: string;
      const aCommit: boolean = True): boolean; // delete

    procedure SQLTransactionEnd(const aCommit: boolean = True);
    // Commit / Rollback update transaction

    function RegisterEvents(aEventNames: ArrayOfString): boolean;
    function ReadMaxDateTimeValuesT(aTable, aGroupField, aTimeStampField: string;
      const aWhere: string): ArrayOfDateTime;
    function PrepareWhereStringFromDateTime(aFilterGroupField: string;
      aFilterGroup: array of TDateTime): string;
    function MakeArrayFromString(aText: string): ArrayOfString;
    // формирует массив строк из строки с разделителями
    function MakeStringFromArray(aTextArray: ArrayOfString): string;
    // формирует строку с разделителями из массива строк
    function MakeStringFromArray(aIntArray: ArrayOfInteger): string;
    // формирует строку с разделителями из массива строк

    function MakeArrayIntegerFromString(aText: string): ArrayOfInteger;
    // формирует массив строк из строки с разделителями
    function MakeArrayArrayIntegerFromString(aText: string;
      const aDefaultVal2: integer = -1): ArrayOfArrayInteger;
    // формирует массив целых чисел из строки с разделителями формат: ЛИСТ|СТРОКА, ЛИСТ|СТРОКА
    function MakeArrayArrayVariantFromString(aText: string;
      const aDelimiter2: char = '=';
      const aDefaultVal2: integer = 0): ArrayOfArrayVariant;
    // формирует массив variant из строки
    function PrepareIDOwnerWhere(aWhereString: string;
      aOwnerFieldValueArr: array of variant): string;
    // подготовка Where для дополнительной фильтрации списков по владельцу
    function PrepareWhereString(aFilterGroupField: string;
      aFilterGroup: ArrayOfInteger; const aANDOR: string = 'OR'): string;
    // подготовка Where
    function WriteWhere(aSQL, aWhere: string; aClearWhere:boolean = false): string;
    function WriteWhereEx(const uSQL: string; aWhere: string; const aClearOldWhere: boolean = false): string;
    function GetWhere(const uSQL: string): string;

    function IsNum(aText: string): boolean;

    function GetRowsCount(aSQL: string): integer;

    function GetCurrencyArray(): ArrayOfCurrency;
    function TextInArray(aText: string; aTextArray: ArrayOfString): boolean;
  end;

  { TDataClearThread }

  TDataClearThread = class(TThread)
    protected
      procedure Execute; override;
    private
      fExceptionEvent: TThreadExceptionEvent;
      fonEndThread: TNotifyEvent;
      fonStatusUpdate: TNotifyEvent;
      fResult: Boolean;
      fBase: TwBase;
      fStatus: String;
      fProgressPosition: integer;

      procedure SetStatus(aText: string);
    public
      Constructor Create(CreateSuspended : boolean);
      property onEndThread: TNotifyEvent read fonEndThread write fonEndThread;
      property onStatusUpdate: TNotifyEvent read fonStatusUpdate write fonStatusUpdate;

      property Status: string read fStatus;
      property Result: boolean read fResult;
      property ProgressPosition: integer read fProgressPosition write fProgressPosition;
      property Base: TwBase write fBase;
  end;

function __wDBaseReadSettings: boolean;

var
  // настройки коннекта
  db_dbName, db_dbLibraryLocation, db_dbHostname, db_dbProtocol,
  db_dbUser, db_dbPassword: string;
  db_dbPort:   integer;
  db_portable: boolean;
  PathApplication_Unsafe: string;
  PathLogFiles_Unsafe: string;
  PathExport_Unsafe: string;
  PathTmp_Unsafe: string;
  PathTemplates_Unsafe: string;
  PathLibreOffice: string;
  FileSettingsPath: string;
  ReportHeaderColor: LongInt;
  CatalogVendorCodeAsNumber: boolean;
 //const
 //  db_dbLibraryLocation ='fbclientd_x64.dll';
 //  db_dbHostname = 'localhost';
 //  db_dbPort = 3050;
 ////  db_dbProtocol = 'firebirdd-2.5';
 //  db_dbName = 'Base\FBBASE.FDB';
 //  db_dbUser = 'SYSDBA';
 //  db_dbPassword ='masterkey';
 //  db_portable = false;
implementation

uses
  wFuncU;

function __wDBaseReadSettings: boolean;
var
  _FileSettings: TINIFile;
begin
  wLog('wDataBase', 'Чтение настроек БД.');

  try
    _FileSettings := TINIFile.Create(FileSettingsPath, True);

    with _FileSettings do
    begin

    {$IFDEF WINDOWS}
    db_dbLibraryLocation := PathApplication_Unsafe+ReadString('dbSettings', 'LibraryLocation',
      'fbclient.dll');
    {$ELSE}
    db_dbLibraryLocation := PathApplication_Unsafe+ReadString('dbSettings', 'LibraryLocation',
      'libfbclient.so');
    {$ENDIF}

      wLog('Main', 'LibraryLocation='+db_dbLibraryLocation);
      db_dbHostname := ReadString('dbSettings', 'Hostname', 'localhost');
      wLog('Main', 'Hostname='+db_dbHostname);
      db_dbPort := ReadInteger('dbSettings', 'Port', 3050);
      wLog('Main', 'Port='+IntToStr(db_dbPort));
      db_dbProtocol := ReadString('dbSettings', 'Protocol', 'firebirdd-3.0');
      wLog('Main', 'Protocol='+db_dbProtocol);
        {$IFDEF WINDOWS}
        if not db_portable then
        {$ELSE}
        if db_portable then
        {$ENDIF}
        db_dbName := SafePath(ReadString('dbSettings',
          'Database', 'FBBASE.FDB'))
      else
        db_dbName := ReadString('dbSettings', 'Database', 'FBBASE.FDB');

      wLog('Main', 'Database='+db_dbName);
      db_dbUser     := ReadString('dbSettings', 'User', 'SYSDBA');
      // wLog('Main', 'User='+wdb_dbUser);
      db_dbPassword := ReadString('dbSettings', 'Password', 'masterkey');
      // wLog('Main', 'Password='+'********');
      FreeAndNil(_FileSettings);
    end;
  except
    on E: Exception do
    begin
      wLog('__wDataBaseReadSettings',
        'Ошибка чтения настроек БД.');
      raise;
    end;
  end;
  Result := True;
end;

{ TDataClearThread }

procedure TDataClearThread.Execute;
var
  _IdOwner: Integer;
begin
  try
    fResult:= false;

      //if not Assigned(fBase) then
      //  Exception.Create('Не найдено подключение к БД!');

        SetStatus('Деактивация индексов...');
        try
          fBase.SQLUpdate('ALTER INDEX PL_ITEMS_VENDORCODE INACTIVE;');
          fBase.SQLUpdate('ALTER INDEX PL_VERSIONS_FTIMESTAMP INACTIVE;');
          fBase.SQLUpdate('ALTER INDEX CTG_VENDORCODE INACTIVE;');
          fBase.SQLUpdate('ALTER INDEX CTG_NAME INACTIVE;');

          fBase.LongTransaction:= true;

          SetStatus('Очистка каталога...');

          fBase.SQLDelete('CATALOG_GROUP','',false);

          SetStatus('Очистка остальных таблиц...');

          fBase.SQLDelete('OWNER','',false);

          fBase.SQLTransactionEnd(true);

          SetStatus('Активация индексов...');

          fBase.SQLUpdate('ALTER INDEX PL_ITEMS_VENDORCODE ACTIVE;');
          fBase.SQLUpdate('ALTER INDEX PL_VERSIONS_FTIMESTAMP ACTIVE;');
          fBase.SQLUpdate('ALTER INDEX CTG_VENDORCODE ACTIVE;');
          fBase.SQLUpdate('ALTER INDEX CTG_NAME ACTIVE;');
        except
          if fBase.LongTransaction then
          fBase.SQLTransactionEnd(false);
          raise;
        end;


        SetStatus('Установка начальных значений генераторов...');

        fBase.SQLUpdate('SET GENERATOR "GEN_PL_ITEMS_ID" TO 0',true);
        fBase.SQLUpdate('SET GENERATOR "GEN_PL_VERSIONS_ID" TO 0',true);
        fBase.SQLUpdate('SET GENERATOR "GEN_PL_GROUP_ID" TO 0',true);
        fBase.SQLUpdate('SET GENERATOR "GEN_PL_SCODS_ID" TO 0',true);

        fBase.SQLUpdate('SET GENERATOR "GEN_OWNER_ID" TO 0',true);
        fBase.SQLUpdate('SET GENERATOR "GEN_FORMATS_ID" TO 0',true);

        fBase.SQLUpdate('SET GENERATOR "GEN_CATALOG_ID" TO 0',true);
        fBase.SQLUpdate('SET GENERATOR "GEN_CATALOG_GROUP_ID" TO 0',true);
        fBase.SQLUpdate('SET GENERATOR "GEN_CATALOG_SCODS_ID" TO 0',true);
        fBase.SQLUpdate('SET GENERATOR "GEN_CATALOG_MATCHING_ID" TO 0',true);

        _IdOwner:= fBase.SQLInsert('OWNER',['NAME','IDPARENT'],['Контрагенты',integer(0)],true);
        _IdOwner:= fBase.SQLInsert('OWNER',['NAME','IDPARENT'],['Свой прайс-лист',_IdOwner],true);
        fBase.SetSettingByName('setDefaultOwner',integer(_IdOwner));
        fBase.SQLInsert('CATALOG_GROUP',['NAME','IDPARENT','IDOWNER'],['Номенклатура',integer(0),_IdOwner],true);

        fBase.SetSettingByName('progCountStart',integer(0));

        fResult:= true;
        SetStatus('Очистка БД... [ЗАВЕРШЕНО]');

        onEndThread(self);
  except
    on E: Exception do begin
      fResult:= false;
      SetStatus('Error: '+E.Message);
      onEndThread(self);
    end;
  end;
end;

procedure TDataClearThread.SetStatus(aText: string);
begin
  fStatus := aText;
  inc(fProgressPosition);
  onStatusUpdate(Self);
end;

constructor TDataClearThread.Create(CreateSuspended: boolean);
begin
  FreeOnTerminate := true;
  fBase:= nil;
  inherited Create(CreateSuspended);
end;

{ TwBase }

procedure TwBase.Log(aText: string);
begin
  if __onLog and assigned (__Log) then
    wLog('[Base] ', aText);
end;

procedure TwBase.SetStatus(aText: string; aLog: boolean);
begin
  if Assigned(fMemo) and not aLog then
    fMemo.Lines.Add(DateTimeToStr(now)+' | '+aText)
  else
    wStatus(FormName, aText, aLog);

  if __onLog then Log(DateTimeToStr(now)+' | '+aText);

  fStatus:= aText;
  if Assigned(onStatusUpdate) then
     onStatusUpdate(self);
end;

function TwBase.GetTransaction: TIBTransaction;
begin
  if fLongTransaction then
    Result := fdbTransactionUpdate
  else
    begin
      if not fdbTransaction.Active then fdbTransaction.StartTransaction;
      Result := fdbTransaction;
    end;
end;

procedure TwBase.SetLongTransaction(AValue: boolean);
begin
  //if fLongTransaction=AValue then Exit;
  if AValue then
    Log('SetLongTransaction:= true')
  else
    Log('SetLongTransaction:= false');

  fLongTransaction := AValue;
  if AValue then
  begin
    if fdbTransactionUpdate.Active then
      fdbTransactionUpdate.Rollback;
    fdbQuery.Close;
    fdbQuery.Transaction := fdbTransactionUpdate;
    fdbTransactionUpdate.StartTransaction;
  end
  else
  begin
    if fdbTransactionUpdate.Active then
      fdbTransactionUpdate.Rollback;
    fdbQuery.Close;
    fdbQuery.Transaction := fdbTransaction;
  end;
end;

constructor TwBase.Create(Sender: TObject);
begin
  try
    fParent    := TComponent(Sender);
    fFormName  := TComponent(Sender).Name;
    fdbConnect := TIBDatabase.Create(fParent);
    fdbTransaction := TIBTransaction.Create(fParent);
    fdbTransactionUpdate := TIBTransaction.Create(fParent);
    fdbQuery   := TIBQuery.Create(fParent);
    fdbQueryUpdate := TIBSQL.Create(fPArent);
    fdbDataSource := TDataSource.Create(fParent);
    fdbEvent:= TIBEvents.Create(fParent);

    fLongTransaction := False;

    with fdbConnect do
    begin
      DefaultTransaction := fdbTransaction;
      DefaultUpdTransaction := fdbTransactionUpdate;
      LoginPrompt := False;
      SQLDialect  := 3;
    end;

    with fdbQuery do
    begin
      Database    := fdbConnect;
      Transaction := fdbConnect.DefaultTransaction;
    end;

    with fdbQueryUpdate do
    begin
      Database    := fdbConnect;
      Transaction := fdbConnect.DefaultUpdTransaction;
    end;

    with fdbTransaction do
    begin
      DefaultDatabase := fdbConnect;
      Params.Add('read');
      Params.Add('read_committed');
      Params.Add('rec_version');
      Params.Add('nowait');
    end;

    with fdbTransactionUpdate do
      DefaultDatabase := fdbConnect;

    with fdbDataSource do
      DataSet := fdbQuery;

    with fdbEvent do
    begin
      Database:= fdbConnect;
      Events.Clear;
    end;

    with fdbConnect do
    begin
      Connected   := False;
      LoginPrompt := False;
      Params.Clear;
      LibraryName := db_dbLibraryLocation;

      if db_portable then
        DatabaseName := db_dbName
      else
        DatabaseName := db_dbHostname+'/'+IntToStr(db_dbPort)+':'+db_dbName;

      Params.Clear;
      Params.Add('user_name='+db_dbUser);
      Params.Add('password='+db_dbPassword);
      Params.Add('lc_ctype=UTF8');
      //if db_portable then
      //   Params.Add('Providers = Engine12, Loopback');

      Connected := True;
      SetStatus('Соединен с БД', False);
    end;
  except
    on E: Exception do
    begin
      Log('Ошибка [BaseCreate]: "'+E.Message+'"');
      SetStatus('Отсоединен от БД', False);
      __Log.SaveLogError(E);
      raise;
    end;
  end;

end;

destructor TwBase.Destroy();
begin
  FreeAndNil(fdbConnect);
  FreeAndNil(fdbTransaction);
  FreeAndNil(fdbTransactionUpdate);
  FreeAndNil(fdbQuery);
  FreeAndNil(fdbQueryUpdate);
  FreeAndNil(fdbDataSource);

  inherited Destroy;
end;

procedure TwBase.BackupBase(aBackupFileName: string);
var
  _LinesCount, i: integer;
begin
  IBBackup := TIBBackupService.Create(fParent);

  try
    with IBBackup do
    begin
    try
      Active := False;
      BackupFile.Clear;
      Params.Clear;
      LoginPrompt := False;
      LibraryName := db_dbLibraryLocation;
      SetStatus('Используемая библиотека: '+LibraryName);

      if db_portable then
      begin
        DatabaseName := db_dbName;
        Protocol:= Local;
        SetStatus('Используемая БД [embedded]: '+DatabaseName);
      end
      else
      begin
        DatabaseName := db_dbName;
        ServerName:=db_dbHostname+'/'+IntToStr(db_dbPort);
        Protocol:= TCP;
        aBackupFileName:= ExtractFileDir(db_dbName)+DirectorySeparator+'Backup_'+ExtractFileName(db_dbName)+'-'+ExtractFileName(aBackupFileName);
        SetStatus('Используемая БД [network]: '+DatabaseName);
        SetStatus('Файл будет сохранен на сверере базы данных по адресу: '+ExtractFileDir(db_dbName));
      end;

      Params.Add('user_name='+db_dbUser);
      Params.Add('password='+db_dbPassword);

      //Params.Add('lc_ctype=UTF8');
      BackupFile.Add(aBackupFileName);
      SetStatus('Сохраняем в файл: '+BackupFile.Text);

        Active := True;
        try

          if Active then
            Detach();

          Attach;
          ServiceStart;

          SetStatus(
            'Резервное копирование... Дождитесь окончания операции...');
          SetStatus('');
          _LinesCount := fMemo.Lines.Count;
          i := 0;
          while IsServiceRunning do
          begin
            Inc(i);
            if i mod 10000 = 1 then
              fMemo.Lines[_LinesCount-1] := fMemo.Lines[_LinesCount-1]+'.';

            Application.ProcessMessages;
          end;
        finally
          if Active then
            Detach();
          //Screen.Cursor := crDefault;
        end;
        SetStatus('Резервное копирование... [ОК]');
      except
        on E: Exception do
        begin
          if IBBackup.Active then
            IBBackup.Detach();
          __Log.SaveLogError(E);
          SetStatus('Соединяюсь с БД... [ОШИБКА]');
          SetStatus('"'+E.Message+'"');
          SetStatus(E.Message);
          raise;
        end;
      end;
    end;
  finally
    IBBackup.Free;
  end;
end;

procedure TwBase.RestoreBase(aBackupFileName: string);
var
  _LinesCount, i: integer;
begin
  IBRestore := TIBRestoreService.Create(fParent);
   try
   with IBRestore do
    begin
     try
      Active := False;
      Params.Clear;
      Options := [CreateNewDB];
      BackupFile.Clear;
      DatabaseName.Clear;
      LoginPrompt := False;
      LibraryName := db_dbLibraryLocation;
      SetStatus('Используемая библиотека: '+LibraryName);

      if db_portable then
      begin
        DatabaseName.Add(db_dbName);
        Protocol:= Local;
        SetStatus('Используемая БД [embedded]: '+DatabaseName.Text);
        if FileExists(db_dbName) then
          DeleteFile(db_dbName);
      end
      else
      begin
        DatabaseName.Add(ExtractFileDir(db_dbName)+DirectorySeparator+'RESTORED_'+ExtractFileName(db_dbName));
        ServerName:=db_dbHostname+'/'+IntToStr(db_dbPort);
        Protocol:= TCP;
        aBackupFileName:= ExtractFileDir(db_dbName)+DirectorySeparator+'RESTORE.fbk';
        SetStatus('ВНИМАНИЕ: Используется удаленная база данных, переименуйте файл с резервной копией в RESTORE.fbk и разместите на сервере базы данных по адресу: '+ExtractFileDir(db_dbName));
        SetStatus('БД будет восстановлена на сервере базы данных в файл: '+DatabaseName.Text);
      end;

      Params.Add('user_name='+db_dbUser);
      Params.Add('password='+db_dbPassword);

      //Params.Add('lc_ctype=UTF8');

        Active := True;
        try

          BackupFile.Add(aBackupFileName);

          SetStatus('Загружаем резервную копию из файла : '+BackupFile.Text);
          if Active then
            Detach();
          Attach;
          ServiceStart;

          SetStatus(
            'Восстановление из резервной копии... Дождитесь окончания операции...');
          SetStatus('');
          _LinesCount := fMemo.Lines.Count;
          i := 0;
          while IsServiceRunning do
          begin
            Inc(i);
            if i mod 10000 = 1 then
              fMemo.Lines[_LinesCount-1] := fMemo.Lines[_LinesCount-1]+'.';

            Application.ProcessMessages;
          end;
        finally
          if Active then
            Detach();
        end;

        if db_portable then
          SetStatus('Восстановление из резервной копии... [ОК]')
        else
          begin
          SetStatus('База данных успешно восстановлена в файл: '+ExtractFileDir(db_dbName)+DirectorySeparator+'RESTORED_'+ExtractFileName(db_dbName)+'.');
          SetStatus('ВНИМАНИЕ: закройте все копии программы, остановите Firebird сервер и переименуйте старую базу данных '+ExtractFileName(db_dbName)+' в OLD_'+ExtractFileName(db_dbName)+', а файл с восстановленной RESTORED_'+ExtractFileName(db_dbName)+' в файл '+ExtractFileName(db_dbName)+'. После этого запустите сарвер Firebird и подключесь к восстановленной базе. Если все прошло успешно и работает — удалите файл '+'OLD_'+ExtractFileName(db_dbName)+'.');
          end;

          db_dbName:=ExtractFileDir(db_dbName)+'RESTORED_'+ExtractFileName(db_dbName);
      except
        on E: Exception do
        begin
          if IBRestore.Active then
            IBRestore.Detach();
          __Log.SaveLogError(E);
          SetStatus('Соединяюсь с БД... [ОШИБКА]');
          SetStatus('"'+E.Message+'"');
          SetStatus(E.Message);
          exit;
        end;
      end;

    end;
  finally
    IBRestore.Free;
  end;
end;

function TwBase.ReadSettingByName(aName: string): variant;
begin
  Result := SQLReadArr('SETTINGS', ['FVALUE'], 'NAME='''+aName+'''', '')[0, 0];
end;

function TwBase.SetSettingByName(_Name: string; _Value: variant): boolean;
begin
  Result := SQLUpdate('SETTINGS', ['FVALUE'], [_Value], 'NAME='''+_Name+'''');
end;

function TwBase.DataToStr(aField: TField; aCSVComma: boolean; aBr: boolean): string;
begin
  case aField.DataType of
    ftSmallInt,
    ftLargeint,
    ftInteger: Result  := IntToStr(aField.AsInteger);
    ftFloat,
    ftCurrency,
    ftBCD: Result      := FloatToStr(aField.AsFloat);
    ftDate: Result     := FormatDateTime('dd.mm.yyyy', aField.AsDateTime);
    ftTime: Result     := FormatDateTime('hh:mm:ss', aField.AsDateTime);
    ftDateTime: Result :=
        FormatDateTime('dd.mm.yyyy hh:mm:ss', aField.AsDateTime);
    ftBoolean: if aField.AsBoolean then
        Result := 'T'
      else
        Result := 'F';
    ftString,
    ftWideString:
    begin
      if aField.AsString<>null then
        Result := aField.AsString
      else
        Result := '';

      if aBr and (Length(Result)>0) then
      begin
        Result := StringReplace(Result,#13#10,'&br;',[rfReplaceAll]);
        Result := StringReplace(Result,#10,'&br;',[rfReplaceAll]);
      end;

      if aCSVComma and (Length(Result)>0) then
        Result := '"'+StringReplace(Result,#34,'&quot;',[rfReplaceAll])+'"';

    end
    else
      WriteStr(Result, aField.DataType);
      //Result := GetEnumName(aField.DataType);
  end;

end;

function TwBase.AddExt(const aFileName, aExt: string):string; // добавление расширения к файлу, если отсутствует
var
  _Ext: String;
begin
  _Ext := ExtractFileExt(aFileName);
  Result:= aFileName;

  if Length(_Ext) = 0 then
    Result := Result+aExt;
end;

function TwBase.ImportTableFromFile(aFileName:string; aCommit:boolean): boolean;
var
  _CSV: TCSVParser;
  _FileStream: TStream;
  _CurrentRow, _ColCount, iRows: integer;
  _Fields: ArrayOfString;
  _Values: ArrayOfVariant;
  _Ext, _TableName: String;

  procedure ClearValues(aColCount: integer);
  begin
    SetLength(_Values,0);
    SetLength(_Values,aColCount);
  end;

begin
  _CurrentRow := 1;
  _ColCount:=0;
  iRows:=0;
  Result:= false;

  SetStatus('Импорт...',true);

  _Ext:= ExtractFileExt(aFileName);
  if Length(_ext)>0 then
     _Ext:='.'+_Ext;

  _TableName:= StringReplace(ExtractFileName(aFileName),_Ext,'',[rfIgnoreCase]);

  try
    _CSV:= TCSVParser.Create;
    _FileStream := TFileStream.Create(aFileName, fmOpenRead+fmShareDenyWrite);
    _CSV.Delimiter:=';';
    _CSV.SetSource(_FileStream);

    SetLength(_Fields,1024);

    while _CSV.ParseNextCell do
    begin
      if _CSV.CurrentRow=0 then
      begin
        _Fields[_CSV.CurrentCol]:= _CSV.CurrentCellText;
        Inc(_ColCount);
      end else
      Break;
    end;

    _ColCount:= _ColCount;

    SetLength(_Values,_ColCount);
    SetLength(_Fields,_ColCount);

    _CSV.Free;
    _FileStream.Free;

    _CSV:= TCSVParser.Create;
    _FileStream := TFileStream.Create(aFileName, fmOpenRead+fmShareDenyWrite);

    _CSV.Delimiter:=';';
    _CSV.SetSource(_FileStream);

    while _CSV.ParseNextCell do
    begin
      if _CSV.CurrentRow>0 then
      begin
         if _CurrentRow <> _CSV.CurrentRow then
         begin
           SQLInsert(_TableName,_Fields,_Values,aCommit);
           inc(iRows);
            if iRows mod 1000 = 0 then
                  SetStatus('Обработано строк: '+IntToStr(iRows),true);

           ClearValues(_ColCount);
           _CurrentRow:= _CSV.CurrentRow;
         end;

         _Values[_CSV.CurrentCol]:= Trim(DecodeHTMLEntrities(_CSV.CurrentCellText));
      end;
    end;

   if _Values[0]<> null then
      SQLInsert(_TableName,_Fields,_Values,aCommit);

   SetStatus('',true);
   Result:= true;

  finally
    FreeAndNil(_FileStream);
    FreeAndNil(_CSV);
  end;

end;

function TwBase.SQLItemGet(aSQL: string; aID: integer): ArrayOfArrayVariant;
var
  _Transaction: TIBTransaction;
  _FieldsCount, i, ifield: integer;
  _RecordCount: longint;
begin
  try
    _Transaction := Transaction;

   //if not fLongTransaction then
   //begin
   //  _Transaction.Active := False;
   //  _Transaction.StartTransaction;
   //end;

    with fdbQuery do
    begin
      fdbQuery.AutoFetchAll := true;
      Close;
      SQL.Clear;
      SQL.Text := aSQL;
      Log(SQL.Text);
      Log(Params[0].Name+'= '+IntToStr(aID));
      Params[0].Value:= aID;
      Open;
    end;

    with fdbQuery.DataSource.DataSet do
    begin
      _FieldsCount := Fields.Count;
      _RecordCount := RecordCount;

      SetLength(Result, _RecordCount, _FieldsCount);
      for i := 0 to _RecordCount-1 do
      begin

        for ifield := 0 to _FieldsCount-1 do
          Result[i, ifield] := Fields[ifield].AsVariant;
        Next;
      end;
      Close;
    end;

    //if not fLongTransaction then
    //  if _Transaction.Active then
    //    _Transaction.Commit;

  except
    on E: Exception do
    begin
      Log('Ошибка [SQLItemGet]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      //_Transaction.Rollback;
      raise;
    end;
  end;
end;

function TwBase.SQLItemGetDS(aSQL: string; aParamValues: ArrayOfVariant): TDataSource;
var
  _Transaction: TIBTransaction;
   i: integer;
begin
  try
    _Transaction := Transaction;

    //if not fLongTransaction then
    //begin
    //  _Transaction.Active := False;
    //  _Transaction.StartTransaction;
    //end;

    with fdbQuery do
    begin
      fdbQuery.AutoFetchAll := true;
      Close;
      SQL.Clear;
      SQL.Text := aSQL;
      Log(SQL.Text);

      if Params.Count <> Length(aParamValues) then
         raise Exception.Create('Error Count Params!');

      for i:=0 to Params.Count-1 do
      begin
        Log(Params[i].Name+'= '+VarToStr(aParamValues[i]));
        Params[i].Value:= aParamValues[i];
      end;

      Open;

      Result:= fdbDataSource;
    end;

  except
    on E: Exception do
    begin
      Log('Ошибка [SQLItemGetDS]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      //_Transaction.Rollback;
      raise;
    end;
  end;

end;

function TwBase.SQLItemUpdate(aSQL: string; aParamValues: ArrayOfVariant; const aReturnID: boolean; const aCommit: boolean): integer;
var
  _Transaction: TIBTransaction;
  i: Integer;
begin
  Result := -1;
  try
    Log('SQLItemUpdate...');

    _Transaction := fdbQueryUpdate.Transaction;

    if not fdbConnect.Connected then
      fdbConnect.Connected := True;

    if not _Transaction.Active and not fLongTransaction then
      _Transaction.StartTransaction;

    with fdbQueryUpdate do
    begin
      ParamCheck := true;
      Close;
      SQL.Clear;
      SQL.Text := aSQL;
      Log(SQL.Text);

      if Params.Count <> Length(aParamValues) then
         raise Exception.Create('Error Count Params!');

      for i:=0 to Params.Count-1 do
        Params[i].Value:=aParamValues[i];

      ExecQuery;

      if aReturnID then
        Result:= FieldByName('ID').AsInteger
      else
        Result:= -1;

      Close;
    end;
    if aCommit and not fLongTransaction then
    begin
      _Transaction.Commit;
      Log('SQLItemUpdate [OK] | [TransactionEnd] = COMMIT');
    end
    else
      Log('SQLItemUpdate [OK] | [TransactionEnd] = NONE');
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLItemUpdate]: "'+E.Message+'"');
      __Log.Add('[Error]'+LineEnding+fdbQueryUpdate.SQL.Text);
      __Log.SaveLogError(E);
      SetStatus('Ошибка [SQLItemUpdate]: "'+E.Message+'"', True);
      //      ShowMessage('Ошибка [SQLInsert]: "'+E.Message+'"');
      _Transaction.Rollback;
      raise;
    end;
  end;
end;

procedure TwBase.ProgressBarInit(aValue:integer);
begin
  if Assigned(onProgressInit) then onProgressInit(pbTop, aValue);
end;

procedure TwBase.ProgressBarPosition(aValue: integer);
begin
  if Assigned(onProgressUpdate) then onProgressUpdate(pbTop, aValue);
end;

function TwBase.ExportTableToFile(aSQL, aFileName: string; aExportFormat: TFileFormat; aTableHeads: ArrayOfString; const aStep: integer): boolean;
begin
  fStopForce:= false;

  case aExportFormat of
    ffCSV: Result:= ExportTableToCSVFile(aSQL,aFileName,aTableHeads,'',aStep);
    ffODS,
    ffXLS,
    ffXLSX : Result:= ExportTableToSpreadSheetFile(aSQL,aFileName,aTableHeads,aExportFormat,aStep)
    else
      Result:= false;
  end;
end;

function TwBase.ExportTableToCSVFile(aSQL, aFileName: string; aTableHeads: ArrayOfString; aTableFields: string; const aStep: integer): boolean;
const
  cCommaChar     = ';';

  procedure WriteHeadFields(const aSQL:string; aPosSelect:integer;var  aFields:ArrayOfString; aCSV:TFileStream);
  var
    _DataSet: TDataSet;
    _CSV_String, _SQL: String;
    i, iFields: Integer;

    procedure WriteField(var iFields: integer; i:integer);
    begin
       aFields[iFields] := _DataSet.Fields[i].FieldName;
       Inc(iFields);
    end;
  begin

    _SQL:= aSQL;
    UTF8Insert(' first 0 ', _SQL, aPosSelect);

    _DataSet := SQLReadDS(_SQL,true).DataSet;
      SetLength(aFields, _DataSet.Fields.Count);

     iFields:= 0;
     for i := 0 to _DataSet.Fields.Count-1 do
     begin
       if Length(aTableFields)>0 then
       begin
         if UTF8Pos(_DataSet.Fields[i].FieldName,aTableFields)>0 then
            WriteField(iFields, i);
       end else
         WriteField(iFields, i);
     end;
    if Length(aTableFields)>0 then
       SetLength(aFields, iFields);

      //try

        _CSV_String := '';
        // формируем заголовок из имен столбцов
        for i := 0 to High(aFields) do
        begin
          if Length(_CSV_String)>0 then
            _CSV_String := _CSV_String+cCommaChar;

          if Assigned(aTableHeads) and (Length(aTableHeads)= Length(aFields)) then // записываем в заголовок наши имена полей (если есть)
            _CSV_String   := _CSV_String+'"'+aTableHeads[i]+'"'
          else
            _CSV_String   := _CSV_String+'"'+aFields[i]+'"';

        end;

        WriteUTF8String(aCSV,_CSV_String);
  end;

var
  _DataSet:  TDataSet;
  i, iDS:    integer;
  _CSV:      TFileStream;

  _Fields:   ArrayOfString;
  _CSV_String: string;
  _CountRows:    integer; // количество строк в таблице
   iRows: integer; // количество пропускаемых строк
  _PosSelect: byte;   // позиция оператора SELECT
begin

  Result    := False;
  _Fields   := nil;

  if FileExists(aFileName) then
        DeleteFile(aFileName);

  _CSV  := TFileStream.Create(aFileName,fmCreate);

  try

   try //finally

    SetStatus('Подготовка экспорта таблицы...',true);

    aSQL := UTF8UpperCase(aSQL);
    _PosSelect := UTF8Pos('SELECT', aSQL)+UTF8Length('SELECT');
    //_PosFrom:= UTF8Pos('FROM', aSQL,_PosSelect);

    _CountRows:= GetRowsCount(aSQL);

    ProgressBarInit(_CountRows);

    WriteHeadFields(aSQL,_PosSelect,_Fields,_CSV); // формирование строки с заголовками

      // модифицируем запрос под выборку партиями
      UTF8Insert(' first %d skip %d ', aSQL, _PosSelect);

      // экспортируем данные
      iRows:=0;
      while iRows<_CountRows do begin
        SetStatus('Выборка данных из БД...',true);

        _DataSet := SQLReadDS(Format(aSQL,[aStep,iRows]),true).DataSet;

        for iDS := 0 to _DataSet.RecordCount-1 do
        begin
          if fStopForce then
             raise Exception.Create('Экспорт отменен!');

          _CSV_String := '';

          for i := 0 to High(_Fields) do
          begin
            if Length(_CSV_String)>0 then
              _CSV_String := _CSV_String+cCommaChar;
            _CSV_String   :=
              _CSV_String+DataToStr(_DataSet.FieldByName(_Fields[i]), True, True);
          end;

          WriteUTF8String(_CSV,_CSV_String);

          if (iDS+iRows) mod 500 = 0 then
          begin
            ProgressBarPosition(iDS+iRows);

            SetStatus('Обработано строк: '+IntToStr(iDS+iRows)+' из '+IntToStr(_CountRows),true);
          end;

          _DataSet.Next;
        end;
        iRows:= iRows+aStep;
      end; //end while

    finally
      FreeAndNil(_CSV);
      SetStatus(' ',true);
    end;

    Result := True;
  except
    Result := False;
    raise;
  end;

end;


procedure TwBase.onWriteCellData(Sender: TsWorksheet; ARow, ACol: cardinal;   // Experemental!!!
  var AValue: variant; var AStyleCell: PCell);
begin
  // Let's handle the header row first:
  if ARow = 0 then begin
    // The value to be written to the spreadsheet is the field name.
    AValue := fDataSet.Fields[ACol].FieldName;
    // Formatting is defined in the HeaderTemplateCell.
    //AStyleCell := MyHeaderTemplateCell;
    // Move to first record
    fDataSet.First;
  end else begin
    // The value to be written to the spreadsheet is the record value in the field corresponding to the column.
    // No special requirements on formatting --> leave AStyleCell at its default (nil).
    AValue := fDataSet.Fields[ACol].AsVariant;
    // Advance database cursor if last field of record has been written
    if ACol = fDataSet.FieldCount-1 then fDataSet.Next;
  end;
end;

function TwBase.ExportTableToSpreadSheetFile(aSQL, aFileName: string; aTableHeads: ArrayOfString; aFileFormat: TFileFormat; const aStep: integer): boolean;
//const
  procedure WriteHeadFields(const aSQL:string; aPosSelect:integer;var _Worksheet: TsWorksheet);
  var
    _DataSet: TDataSet;
    _SQL: String;
    i: Integer;
    _UserFieldsName: Boolean;
  begin

    _SQL:= aSQL;

     UTF8Insert(' first 0 ', _SQL, aPosSelect);

     _DataSet := SQLReadDS(_SQL,true).DataSet;

     //SetLength(aFields, _DataSet.Fields.Count);
     if Assigned(aTableHeads) and (Length(aTableHeads)= _DataSet.Fields.Count) then
     _UserFieldsName:= true else
     _UserFieldsName:= false;

     for i := 0 to _DataSet.Fields.Count - 1 do
     begin
       if _UserFieldsName then
         _Worksheet.WriteText(0, i, aTableHeads[i])
       else
         _Worksheet.WriteText(0, i, _DataSet.Fields[i].FieldName);

       _Worksheet.WriteFontStyle (0, i, [fssBold]);
     end;
    _DataSet.Close;
  end;

var
  _WorkBook: TsWorkbook;
  _Worksheet: TsWorksheet;

  _CountRows, iRows, iDS, i:    integer; // количество строк в таблице
  _PosSelect: byte;   // позиция оператора SELECT
  _SpreadsheetFormat: TsSpreadsheetFormat;
  _PosFrom: PtrInt;
begin

  Result    := False;

  if FileExists(aFileName) then
        DeleteFile(aFileName);

  _WorkBook  := TsWorkbook.Create();
  _WorkBook.Options:= [boBufStream];

  _Worksheet := _WorkBook.AddWorksheet('Sheet1');

  try

   try //finally

    SetStatus('Подготовка экспорта таблицы...',true);

    aSQL := UTF8UpperCase(aSQL);
    _PosSelect:= UTF8Pos('SELECT', aSQL)+UTF8Length('SELECT');
    _PosFrom:= UTF8Pos('FROM', aSQL,_PosSelect);

    _CountRows:= GetRowsCount(aSQL);

    ProgressBarInit(_CountRows);

    // определяем формат
    case aFileFormat of
        ffODS: _SpreadsheetFormat:= sfOpenDocument;
        ffXLS: _SpreadsheetFormat:= sfExcel8;
        ffXLSX: _SpreadsheetFormat:= sfOOXML;
      else
        _SpreadsheetFormat:= sfUser;
    end;


    if _SpreadsheetFormat = sfUser then
      raise Exception.Create('Формат экспорта не определен!');


    //fOutStringArr

    WriteHeadFields(aSQL,_PosSelect,_Worksheet); // формирование строки с заголовками + подсчет колонок

      // модифицируем запрос под выборку партиями
      UTF8Insert(' first %d skip %d ', aSQL, _PosSelect);

      // экспортируем данные
      iRows:=0;
      while iRows<_CountRows do begin
        SetStatus('Выборка данных из БД...',true);

        fDataSet := SQLReadDS(Format(aSQL,[aStep,iRows]),true).DataSet;

        for iDS := 0 to fDataSet.RecordCount-1 do
        begin
          if fStopForce then
             raise Exception.Create('Экспорт отменен!');

          for i := 0 to fDataSet.Fields.Count-1 do
          begin
            WriteDataToCell(fDataSet.Fields[i],iDS+1,i,_Worksheet);
          end;


          if (iDS+iRows) mod 100 = 0 then
          begin
            ProgressBarPosition(iDS+iRows);

            SetStatus('Обработано строк: '+IntToStr(iDS+iRows)+' из '+IntToStr(_CountRows),true);
          end;

          fDataSet.Next;
        end;
        iRows:= iRows+aStep;
      end; //end while

      if Assigned(fOutStringArr) and (Length(fOutStringArr)>0) then
       begin
         for i:= High(fOutStringArr) downto 0 do begin
           _Worksheet.InsertRow(0);
           _Worksheet.WriteText(0,1,fOutStringArr[i]);
         end;

         _Worksheet.InsertRow(High(fOutStringArr)+1);
       end;
//записываем в файл
       _Workbook.WriteToFile(aFileName,_SpreadsheetFormat,true);

    finally
      fDataSet:= nil;
      _WorkBook.Free;
      SetStatus(' ',true);
    end;

    Result := True;
  except
    Result := False;
    raise;
  end;

end;

function TwBase.SQLReadArr(aSQL: string): ArrayOfArrayVariant;
var
  _Transaction: TIBTransaction;
  _FieldsCount, i, ifield: integer;
  _RecordCount: longint;
  _DS: TDataSource;
begin
  try

    _Transaction := Transaction;

    _DS := SQLReadDS(aSQL,true);

    with _DS.DataSet do
    begin
      _FieldsCount := Fields.Count;
      _RecordCount := RecordCount;

      SetLength(Result, _RecordCount, _FieldsCount);
      for i := 0 to _RecordCount-1 do
      begin

        for ifield := 0 to _FieldsCount-1 do
          Result[i, ifield] := Fields[ifield].AsVariant;
        Next;
      end;
      Close;
    end;

  except
    on E: Exception do
    begin
      Log('Ошибка [SQLReadArr]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      //_Transaction.Rollback;
      raise;
    end;
  end;
end;

function TwBase.SQLReadArr(aTableName: string; aFieldsArray: array of string;
  aWhere, aOrderBy: string): ArrayOfArrayVariant;
var
  _Fields: string;
  ifield:  integer;
begin
  try

    _Fields := '';
    if Length(aFieldsArray) = 0 then
      exit;

    for ifield := 0 to Length(aFieldsArray)-1 do
    begin
      if ifield>0 then
        _Fields := _Fields+',';
      _Fields   := _Fields+aFieldsArray[ifield];
    end;

    if Length(aWhere)>0 then
      aWhere := 'WHERE '+aWhere;
    if Length(aOrderBy)>0 then
      aOrderBy := 'ORDER BY '+aOrderBy;


    Result := SQLReadArr('SELECT '+_Fields+' FROM "'+aTableName +
      '" '+aWhere+' '+aOrderBy+';');

  except
    on E: Exception do
    begin
      Log('Ошибка [SQLReadArr]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      raise;
    end;
  end;
end;

function TwBase.SQLReadDS(aSQL: string; aFetchAll:boolean = false): TDataSource;
var
  _Transaction: TIBTransaction;
begin
  try
    Log('SQLReadDS...');

    if not fdbConnect.Connected then
      fdbConnect.Connected := True;

    _Transaction := Transaction;

    //if not fLongTransaction then
    //begin
    //  _Transaction.Active := False;
    //  _Transaction.StartTransaction;
    //end;

    with fdbQuery do
    begin
      fdbQuery.AutoFetchAll := aFetchAll;
      Close;
      SQL.Clear;
      SQL.Text := aSQL;
      Log(SQL.Text);
      Open;
      Result := fdbDataSource;
    end;

    Log('SQLReadDS [OK]');
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLReadDS]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      //_Transaction.Rollback;
      raise;
    end;
  end;
end;

procedure TwBase.SQLReadDS(aDataSet: TIBDataSet; aTransaction: TIBTransaction;
  aDataSource: TDataSource; aSQL: string);
var
  _Transaction: TIBTransaction;
  _Field: TFloatField;
  i:      integer;
begin
  try
    Log('SQLReadDS...');
    screen.Cursor := crSQLWait;
    //Application.ProcessMessages;

    if not fdbConnect.Connected then
      fdbConnect.Connected := True;

    if not Assigned(aTransaction) or fLongTransaction then
      _Transaction := Transaction
    else
      _Transaction := aTransaction;

    aDataSet.Close;
    aDataSet.Transaction := _Transaction;

    //if not fLongTransaction then
    //begin
    //  _Transaction.Active := False;
    //  _Transaction.StartTransaction;
    //end;

    with aDataSet do
    begin
      //Close;

      if Length(aSQL)>0 then
      begin
        SelectSQL.Clear;
        SelectSQL.Text := aSQL;
      end;

      Log(SelectSQL.Text);

      try
        Open;
      finally
        screen.Cursor := crDefault;
      end;
    end;

    Log('SQLReadDS [OK]');
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLReadDS]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      //_Transaction.Rollback;
      raise;
    end;
  end;
end;

function TwBase.SQLReadDS(aTableName: string; aFieldsArray: array of string;
  aWhere, aOrderBy: string): TDataSource;
var
  _Fields: string;
  ifield:  integer;
begin
  try
    Log('SQLReadDS...');

    _Fields := '';
    if Length(aFieldsArray) = 0 then
      exit;

    for ifield := 0 to Length(aFieldsArray)-1 do
    begin
      if ifield>0 then
        _Fields := _Fields+',';
      _Fields   := _Fields+aFieldsArray[ifield];
    end;

    if Length(aWhere)>0 then
      aWhere := 'WHERE '+aWhere;
    if Length(aOrderBy)>0 then
      aOrderBy := 'ORDER BY '+aOrderBy;

    Result := SQLReadDS('SELECT '+_Fields+' FROM "'+aTableName +
      '" '+aWhere+' '+aOrderBy+';');

    Log('SQLReadDS [OK]');
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLReadDS]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      raise;
    end;
  end;
end;

function TwBase.PrintQueryParams(adbQueryUpdate:TIBSQL):string;
var
  i: Integer;
begin
  Result:='';
     for i:=0 to adbQueryUpdate.Params.Count-1 do
       Result:= Result+':'+adbQueryUpdate.Params[i].Name+'='+VarToStr(adbQueryUpdate.Params[i].Value)+LineEnding;
end;

function TwBase.SQLInsert(aTableName: string; aFieldsArray: array of string;
  aValuesArray: array of variant; const aCommit: boolean): integer;
var
  _Transaction: TIBTransaction;
  _Fields, _Params: string;
  i: integer;
begin
  Result := -1;
  try
    Log('---------------------------------------------------');
    Log('SQLInsert...');

    if (Length(aFieldsArray)<0) or (Length(aValuesArray)<0) then
    begin
      Result := -1;
      exit;
    end;

    _Fields := '';
    _Params := '';


    _Transaction := fdbQueryUpdate.Transaction;

    if not fdbConnect.Connected then
      fdbConnect.Connected := True;

    if not _Transaction.Active and not fLongTransaction then
      _Transaction.StartTransaction;

    with fdbQueryUpdate do
    begin
      ParamCheck := True;

      Close;
      SQL.Clear;

      for i := 0 to Length(aFieldsArray)-1 do
      begin
        if i>0 then
        begin
          _Fields := _Fields+',';
          _Params := _Params+',';
        end;
        _Fields := _Fields+aFieldsArray[i];
        _Params := _Params+':'+aFieldsArray[i];

      end;

      SQL.Text := 'INSERT INTO "'+aTableName+'" ('+_Fields +
        ') VALUES ('+_Params+') RETURNING ID;';
      Log(SQL.Text);

      for i := 0 to Length(aFieldsArray)-1 do
      begin
        ParamByName(aFieldsArray[i]).Value := aValuesArray[i];
        Log(':'+aFieldsArray[i]+'='+string(aValuesArray[i]));
      end;

      ExecQuery;
      Result := FieldByName('ID').AsInteger;
      Close;
    end;
    if aCommit and not fLongTransaction then
    begin
      _Transaction.Commit;
      Log('SQLInsert [OK] | [TransactionEnd] = COMMIT');
    end
    else
      Log('SQLInsert [OK] | [TransactionEnd] = NONE');
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLInsert]: "'+E.Message+'"');
      __Log.Add('[Error]'+LineEnding+PrintQueryParams(fdbQueryUpdate));
      __Log.SaveLogError(E);
      SetStatus('Ошибка [SQLInsert]: "'+E.Message+'"', True);
      //      ShowMessage('Ошибка [SQLInsert]: "'+E.Message+'"');
      _Transaction.Rollback;
      raise;
    end;
  end;
end;

function TwBase.SQLInsert(aTableName: string; aFieldsArray: array of string;
  aValuesArray: array of variant; aWhereFields: string;
  const aCommit: boolean): integer;
var
  _Transaction: TIBTransaction;
  _Fields, _Params: string;
  i: integer;
begin
  Result := -1;
  try
    Log('---------------------------------------------------');
    Log('SQLInsert...');

    if (Length(aFieldsArray)<0) or (Length(aValuesArray)<0) then
    begin
      Result := -1;
      exit;
    end;

    _Fields := '';
    _Params := '';

    _Transaction := fdbQueryUpdate.Transaction;

    if not fdbConnect.Connected then
      fdbConnect.Connected := True;

    if not _Transaction.Active and not fLongTransaction then
      _Transaction.StartTransaction;

    with fdbQueryUpdate do
    begin
      ParamCheck := True;

      Close;
      SQL.Clear;

      for i := 0 to Length(aFieldsArray)-1 do
      begin
        if i>0 then
        begin
          _Fields := _Fields+',';
          _Params := _Params+',';
        end;
        _Fields := _Fields+aFieldsArray[i];
        _Params := _Params+':'+aFieldsArray[i];
        Log(':'+aFieldsArray[i]+'='+string(aValuesArray[i]));
      end;

        Log('MATCHING ('+aWhereFields+')');

      SQL.Text := 'UPDATE OR INSERT INTO "'+aTableName+'" (' +
        _Fields+') VALUES ('+_Params+') MATCHING ('+aWhereFields +
        ') RETURNING ID;';
      Log(SQL.Text);

      for i := 0 to Length(aFieldsArray)-1 do
        ParamByName(aFieldsArray[i]).Value := aValuesArray[i];

      ExecQuery;
      Result := FieldByName('ID').AsInteger;
      Close;
    end;
    if aCommit and not fLongTransaction then
    begin
      _Transaction.Commit;
      Log('SQLInsert [OK] | [TransactionEnd] = COMMIT');
    end
    else
      Log('SQLInsert [OK] | [TransactionEnd] = NONE');
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLInsert]: "'+E.Message+'"');
      __Log.Add('[Error]'+LineEnding+'MATCHING='+aWhereFields+LineEnding+PrintQueryParams(fdbQueryUpdate));
      __Log.SaveLogError(E);
      SetStatus('Ошибка [SQLInsert]: "'+E.Message+'"', True);
      //      ShowMessage('Ошибка [SQLInsert]: "'+E.Message+'"');
      _Transaction.Rollback;
      raise;
    end;
  end;
end;

function TwBase.SQLInsert(aSQLText: string; const aCommit: boolean): integer;
var
  _Transaction: TIBTransaction;
begin
  Result := -1;
  try
    Log('SQLInsert...');

    _Transaction := fdbQueryUpdate.Transaction;

    if not fdbConnect.Connected then
      fdbConnect.Connected := True;

    if not _Transaction.Active and not fLongTransaction then
      _Transaction.StartTransaction;

    with fdbQueryUpdate do
    begin
      //ParamCheck:=true;
      Close;
      SQL.Clear;
      SQL.Text := aSQLText+' RETURNING ID;';
      Log(SQL.Text);
      ExecQuery;
      Result := FieldByName('ID').AsInteger;
      Close;
    end;
    if aCommit and not fLongTransaction then
    begin
      _Transaction.Commit;
      Log('SQLInsert [OK] | [TransactionEnd] = COMMIT');
    end
    else
      Log('SQLInsert [OK] | [TransactionEnd] = NONE');
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLInsert]: "'+E.Message+'"');
      __Log.Add('[Error]'+LineEnding+PrintQueryParams(fdbQueryUpdate));
      __Log.SaveLogError(E);
      SetStatus('Ошибка [SQLInsert]: "'+E.Message+'"', True);
      //      ShowMessage('Ошибка [SQLInsert]: "'+E.Message+'"');
      _Transaction.Rollback;
      raise;
    end;
  end;
end;

function TwBase.SQLExecStringList(aStringList: TStringList;
  const aCommit: boolean): boolean;
var
  i, _cnt: integer;
begin
  Result := False;
  if aStringList.Count = 0 then
    exit;
  _cnt := aStringList.Count;
  try
    Log('SQLExecStringList...');


    for i := 0 to _cnt-1 do
    begin
      SQLInsert(aStringList[i], aCommit);
      if i mod 300 = 0 then
        SetStatus('Обработано: '+IntToStr(i+1)+' из ' +
          IntToStr(_cnt), True);
    end;
    SetStatus('...', True);
    Result := True;
  except
    on E: Exception do
    begin
      Result := False;
      Log('Ошибка [SQLExecStringList]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      SetStatus('Ошибка [SQLExecStringList]: "'+E.Message+'"', True);
      //      ShowMessage('Ошибка [SQLInsert]: "'+E.Message+'"');
      raise;
    end;
  end;
end;

function TwBase.SQLUpdate(aTableName: string; aFieldsArray: array of string;
  aValuesArray: array of variant; aWhereFields: string; const aCommit: boolean): boolean;
var
  _SetString: string;
  _Transaction: TIBTransaction;
  i: integer;
begin
  Result := False;
  try
    Log('SQLUpdate...');

    if (Length(aFieldsArray)<0) or (Length(aValuesArray)<0) then
    begin
      Result := False;
      exit;
    end;

    _SetString := '';

    if Length(aWhereFields)>0 then
      aWhereFields := 'WHERE '+aWhereFields;

    _Transaction := fdbQueryUpdate.Transaction;

    if not fdbConnect.Connected then
      fdbConnect.Connected := True;

    if not _Transaction.Active and not fLongTransaction then
      _Transaction.StartTransaction;

    with fdbQueryUpdate do
    begin
      ParamCheck := True;

      Close;
      SQL.Clear;

      for i := 0 to Length(aFieldsArray)-1 do
      begin
        if i>0 then
          _SetString := _SetString+',';

        _SetString := _SetString+aFieldsArray[i]+'='+':'+aFieldsArray[i];

        Log(':'+aFieldsArray[i]+'='+string(aValuesArray[i]));
      end;

      SQL.Text := 'UPDATE "'+aTableName+'" SET '+_SetString +
        ' '+aWhereFields+';';

      Log(SQL.Text);

      for i := 0 to Length(aFieldsArray)-1 do
      begin
        ParamByName(aFieldsArray[i]).Value := aValuesArray[i];
        Log(':'+aFieldsArray[i]+'='+string(aValuesArray[i]));
      end;
      ExecQuery;
      Close;
    end;
    if aCommit and not fLongTransaction then
    begin
      _Transaction.Commit;
      Log('SQLUpdate [OK] | [TransactionEnd] = COMMIT');
    end
    else
      Log('SQLUpdate [OK] | [TransactionEnd] = NONE');
    Result := True;
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLUpdate]: "'+E.Message+'"');
      __Log.Add('[Error]'+LineEnding+fdbQueryUpdate.SQL.Text);
      __Log.SaveLogError(E);
      SetStatus('Ошибка [SQLUpdate]: "'+E.Message+'"', True);
      //      ShowMessage('Ошибка [SQLInsert]: "'+E.Message+'"');
      _Transaction.Rollback;
      raise;
    end;
  end;
end;

function TwBase.SQLUpdate(aSQLText: string; const aCommit: boolean;
  const aParamCheck: boolean): boolean;
var
  _Transaction: TIBTransaction;
begin
  Result := False;
  try
    Log('SQLUpdate...');

    _Transaction := fdbQueryUpdate.Transaction;

    if not fdbConnect.Connected then
      fdbConnect.Connected := True;

    if not _Transaction.Active and not fLongTransaction then
      _Transaction.StartTransaction;

    with fdbQueryUpdate do
    begin
      ParamCheck := aParamCheck;
      Close;
      SQL.Clear;
      SQL.Text := aSQLText;
      Log(SQL.Text);
      ExecQuery;
      Result := True;
      Close;
    end;
    if aCommit and not fLongTransaction then
    begin
      _Transaction.Commit;
      Log('SQLUpdate [OK] | [TransactionEnd] = COMMIT');
    end
    else
      Log('SQLUpdate [OK] | [TransactionEnd] = NONE');
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLUpdate]: "'+E.Message+'"');
      __Log.Add('[Error]'+LineEnding+fdbQueryUpdate.SQL.Text);
      __Log.SaveLogError(E);
      SetStatus('Ошибка [SQLUpdate]: "'+E.Message+'"', True);
      //      ShowMessage('Ошибка [SQLInsert]: "'+E.Message+'"');
      _Transaction.Rollback;
      raise;
    end;
  end;

end;

function TwBase.SQLDelete(aTableName: string; aWhereFields: string;
  const aCommit: boolean): boolean;
var
  _Transaction: TIBTransaction;
begin
  Result := False;
  try
    Log('SQLDelete...');

    if Length(aTableName) = 0 then
    begin
      Result := False;
      exit;
    end;

    if Length(aWhereFields)>0 then
      aWhereFields := 'WHERE '+aWhereFields;

    _Transaction := fdbQueryUpdate.Transaction;

    if not fdbConnect.Connected then
      fdbConnect.Connected := True;

    if not _Transaction.Active and not fLongTransaction then
      _Transaction.StartTransaction;

    with fdbQueryUpdate do
    begin
      ParamCheck := True;

      Close;
      SQL.Clear;

      SQL.Text := 'DELETE FROM "'+aTableName+'" '+aWhereFields+';';

      Log(SQL.Text);

      ExecQuery;
      Close;
    end;

    if aCommit and not fLongTransaction then
    begin
      _Transaction.Commit;
      Log('SQLDelete [OK] | [TransactionEnd] = COMMIT');
    end
    else
      Log('SQLDelete [OK] | [TransactionEnd] = NONE');

    Result := True;
  except
    on E: Exception do
    begin
      Log('Ошибка [SQLDelete]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      SetStatus('Ошибка [SQLDelete]: "'+E.Message+'"', True);
      //      ShowMessage('Ошибка [SQLInsert]: "'+E.Message+'"');
      _Transaction.Rollback;
      raise;
    end;
  end;
end;

procedure TwBase.SQLTransactionEnd(const aCommit: boolean);
begin
  if aCommit then
  begin
    if fdbTransactionUpdate.Active then
      fdbTransactionUpdate.Commit;
    Log('[TransactionEnd] = COMMIT');
    LongTransaction := False;
  end
  else
  begin
    if fdbTransactionUpdate.Active then
      fdbTransactionUpdate.Rollback;
    Log('[TransactionEnd] = ROLBACK');
    LongTransaction := False;
  end;

end;

function TwBase.RegisterEvents(aEventNames: ArrayOfString): boolean;
var
  i: Integer;
  EventName: string;
begin
  Result:= false;
  if not Assigned(aEventNames) then exit;

  if EventDB.Registered then
         EventDB.UnRegisterEvents;

  for i := 0 to High(aEventNames) do
  begin
    EventName:= aEventNames[i];

    if EventDB.Events.IndexOf(EventName) = -1 then
    begin
     EventDB.Events.Add(EventName);
     Result:= true;
    end else
     Result:= false;
  end;
  EventDB.RegisterEvents;
end;

function TwBase.ReadMaxDateTimeValuesT(aTable, aGroupField, aTimeStampField: string;
  const aWhere: string): ArrayOfDateTime;
var
  _DataSet: TDataSet;
  i:      integer;
  _Where: string;
  _OrderBy: string;
begin
  try
    if Length(aWhere)>0 then
      _Where := ' where '+aWhere;

    _DataSet := SQLReadDS('select '+aGroupField+', max(' +
      aTimeStampField+') from "'+aTable+'" '+_Where+' group by 1;').DataSet;
    _DataSet.Last;
    _DataSet.First;
    SetLength(Result, _DataSet.RecordCount);

    for i := 0 to _DataSet.RecordCount-1 do
    begin
      Result[i] := _DataSet.FieldByName('MAX').AsDateTime;
      _DataSet.Next;
    end;
    _DataSet.Close;
  except
    on E: Exception do
    begin
      Log('Ошибка [ReadMaxDateTimeValues]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      raise;
    end;
  end;

end;

function TwBase.PrepareWhereStringFromDateTime(aFilterGroupField: string;
  aFilterGroup: array of TDateTime): string;
var
  i: integer;
begin
  try
    Result := '';
    // настраиваем параметры фильтрации
    if Length(aFilterGroup)>0 then
    begin
      for i := 0 to Length(aFilterGroup)-1 do
      begin
        if i>0 then
          Result := Result+' OR ';
        Result   := Result+aFilterGroupField+'=' +
          QuotedStr(DateTimeToStr(aFilterGroup[i]));
      end;

    end
    else
      Result := '';
  except
    on E: Exception do
    begin
      Log('Ошибка [PrepareWhereStringFromDateTime]: "'+E.Message+'"');
      __Log.SaveLogError(E);
      raise;
    end;
  end;
end;

function TwBase.MakeArrayFromString(aText: string): ArrayOfString;
var
  i:     integer;
  _List: TStringList;
begin
  Result := nil;

  if Length(aText)>0 then
  begin
    aText := StringReplace(aText, #39, #39+#39, [rfReplaceAll]);
    aText := StringReplace(aText, ' ', ',', [rfReplaceAll]);

    _List := TStringList.Create;
    _List.Delimiter := ',';
    _list.DelimitedText := aText;

    for i := _List.Count-1 downto 0 do
      if Length(_List[i]) = 0 then
        _List.Delete(i);

    SetLength(Result, _List.Count);

    for i := 0 to _List.Count-1 do
      Result[i] := _List[i];

    FreeAndNil(_List);

  end;
end;

function TwBase.MakeStringFromArray(aTextArray: ArrayOfString): string;
var
  i: integer;
begin
  Result := '';
  if not Assigned(aTextArray) then
    exit;
  for i := 0 to High(aTextArray) do
    if Length(Result)>0 then
      Result := Result+','+aTextArray[i]
    else
      Result := aTextArray[i];
end;

function TwBase.MakeStringFromArray(aIntArray: ArrayOfInteger): string;
var
  i: integer;
begin
  Result := '';
  if not Assigned(aIntArray) then
    exit;
  for i := 0 to High(aIntArray) do
    if Length(Result)>0 then
      Result := Result+','+IntToStr(aIntArray[i])
    else
      Result := IntToStr(aIntArray[i]);
end;

function TwBase.MakeArrayIntegerFromString(aText: string): ArrayOfInteger;
var
  i:     integer;
  _List: TStringList;
begin
  Result := nil;

  if Length(aText)>0 then
  begin
    aText := StringReplace(aText, #39, #39+#39, [rfReplaceAll]);
    aText := StringReplace(aText, ' ', ',', [rfReplaceAll]);

    _List := TStringList.Create;
    _List.Delimiter := ',';
    _list.DelimitedText := aText;

    for i := _List.Count-1 downto 0 do
      if Length(_List[i]) = 0 then
        _List.Delete(i);

    SetLength(Result, _List.Count);

    for i := 0 to _List.Count-1 do
      Result[i] := StrToInt(_List[i]);

    FreeAndNil(_List);

  end;

end;

function TwBase.MakeArrayArrayIntegerFromString(aText: string;
  const aDefaultVal2: integer = -1): ArrayOfArrayInteger;
var
  i, _Pos: integer;
  _List:   TStringList;
begin
  Result := nil;

  if Length(aText)>0 then
  begin
    aText := StringReplace(aText, #39, #39+#39, [rfReplaceAll]);
    aText := StringReplace(aText, ' ', '', [rfReplaceAll]);
    _List := TStringList.Create;
    _List.Delimiter := ',';
    _list.DelimitedText := aText;

    for i := _List.Count-1 downto 0 do
      if Length(_List[i]) = 0 then
        _List.Delete(i);

    SetLength(Result, _List.Count, 2);

    for i := 0 to _List.Count-1 do
    begin
      _Pos := UTF8Pos('|', _List[i]);
      if _Pos = 0 then
      begin
        Result[i, 0] := StrToInt(_List[i]);
        Result[i, 1] := -1;
      end
      else
      begin
        Result[i, 0] := StrToInt(UTF8Copy(_List[i], 1, _Pos-1));
        Result[i, 1] := StrToInt(UTF8Copy(_List[i], _Pos+1,
          UTF8Length(_List[i])-_Pos));
      end;
      if Result[i, 1] = -1 then
        Result[i, 1] := aDefaultVal2;
    end;


    FreeAndNil(_List);

  end;

end;

function TwBase.MakeArrayArrayVariantFromString(aText: string;
  const aDelimiter2: char = '='; const aDefaultVal2: integer = 0): ArrayOfArrayVariant;
var
  i, _Pos: integer;
  _List:   TStringList;
begin
  Result := nil;

  if Length(aText)>0 then
  begin
    //aText:= StringReplace(aText, #39,#39+#39,[rfReplaceAll]);
    //aText:= StringReplace(aText, ' ','',[rfReplaceAll]);
    _List := TStringList.Create;
    _List.Delimiter := ',';
    _List.StrictDelimiter := True;
    _list.DelimitedText := aText;

    for i := _List.Count-1 downto 0 do
      if Length(_List[i]) = 0 then
        _List.Delete(i);

    SetLength(Result, _List.Count, 2);

    for i := 0 to _List.Count-1 do
    begin
      _Pos := UTF8Pos(aDelimiter2, _List[i]);
      if _Pos = 0 then
      begin
        Result[i, 0] := StrToInt(_List[i]);
        Result[i, 1] := -1;
      end
      else
      begin
        Result[i, 0] := UTF8Copy(_List[i], 1, _Pos-1);
        Result[i, 1] := UTF8Copy(_List[i], _Pos+1, UTF8Length(_List[i])-_Pos);
      end;
      if Result[i, 1] = 0 then
        Result[i, 1] := aDefaultVal2;
    end;


    FreeAndNil(_List);

  end;

end;

function TwBase.PrepareIDOwnerWhere(aWhereString: string;
  aOwnerFieldValueArr: array of variant): string;
begin
  Result := '';
  if (Length(aOwnerFieldValueArr)>0) and (Length(aOwnerFieldValueArr)<3) then
    if (aOwnerFieldValueArr[1]<>0) then
    begin
      if Length(aWhereString)>0 then
        Result := '('+aWhereString+') AND ';

      Result := Result+' '+string(aOwnerFieldValueArr[0]) +
        '='+string(aOwnerFieldValueArr[1])+' ';

    end
    else
      Result := aWhereString;
end;

function TwBase.PrepareWhereString(aFilterGroupField: string;
  aFilterGroup: ArrayOfInteger; const aANDOR: string): string;
var
  i: integer;
begin
  Result := '';
  // настраиваем параметры фильтрации
  if Length(aFilterGroup)>0 then
  begin
    for i := 0 to Length(aFilterGroup)-1 do
    begin
      if i>0 then
        Result := Result+' '+aANDOR+' ';
      Result   := Result+aFilterGroupField+'='+IntToStr(aFilterGroup[i]);
    end;
  end
  else
    Result := '';
end;

function TwBase.GetWhere(const uSQL: string): string;
var
  _PosWhere, aPosWhereRes, aPosWhereEnd: PtrInt;
  aSQL: String;
begin

   aSQL:= UTF8UpperCase(uSQL);
   aPosWhereRes:= 0;

   _PosWhere:=1;
   while _PosWhere>0 do begin
   _PosWhere:= UTF8Pos('WHERE',aSQL,_PosWhere+1);
   if (_PosWhere<>0) and (getUTFSymbol(aSQL,_PosWhere-1) = ' ') and (getUTFSymbol(aSQL,_PosWhere+Length('WHERE')) = ' ')then
      aPosWhereRes:= _PosWhere+Length('WHERE');
  end;

   aPosWhereEnd:= UTF8Pos('GROUP',aSQL,aPosWhereRes);

   if (aPosWhereEnd>0) and not ((getUTFSymbol(aSQL,aPosWhereEnd-1) = ' ') and (getUTFSymbol(aSQL,aPosWhereEnd+5) = ' ')) then
     aPosWhereEnd:= 0;

   if aPosWhereEnd = 0 then
      aPosWhereEnd:= UTF8Pos('ORDER',aSQL,aPosWhereRes);

   if (aPosWhereEnd>0) and not ((getUTFSymbol(aSQL,aPosWhereEnd-1) = ' ') and (getUTFSymbol(aSQL,aPosWhereEnd+5) = ' ')) then
     aPosWhereEnd:= 0;

   if (aPosWhereEnd = 0) then aPosWhereEnd:= UTF8Length(aSQL)+1;

  Result:= UTF8Copy(uSQL, aPosWhereRes, aPosWhereEnd-aPosWhereRes);
end;

function TwBase.WriteWhere(aSQL, aWhere: string; aClearWhere: boolean): string;
var
  _PosWhere, _PosWhereRes: PtrInt;
  _SQL: String;
begin

   _SQL:= UTF8UpperCase(aSQL);
   _PosWhereRes:= 0;

   _PosWhere:=1;
   while _PosWhere>0 do begin
   _PosWhere:= UTF8Pos('WHERE',_SQL,_PosWhere+1);
   if (_PosWhere<>0) and (getUTFSymbol(aSQL,_PosWhere-1) = ' ') and (getUTFSymbol(aSQL,_PosWhere+Length('WHERE')) = ' ')then
      _PosWhereRes:= _PosWhere+Length('WHERE');
  end;

   if aClearWhere then
     UTF8Delete(_SQL,_PosWhereRes,Length(_SQL)-_PosWhereRes+1);

   if aClearWhere then
     UTF8Insert(' ('+aWhere+') ',_SQL,_PosWhereRes)
   else
     UTF8Insert(' ('+aWhere+') AND ',_SQL,_PosWhereRes);

   Result:= _SQL;

end;

function TwBase.WriteWhereEx(const uSQL: string; aWhere: string; const aClearOldWhere: boolean): string;
var
  aPosWhere, aPosWhereRes, aPosWhereEnd, aLengthWhere: Integer;
  aWhereOld: String;
begin
 Result:= UTF8UpperCase(uSQL);
 aPosWhere:= 1;
 aPosWhereRes:= 0;
 aPosWhereEnd:= 0;
 aLengthWhere:= Length(' WHERE ');
 aWhereOld:= '';

 while aPosWhere>0 do begin
   aPosWhere:= UTF8Pos(' WHERE ',Result,aPosWhere+1);

   if (aPosWhere<>0) then
      aPosWhereRes:= aPosWhere;
 end;

 aPosWhereEnd:= UTF8Pos(' GROUP ',Result,aPosWhereRes+1);

 if aPosWhereEnd = 0 then
    aPosWhereEnd:= UTF8Pos(' ORDER ',Result,aPosWhereRes+1);


 if aPosWhereEnd = 0 then aPosWhereEnd:= UTF8Length(Result)+1;

 //wfLog(Format('aPosWhereRes %d | aLengthWhere %d | aPosWhereEnd %d',[aPosWhereRes,aLengthWhere,aPosWhereEnd]));

 if aPosWhereRes>0 then
  begin
   if not aClearOldWhere then
      aWhereOld:= Trim(UTF8Copy(Result, aPosWhereRes+aLengthWhere, aPosWhereEnd-(aPosWhereRes+aLengthWhere)));

   UTF8Delete(Result,aPosWhereRes,aPosWhereEnd-aPosWhereRes)
  end
 else
   aPosWhereRes:= aPosWhereEnd;

 if (Length(aWhereOld)>0) and (Length(aWhere)>0) then
    aWhere:= '('+aWhereOld+') AND ('+aWhere+')'
  else
    if (Length(aWhereOld)>0) and (Length(aWhere)=0) then
      aWhere:= aWhereOld;

 if Length(aWhere)>0 then
   UTF8Insert(' WHERE '+aWhere,Result,aPosWhereRes);

end;

function TwBase.GetRowsCount(aSQL: string): integer;
var
  _PosFrom, _PosSelect, _PosFromRes, _PosOrderBy, _PosJoin: PtrInt;
  _CurrentCursor: TCursor;

begin
   _PosFromRes:= 0;
   _CurrentCursor:= screen.Cursor;
   screen.Cursor:= crSQLWait;

   //aSQL:= UTF8UpperCase(aSQL);
   _PosSelect:= UTF8Pos('SELECT',aSQL,1)+Length('SELECT');

   _PosFrom:=_PosSelect;
   _PosJoin:= UTF8Pos('JOIN',aSQL,1);
   if _PosJoin = 0 then
      _PosJoin:= Length(aSQL);

   while _PosFrom>0 do begin
   _PosFrom:= UTF8Pos('FROM',aSQL,_PosFrom+1);
   if (_PosFrom<>0) and (_PosFrom<_PosJoin) then
      _PosFromRes:= _PosFrom;
  end;


   UTF8Delete(aSQL,_PosSelect,_PosFromRes-_PosSelect);
   UTF8Insert(' COUNT(*) ',aSQL,_PosSelect);

   _PosOrderBy:= UTF8Pos('ORDER BY',aSQL,1);

   UTF8Delete(aSQL,_PosOrderBy,Length(aSQL)-_PosOrderBy+1);

   Result:= SQLReadDS(aSQL).DataSet.FieldByName('COUNT').AsInteger;

   screen.Cursor:= _CurrentCursor;
end;

function TwBase.IsNum(aText: string): boolean;
var
  _i: integer;
  _d: double;
begin

  if TryStrToInt(aText, _i) or TryStrToFloat(aText, _d) then
  begin
    if Length(aText)>1 then
    begin
      if (aText[1] = '0') and (aText[2]<>DecimalSeparator) then
        Result := False
      else
        Result := True;
    end
    else
      Result := True;
  end
  else
    Result := False;

end;

function TwBase.GetCurrencyArray(): ArrayOfCurrency;
var
  i: Integer;
  _DS: TDataSet;
begin
  _DS:= self.SQLReadDS('CURRENCY',['KURS'],'','').DataSet;

  SetLength(Result,_DS.RecordCount);
  for i:=0 to _DS.RecordCount-1 do
  begin
      Result[i]:=_DS.Fields[0].AsCurrency;
      _DS.Next;
  end;
  _DS.Close;

end;

function TwBase.TextInArray(aText: string; aTextArray: ArrayOfString): boolean;
var
  i: Integer;
begin
 Result:= false;
 for i:=0 to High(aTextArray) do
     if aText = aTextArray[i] then
        begin
          Result:= true;
          Break;
        end;
end;

end.
