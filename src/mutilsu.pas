unit mUtilsU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, Clipbrd,
  ComCtrls, Controls, CSVdocument,
  db, DBGrids, Dialogs, FileUtil, Forms, IBCustomDataSet, IBDatabase, LazFileUtils,
  LaZUTF8, Menus, StdCtrls, LCLIntf,
  SysUtils, UtilsU, wBaseU, wDBGridU, wDBImportU, wDBTreeU, wFuncU, wLogU, wZipperU, wTProgressU, wTypesU,
  wCustomClassThreadU;

type
  TExportObject = (eoOwner, eoFormats, eoMatchings, eoPrice, eoCatalog, eoPriceVersions, eoCustomObject);
  TExportObjects = set of TExportObject;

  TExportMode  = (emData, emTemplate, emOpenViewver, emSaveFile);
  TExportModes = set of TExportMode;

  TImportMode  = (imReplace);
  TImportModes = set of TImportMode;

  { TDataExportThread }
  TDataExportThread = class(TwCustomThreadWithProgressBar)
  protected
    procedure Execute; override;
  private
    fPathExport: string;
    fBase:   TwBase;
    fFileName: string;
    fFileFormat: TFileFormat;

    fExportModes: TExportModes;

    fSQLCustomObject:string;

    fExportObjects: TExportObjects;
    fTableHeads: ArrayOfString;
    fWhereString: string;
    fOutStringArr: ArrayOfString;

    procedure ExportRun(const aSQL, aTableName: string; const aStep: integer=10000);
    procedure Export_CATALOG();
    procedure Export_CATALOG_GROUP();
    procedure Export_CATALOG_MATCHING();
    procedure Export_CATALOG_SCODS();
    procedure Export_CUSTOM();
    procedure Export_FORMATS();
    procedure Export_OWNER();
    procedure Export_SETTINGS();
    procedure Export_PL_GROUP();
    procedure Export_PL_ITEMS();
    procedure Export_PL_SCODS();
    procedure Export_PL_VERSIONS();
    procedure PackFiles(aZipFile: string);

  public
    constructor Create(CreateSuspended: boolean);
    destructor Destroy(); override;
  end;

  TImportTables = record
    fOWNER:    integer;
    fFORMATS:  integer;
    fCATALOG_GROUP: integer;
    fCATALOG:  integer;
    fCATALOG_SCODS: integer;
    fCATALOG_MATCHING: integer;
    fPL_GROUP: integer;
    fPL_ITEMS: integer;
    fPL_SCODS: integer;
    fPL_VERSIONS: integer;
    fSETTINGS: integer;
    fPRICELISTS_TIMESTAMPS: integer;
    fVERSION: integer;
  end;
  //
  //TImportTable: array of TImportTableField;

  { TDataImportThread }

  TDataImportThread = class(TwCustomThread)
  protected
    procedure Execute; override;
  private
    fFileList: TStringList;
    fBase: TwBase;
    fImportModes: TImportModes;
    fImportTables: TImportTables;

    fFileName: string;
    fZipper: TwZipper;

    procedure Import_Run();
  public
    constructor Create(CreateSuspended: boolean);
    destructor Destroy(); override;
  end;

  { TUtils }

  TUtils = class(TObject)
  private
    fBase:   TwBase;
    fDefaultFilterIndex: integer;
    fDefaultTemplateFileName: String;
    fonEndOperation: TNotifyEvent;
    fOutStringArr: ArrayOfString;
    fOwner:  TComponent;
    fMemo:   TMemo;
    fDataExport: TDataExportThread;
    fDataImport: TDataImportThread;
    fProgress: TProgress;
    //fProgress: TProgress;
    FResult: boolean;
    fSQLCustomObject: string;
    fWhereString: string;

    procedure onProgressInit(const aProgressBar: TProgressBarName; aMax: integer);
    procedure onProgressUpdate(const aProgressBar: TProgressBarName; aValue: integer);
    procedure onStatusUpdate(Sender: TObject);
    procedure onEndThread(Sender: TObject);

    procedure ExportData(aFileName: string; aExportObjects: TExportObjects; const aFileFormat: TFileFormat; aExportModes: TExportModes);
    procedure onStopForce(Sender: TObject);
  public

    constructor Create(Sender: TObject; aBase: TwBase);
    destructor Destroy();

    property Base: TwBase read fBase write fBase;
    property Memo: TMemo read fMemo write fMemo;
    property Result: boolean read FResult write FResult;
    property WhereString: string read fWhereString write fWhereString;
    property SQLCustomObject:string read fSQLCustomObject write fSQLCustomObject;
    property DefaultFilterIndex:integer read fDefaultFilterIndex write fDefaultFilterIndex;
    property DefaultTemplateFileName: string read fDefaultTemplateFileName write fDefaultTemplateFileName;
    property OutStringArr: ArrayOfString read fOutStringArr write fOutStringArr;

    property onEndOperation: TNotifyEvent read fonEndOperation write fonEndOperation;

    procedure ExportData(aExportObjects: TExportObjects; const aExportModes: TExportModes = [emTemplate, emSaveFile]);
    procedure ImportData(aFileName: string);
    //procedure
    procedure SetStatus(aText: string; const aStatus: boolean;
      const aLogSection: boolean);
  end;


implementation

const
  cExportTempDir = 'ExportData';

{ TDataImportThread }

procedure TDataImportThread.Execute;
procedure ClearImportTablesList();
begin
    fImportTables.fOWNER:=-1;
    fImportTables.fFORMATS:=-1;
    fImportTables.fCATALOG_GROUP:=-1;
    fImportTables.fCATALOG:=-1;
    fImportTables.fCATALOG_SCODS:=-1;
    fImportTables.fCATALOG_MATCHING:=-1;
    fImportTables.fPL_GROUP:=-1;
    fImportTables.fPL_ITEMS:=-1;
    fImportTables.fPL_SCODS:=-1;
    fImportTables.fPL_VERSIONS:=-1;
    fImportTables.fSETTINGS:=-1;
    fImportTables.fPRICELISTS_TIMESTAMPS:=-1;
    fImportTables.fVERSION:=-1;
end;

var
  i: integer;
  _GenValue: integer;
begin
    fFileList := fZipper.ReadFileList(fFileName);
    _GenValue:=0;

    ClearImportTablesList();

    SetStatus('Подготовка импорта...');

    for i := 0 to fFileList.Count-1 do
      case StringReplace(fFileList[i], cExportTempDir+'/', '', [rfIgnoreCase]) of
        'OWNER':
          fImportTables.fOWNER := i;

        'FORMATS':
          fImportTables.fFORMATS := i;

        'CATALOG_GROUP':
          fImportTables.fCATALOG_GROUP := i;

        'CATALOG':
          fImportTables.fCATALOG := i;

        'CATALOG_SCODS':
          fImportTables.fCATALOG_SCODS := i;

        'CATALOG_MATCHING':
          fImportTables.fCATALOG_MATCHING := i;

        'PL_GROUP':
          fImportTables.fPL_GROUP := i;

        'PL_ITEMS':
          fImportTables.fPL_ITEMS := i;

        'PL_SCODS':
          fImportTables.fPL_SCODS := i;

        'PL_VERSIONS':
          fImportTables.fPL_VERSIONS := i;

        'SETTINGS':
          fImportTables.fSETTINGS := i;

        'PRICELISTS_TIMESTAMPS':
          fImportTables.fPRICELISTS_TIMESTAMPS := i;

        'VERSION':
          fImportTables.fVERSION:= i;
      end;

    try

        fBase.LongTransaction:= true;

        SetStatus('Проверка версии архива...');
        // задел на будущее - при изменении таблиц в зависимости от версии будем что-то менять в импорте
        if fImportTables.fVERSION>-1 then
        begin
           Result:= CheckVersionFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fVERSION],PathTmp_Unsafe));
        end else
          raise Exception.Create('Ошибка определения версии пакета! Импорт невозможен.');

        SetStatus('Импорт контрагентов...');
          if (fImportTables.fOWNER>-1) and (fImportTables.fSETTINGS>-1) then
          begin
            fBase.SQLDelete('OWNER','',false);
            Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fOWNER],PathTmp_Unsafe),false);

            fBase.SQLDelete('SETTINGS','',false);
            Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fSETTINGS],PathTmp_Unsafe),false);
            fBase.SQLUpdate('SETTINGS',['FVALUE'],[GetVersion],'NAME LIKE ''dbVersion''',false);
            fBase.SQLUpdate('SETTINGS',['FVALUE'],[0],'NAME LIKE ''progCountStart''',false);
          end
          else
             raise Exception.Create('Ошибка! Таблица "Контрагенты" не найдена! Дальнейший импорт невозможен');
        SetStatus('Импорт контрагентов...[OK]');

        SetStatus('Импорт форматов...');
          if fImportTables.fFORMATS>-1 then
            Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fFORMATS],PathTmp_Unsafe),false)
          else
             raise Exception.Create('Ошибка! Таблица "Форматы" не найдена! Дальнейший импорт невозможен');
        SetStatus('Импорт форматов...[OK]');

          if fImportTables.fCATALOG_GROUP>-1 then
          begin
            SetStatus('Импорт групп каталога...');
            Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fCATALOG_GROUP],PathTmp_Unsafe),false);
            SetStatus('Импорт групп каталога...[OK]');
          end;

          if (fImportTables.fCATALOG>-1) and (fImportTables.fCATALOG_GROUP>-1) then
          begin
            SetStatus('Импорт каталога...');
            Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fCATALOG],PathTmp_Unsafe),false);
            SetStatus('Импорт каталога...[OK]');
          end;

          if (fImportTables.fCATALOG_SCODS>-1) and (fImportTables.fCATALOG>-1) and (fImportTables.fCATALOG_GROUP>-1) then
          begin
            SetStatus('Импорт штрих-кодов каталога...');
            Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fCATALOG_SCODS],PathTmp_Unsafe),false);
            SetStatus('Импорт штрих-кодов каталога...[OK]');
          end;


          if fImportTables.fPL_GROUP>-1 then
          begin
            SetStatus('Импорт групп прайс-листов...');
            Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fPL_GROUP],PathTmp_Unsafe),false);
            SetStatus('Импорт групп прайс-листов...[OK]');
          end;


          if (fImportTables.fPL_ITEMS>-1) and (fImportTables.fPL_GROUP>-1) then
          begin
            SetStatus('Импорт прайс-листов...');
            Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fPL_ITEMS],PathTmp_Unsafe),false);
            SetStatus('Импорт прайс-листов...[OK]');
          end;

        if (fImportTables.fPL_SCODS>-1) and (fImportTables.fPL_ITEMS>-1) and (fImportTables.fPL_GROUP>-1) then
        begin
          SetStatus('Импорт штрих-кодов прайс-листа...');
            Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fPL_SCODS],PathTmp_Unsafe),false);
          SetStatus('Импорт штрих-кодов прайс-листа...[OK]');
        end;

        if (fImportTables.fCATALOG_MATCHING>-1) and (fImportTables.fPL_ITEMS>-1) and (fImportTables.fCATALOG>-1) then
        begin
          SetStatus('Импорт соответствий...');
           Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fCATALOG_MATCHING],PathTmp_Unsafe),false);
          SetStatus('Импорт соответствий...[OK]');
        end;

        if (fImportTables.fPL_ITEMS>-1) and (fImportTables.fPL_VERSIONS>-1) then
        begin
          SetStatus('Импорт архива прайс-листов...');
          Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fPL_VERSIONS],PathTmp_Unsafe),false);
          SetStatus('Импорт архива прайс-листов...[OK]');
        end;


        if (fImportTables.fPL_ITEMS>-1) and (fImportTables.fPL_VERSIONS>-1) and (fImportTables.fPRICELISTS_TIMESTAMPS>-1) then
           Result:= fBase.ImportTableFromFile(fZipper.ExtractOneFile(fFileName,fFileList[fImportTables.fPRICELISTS_TIMESTAMPS],PathTmp_Unsafe),false);


      fBase.SQLTransactionEnd(true);

      SetStatus('Синхронизация генераторов...');
      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM PL_ITEMS')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_PL_ITEMS_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM PL_VERSIONS')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_PL_VERSIONS_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM PL_GROUP')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_PL_GROUP_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM PL_SCODS')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_PL_SCODS_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM OWNER')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_OWNER_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM CATALOG')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_CATALOG_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM CATALOG_GROUP')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_CATALOG_GROUP_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM CATALOG_SCODS')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_CATALOG_SCODS_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM FORMATS')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_FORMATS_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM CATALOG_MATCHING')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_CATALOG_MATCHING_ID" TO '+IntToStr(_GenValue),true);

      _GenValue:= fBase.SQLReadArr('SELECT CASE WHEN MAX(ID) IS NULL THEN 0 ELSE MAX(ID) END FROM SETTINGS')[0,0];
      fBase.SQLUpdate('SET GENERATOR "GEN_SETTINGS_ID" TO '+IntToStr(_GenValue),true);

      SetStatus('Синхронизация генераторов... [OK]');

      SetStatus('Пересчет индексов...');
      fBase.SQLUpdate('ALTER INDEX PL_ITEMS_VENDORCODE INACTIVE;');
      fBase.SQLUpdate('ALTER INDEX PL_VERSIONS_FTIMESTAMP INACTIVE;');
      fBase.SQLUpdate('ALTER INDEX CTG_VENDORCODE INACTIVE;');
      fBase.SQLUpdate('ALTER INDEX CTG_NAME INACTIVE;');

      fBase.SQLUpdate('ALTER INDEX PL_ITEMS_VENDORCODE ACTIVE;');
      fBase.SQLUpdate('ALTER INDEX PL_VERSIONS_FTIMESTAMP ACTIVE;');
      fBase.SQLUpdate('ALTER INDEX CTG_VENDORCODE ACTIVE;');
      fBase.SQLUpdate('ALTER INDEX CTG_NAME ACTIVE;');
      SetStatus('Пересчет индексов... [OK]');

      if Assigned(onEndThread) then onEndThread(self);
    except
      on E: Exception do
      begin
        Result:= false;
        if  fBase.LongTransaction then fBase.SQLTransactionEnd(false);
        SetStatus(E.Message);
        if Assigned(onEndThread) then onEndThread(self);
      end;
    end;

end;

procedure TDataImportThread.Import_Run();
begin

end;


constructor TDataImportThread.Create(CreateSuspended: boolean);
begin
  FreeOnTerminate := True;
  fZipper := TwZipper.Create();
  inherited Create(CreateSuspended);
end;

destructor TDataImportThread.Destroy();
begin
  fZipper.Destroy();
  inherited Destroy();
end;

{ TDataExportThread }

procedure TDataExportThread.Export_CATALOG_GROUP();
var
  _SQL, _TableName: string;
begin
  _TableName := 'CATALOG_GROUP';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT * FROM CATALOG_GROUP ORDER BY ID';

  ExportRun(_SQL, _TableName, 20000); // запуск экспорта
end;

procedure TDataExportThread.Export_CATALOG();
var
  _SQL, _TableName: string;
begin
  _TableName := 'CATALOG';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT CTG.* FROM CATALOG CTG ORDER BY CTG.ID';

  ExportRun(_SQL, _TableName); // запуск экспорта
end;

procedure TDataExportThread.Export_CATALOG_SCODS();
var
  _SQL, _TableName: string;
begin
  _TableName := 'CATALOG_SCODS';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT CTGS.* FROM CATALOG_SCODS CTGS';

  ExportRun(_SQL, _TableName, 100000); // запуск экспорта
end;

procedure TDataExportThread.Export_CATALOG_MATCHING();
var
  _SQL, _TableName: string;
begin
  _TableName := 'CATALOG_MATCHING';
  _SQL := '';
  fTableHeads:= nil;

  if emData in fExportModes then
    _SQL := 'SELECT CTGMTH.* FROM CATALOG_MATCHING CTGMTH';

  if emTemplate in fExportModes then
  begin
    _SQL := 'SELECT CTGMTH.ID, OWNCTG.ID CATALOG_IDOWNER,'+
      ' OWNCTG.NAME CATALOG_OWNER, '
      +' CTG.VENDORCODE CATALOG_VENDORCODE, '
      +' CTGMTH.QUANTITYINPACKING, '
      +' PLI.VENDORCODE PRICE_VENDORCODE, '
      +' OWNPLI.ID PRICE_IDOWNER, '
      +' OWNPLI.NAME PRICE_OWNER '
      +' FROM CATALOG_MATCHING CTGMTH '
      +' LEFT JOIN CATALOG CTG ON (CTG.ID=CTGMTH.IDCATALOG) '
      +' LEFT JOIN PL_ITEMS PLI ON (PLI.ID=CTGMTH.IDPL_ITEMS) '
      +' LEFT JOIN OWNER OWNPLI ON (OWNPLI.ID=PLI.IDOWNER) '
      +' LEFT JOIN OWNER OWNCTG ON (OWNCTG.ID=CTG.IDOWNER) '
      +' /*where_string*/ '
      +' ORDER BY OWNCTG.ID,OWNPLI.ID';
    fTableHeads:=['ID','Каталог.Контр.Идент','Каталог.Контр.Наименование','Каталог.Идентификатор','Фасовка','Прайс.Идентификатор','Прайс.Контр.Идент','Прайс.Контр.Наименование'];
  end;

  ExportRun(_SQL, _TableName); // запуск экспорта
end;

procedure TDataExportThread.Export_PL_GROUP();
var
  _SQL, _TableName: string;
begin
  _TableName := 'PL_GROUP';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT * FROM PL_GROUP ORDER BY IDOWNER';

  ExportRun(_SQL, _TableName, 20000); // запуск экспорта
end;

procedure TDataExportThread.Export_PL_ITEMS();
var
  _SQL, _TableName: string;
begin
  _TableName := 'PL_ITEMS';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT PLI.* FROM PL_ITEMS PLI';

  ExportRun(_SQL, _TableName); // запуск экспорта
end;

procedure TDataExportThread.Export_PL_SCODS();
var
  _SQL, _TableName: string;
begin
  _TableName := 'PL_SCODS';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT PLIS.* FROM PL_SCODS PLIS';

  ExportRun(_SQL, _TableName, 100000); // запуск экспорта
end;

procedure TDataExportThread.Export_PL_VERSIONS();
var
  _SQL, _TableName: string;
begin
  _TableName := 'PL_VERSIONS';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT * FROM PL_VERSIONS';

  ExportRun(_SQL, _TableName, 150000); // запуск экспорта

  _TableName := 'PRICELISTS_TIMESTAMPS';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT * FROM PRICELISTS_TIMESTAMPS';

  ExportRun(_SQL, _TableName, 150000); // запуск экспорта


end;

procedure TDataExportThread.Export_OWNER();
var
  _SQL, _TableName: string;
begin
  _TableName := 'OWNER';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT * FROM OWNER ORDER BY IDPARENT,ID';

  ExportRun(_SQL, _TableName); // запуск экспорта
end;

procedure TDataExportThread.Export_SETTINGS();
var
  _SQL, _TableName: string;
begin
  _TableName := 'SETTINGS';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT * FROM SETTINGS ORDER BY ID';

  ExportRun(_SQL, _TableName); // запуск экспорта
end;

procedure TDataExportThread.Export_FORMATS();
var
  _SQL, _TableName: string;
begin
  _TableName := 'FORMATS';
  _SQL := '';

  if emData in fExportModes then
    _SQL := 'SELECT * FROM FORMATS ORDER BY IDOWNER,ID';

  ExportRun(_SQL, _TableName); // запуск экспорта
end;

procedure TDataExportThread.Export_CUSTOM();
begin

  ExportRun(fSQLCustomObject, ''); // запуск экспорта
end;

procedure TDataExportThread.PackFiles(aZipFile: string);
var
  _Zipper: TwZipper;
  _FileVersion: TFileStream;
begin

  _Zipper := TwZipper.Create();

  _FileVersion:= TFileStream.Create(fPathExport+'VERSION',fmCreate);

  try
    WriteUTF8String(_FileVersion,GetVersion);
  finally
    FreeAndNil(_FileVersion);
  end;

  try
    _Zipper.PackAllFiles(aZipFile, fPathExport);

  finally
    _Zipper.Destroy();
    if DeleteDirectory(fPathExport, True) then
      RemoveDirUTF8(fPathExport);
  end;
end;

procedure TDataExportThread.ExportRun(const aSQL, aTableName: string;
  const aStep: integer = 10000);
var
  _FileName: string;
  _SQL: string;

begin
  _SQL:= '';
  if emData in fExportModes then
    _FileName := fPathExport+aTableName
    else
    _FileName:= fFileName;

  //fBase.AddExt(fPathExport+aTableName,aExportFormat);
  _SQL:= aSQL;
  if Length(fWhereString)>0 then
     _SQL:= StringReplace(_SQL,'/*where_string*/',fWhereString,[rfIgnoreCase]);
  fBase.OutStringArr:= fOutStringArr;

  Result := fBase.ExportTableToFile(_SQL, _FileName, fFileFormat, fTableHeads, aStep);

end;

procedure TDataExportThread.Execute;
var
  _CountObjects: integer;
begin
  try
    Result := False;

    fPathExport := PathTmp_Unsafe+cExportTempDir+DirectorySeparator;

    fBase.onProgressInit:= onProgressInit;
    fBase.onProgressUpdate:= onProgressUpdate;

    _CountObjects:= BitTest(fExportObjects,sizeof(fExportObjects)*8);

    ProgressInit(pbBottom, _CountObjects);

    if not DirectoryExistsUTF8(fPathExport) then
      ForceDirectoriesUTF8(fPathExport);

      if eoCustomObject in fExportObjects then
        begin

          SetStatus('Экспорт данных...');
          Export_CUSTOM();
          SetStatus('Экспорт данных...[OK]');
          ProgressUpdate(pbBottom);
        end;

      if eoOwner in fExportObjects then
        begin
          SetStatus('Экспорт контрагентов...');
          Export_SETTINGS();
          Export_OWNER();
          SetStatus('Экспорт контрагентов...[OK]');

          ProgressUpdate(pbBottom);
        end;

      if eoFormats in fExportObjects then
        begin
          SetStatus('Экспорт форматов...');
          Export_FORMATS();
          SetStatus('Экспорт форматов...[OK]');

          ProgressUpdate(pbBottom);
        end;

      if eoCatalog in fExportObjects then
        begin
          SetStatus('Экспорт групп каталога...');
          Export_CATALOG_GROUP(); //
          SetStatus('Экспорт групп каталога...[OK]');

          SetStatus('Экспорт каталога...');
          Export_CATALOG();
          SetStatus('Экспорт каталога...[OK]');

          SetStatus('Экспорт штрих-кодов каталога...');
          Export_CATALOG_SCODS();
          SetStatus('Экспорт штрих-кодов каталога...[OK]');

          ProgressUpdate(pbBottom);
        end;

      if eoPrice in fExportObjects then
        begin
          SetStatus('Экспорт групп прайс-листов...');
          Export_PL_GROUP();
          SetStatus('Экспорт групп прайс-листов...[OK]');

          SetStatus('Экспорт прайс-листов...');
          Export_PL_ITEMS();
          SetStatus('Экспорт прайс-листов...[OK]');

          SetStatus('Экспорт штрих-кодов прайс-листа...');
          Export_PL_SCODS();
          SetStatus('Экспорт штрих-кодов прайс-листа...[OK]');

          ProgressUpdate(pbBottom);
        end;

      if eoMatchings in fExportObjects then
        begin
          SetStatus('Экспорт соответствий...');
          Export_CATALOG_MATCHING();
          SetStatus('Экспорт соответствий...[OK]');

          ProgressUpdate(pbBottom);
        end;

      if eoPriceVersions in fExportObjects then
        begin
          SetStatus('Экспорт архива прайс-листов...');
          Export_PL_VERSIONS();
          SetStatus('Экспорт архива прайс-листов...[OK]');

          ProgressUpdate(pbBottom);
        end;


    if emData in fExportModes then
      begin
        SetStatus('Упаковка данных...');
        PackFiles(fFileName); // упаковываем данные
        SetStatus('Упаковка данных... [OK]');
      end;

    ProgressUpdate(pbBottom);
    if Assigned(onEndThread) then onEndThread(self);

  except
    on E: Exception do begin
      Result:= false;
      SetStatus('Error: '+E.Message);
      if Assigned(onEndThread) then onEndThread(self);
    end;
  end;

end;

constructor TDataExportThread.Create(CreateSuspended: boolean);
begin
  fTableHeads:= nil;
  FreeOnTerminate := True;
  inherited Create(CreateSuspended);
end;

destructor TDataExportThread.Destroy();
begin
  inherited Destroy();
end;

{ TUtils }

procedure TUtils.onStatusUpdate(Sender: TObject);
begin
  if Assigned(fDataExport) then
    SetStatus(fDataExport.Status, False, False);

  if Assigned(fDataImport) then
    SetStatus(fDataImport.Status, False, False);
end;

procedure TUtils.onEndThread(Sender: TObject);
var
  _Status: string;
begin
  if Assigned(fDataExport) then
  begin
    FResult := fDataExport.Result;
    fDataExport.Terminate;
    fDataExport := nil;
  end;

  if Assigned(fDataImport) then
  begin
    FResult := fDataImport.Result;
    fDataImport.Terminate;
    fDataImport := nil;
  end;

  if FResult then
    _Status := 'Операция успешно завершена.'
  else
    _Status := 'Произошла ошибка при выполнении операции!';

  fProgress.ForceClose;

  SetStatus(_Status, False, False);
  //SetStatus(_Status, False, true);

  if Assigned(onEndOperation) then onEndOperation(self);
end;

constructor TUtils.Create(Sender: TObject; aBase: TwBase);
begin
  fBase  := aBase;
  fOwner := TComponent(Sender);
  fDataExport := nil;
  fDataImport := nil;
  fMemo:= nil;
  fProgress:= nil;
  fDefaultFilterIndex:= 1;
  fDefaultTemplateFileName:= 'Экспорт данных';
end;

destructor TUtils.Destroy();
begin
  inherited Destroy();
end;


procedure TUtils.onProgressUpdate(const aProgressBar: TProgressBarName; aValue: integer);
begin
  fProgress.SetBar(aProgressBar, aValue);
end;

procedure TUtils.onProgressInit(const aProgressBar: TProgressBarName; aMax: integer);
begin
  fProgress.InitBar(aProgressBar, aMax);
end;

procedure TUtils.ImportData(aFileName: string);
begin
  if Assigned(fDataImport) then
  begin
    ShowMessage('Импорт уже запущен!');
    exit;
  end;

  fDataImport := TDataImportThread.Create(True);
  fDataImport.fFileName := aFileName;
  fDataImport.fBase := fBase;
  fDataImport.onStatusUpdate := @onStatusUpdate;
  fDataImport.onEndThread := @onEndThread;

  fDataImport.start;
end;

procedure TUtils.ExportData(aExportObjects: TExportObjects; const aExportModes: TExportModes);
var
  SaveDialog: TSaveDialog;
  _FileFormat: TFileFormat;
begin
  SaveDialog:= TSaveDialog.Create(fOwner);
  SaveDialog.Options:= [ofOverwritePrompt];

try
  if emData in aExportModes then
  begin
    if MessageDlg(
      'Экспортировать выбранные данные во внешний файл?',
      mtConfirmation, mbOKCancel, 0) = mrCancel then
      exit;

    //_ExportObjects:= [];
    SaveDialog.FileName:='';
    SaveDialog.FileName:='iPriceSEBackupDataPackage.bdpkg';
    SaveDialog.Filter:='Backup Data Package|*.bdpkg';
    _FileFormat:= ffCSV;
  end else
  begin
    SaveDialog.Filter:='OpenDocument (*.ods)|*.ods| Excel (*.xls)|*.xls| Excel (*.xlsx)|*.xlsx|Comma Text (*.csv)|*.csv';
    SaveDialog.FilterIndex:= fDefaultFilterIndex;

  if eoMatchings in aExportObjects then
    SaveDialog.FileName:='Соответствия'
  else
    SaveDialog.FileName:=fDefaultTemplateFileName;

  end;

  if SaveDialog.Execute then
  begin
    if emTemplate in aExportModes then
      case SaveDialog.FilterIndex of
        1: _FileFormat:= ffODS;
        2: _FileFormat:= ffXLS;
        3: _FileFormat:= ffXLSX;
        4: _FileFormat:= ffCSV;
      end;

    ExportData(SaveDialog.FileName, aExportObjects, _FileFormat, aExportModes);
  end else
    onEndOperation(self);

finally
  SaveDialog.Free;
end;

end;

procedure TUtils.ExportData(aFileName: string; aExportObjects: TExportObjects; const aFileFormat: TFileFormat; aExportModes: TExportModes);
begin
  if Assigned(fDataExport) then
  begin
    ShowMessage('Экспорт уже запущен!');
    exit;
  end;

  fProgress:= TProgress.Create(fOwner);
  fProgress.Caption:= 'Экспорт...';
  fProgress.ShowLog:= false;
  fProgress.ShowBottom:= not (eoCustomObject in aExportObjects);
  fProgress.onStopForce:= @onStopForce;

  screen.Cursor:=crSQLWait;

  fDataExport := TDataExportThread.Create(True);
  fDataExport.fFileName := aFileName;
  fDataExport.fExportObjects := aExportObjects;
  fDataExport.fBase := fBase;
  fDataExport.onStatusUpdate := @onStatusUpdate;
  fDataExport.onEndThread := @onEndThread;
  fDataExport.fExportModes := aExportModes;
  fDataExport.fFileFormat:= aFileFormat;
  fDataExport.fWhereString:= fWhereString;
  fDataExport.fSQLCustomObject:= fSQLCustomObject; // для eoCustomObject;
  fDataExport.fOutStringArr:= OutStringArr;
  fDataExport.onProgressInit:= @onProgressInit;
  fDataExport.onProgressUpdate:= @onProgressUpdate;

  fDataExport.start;

  try
    fProgress.ShowModal;
  finally
    screen.Cursor:=crDefault;
    if Assigned(fProgress) then
      fProgress.Free;
  end;

  if MessageDlg('Открыть полученный файл в программе просмотра?',
        mtConfirmation, mbOKCancel, 0) = mrOK then
        OpenDocument(aFileName);
end;

procedure TUtils.onStopForce(Sender: TObject);
begin
  fBase.StopForce:= true;
end;

procedure TUtils.SetStatus(aText: string; const aStatus: boolean;
  const aLogSection: boolean);
begin
  if Memo = nil then
    wStatus('', aText, true)
  else
  begin
    if aStatus then
      wStatus('', aText, aLogSection)
    else
      Memo.Lines.Add(DateTimeToStr(now())+' | '+aText);
    wLog('[console.log] '+'[Utils] ', aText);
  end;
end;

end.



