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
  db, Graphics, ColorBox, LazUTF8, LazFileUtils, FileUtil,
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
    fIdMainOwner: Integer;
    { private declarations }
    FormIDent: string;
    fBase: TwBase;
    fReport: TwReport;
    fSilentMode: boolean;
    _Version,_VersionRaw: string;
    procedure DBUpdate(aSelectedPrices: ArrayOfInteger);
    procedure ExportCatalog(aPatch: string; aStocks:string; aPrices: string; aType: TwfExportFormat);

    procedure DBBackup();
    procedure EnabledFuncional(AValue: boolean);
    procedure Init(Sender: TObject);
    procedure mKursClick(Sender: TObject);
    procedure mm_ClearDB_onEnd(Sender: TObject);
    procedure mm_ClearDB_onStatusUpdate(Sender: TObject);
    procedure mTrayClick(Sender: TObject);
    procedure mTrayMenuFill(Sender: TObject);
    procedure UpdateKurs(const aSilent: boolean=false);
    procedure ClearProps();

    property wFormID: string read FormIDent write FormIDent;
  public
    { public declarations }
    procedure SetStatus(_Text:string;_Log:boolean);
    function GetStatus(_Log:boolean):string;
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
  for i:=0 to tbPluginBtn.ButtonCount-1 do
          tbPluginBtn.Buttons[i].Enabled:=AValue;

  tbPluginBtnControl.Enabled:= true;
  MainMenu.Items[1].Enabled:=AValue;
  MainMenu.Items[2].Enabled:=AValue;
end;

procedure TFmMain.CheckDBVersion(ADBase: TwBase);
var
  _Invoce: TInvoce;
  _Invoce_CreateTable: ArrayOfString;
  i: Integer;
  _CheckVersion, aSQL, aNewVersion: String;
begin
 try
      if ADBase.ReadSettingByName('dbVersion')<>_VersionRaw then
      begin
        wLog('Main','Рефакторинг метаданных БД...');

      aNewVersion:= '0.0.3.39';

      if  CheckLessVersion(ADBase.ReadSettingByName('dbVersion'),aNewVersion) then
      begin

        FmWait:= TFmWait.Create(self);

        FmWait.Show;
        //FmWait.Height:=250;
        //FmWait.Width:=540;
        //FmWait.mStatus.Alignment:=taLeftJustify;
        FMWait.InitBar(3,0);
        wLog('Main','Обновление метаданных БД... до версии '+aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus('Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate('create or alter procedure ANALIS_SEL_ALL_ANALOG '
            +' as '
            +' begin '
            +'  suspend; '
            +' end',true,false);

          ADBase.SQLUpdate('create or alter procedure CATALOG_PL_ITEMS_PRICE '
            +' as '
            +' begin '
            +'  suspend; '
            +' end',true,false);

          ADBase.SQLUpdate('create or alter procedure CATALOG_PL_MIN_PRICE '
            +' as '
            +' begin '
            +'  suspend; '
            +' end',true,false);

          ADBase.SQLUpdate('create or alter procedure PL_GROUP_TO_CATALOG '
            +' as '
            +' begin '
            +'  suspend; '
            +' end',true,false);

          FMWait.SetBar(1);

          ADBase.SQLUpdate('ALTER TABLE CATALOG_MATCHING ALTER QUANTITYINPACKING TO QUANTITYINPACKING_OLD',true,false);
          ADBase.SQLUpdate('ALTER TABLE CATALOG_MATCHING ADD QUANTITYINPACKING DOUBLE PRECISION',true,false);
          ADBase.SQLUpdate('UPDATE CATALOG_MATCHING CM SET CM.QUANTITYINPACKING=CM.QUANTITYINPACKING_OLD',true,false);
          ADBase.SQLUpdate('ALTER TABLE CATALOG_MATCHING DROP QUANTITYINPACKING_OLD',true,false);

          ADBase.SQLUpdate('ALTER TABLE W_TMP_ANALIS_ALL_ANALOG DROP QUANTITYINPACKING',true,false);
          ADBase.SQLUpdate( 'ALTER TABLE W_TMP_ANALIS_ALL_ANALOG ADD QUANTITYINPACKING DOUBLE PRECISION',true,false);

          FMWait.SetBar(2);

          ADBase.SQLUpdate('create or alter procedure ANALIS_SEL_ALL_ANALOG ( '
            +'    PL_ID bigint, '+ wfLineEnding
            +'    HIDE_POS_PL_ID boolean) '+ wfLineEnding
            +'returns ( '+ wfLineEnding
            +'    IDOWNER bigint,'+ wfLineEnding
            +'    OWNERNAME varchar(150),'+ wfLineEnding
            +'    VENDORCODE varchar(300),'+ wfLineEnding
            +'    SCOD varchar(1024),'+ wfLineEnding
            +'    PLNAME varchar(500),'+ wfLineEnding
            +'    PRICE numeric(15,10),'+ wfLineEnding
            +'    PRICE2 numeric(15,10),'+ wfLineEnding
            +'    PRICE3 numeric(15,10),'+ wfLineEnding
            +'    PRICE4 numeric(15,10),'+ wfLineEnding
            +'    PRICE5 numeric(15,10),'+ wfLineEnding
            +'    PRICE6 numeric(15,10),'+ wfLineEnding
            +'    PRICE7 numeric(15,10),'+ wfLineEnding
            +'    PRICE8 numeric(15,10),'+ wfLineEnding
            +'    PRICE9 numeric(15,10),'+ wfLineEnding
            +'    PRICE10 numeric(15,10),'+ wfLineEnding
            +'    STOCK bigint,'+ wfLineEnding
            +'    QUANTITYINPACKINGTEXT varchar(100),'+ wfLineEnding
            +'    ID bigint,'+ wfLineEnding
            +'    FTIMESTAMP timestamp,'+ wfLineEnding
            +'    UNIT varchar(15),'+ wfLineEnding
            +'    STOCKONLYINFO smallint,'+ wfLineEnding
            +'    FCOLOR smallint,'+ wfLineEnding
            +'    QUANTITYINPACKING numeric(15,10),'+ wfLineEnding
            +'    LABEL varchar(255))'+ wfLineEnding
            +'as '+ wfLineEnding
            +'declare variable V_CTG_ID bigint;'+ wfLineEnding
            +'declare variable V_PL_IDOWNER bigint;'+ wfLineEnding
            +'declare variable V_QUANTITYINPACKING_PL_ID numeric(15,10);'+ wfLineEnding
            +'declare variable V_PL_VENDORCODE varchar(300);'+ wfLineEnding
            +'BEGIN '+ wfLineEnding
            +'   DELETE FROM W_TMP_ANALIS_ALL_ANALOG; '+ wfLineEnding
            +' '+ wfLineEnding
            +'   V_CTG_ID = (select MTH.IDCATALOG FROM CATALOG_MATCHING MTH WHERE MTH.IDPL_ITEMS=:PL_ID rows 1); '+ wfLineEnding
            +' '+ wfLineEnding
            +'  if (V_CTG_ID IS NULL) then '+ wfLineEnding
            +'  begin '+ wfLineEnding
            +'      SELECT PL.IDOWNER, PL.VENDORCODE FROM PL_ITEMS PL WHERE PL.ID=:PL_ID '+ wfLineEnding
            +'      INTO '+ wfLineEnding
            +'       :V_PL_IDOWNER, '+ wfLineEnding
            +'       :V_PL_VENDORCODE; '+ wfLineEnding
            +' '+ wfLineEnding
            +'    V_CTG_ID = (select CTG.ID CTGID FROM CATALOG CTG WHERE CTG.VENDORCODE=:V_PL_VENDORCODE AND CTG.IDOWNER=:V_PL_IDOWNER rows 1); '+ wfLineEnding
            +'    INSERT INTO W_TMP_ANALIS_ALL_ANALOG VALUES (:PL_ID,1); '+ wfLineEnding
            +'    V_QUANTITYINPACKING_PL_ID = 1; '+ wfLineEnding
            +'  end else '+ wfLineEnding
            +'  begin '+ wfLineEnding
            +'    INSERT INTO W_TMP_ANALIS_ALL_ANALOG '+ wfLineEnding
            +'    select PL.ID, 1 FROM CATALOG_MATCHING MTH '+ wfLineEnding
            +'    inner join CATALOG CTG on (CTG.ID=MTH.IDCATALOG) '+ wfLineEnding
            +'    inner join PL_ITEMS PL on (PL.VENDORCODE=CTG.VENDORCODE and PL.IDOWNER= CTG.IDOWNER) '+ wfLineEnding
            +'    WHERE MTH.IDPL_ITEMS=:PL_ID; '+ wfLineEnding
            +' '+ wfLineEnding
            +'    V_QUANTITYINPACKING_PL_ID= (SELECT MTH.QUANTITYINPACKING FROM CATALOG_MATCHING MTH WHERE MTH.IDPL_ITEMS=:PL_ID); '+ wfLineEnding
            +'  end '+ wfLineEnding
            +' '+ wfLineEnding
            +'    INSERT INTO W_TMP_ANALIS_ALL_ANALOG '+ wfLineEnding
            +'    SELECT MTH.IDPL_ITEMS,MTH.QUANTITYINPACKING FROM CATALOG_MATCHING MTH WHERE MTH.IDCATALOG=:V_CTG_ID; '+ wfLineEnding
            +' '+ wfLineEnding
            +'  if (HIDE_POS_PL_ID) then '+ wfLineEnding
            +'  begin '+ wfLineEnding
            +'    DELETE FROM W_TMP_ANALIS_ALL_ANALOG WHERE IDPL_ITEMS=:PL_ID; '+ wfLineEnding
            +'  end '+ wfLineEnding
            +' '+ wfLineEnding
            +'      FOR '+ wfLineEnding
            +'        SELECT '+ wfLineEnding
            +'        PL.ID,  '+ wfLineEnding
            +'        PL.IDOWNER, '+ wfLineEnding
            +'        OWN.NAME OWNERNAME, '+ wfLineEnding
            +'        INT_QUANTITYINPACKINGTEXT(WTMP.QUANTITYINPACKING/:V_QUANTITYINPACKING_PL_ID, true) AS QUANTITYINPACKINGTEXT, '+ wfLineEnding
            +'        PL.VENDORCODE, '+ wfLineEnding
            +'        PL.LABEL, '+ wfLineEnding
            +'        (SELECT VSCOD FROM PL_GET_SCOD(PL.ID, true)) SCOD, '+ wfLineEnding
            +'        PL.NAME, '+ wfLineEnding
            +'        PL.UNIT, '+ wfLineEnding
            +'        PL.PRICECALC/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE, '+ wfLineEnding
            +'        PL.PRICECALC2/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE2, '+ wfLineEnding
            +'        PL.PRICECALC3/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE3, '+ wfLineEnding
            +'        PL.PRICECALC4/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE4, '+ wfLineEnding
            +'        PL.PRICECALC5/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE5, '+ wfLineEnding
            +'        PL.PRICECALC6/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE6, '+ wfLineEnding
            +'        PL.PRICECALC7/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE7, '+ wfLineEnding
            +'        PL.PRICECALC8/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE8, '+ wfLineEnding
            +'        PL.PRICECALC9/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE9, '+ wfLineEnding
            +'        PL.PRICECALC10/WTMP.QUANTITYINPACKING*:V_QUANTITYINPACKING_PL_ID PRICE10, '+ wfLineEnding
            +'        (PL.STOCK+PL.STOCK2+PL.STOCK3+PL.STOCK4+PL.STOCK5)*WTMP.QUANTITYINPACKING/:V_QUANTITYINPACKING_PL_ID STOCK, '+ wfLineEnding
            +'        PL.FTIMESTAMP, '+ wfLineEnding
            +'        PL.FCOLOR, '+ wfLineEnding
            +'        FMTS.STOCKONLYINFO AS STOCKONLYINFO, '+ wfLineEnding
            +'        (WTMP.QUANTITYINPACKING/:V_QUANTITYINPACKING_PL_ID) QUANTITYINPACKING '+ wfLineEnding
            +'        FROM PL_ITEMS PL '+ wfLineEnding
            +'        INNER JOIN OWNER OWN ON (OWN.ID=PL.IDOWNER) '+ wfLineEnding
            +'        INNER JOIN W_TMP_ANALIS_ALL_ANALOG WTMP ON WTMP.IDPL_ITEMS = PL.ID '+ wfLineEnding
            +'        LEFT JOIN FORMATS FMTS ON (FMTS.ID=PL.IDFORMATS) '+ wfLineEnding
            +' '+ wfLineEnding
            +'        INTO :ID, '+ wfLineEnding
            +'             :IDOWNER, '+ wfLineEnding
            +'             :OWNERNAME, '+ wfLineEnding
            +'             :QUANTITYINPACKINGTEXT, '+ wfLineEnding
            +'             :VENDORCODE, '+ wfLineEnding
            +'             :LABEL, '+ wfLineEnding
            +'             :SCOD, '+ wfLineEnding
            +'             :PLNAME, '+ wfLineEnding
            +'             :UNIT, '+ wfLineEnding
            +'             :PRICE, '+ wfLineEnding
            +'             :PRICE2, '+ wfLineEnding
            +'             :PRICE3, '+ wfLineEnding
            +'             :PRICE4, '+ wfLineEnding
            +'             :PRICE5, '+ wfLineEnding
            +'             :PRICE6, '+ wfLineEnding
            +'             :PRICE7, '+ wfLineEnding
            +'             :PRICE8, '+ wfLineEnding
            +'             :PRICE9, '+ wfLineEnding
            +'             :PRICE10, '+ wfLineEnding
            +'             :STOCK, '+ wfLineEnding
            +'             :FTIMESTAMP, '+ wfLineEnding
            +'             :FCOLOR, '+ wfLineEnding
            +'             :STOCKONLYINFO, '+ wfLineEnding
            +'             :QUANTITYINPACKING '+ wfLineEnding
            +' '+ wfLineEnding
            +'      DO '+ wfLineEnding
            +'      BEGIN '+ wfLineEnding
            +'        SUSPEND; '+ wfLineEnding
            +'      END '+ wfLineEnding
            +'END ',true,false);

          ADBase.SQLUpdate('create or alter procedure CATALOG_PL_ITEMS_PRICE ( '+ wfLineEnding
            +'     IDCATALOG bigint) '+ wfLineEnding
            +' returns ( '+ wfLineEnding
            +'     ID bigint, '+ wfLineEnding
            +'     PRICEPL numeric(15,10)) '+ wfLineEnding
            +' as '+ wfLineEnding
            +' BEGIN  '+ wfLineEnding
            +'    FOR  '+ wfLineEnding
            +'      SELECT  '+ wfLineEnding
            +'      PL.ID,  '+ wfLineEnding
            +'      (PL.PRICECALC/MTH.QUANTITYINPACKING) AS PRICEPL  '+ wfLineEnding
            +'      FROM "PL_ITEMS" PL   '+ wfLineEnding
            +'      INNER JOIN "CATALOG_MATCHING" MTH ON (PL.ID = MTH.IDPL_ITEMS AND MTH.IDCATALOG=:IDCATALOG)  '+ wfLineEnding
            +'      WHERE (PL.STOCK>0 OR PL.STOCK2>0 OR PL.STOCK3>0 OR PL.STOCK4>0 OR PL.STOCK5>0)  '+ wfLineEnding
            +'      INTO :ID,  '+ wfLineEnding
            +'           :PRICEPL  '+ wfLineEnding
            +'    DO  '+ wfLineEnding
            +'    BEGIN  '+ wfLineEnding
            +'      SUSPEND;  '+ wfLineEnding
            +'    END  '+ wfLineEnding
            +'  END ',true,false);

          ADBase.SQLUpdate('create or alter procedure CATALOG_PL_MIN_PRICE ( '+ wfLineEnding
            +'     IDCATALOG bigint) '+ wfLineEnding
            +' returns ( '+ wfLineEnding
            +'     IDFORMATS bigint, '+ wfLineEnding
            +'     PRICEPL numeric(15,10), '+ wfLineEnding
            +'     PRICEPL2 numeric(15,10), '+ wfLineEnding
            +'     PRICEPL3 numeric(15,10), '+ wfLineEnding
            +'     PRICEPL4 numeric(15,10), '+ wfLineEnding
            +'     PRICEPL5 numeric(15,10), '+ wfLineEnding
            +'     PRICEPL6 numeric(15,10), '+ wfLineEnding
            +'     PRICEPL7 numeric(15,10), '+ wfLineEnding
            +'     PRICEPL8 numeric(15,10), '+ wfLineEnding
            +'     PRICEPL9 numeric(15,10), '+ wfLineEnding
            +'     PRICEPL10 numeric(15,10)) '+ wfLineEnding
            +' as '+ wfLineEnding
            +' declare variable PLID bigint; '+ wfLineEnding
            +' begin   '+ wfLineEnding
            +'     '+ wfLineEnding
            +'     SELECT ID, PRICEPL FROM CATALOG_PL_ITEMS_PRICE(:IDCATALOG)  '+ wfLineEnding
            +'     WHERE PRICEPL=(SELECT MIN(PRICEPL) FROM CATALOG_PL_ITEMS_PRICE(:IDCATALOG) ROWS 1) ROWS 1  '+ wfLineEnding
            +'     INTO :PLID,:PRICEPL;  '+ wfLineEnding
            +'     '+ wfLineEnding
            +'     if (PLID IS NULL) then  '+ wfLineEnding
            +'     begin  '+ wfLineEnding
            +'            :PRICEPL = 0;  '+ wfLineEnding
            +'            :PRICEPL2 = 0;  '+ wfLineEnding
            +'            :PRICEPL3 = 0;  '+ wfLineEnding
            +'            :PRICEPL4 = 0;  '+ wfLineEnding
            +'            :PRICEPL5 = 0;  '+ wfLineEnding
            +'            :PRICEPL6 = 0;  '+ wfLineEnding
            +'            :PRICEPL7 = 0;  '+ wfLineEnding
            +'            :PRICEPL8 = 0;  '+ wfLineEnding
            +'            :PRICEPL9 = 0;  '+ wfLineEnding
            +'            :PRICEPL10 = 0;  '+ wfLineEnding
            +'            :IDFORMATS = 0;  '+ wfLineEnding
            +'     end else  '+ wfLineEnding
            +'     begin  '+ wfLineEnding
            +'       SELECT  '+ wfLineEnding
            +'       (PL.PRICECALC/MTH.QUANTITYINPACKING) AS PRICEPL,  '+ wfLineEnding
            +'       (PL.PRICECALC2/MTH.QUANTITYINPACKING) AS PRICEPL2,  '+ wfLineEnding
            +'       (PL.PRICECALC3/MTH.QUANTITYINPACKING) AS PRICEPL3,  '+ wfLineEnding
            +'       (PL.PRICECALC4/MTH.QUANTITYINPACKING) AS PRICEPL4,  '+ wfLineEnding
            +'       (PL.PRICECALC5/MTH.QUANTITYINPACKING) AS PRICEPL5,  '+ wfLineEnding
            +'       (PL.PRICECALC6/MTH.QUANTITYINPACKING) AS PRICEPL6,  '+ wfLineEnding
            +'       (PL.PRICECALC7/MTH.QUANTITYINPACKING) AS PRICEPL7,  '+ wfLineEnding
            +'       (PL.PRICECALC8/MTH.QUANTITYINPACKING) AS PRICEPL8,  '+ wfLineEnding
            +'       (PL.PRICECALC9/MTH.QUANTITYINPACKING) AS PRICEPL9,  '+ wfLineEnding
            +'       (PL.PRICECALC10/MTH.QUANTITYINPACKING) AS PRICEPL10, '+ wfLineEnding
            +'       PL.IDFORMATS  '+ wfLineEnding
            +'      FROM "PL_ITEMS" PL  '+ wfLineEnding
            +'       INNER JOIN "CATALOG_MATCHING" MTH ON (PL.ID = MTH.IDPL_ITEMS AND MTH.IDCATALOG=:IDCATALOG)  '+ wfLineEnding
            +'       WHERE PL.ID=:PLID  '+ wfLineEnding
            +'       INTO :PRICEPL,  '+ wfLineEnding
            +'            :PRICEPL2,  '+ wfLineEnding
            +'            :PRICEPL3,  '+ wfLineEnding
            +'            :PRICEPL4,  '+ wfLineEnding
            +'            :PRICEPL5,  '+ wfLineEnding
            +'            :PRICEPL6,  '+ wfLineEnding
            +'            :PRICEPL7,  '+ wfLineEnding
            +'            :PRICEPL8,  '+ wfLineEnding
            +'            :PRICEPL9,  '+ wfLineEnding
            +'            :PRICEPL10, '+ wfLineEnding
            +'            :IDFORMATS;  '+ wfLineEnding
            +'     end  '+ wfLineEnding
            +'     suspend;  '+ wfLineEnding
            +'   END ',true,false);

          ADBase.SQLUpdate('create or alter function INT_QUANTITYINPACKINGTEXT ( ' + wfLineEnding
            +'     QUANTITYINPACKING numeric(15,10), '+ wfLineEnding
            +'     HIDE_1 boolean) '+ wfLineEnding
            +' returns varchar(100) as ' + wfLineEnding
            +' begin     ' + wfLineEnding
            +'     if ((HIDE_1) and (:QUANTITYINPACKING=1.00))   ' + wfLineEnding
            +'     then     '+ wfLineEnding
            +'         return '''';     '+ wfLineEnding
            +'  '
            +'     if (:QUANTITYINPACKING<1.0)  ' + wfLineEnding
            +'     then   '+ wfLineEnding
            +'         begin '+ wfLineEnding
            +'             return ''1 к ''||CAST(INT_ROUND(1/:QUANTITYINPACKING,0) AS INTEGER);   '+ wfLineEnding
            +'         end  '+ wfLineEnding
            +'     else   '+ wfLineEnding
            +'         begin    '+ wfLineEnding
            +'             return CAST(:QUANTITYINPACKING AS INTEGER)||'' к 1'';   '+ wfLineEnding
            +'         end     '+ wfLineEnding
            +' end ',true,false);


          ADBase.SQLUpdate('create or alter procedure INVOCE_SEARCH_OUR_POS_IN_PRICES ( ' + wfLineEnding
            +'     OUR_IDOWNER bigint, ' + wfLineEnding
            +'     OUR_VENDORCODE varchar(300), ' + wfLineEnding
            +'     IGNORESTOCKPRICE boolean = false) ' + wfLineEnding
            +' returns ( ' + wfLineEnding
            +'     QUANTITYINPACKING numeric(15,10), ' + wfLineEnding
            +'     PL_IDOWNER bigint, ' + wfLineEnding
            +'     PL_ID bigint) ' + wfLineEnding
            +' as ' + wfLineEnding
            +' declare variable ID bigint; ' + wfLineEnding
            +' BEGIN  ' + wfLineEnding
            +'    FOR  ' + wfLineEnding
            +'      SELECT PL.ID FROM PL_ITEMS PL WHERE PL.IDOWNER=:OUR_IDOWNER AND PL.VENDORCODE=:OUR_VENDORCODE  ' + wfLineEnding
            +'      INTO :ID  ' + wfLineEnding
            +'    DO  ' + wfLineEnding
            +'    BEGIN  ' + wfLineEnding
            +'      if (:ID IS NOT NULL) then  ' + wfLineEnding
            +'      begin  ' + wfLineEnding
            +'          if (IGNORESTOCKPRICE) then  ' + wfLineEnding
            +'          begin  ' + wfLineEnding
            +'              SELECT AP.IDOWNER, AP.ID, AP.QUANTITYINPACKING  FROM ANALIS_SEL_ALL_ANALOG(:ID, true) AP ORDER BY AP.PRICE ASC ROWS 1 ' + wfLineEnding
            +'              INTO :PL_IDOWNER,  ' + wfLineEnding
            +'                   :PL_ID, ' + wfLineEnding
            +'                   :QUANTITYINPACKING; ' + wfLineEnding
            +'          end else  ' + wfLineEnding
            +'          begin  ' + wfLineEnding
            +'              SELECT AP.IDOWNER, AP.ID, AP.QUANTITYINPACKING  FROM ANALIS_SEL_ALL_ANALOG(:ID, true) AP WHERE AP.PRICE>0 AND AP.STOCK>0 ORDER BY AP.PRICE ASC ROWS 1 ' + wfLineEnding
            +'              INTO :PL_IDOWNER,  ' + wfLineEnding
            +'                   :PL_ID, ' + wfLineEnding
            +'                   :QUANTITYINPACKING; ' + wfLineEnding
            +'          end  ' + wfLineEnding
            +'      end  ' + wfLineEnding
            +'      SUSPEND;  ' + wfLineEnding
            +'    END  ' + wfLineEnding
            +'  END',true,false);


          ADBase.SQLUpdate('create or alter procedure INVOCE_ADD_ALANOG_FROM_PRICE ( ' + wfLineEnding
            +'     IGNORESTOCKPRICE boolean = false) ' + wfLineEnding
            +' as ' + wfLineEnding
            +' declare variable QUANTITYINPACKING numeric(15,10); ' + wfLineEnding
            +' declare variable IDOWNER bigint; ' + wfLineEnding
            +' declare variable IDINVOCE bigint; ' + wfLineEnding
            +' declare variable ID bigint; ' + wfLineEnding
            +' declare variable VENDORCODE varchar(300); ' + wfLineEnding
            +' declare variable QUANTITY numeric(15,2); ' + wfLineEnding
            +' declare variable PL_IDOWNER bigint; ' + wfLineEnding
            +' declare variable PL_ID bigint; ' + wfLineEnding
            +' BEGIN  ' + wfLineEnding
            +'  FOR  ' + wfLineEnding
            +'  SELECT ORD.ID, ORD.ORDOWNER, ORD.ORDVENDORCODE, ORD.ORDQUANTITY FROM W_TMP_ORDERS_IMPORT ORD  ' + wfLineEnding
            +'  INTO :ID,  ' + wfLineEnding
            +'       :IDOWNER,  ' + wfLineEnding
            +'       :VENDORCODE,  ' + wfLineEnding
            +'       :QUANTITY  ' + wfLineEnding
            +'  DO  ' + wfLineEnding
            +'  BEGIN  ' + wfLineEnding
            +' :PL_IDOWNER = NULL;  ' + wfLineEnding
            +' :PL_ID = NULL;  ' + wfLineEnding
            +' :QUANTITYINPACKING = NULL;  ' + wfLineEnding
            +'      SELECT INVS.PL_IDOWNER, INVS.PL_ID, INVS.QUANTITYINPACKING FROM INVOCE_SEARCH_OUR_POS_IN_PRICES(:IDOWNER, :VENDORCODE, :IGNORESTOCKPRICE) INVS ' + wfLineEnding
            +'      INTO :PL_IDOWNER,  ' + wfLineEnding
            +'           :PL_ID, ' + wfLineEnding
            +'           :QUANTITYINPACKING; ' + wfLineEnding
            +'            ' + wfLineEnding
            +'      if (PL_ID IS NOT NULL) then  ' + wfLineEnding
            +'      begin  ' + wfLineEnding
            +'         INSERT INTO INVOCES (IDOWNER, IDPL_ITEMS, QUANTITY, REMARK) VALUES (:PL_IDOWNER, :PL_ID, INT_ROUND(:QUANTITY/:QUANTITYINPACKING,0), ''auto'') RETURNING ID INTO :IDINVOCE; ' + wfLineEnding
            +'          ' + wfLineEnding
            +'         UPDATE W_TMP_ORDERS_IMPORT ORD SET ORD.FPASSED=1, ORD.MTHID=:IDINVOCE WHERE ORD.ID=:ID;  ' + wfLineEnding
            +'      end  ' + wfLineEnding
            +'  END  ' + wfLineEnding
            +'        ' + wfLineEnding
            +'  END ',true,false);


          ADBase.SQLUpdate('create or alter procedure CATALOG_PL_ITEMS_PRICE ( '+ wfLineEnding
            +'     IDCATALOG bigint) '+ wfLineEnding
            +' returns ( '+ wfLineEnding
            +'     ID bigint, '+ wfLineEnding
            +'     PRICEPL numeric(15,10)) '+ wfLineEnding
            +' as '+ wfLineEnding
            +' BEGIN   '+ wfLineEnding
            +'     FOR   '+ wfLineEnding
            +'       SELECT   '+ wfLineEnding
            +'       PL.ID,   '+ wfLineEnding
            +'       (PL.PRICECALC/MTH.QUANTITYINPACKING) AS PRICEPL   '+ wfLineEnding
            +'       FROM "PL_ITEMS" PL    '+ wfLineEnding
            +'       INNER JOIN "CATALOG_MATCHING" MTH ON (PL.ID = MTH.IDPL_ITEMS AND MTH.IDCATALOG=:IDCATALOG)   '+ wfLineEnding
            +'       WHERE (PL.STOCK>0 OR PL.STOCK2>0 OR PL.STOCK3>0 OR PL.STOCK4>0 OR PL.STOCK5>0) AND (PL.PRICECALC>0) '+ wfLineEnding
            +'       INTO :ID,   '+ wfLineEnding
            +'            :PRICEPL   '+ wfLineEnding
            +'     DO   '+ wfLineEnding
            +'     BEGIN   '+ wfLineEnding
            +'       SUSPEND;   '+ wfLineEnding
            +'     END   '+ wfLineEnding
            +'   END ',true,false);

          FMWait.SetBar(3);

          wLog('Main','Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion',aNewVersion);

          FMWait.Free;
        except
          wLog('Main','Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;

      aNewVersion:= '0.0.3.47';
      if  CheckLessVersion(ADBase.ReadSettingByName('dbVersion'),aNewVersion) then
      begin

        FmWait:= TFmWait.Create(self);

        FmWait.Show;
        //FmWait.Height:=250;
        //FmWait.Width:=540;
        //FmWait.mStatus.Alignment:=taLeftJustify;
        FMWait.InitBar(2,0);
        wLog('Main','Обновление метаданных БД... до версии '+aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus('Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate('create or alter procedure PL_GROUP_TO_CATALOG ( '+ wfLineEnding
          +'    IDOWNER_CATALOG bigint, '+ wfLineEnding
          +'    ID_CATALOG bigint, '+ wfLineEnding
          +'    ID_NODE bigint, '+ wfLineEnding
          +'    PL_GROUP_NAME varchar(255), '+ wfLineEnding
          +'    CTG_PRICE numeric(15,2), '+ wfLineEnding
          +'    CTG_PN numeric(15,2), '+ wfLineEnding
          +'    CTG_PM numeric(15,2), '+ wfLineEnding
          +'    CTG_PD numeric(15,2), '+ wfLineEnding
          +'    CTG_PC numeric(15,2), '+ wfLineEnding
          +'    CTG_PK numeric(15,2)) '+ wfLineEnding
          +'as '+ wfLineEnding
          +'declare variable INBASE bigint not null; '+ wfLineEnding
          +'declare variable CTG_ID bigint; '+ wfLineEnding
          +'declare variable VSCODS varchar(1024); '+ wfLineEnding
          +'declare variable PL_ID_KID bigint; '+ wfLineEnding
          +'declare variable PLI_ID_SELECTED bigint; '+ wfLineEnding
          +'declare variable PL_NAME varchar(255); '+ wfLineEnding
          +'declare variable PL_FTIMESTAMP timestamp; '+ wfLineEnding
          +'declare variable PLI_ID bigint; '+ wfLineEnding
          +'declare variable PLI_IDOWNER bigint; '+ wfLineEnding
          +'declare variable PLI_NAME varchar(500); '+ wfLineEnding
          +'declare variable PLI_FTIMESTAMP timestamp; '+ wfLineEnding
          +'declare variable PLI_UNIT varchar(15); '+ wfLineEnding
          +'declare variable PLI_VENDORCODE varchar(300); '+ wfLineEnding
          +'declare variable PLI_LABEL varchar(255); '+ wfLineEnding
          +'declare variable PLI_SCOD varchar(300); '+ wfLineEnding
          +'declare variable PLI_REMARK varchar(3000); '+ wfLineEnding
          +'declare variable PLI_FURL varchar(1000); '+ wfLineEnding
          +'declare variable PLI_FURLPICTURE varchar(1000); '+ wfLineEnding
          +'declare variable PLI_FCOLOR integer; '+ wfLineEnding
          +'declare variable NOW_FTIMESTAMP timestamp; '+ wfLineEnding
          +'begin '+ wfLineEnding

          +'end ',true,false);

          FMWait.SetBar(2);

          ADBase.SQLUpdate('create or alter procedure PL_GROUP_TO_CATALOG ( '+ wfLineEnding
          +'    IDOWNER_CATALOG bigint, '+ wfLineEnding
          +'    ID_CATALOG bigint, '+ wfLineEnding
          +'    ID_NODE bigint, '+ wfLineEnding
          +'    PL_GROUP_NAME varchar(255), '+ wfLineEnding
          +'    CTG_PRICE numeric(15,2), '+ wfLineEnding
          +'    CTG_PN numeric(15,2), '+ wfLineEnding
          +'    CTG_PM numeric(15,2), '+ wfLineEnding
          +'    CTG_PD numeric(15,2), '+ wfLineEnding
          +'    CTG_PC numeric(15,2), '+ wfLineEnding
          +'    CTG_PK numeric(15,2)) '+ wfLineEnding
          +'as '+ wfLineEnding
          +'declare variable INBASE bigint not null; '+ wfLineEnding
          +'declare variable CTG_ID bigint; '+ wfLineEnding
          +'declare variable VSCODS varchar(1024); '+ wfLineEnding
          +'declare variable PL_ID_KID bigint; '+ wfLineEnding
          +'declare variable PLI_ID_SELECTED bigint; '+ wfLineEnding
          +'declare variable PL_NAME varchar(255); '+ wfLineEnding
          +'declare variable PL_FTIMESTAMP timestamp; '+ wfLineEnding
          +'declare variable PLI_ID bigint; '+ wfLineEnding
          +'declare variable PLI_IDOWNER bigint; '+ wfLineEnding
          +'declare variable PLI_NAME varchar(500); '+ wfLineEnding
          +'declare variable PLI_FTIMESTAMP timestamp; '+ wfLineEnding
          +'declare variable PLI_UNIT varchar(15); '+ wfLineEnding
          +'declare variable PLI_VENDORCODE varchar(300); '+ wfLineEnding
          +'declare variable PLI_LABEL varchar(255); '+ wfLineEnding
          +'declare variable PLI_SCOD varchar(300); '+ wfLineEnding
          +'declare variable PLI_REMARK varchar(3000); '+ wfLineEnding
          +'declare variable PLI_FURL varchar(1000); '+ wfLineEnding
          +'declare variable PLI_FURLPICTURE varchar(1000); '+ wfLineEnding
          +'declare variable PLI_FCOLOR integer; '+ wfLineEnding
          +'declare variable NOW_FTIMESTAMP timestamp; '+ wfLineEnding
          +'begin '+ wfLineEnding
          +' '+ wfLineEnding
          +'             if (:ID_NODE IS NOT NULL) then '+ wfLineEnding
          +'             begin    /* если ID группы корректно, то добавляем (обновляем) соответствующую группу в каталоге */'+ wfLineEnding
          +'                   UPDATE OR INSERT INTO CATALOG_GROUP (IDPARENT,IDOWNER,NAME,FTIMESTAMP) VALUES (:ID_CATALOG,:IDOWNER_CATALOG,:PL_GROUP_NAME,:PL_FTIMESTAMP) MATCHING (IDPARENT,IDOWNER,NAME) RETURNING ID INTO :INBASE; '+ wfLineEnding
          +'                   NOW_FTIMESTAMP = CURRENT_TIMESTAMP; '+ wfLineEnding
          +'                  /* ВЫБИРАЕМ ТОВАРЫ, принадлежащие этой группе*/ '+ wfLineEnding
          +'                  FOR SELECT ID,IDOWNER,NAME,UNIT,VENDORCODE,LABEL,REMARK,FURL,FURLPICTURE,FCOLOR,FTIMESTAMP '+ wfLineEnding
          +'                     FROM PL_ITEMS '+ wfLineEnding
          +'                     WHERE IDPL_GROUP = :ID_NODE '+ wfLineEnding
          +'                     INTO :PLI_ID,:PLI_IDOWNER,:PLI_NAME,:PLI_UNIT,:PLI_VENDORCODE,:PLI_LABEL,:PLI_REMARK,:PLI_FURL,:PLI_FURLPICTURE,:PLI_FCOLOR,:PLI_FTIMESTAMP '+ wfLineEnding
          +' '+ wfLineEnding
          +'                    DO BEGIN '+ wfLineEnding
          +' '+ wfLineEnding
          +'                         if (:PLI_ID IS NOT NULL) then  /* если товар существует, то */'+ wfLineEnding
          +'                         begin '+ wfLineEnding
          +'                             /* смотрим - есть ли к нему соответствия */'+ wfLineEnding
          +'                             PLI_ID_SELECTED = (SELECT ID FROM CATALOG_MATCHING WHERE IDPL_ITEMS=:PLI_ID); '+ wfLineEnding
          +' '+ wfLineEnding
          +'                            if (:PLI_ID_SELECTED IS NULL) then '+ wfLineEnding
          +'                            begin '+ wfLineEnding
          +'                                /* если соответствия нет, то значит добавляем товар впервые, следовательно */'+ wfLineEnding
          +'                                /* добавляем позицию в каталог */'+ wfLineEnding
          +'                                  INSERT INTO CATALOG ( '+ wfLineEnding
          +'                                  IDCTG_GROUP,IDOWNER,NAME,UNIT,VENDORCODE,PRICE,PN,PM,PD,PC,PK,LABEL,REMARK,FURL,FURLPICTURE,FCOLOR,FTIMESTAMP '+ wfLineEnding
          +'                                  ) '+ wfLineEnding
          +'                                    VALUES ( '+ wfLineEnding
          +'                                    :INBASE,:IDOWNER_CATALOG,:PLI_NAME,:PLI_UNIT,'''',:CTG_PRICE,:CTG_PN,:CTG_PM,:CTG_PD,:CTG_PC,:CTG_PK,:PLI_LABEL,:PLI_REMARK,:PLI_FURL,:PLI_FURLPICTURE,:PLI_FCOLOR,:NOW_FTIMESTAMP '+ wfLineEnding
          +'                                    ) '+ wfLineEnding
          +'                                    RETURNING ID INTO :CTG_ID; '+ wfLineEnding
          +' '+ wfLineEnding
          +'                                    VSCODS = (SELECT VSCOD FROM PL_GET_SCOD(:PLI_ID,true)); '+ wfLineEnding
          +'                                    EXECUTE PROCEDURE CTG_SET_SCOD(:IDOWNER_CATALOG,:CTG_ID,:VSCODS,'',''); '+ wfLineEnding
          +' '+ wfLineEnding
          +'                                /* добавляем позицию в таблицу соответствий */'+ wfLineEnding
          +'                                 UPDATE OR INSERT INTO "CATALOG_MATCHING" (IDOWNER, IDCATALOG, IDPL_ITEMS, QUANTITYINPACKING, IDUSER, FTIMESTAMP) '+ wfLineEnding
          +'                                     VALUES (:PLI_IDOWNER,:CTG_ID,:PLI_ID,1,1,:NOW_FTIMESTAMP) '+ wfLineEnding
          +'                                     MATCHING (IDOWNER,IDCATALOG,IDPL_ITEMS); '+ wfLineEnding
          +'                            end '+ wfLineEnding
          +'                         end '+ wfLineEnding
          +'                    END '+ wfLineEnding
          +' '+ wfLineEnding
          +'             end '+ wfLineEnding
          +' '+ wfLineEnding
          +'  /* ВЫБИРАЕМ ГРУППЫ, перебор подгрупп */ '+ wfLineEnding
          +'  FOR SELECT ID,NAME,FTIMESTAMP '+ wfLineEnding
          +'     FROM PL_GROUP '+ wfLineEnding
          +'     WHERE IDPARENT = :ID_NODE '+ wfLineEnding
          +'     INTO :PL_ID_KID,:PL_NAME,:PL_FTIMESTAMP '+ wfLineEnding
          +' '+ wfLineEnding
          +'    DO BEGIN '+ wfLineEnding
          +'                /* ЗАПУСКАЕМ ПРОЦЕДУРУ ДЛЯ ВСЕХ ВЛОЖЕННЫХ ГРУПП ПООЧЕРЕДИ (РЕКУРСИЯ) */ '+ wfLineEnding
          +'                EXECUTE PROCEDURE PL_GROUP_TO_CATALOG(IDOWNER_CATALOG,INBASE,PL_ID_KID,PL_NAME,:CTG_PRICE,:CTG_PN,:CTG_PM,:CTG_PD,:CTG_PC,:CTG_PK); '+ wfLineEnding
          +'    END '+ wfLineEnding
          +'end ',true,false);

          wLog('Main','Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion',aNewVersion);

          FMWait.Free;
        except
          wLog('Main','Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;

      aNewVersion:= '0.0.3.50';
      if  CheckLessVersion(ADBase.ReadSettingByName('dbVersion'),aNewVersion) then
      begin

        FmWait:= TFmWait.Create(self);

        FmWait.Show;
        //FmWait.Height:=250;
        //FmWait.Width:=540;
        //FmWait.mStatus.Alignment:=taLeftJustify;
        FMWait.InitBar(2,0);
        wLog('Main','Обновление метаданных БД... до версии '+aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus('Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate('ALTER TABLE CATALOG ADD FTIMESTAMPCREATED TIMESTAMP;',true,false);
          ADBase.SQLUpdate('ALTER TABLE CATALOG ALTER COLUMN FTIMESTAMPCREATED SET DEFAULT CURRENT_TIMESTAMP;',true,false);
          FMWait.SetBar(2);
          ADBase.SQLUpdate('update CATALOG SET FTIMESTAMPCREATED=FTIMESTAMP;',true,false);

          wLog('Main','Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion',aNewVersion);

          FMWait.Free;
        except
          wLog('Main','Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;

      aNewVersion:= '0.0.3.82';
      if  CheckLessVersion(ADBase.ReadSettingByName('dbVersion'),aNewVersion) then
      begin

        FmWait:= TFmWait.Create(self);

        FmWait.Show;
        //FmWait.Height:=250;
        //FmWait.Width:=540;
        //FmWait.mStatus.Alignment:=taLeftJustify;
        FMWait.InitBar(2,0);
        wLog('Main','Обновление метаданных БД... до версии '+aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus('Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate('CREATE GLOBAL TEMPORARY TABLE W_TMP_PL_VERSIONS ( '+ wfLineEnding
            +' ID           BIGINT NOT NULL, '+ wfLineEnding
            +' IDPL_ITEMS   BIGINT NOT NULL, '+ wfLineEnding
            +' IDOWNER      BIGINT NOT NULL, '+ wfLineEnding
            +' IDFORMATS    BIGINT, '+ wfLineEnding
            +' FTIMESTAMP   TIMESTAMP, '+ wfLineEnding
            +' PRICE        DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC    DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICE2       DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC2   DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICE3       DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC3   DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' STOCK        DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' STOCK2       DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' STOCK3       DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' TRANSIT      VARCHAR(120), '+ wfLineEnding
            +' PRICE4       DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC4   DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICE5       DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC5   DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' STOCK4       DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' STOCK5       DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICE6       DECIMAL(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC6   NUMERIC(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICE7       NUMERIC(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC7   NUMERIC(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICE8       NUMERIC(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC8   NUMERIC(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICE9       NUMERIC(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC9   NUMERIC(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICE10      NUMERIC(15,2) DEFAULT 0, '+ wfLineEnding
            +' PRICECALC10  NUMERIC(15,2) DEFAULT 0 '+ wfLineEnding
            +' ) ON COMMIT PRESERVE ROWS;',true,false);

          FMWait.SetBar(2);

          ADBase.SQLUpdate('CREATE INDEX W_TMP_PL_VERSIONS_IDPL ON W_TMP_PL_VERSIONS (IDPL_ITEMS);',true,false);
          ADBase.SQLUpdate('CREATE INDEX W_TMP_PL_VERSIONS_ID_PRCALC ON W_TMP_PL_VERSIONS (IDPL_ITEMS);',true,false);

          wLog('Main','Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion',aNewVersion);

          FMWait.Free;
        except
          wLog('Main','Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;

      aNewVersion:= '0.0.3.94';
      if  CheckLessVersion(ADBase.ReadSettingByName('dbVersion'),aNewVersion) then
      begin

        FmWait:= TFmWait.Create(self);

        FmWait.Show;
        FMWait.InitBar(2,0);
        wLog('Main','Обновление метаданных БД... до версии '+aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus('Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate('ALTER TABLE W_TMP_ORDERS_GET_MTH'+ wfLE
            +'ADD VENDORPRICE NUMERIC(15,2)',true,false);

          FMWait.SetBar(2);

          ADBase.SQLUpdate('CREATE OR ALTER procedure WTOI_GET_ANALOGS ('+ wfLE
              +'    WTOIIDOWNER bigint)'+ wfLE
              +'as'+ wfLE
              +'declare variable ORDVENDORCODE varchar(300);'+ wfLE
              +'declare variable ORDLABEL varchar(255);'+ wfLE
              +'declare variable ORDSCOD varchar(120);'+ wfLE
              +'declare variable ORDNAME varchar(500);'+ wfLE
              +'declare variable ORDUNIT varchar(60);'+ wfLE
              +'declare variable ORDQUANTITY numeric(18,2);'+ wfLE
              +'declare variable ORDPRICE numeric(18,2);'+ wfLE
              +'declare variable ORDSUM numeric(18,2);'+ wfLE
              +'declare variable ORDREMARK varchar(4000);'+ wfLE
              +'declare variable ORDID bigint;'+ wfLE
              +'declare variable IDOWNER bigint;'+ wfLE
              +'declare variable OWNERNAME varchar(150);'+ wfLE
              +'declare variable VENDORCODE varchar(300);'+ wfLE
              +'declare variable SCOD varchar(1024);'+ wfLE
              +'declare variable PLNAME varchar(500);'+ wfLE
              +'declare variable PRICE numeric(15,2);'+ wfLE
              +'declare variable PRICE2 numeric(15,2);'+ wfLE
              +'declare variable PRICE3 numeric(15,2);'+ wfLE
              +'declare variable PRICE4 numeric(15,2);'+ wfLE
              +'declare variable PRICE5 numeric(15,2);'+ wfLE
              +'declare variable PRICE6 numeric(15,2);'+ wfLE
              +'declare variable PRICE7 numeric(15,2);'+ wfLE
              +'declare variable PRICE8 numeric(15,2);'+ wfLE
              +'declare variable PRICE9 numeric(15,2);'+ wfLE
              +'declare variable PRICE10 numeric(15,2);'+ wfLE
              +'declare variable STOCK bigint;'+ wfLE
              +'declare variable QUANTITYINPACKINGTEXT varchar(100);'+ wfLE
              +'declare variable ID bigint;'+ wfLE
              +'declare variable FTIMESTAMP timestamp;'+ wfLE
              +'declare variable UNIT varchar(15);'+ wfLE
              +'declare variable STOCKONLYINFO smallint;'+ wfLE
              +'declare variable FCOLOR smallint;'+ wfLE
              +'declare variable QUANTITYINPACKING numeric(15,2);'+ wfLE
              +'declare variable LABEL varchar(255);'+ wfLE
              +'declare variable FLAGFINDEDMATCHING boolean;'+ wfLE
              +'BEGIN '+ wfLE
              +'   FOR '+ wfLE
              +'     SELECT '+ wfLE
              +'     WTOI.ID, '+ wfLE
              +'     WTOI.ORDVENDORCODE, '+ wfLE
              +'     WTOI.ORDLABEL, '+ wfLE
              +'     WTOI.ORDSCOD, '+ wfLE
              +'     WTOI.ORDNAME, '+ wfLE
              +'     WTOI.ORDUNIT, '+ wfLE
              +'     WTOI.ORDQUANTITY, '+ wfLE
              +'     WTOI.ORDPRICE, '+ wfLE
              +'     WTOI.ORDSUM, '+ wfLE
              +'     WTOI.ORDREMARK '+ wfLE
              +'     FROM W_TMP_ORDERS_IMPORT WTOI '+ wfLE
              +'     INTO :ORDID, '+ wfLE
              +'          :ORDVENDORCODE, '+ wfLE
              +'          :ORDLABEL, '+ wfLE
              +'          :ORDSCOD, '+ wfLE
              +'          :ORDNAME, '+ wfLE
              +'          :ORDUNIT, '+ wfLE
              +'          :ORDQUANTITY, '+ wfLE
              +'          :ORDPRICE, '+ wfLE
              +'          :ORDSUM, '+ wfLE
              +'          :ORDREMARK'+ wfLE
              +'   DO '+ wfLE
              +'   BEGIN '+ wfLE
              +'     FLAGFINDEDMATCHING = false; '+ wfLE
              +'     FOR SELECT '+ wfLE
              +'             ID,  '+ wfLE
              +'             IDOWNER, '+ wfLE
              +'             OWNERNAME,  '+ wfLE
              +'             QUANTITYINPACKINGTEXT,  '+ wfLE
              +'             VENDORCODE,  '+ wfLE
              +'             LABEL,  '+ wfLE
              +'             SCOD,  '+ wfLE
              +'             PLNAME,  '+ wfLE
              +'             UNIT,  '+ wfLE
              +'             PRICE,  '+ wfLE
              +'             PRICE2,  '+ wfLE
              +'             PRICE3,  '+ wfLE
              +'             PRICE4,  '+ wfLE
              +'             PRICE5,  '+ wfLE
              +'             PRICE6,  '+ wfLE
              +'             PRICE7,  '+ wfLE
              +'             PRICE8,  '+ wfLE
              +'             PRICE9,  '+ wfLE
              +'             PRICE10,  '+ wfLE
              +'             STOCK,  '+ wfLE
              +'             FTIMESTAMP,  '+ wfLE
              +'             FCOLOR,  '+ wfLE
              +'             STOCKONLYINFO,  '+ wfLE
              +'             QUANTITYINPACKING '+ wfLE
              +'          FROM WTOI_GET_MATCHING(:ORDID) WHERE IDOWNER =:WTOIIDOWNER '+ wfLE
              +'        INTO :ID, '+ wfLE
              +'             :IDOWNER, '+ wfLE
              +'             :OWNERNAME,  '+ wfLE
              +'             :QUANTITYINPACKINGTEXT,  '+ wfLE
              +'             :VENDORCODE,  '+ wfLE
              +'             :LABEL,  '+ wfLE
              +'             :SCOD,  '+ wfLE
              +'             :PLNAME,  '+ wfLE
              +'             :UNIT,  '+ wfLE
              +'             :PRICE,  '+ wfLE
              +'             :PRICE2,  '+ wfLE
              +'             :PRICE3,  '+ wfLE
              +'             :PRICE4,  '+ wfLE
              +'             :PRICE5,  '+ wfLE
              +'             :PRICE6,  '+ wfLE
              +'             :PRICE7,  '+ wfLE
              +'             :PRICE8,  '+ wfLE
              +'             :PRICE9,  '+ wfLE
              +'             :PRICE10,  '+ wfLE
              +'             :STOCK,  '+ wfLE
              +'             :FTIMESTAMP,  '+ wfLE
              +'             :FCOLOR,  '+ wfLE
              +'             :STOCKONLYINFO,  '+ wfLE
              +'             :QUANTITYINPACKING '+ wfLE
              +'     DO '+ wfLE
              +'     BEGIN '+ wfLE
              +'     FLAGFINDEDMATCHING = true; '+ wfLE
              +'     INSERT INTO W_TMP_ORDERS_GET_MTH '+ wfLE
              +'        (ORDVENDORCODE, '+ wfLE
              +'         ORDLABEL, '+ wfLE
              +'         ORDSCOD, '+ wfLE
              +'         ORDNAME, '+ wfLE
              +'         ORDUNIT, '+ wfLE
              +'         ORDQUANTITY, '+ wfLE
              +'         ORDREMARK, '+ wfLE
              +'         VENDORCODE,'+ wfLE
              +'         VENDORPRICE,'+ wfLE
              +'         OWNERNAME, '+ wfLE
              +'         OWNERSEARCH) VALUES '+ wfLE
              +'         (:ORDVENDORCODE, '+ wfLE
              +'          :ORDLABEL, '+ wfLE
              +'          :ORDSCOD, '+ wfLE
              +'          :ORDNAME, '+ wfLE
              +'          :ORDUNIT, '+ wfLE
              +'          (:ORDQUANTITY/:QUANTITYINPACKING), '+ wfLE
              +'          :ORDREMARK, '+ wfLE
              +'          :VENDORCODE,'+ wfLE
              +'          :PRICE,'+ wfLE
              +'          :OWNERNAME, '+ wfLE
              +'          :WTOIIDOWNER '+ wfLE
              +'         ); '+ wfLE
              +'     END '+ wfLE
              +'  '+ wfLE
              +'     if (NOT FLAGFINDEDMATCHING) then '+ wfLE
              +'     INSERT INTO W_TMP_ORDERS_GET_MTH '+ wfLE
              +'        (ORDVENDORCODE, '+ wfLE
              +'         ORDLABEL, '+ wfLE
              +'         ORDSCOD, '+ wfLE
              +'         ORDNAME, '+ wfLE
              +'         ORDUNIT, '+ wfLE
              +'         ORDQUANTITY, '+ wfLE
              +'         ORDREMARK, '+ wfLE
              +'         VENDORCODE,'+ wfLE
              +'         VENDORPRICE,'+ wfLE
              +'         OWNERNAME, '+ wfLE
              +'         OWNERSEARCH) VALUES '+ wfLE
              +'         (:ORDVENDORCODE, '+ wfLE
              +'          :ORDLABEL, '+ wfLE
              +'          :ORDSCOD, '+ wfLE
              +'          :ORDNAME, '+ wfLE
              +'          :ORDUNIT, '+ wfLE
              +'          :ORDQUANTITY, '+ wfLE
              +'          :ORDREMARK, '+ wfLE
              +'          '''','+ wfLE
              +'          0,'+ wfLE
              +'          '''', '+ wfLE
              +'          :WTOIIDOWNER '+ wfLE
              +'         ); '+ wfLE
              +'  '+ wfLE
              +'   END '+ wfLE
              +' END ',true,false);

          wLog('Main','Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion',aNewVersion);

          FMWait.Free;
        except
          wLog('Main','Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;

      aNewVersion:= '0.0.3.104';
      if  CheckLessVersion(ADBase.ReadSettingByName('dbVersion'),aNewVersion) then
      begin

        FmWait:= TFmWait.Create(self);

        FmWait.Show;
        FMWait.InitBar(2,0);
        wLog('Main','Обновление метаданных БД... до версии '+aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus('Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate(' create or alter procedure CATALOG_PL_MIN_PRICE ( '
              +'    IDCATALOG bigint) '+ wfLE
              +' returns ( '+ wfLE
              +'     IDFORMATS bigint, '+ wfLE
              +'     PRICEPL numeric(15,10), '+ wfLE
              +'     PRICEPL2 numeric(15,10), '+ wfLE
              +'     PRICEPL3 numeric(15,10), '+ wfLE
              +'     PRICEPL4 numeric(15,10), '+ wfLE
              +'     PRICEPL5 numeric(15,10), '+ wfLE
              +'     PRICEPL6 numeric(15,10), '+ wfLE
              +'     PRICEPL7 numeric(15,10), '+ wfLE
              +'     PRICEPL8 numeric(15,10), '+ wfLE
              +'     PRICEPL9 numeric(15,10), '+ wfLE
              +'     PRICEPL10 numeric(15,10), '+ wfLE
              +'     PDATE timestamp) '+ wfLE
              +' as '+ wfLE
              +' declare variable PLID bigint; '+ wfLE
              +' begin    '+ wfLE
              +'       '+ wfLE
              +'      SELECT ID, PRICEPL FROM CATALOG_PL_ITEMS_PRICE(:IDCATALOG)   '+ wfLE
              +'      WHERE PRICEPL=(SELECT MIN(PRICEPL) FROM CATALOG_PL_ITEMS_PRICE(:IDCATALOG) ROWS 1) ROWS 1   '+ wfLE
              +'      INTO :PLID,:PRICEPL;   '+ wfLE
              +'       '+ wfLE
              +'      if (PLID IS NULL) then   '+ wfLE
              +'      begin   '+ wfLE
              +'             :PRICEPL = 0;   '+ wfLE
              +'             :PRICEPL2 = 0;   '+ wfLE
              +'             :PRICEPL3 = 0;   '+ wfLE
              +'             :PRICEPL4 = 0;   '+ wfLE
              +'             :PRICEPL5 = 0;   '+ wfLE
              +'             :PRICEPL6 = 0;   '+ wfLE
              +'             :PRICEPL7 = 0;   '+ wfLE
              +'             :PRICEPL8 = 0;   '+ wfLE
              +'             :PRICEPL9 = 0;   '+ wfLE
              +'             :PRICEPL10 = 0;   '+ wfLE
              +'             :IDFORMATS = 0;   '+ wfLE
              +'             :PDATE = ''01.01.0001 00:00''; '+ wfLE
              +'      end else   '+ wfLE
              +'      begin   '+ wfLE
              +'        SELECT   '+ wfLE
              +'        (PL.PRICECALC/MTH.QUANTITYINPACKING) AS PRICEPL,   '+ wfLE
              +'        (PL.PRICECALC2/MTH.QUANTITYINPACKING) AS PRICEPL2,   '+ wfLE
              +'        (PL.PRICECALC3/MTH.QUANTITYINPACKING) AS PRICEPL3,   '+ wfLE
              +'        (PL.PRICECALC4/MTH.QUANTITYINPACKING) AS PRICEPL4,   '+ wfLE
              +'        (PL.PRICECALC5/MTH.QUANTITYINPACKING) AS PRICEPL5,   '+ wfLE
              +'        (PL.PRICECALC6/MTH.QUANTITYINPACKING) AS PRICEPL6,   '+ wfLE
              +'        (PL.PRICECALC7/MTH.QUANTITYINPACKING) AS PRICEPL7,   '+ wfLE
              +'        (PL.PRICECALC8/MTH.QUANTITYINPACKING) AS PRICEPL8,   '+ wfLE
              +'        (PL.PRICECALC9/MTH.QUANTITYINPACKING) AS PRICEPL9,   '+ wfLE
              +'        (PL.PRICECALC10/MTH.QUANTITYINPACKING) AS PRICEPL10,  '+ wfLE
              +'        PL.IDFORMATS, '+ wfLE
              +'        PL.FTIMESTAMP '+ wfLE
              +'       FROM "PL_ITEMS" PL   '+ wfLE
              +'        INNER JOIN "CATALOG_MATCHING" MTH ON (PL.ID = MTH.IDPL_ITEMS AND MTH.IDCATALOG=:IDCATALOG)   '+ wfLE
              +'        WHERE PL.ID=:PLID   '+ wfLE
              +'        INTO :PRICEPL,   '+ wfLE
              +'             :PRICEPL2,   '+ wfLE
              +'             :PRICEPL3,   '+ wfLE
              +'             :PRICEPL4,   '+ wfLE
              +'             :PRICEPL5,   '+ wfLE
              +'             :PRICEPL6,   '+ wfLE
              +'             :PRICEPL7,   '+ wfLE
              +'             :PRICEPL8,   '+ wfLE
              +'             :PRICEPL9,   '+ wfLE
              +'             :PRICEPL10,  '+ wfLE
              +'             :IDFORMATS, '+ wfLE
              +'             :PDATE; '+ wfLE
              +'      end   '+ wfLE
              +'      suspend;   '+ wfLE
              +'    END ',true,false);

          wLog('Main','Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion',aNewVersion);

          FMWait.Free;
        except
          wLog('Main','Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;

      aNewVersion:= '0.0.3.116';
      if  CheckLessVersion(ADBase.ReadSettingByName('dbVersion'),aNewVersion) then
      begin

        FmWait:= TFmWait.Create(self);

        FmWait.Show;
        FMWait.InitBar(2,0);
        wLog('Main','Обновление метаданных БД... до версии '+aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus('Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate('CREATE OR ALTER procedure CATALOG_PL_ITEMS_PRICE (  '+ wfLE
            +'     IDCATALOG bigint)  '+ wfLE
            +' returns (  '+ wfLE
            +'     ID bigint, '+ wfLE
            +'     PRICEPL numeric(15,10))  '+ wfLE
            +' as '+ wfLE
            +' BEGIN    '+ wfLE
            +'      FOR     '+ wfLE
            +'        SELECT    '+ wfLE
            +'        PL.ID,    '+ wfLE
            +'        (PL.PRICECALC/MTH.QUANTITYINPACKING) AS PRICEPL     '+ wfLE
            +'        FROM "PL_ITEMS" PL      '+ wfLE
            +'        INNER JOIN "CATALOG_MATCHING" MTH ON (PL.ID = MTH.IDPL_ITEMS AND MTH.IDCATALOG=:IDCATALOG)  '+ wfLE
            +'        LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PL.IDFORMATS) '+ wfLE
            +'        WHERE (PL.STOCK>0 OR PL.STOCK2>0 OR PL.STOCK3>0 OR PL.STOCK4>0 OR PL.STOCK5>0) AND (PL.PRICECALC>0) AND FMTS.FCLOSE = 0 '+ wfLE
            +'        INTO :ID,     '+ wfLE
            +'             :PRICEPL     '+ wfLE
            +'      DO    '+ wfLE
            +'      BEGIN     '+ wfLE
            +'        SUSPEND;    '+ wfLE
            +'      END     '+ wfLE
            +'    END ',true,false);

          wLog('Main','Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion',aNewVersion);

          FMWait.Free;
        except
          wLog('Main','Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;

      aNewVersion:= '0.0.4.14';
      if  CheckLessVersion(ADBase.ReadSettingByName('dbVersion'),aNewVersion) then
      begin

        FmWait:= TFmWait.Create(self);

        FmWait.Show;
        FMWait.InitBar(2,0);
        wLog('Main','Обновление метаданных БД... до версии '+aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus('Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate('CREATE OR ALTER TRIGGER INVOCES_AIUDE0 FOR INVOCES '+ wfLE
            +' ACTIVE AFTER INSERT OR UPDATE OR DELETE POSITION 0 '+ wfLE
            +' AS  begin    post_event ''INVOCES_Change'';  end ', true,false);
          wLog('Main','Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion',aNewVersion);

          FMWait.Free;
        except
          wLog('Main','Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;

      aNewVersion:= '0.0.6.0';
      if  CheckLessVersion(ADBase.ReadSettingByName('dbVersion'),aNewVersion) then
      begin

        FmWait:= TFmWait.Create(self);

        FmWait.Show;
        FMWait.InitBar(4,0);
        wLog('Main','Обновление метаданных БД... до версии '+aNewVersion);
        FmWait.SetStatus('--=== iPriceSE ===--');
        FmWait.SetStatus('Обновление метаданных БД...');
        FmWait.SetStatus('Дождитесь окончания всех операций. Изменение БД может занять несколько минут...');
        FMWait.SetBar(1);

        try
          ADBase.SQLUpdate('ALTER TABLE FORMATS ADD ACTUALDAYS INTEGER DEFAULT 3 ', true,false);
          ADBase.SQLUpdate('ALTER TABLE FORMATS ADD NOMINPRICE INTEGER DEFAULT 0 ', true,false);
          FMWait.SetBar(2);
          ADBase.SQLUpdate('UPDATE FORMATS SET ACTUALDAYS=3, NOMINPRICE=0', true,false);
          FMWait.SetBar(3);
          ADBase.SQLUpdate('CREATE OR ALTER procedure CATALOG_PL_ITEMS_PRICE (  '+ wfLE
            +'     IDCATALOG bigint)  '+ wfLE
            +' returns (  '+ wfLE
            +'     ID bigint, '+ wfLE
            +'     PRICEPL numeric(15,10))  '+ wfLE
            +' as '+ wfLE
            +' BEGIN    '+ wfLE
            +'      FOR     '+ wfLE
            +'        SELECT    '+ wfLE
            +'        PL.ID,    '+ wfLE
            +'        (PL.PRICECALC/MTH.QUANTITYINPACKING) AS PRICEPL     '+ wfLE
            +'        FROM "PL_ITEMS" PL      '+ wfLE
            +'        INNER JOIN "CATALOG_MATCHING" MTH ON (PL.ID = MTH.IDPL_ITEMS AND MTH.IDCATALOG=:IDCATALOG)  '+ wfLE
            +'        LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PL.IDFORMATS) '+ wfLE
            +'        WHERE (PL.STOCK>0 OR PL.STOCK2>0 OR PL.STOCK3>0 OR PL.STOCK4>0 OR PL.STOCK5>0) AND (PL.PRICECALC>0) AND FMTS.FCLOSE = 0  AND FMTS.NOMINPRICE = 0 AND IIF(FMTS.ACTUALDAYS>0,(DATEDIFF(day, PL.FTIMESTAMP, CURRENT_TIMESTAMP)<=FMTS.ACTUALDAYS),true)'+ wfLE
            +'        INTO :ID,     '+ wfLE
            +'             :PRICEPL     '+ wfLE
            +'      DO    '+ wfLE
            +'      BEGIN     '+ wfLE
            +'        SUSPEND;    '+ wfLE
            +'      END     '+ wfLE
            +'    END ',true,false);

          wLog('Main','Обновление метаданных БД [ОК]');

          ADBase.SetSettingByName('dbVersion',aNewVersion);

          FMWait.Free;
        except
          wLog('Main','Обновление метаданных БД [ОШИБКА]');
          FMWait.Free;
          raise;
        end;
      end;
  end;

    ADBase.SetSettingByName('dbVersion',_VersionRaw);
 except

 end;
end;

procedure TFmMain.Init(Sender:TObject);
var
  i, _progCountStart: integer;
  _DbIsClear: boolean;
  _LogUpdate: TStringList;
  _LogUpdateFile: String;
begin
  try

            fBase:= TwBase.Create(Sender);

            if fBase.SQLReadDS('OWNER',['ID'],'','').DataSet.RecordCount=1 then
               begin
                 ShowMessage('Список контрагентов пуст! Для работы приложения добавьте хотя бы одного контрагента!');
                 _DbIsClear:= true;
               end else
               begin
                 _DbIsClear:= false;
                 SetStatus('Сервер БД доступен и ожидает соединения',true);
                 SetStatus('Отсоединен от БД',false);
               end;

                if CheckLessVersion(_VersionRaw,fBase.ReadSettingByName('dbVersion')) then // проверка версии не ниже
                begin
                  ShowMessage('Версия БД новее используемой версии программы! Для продолжения работы используйте версию приложения не ниже v. '+fBase.ReadSettingByName('dbVersion')+'!');
                  fBase.Free;
                  exit;
                end;

                if CheckLessVersion(fBase.ReadSettingByName('dbVersion'),'0.0.1.27') then // был сильный рефакторинг БД ниже не поддерживается
               begin
                 ShowMessage('Версия БД несовместима с текущей версией приложения!');
                 fBase.Free;
                 exit;
               end;

                CheckDBVersion(fBase);

    //End Проверка версии БД

    // обслуживание индексов CATALOG каждые 100 запусков

    TryStrToInt(fBase.ReadSettingByName('progCountStart'),_progCountStart);

    if _progCountStart > 300 then
       begin
         FmWait:= TFmWait.Create(self);

         FmWait.Show;
         FMWait.InitBar(5,0);
         wLog('Main','Обслуживание БД...');
         FmWait.SetStatus('--=== iPriceSE ===--');
         FmWait.SetStatus('Обновление индексов...');
         FmWait.SetStatus('Это может занять некоторое время...');
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
           fBase.SetSettingByName('progCountStart',0);

           FmWait.SetStatus('Обновление индексов... [ОК]');
           FMWait.SetBar(5);
           FMWait.Free;
         end;
       end else
       begin
         // счетчик запусков программы
         fBase.SetSettingByName('progCountStart',_progCountStart+1);
       end;

    //End обслуживание индексов CATALOG каждые 100 запусков

          wLog('Main','Формирование списка плагинов...');

          if _DbIsClear then
               __wPluginSettings([TFmCatalog,TFmPrices,TFmFormats,TFmUtils,TFmAnalisis,TFmOrders],0)
               else
                __wPluginSettings([TFmCatalog,TFmPrices,TFmFormats,TFmUtils,TFmAnalisis,TFmOrders],0);

            __wPluginInit(FmMain,pcPlugins,self);

          if _DbIsClear then
             begin
              if Plugin <> nil then
                   Plugin.Add(TwPlugin.Create(2)); // Formats
             end;

          wLog('Main','Формирование списка плагинов успешно завершено.');

            fBase.free;
          except
            SetStatus('Отсоединен от БД',false);

            EnabledFuncional(false); // ограничиваем функционал

            //В случае отсутствия БД запускаем плагин Утилиты
            __wPluginSettings([TFmUtils],0);
            __wPluginInit(FmMain,pcPlugins,self);
            if Plugin <> nil then
                Plugin.Add(TwPlugin.Create(0));
            if Assigned(fBase) then
              fBase.Free;

            raise Exception.Create('Ошибка соединения с БД. Проверьте настройки соединения');
          end;

 _LogUpdateFile:= PathLogFiles_Unsafe+'log-update.txt';
 if FileExistsUTF8(_LogUpdateFile) then
    begin
        _LogUpdate  := TStringList.Create;
      try
        _LogUpdate.LoadFromFile(_LogUpdateFile);

        if (UTF8Pos('error',UTF8LowerCase(_LogUpdate.Text))>0) then
        if MessageDlg('Внимание! При последнем автоматическом обновлении были обнаружены ошибки!'+LineEnding+'Открыть файл в программе просмотра?',
              mtWarning, mbOKCancel, 0) = mrOK then
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
  fBase:= TwBase.Create(self);
  _Form:= TFmWait.Create(self);

  fDBImport:= TwDBImport.Create(self,_Form.Memo);

  _FieldsArray:= nil;

  _FieldsArray:=fBase.MakeArrayFromString(FormatImportFields);

  if Assigned(aSelectedPrices) then
  _Where:= ' IDFMTS_CATEGORY=1 AND FCLOSE=0 AND IDOWNER IN('+fBase.MakeStringFromArray(aSelectedPrices)+')' else// только прайс-листы
  _Where:= ' IDFMTS_CATEGORY=1 AND FCLOSE=0 ';// только прайс-листы

  _DataSet:= fBase.SQLReadDS('FORMATS',_FieldsArray,_Where,'IDOWNER, PRIORITY, NAME').DataSet;
  _DataSet.Last;
  _DataSet.First;
 fDBImport.FormatsPrice.Clear;
 for i:=0 to _DataSet.RecordCount-1 do
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
         fBase.MakeArrayArrayVariantFromString(_DataSet.FieldByName('STOCKSYMBOLS').AsString),
         _DataSet.FieldByName('STOCKONLYINFO').AsInteger,
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
         fBase.MakeArrayArrayIntegerFromString(_DataSet.FieldByName('SPREADSHEET').AsString,_DataSet.FieldByName('FIRSTLINE').AsInteger),
         _DataSet.FieldByName('IDVENDORCODEVARIANT').AsInteger,
         _DataSet.FieldByName('IDCSVDELIMITER').AsInteger,
         _DataSet.FieldByName('IDSTOCKVARIANT').AsInteger,
         _DataSet.FieldByName('IDPRICEVARIANT').AsInteger
         )
     );

     _DataSet.Next;
   end;


    fDBImport.IgnoreVersion:= false;
    _Form.BorderStyle:= bsSizeable;
    _Form.Height:=500;
    _Form.Width:=600;
    _Form.Memo.Alignment:= taLeftJustify;
    _Form.Memo.Font.Style:=_Form.Memo.Font.Style-[fsBold];
    _Form.Show;
    _Form.NoClose:= true;
   try
      fDBImport.Import();
      while not fDBImport.EndThread do
      begin
        Application.ProcessMessages;
      end;
      _Form.NoClose:= false;
   finally
     if Assigned(aSelectedPrices) and (aSelectedPrices[0]=fIdMainOwner) then
       _LogUpdateFile:= PathLogFiles_Unsafe+'log-update_ourprice.txt'
     else
       _LogUpdateFile:= PathLogFiles_Unsafe+'log-update.txt';
       if FileExistsUTF8(_LogUpdateFile) then DeleteFileUTF8(_LogUpdateFile);

       if FileExistsUTF8(PathLogFiles_Unsafe+'log-crash.txt') then
       begin
         _Form.Memo.Lines.Append('');
         _Form.Memo.Lines.Append('======= log-crash.txt =======');
         _Form.Memo.Lines.Append(__Log.Text);
       end;

       _Form.Memo.Lines.SaveToFile(_LogUpdateFile);
       _Form.Free;
       _DataSet:= nil;
       _FieldsArray:= nil;

       fBase.Destroy();
       fDBImport.Destroy();
       SetStatus('Импорт завершен.',false);
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
  i: Integer;
  //aLogExportFileSpreadSheet= PathLogFiles_Unsafe+'log-exportcatalogSpreadSheet.txt';

procedure WriteLog(aText: string);
begin
  aLog.Append(DateTimeToStr(now())+' | '+aText);
end;

begin
  fBase:= TwBase.Create(self);
  aCatalog:= TCatalog.Create(self, fBase, true);
  aLog:= TStringList.Create;

  try

case aType of
  eftCSV:
      begin
        aLogExportFile:= PathLogFiles_Unsafe+'log-exportcatalogCSV.txt';
        WriteLog('ExportCatalog in CSV...');
        aCatalog.ExportCatalogInCSV(aPatch, true);
        WriteLog('ExportCatalog in CSV... [OK]');
      end;
  eftSpreadSheet:
      begin
        aLogExportFile:= PathLogFiles_Unsafe+'log-exportcatalogSpreadSheet.txt';
        WriteLog('ExportCatalog in SpreadSheet...');

        aArr:= nil;
        aArr:= fBase.SQLReadArr('SELECT ID FROM PRICEFIELD WHERE PRIORITY IN ('+aPrices+')');
        SetLength(aPriceArr, Length(aArr));
        for i:=0 to High(aArr) do
          aPriceArr[i]:= aArr[i,0];
        aArr:= nil;

        aStockArr:= fBase.MakeArrayIntegerFromString(aStocks);
        aCatalog.ExportCatalogInSpreadsheet(aPatch, aStockArr, aPriceArr, true);
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
   fBase:= TwBase.Create(self);
   _Form:= TFmWait.Create(self);
   fBase.Memo:= _Form.Memo;

   try
     _Form.Memo.Lines.Add(DateTimeToStr(now)+'|Запущена операция резервного копирования...');
     _Form.Memo.Lines.Add(DateTimeToStr(now)+'|Соединяюсь с БД...');

     DateTimeToString(_DateTimeStr, 'dd_mm_yy_hh-mm-ss', now);

     if not DirectoryExistsUTF8(PathApplication_Unsafe+'BackupDB') then ForceDirectoriesUTF8(PathApplication_Unsafe+'BackupDB');

     _BackupFileName:= PathApplication_Unsafe+includeTrailingPathDelimiter('BackupDB')+_DateTimeStr+'_db.fbk';

     _Form.BorderStyle:= bsSizeable;
     _Form.Height:=500;
     _Form.Width:=600;
     _Form.Memo.Alignment:= taLeftJustify;
     _Form.Memo.Font.Style:=_Form.Memo.Font.Style-[fsBold];
     _Form.Show;
     _Form.NoClose:= true;

     try
       fBase.BackupBase(_BackupFileName);
       _Form.Memo.Lines.Add(DateTimeToStr(now)+'|Резервное копирование успешно завершено.');
     except
       raise;
     end;

   finally
     _LogBackupFile:= PathLogFiles_Unsafe+'log-backup.txt';
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
  i: Integer;
begin
 fBase:= TwBase.Create(self);

 _arr:= fBase.SQLReadArr('CURRENCY',['ID','KURS','FTIMESTAMP'],'ID<>1','ID');
 try

   if Assigned(_arr) and Assigned(TrayIcon.PopupMenu) then
   begin
      _PopupMenu:= TrayIcon.PopUpMenu;
      _PopupMenu.Items.Clear;

      _MenuItem:= TMenuItem.Create(_PopupMenu);
      _MenuItem.Name:='ShowMainForm';
      _MenuItem.Caption:= 'Восстановить/Спрятать';
      _MenuItem.ImageIndex:=0;
      _MenuItem.OnClick:=@mShowMainFormClick;
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem:= TMenuItem.Create(_PopupMenu);
      _MenuItem.Name:='RestoreFormSize';
      _MenuItem.Caption:= 'Восстановить размер окна по-умолчанию';
      _MenuItem.ImageIndex:=5;
      _MenuItem.OnClick:= @mTrayClick;
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem:= TMenuItem.Create(_PopupMenu);
      _MenuItem.Caption:= '-';
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem:= TMenuItem.Create(_PopupMenu);
      _MenuItem.Name:='Data';
      _MenuItem.Caption:= _arr[0,2];
      _MenuItem.ImageIndex:=1;
      _MenuItem.OnClick:=@mKursClick;
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem:= TMenuItem.Create(_PopupMenu);
      _MenuItem.Caption:= '-';
      _PopupMenu.Items.Add(_MenuItem);

      for i:=0 to High(_arr) do
          begin
             case i of
               0:
                 begin
                   _MenuItem:= TMenuItem.Create(_PopupMenu);
                   _MenuItem.Name:='USD';
                   _MenuItem.Caption:= _arr[i,1];
                   _MenuItem.ImageIndex:= 2;
                   _MenuItem.OnClick:=@mKursClick;
                   _PopupMenu.Items.Add(_MenuItem);
                 end;
               1:
                 begin
                   _MenuItem:= TMenuItem.Create(_PopupMenu);
                   _MenuItem.Name:='EUR';
                   _MenuItem.Caption:= _arr[i,1];
                   _MenuItem.ImageIndex:= 3;
                   _MenuItem.OnClick:=@mKursClick;
                   _PopupMenu.Items.Add(_MenuItem);
                 end;
             end;
          end;

      _MenuItem:= TMenuItem.Create(_PopupMenu);
      _MenuItem.Caption:= '-';
      _PopupMenu.Items.Add(_MenuItem);

      _MenuItem:= TMenuItem.Create(_PopupMenu);
      _MenuItem.Name:='CloseMainForm';
      _MenuItem.Caption:= 'Выйти';
      //_MenuItem.Sty
      _MenuItem.ImageIndex:=4;
      _MenuItem.OnClick:=@mCloseClick;
      _PopupMenu.Items.Add(_MenuItem);
   end;

   //mClose
 finally
   fBase.Destroy();
 end;
end;

procedure TFmMain.FormCreate(Sender: TObject);
var
  _FileSettings:TINIFile;
  i: Integer;
  _res: string;
  aTmpArray: ArrayOfInteger;
  fClearProps: Boolean;
begin
  try
      TrayIcon.Show;
      wFormID:=Self.Name;
      fSilentMode:= false;
      DefaultFormatSettings.ThousandSeparator:= ' ';
      DefaultFormatSettings.DecimalSeparator:= ',';

      PathApplication_Unsafe:= SysToUTF8(ExtractFilePath(ParamStr(0)));
      PathLogFiles_Unsafe:=PathApplication_Unsafe+'logs'+DirectorySeparator;
      if not DirectoryExistsUTF8(PathLogFiles_Unsafe) then ForceDirectoriesUTF8(PathLogFiles_Unsafe);

      PathExport_Unsafe:=PathApplication_Unsafe+DirectorySeparator+'export'+DirectorySeparator;
      if not DirectoryExistsUTF8(PathExport_Unsafe) then ForceDirectoriesUTF8(PathExport_Unsafe);

      PathTmp_Unsafe:=PathApplication_Unsafe+'tmp'+DirectorySeparator;

      PathTemplates_Unsafe:=PathApplication_Unsafe+'templates'+DirectorySeparator;

      DeleteDirectory(PathTmp_Unsafe,true);

      if not DirectoryExistsUTF8(PathTmp_Unsafe) then ForceDirectoriesUTF8(PathTmp_Unsafe);

       //FileExistsUTF8
      if FileExists (PathLogFiles_Unsafe+'log-crash.txt') then
         DeleteFile(PathLogFiles_Unsafe+'log-crash.txt');

      if FileExists (PathLogFiles_Unsafe+'log.txt') then
               DeleteFile(PathLogFiles_Unsafe+'log.txt');

      FileSettingsPath:='';
      // проверка пути к конфигурационному файлу
      for i:=1 to ParamCount do
      begin
        if UTF8Pos('-settings=',ParamStr(i))>0 then
        begin
          FileSettingsPath:= StringReplace(ParamStr(i),'-settings=','',[rfReplaceAll, rfIgnoreCase]);
          FileSettingsPath:= SafePath(StringReplace(FileSettingsPath,#34,'',[rfReplaceAll, rfIgnoreCase]));
        end;
      end;

      if Length(FileSettingsPath)=0 then
            FileSettingsPath:= PathApplication_Unsafe+'dbconfig.ini';


      try
        _FileSettings:=TINIFile.Create(FileSettingsPath,true);
        __onLog:= _FileSettings.ReadBool('Others','LogFileON',false);

        CatalogVendorCodeAsNumber:= _FileSettings.ReadBool('Others','CatalogVendorCodeAsNumber',false);

        fClearProps:= _FileSettings.ReadBool('Others','ClearProps',false);
        if fClearProps then
        begin
         DeleteFileUTF8(PathApplication_Unsafe+'ipricese.xml');
         _FileSettings.WriteString('Others','ClearProps', '0');
        end;

        {$IFDEF WINDOWS}
          PathLibreOffice:= _FileSettings.ReadString('Others','LibreOffice','');
          if Length(PathLibreOffice)=0 then PathLibreOffice:= GetLibreOfficeInstallation;
        {$ENDIF}


        db_portable:= _FileSettings.ReadBool('Others','Portable',true); // узнаем - портабле ли мы
        ReportHeaderColor:= RGBtoBGR(_FileSettings.ReadString('Others','ReportHeaderColor','#FFFBF0'));//$F0FBFF_
        FreeAndNil(_FileSettings);
      finally

      __Log:=TwLog.Create;

      _VersionRaw:= GetVersion;
      _Version:= _VersionRaw+' ['+wTargetOS+'] ';

      __Log.Add('Main','-= '+wProgName+' | версия: '+_Version+' =-');
      FmMain.Caption:= wProgName+' | версия: '+_Version;

      if FileExists(PathLibreOffice) then
         __Log.Add('Main','Найден LibreOffice '+PathLibreOffice) else
         __Log.Add('Main','LibreOffice '+PathLibreOffice+' НЕ НАЙДЕН!');

      __Log.Add('Main','Используемый файл настроек: '+FileSettingsPath);

      if __onLog then
         __Log.Add('Main','Ведение лог-файла [ON]') else
           __Log.Add('Main','Ведение лог-файла [OFF]');
      if db_portable then
            __Log.Add('Main','Приложение запущено в режиме [Portable]') else
                 __Log.Add('Main','Приложение запущено в режиме [Network]')
      end;

      __Log.Add('Main','Инициализация приложения...');

          //__dbReadSettings();

          //DBase
          __wDBaseReadSettings();


      fBase:= TwBase.Create(self);
      try
        fIdMainOwner:= fBase.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист
      finally
        fBase.Free;
      end;

      // проверка остальных параметров
      i:= 1;
      while i< ParamCount+1 do
      begin
        case ParamStr(i) of
           '-update':
                   begin
                     fSilentMode:= true;
                     DBUpdate(nil);
                   end;
           '-updateour':                    { TODO : выбор прайс-листов для загрузки }
                   begin
                     fSilentMode:= true;
                     DBUpdate([fIdMainOwner]);
                   end;
           '-updateis':
                   begin
                     fSilentMode:= true;
                     inc(i);
                     fBase:= TwBase.Create(self);
                     try
                       aTmpArray:= fBase.MakeArrayIntegerFromString(ParamStr(i));
                     finally
                       fBase.Free;
                     end;
                     DBUpdate(aTmpArray);
                   end;
           '-updatekurs':
                   begin
                      fSilentMode:= true;
                      UpdateKurs(true);
                   end;
           '-exportcatalogcsv':
                   begin
                     fSilentMode:= true;
                     inc(i);
                     ExportCatalog(ParamStr(i), '', '', eftCSV);
                   end;
           '-exportcatalogxls':
                   begin
                     fSilentMode:= true;
                     inc(i);
                     ExportCatalog(ParamStr(i),ParamStr(i+1), ParamStr(i+2), eftSpreadSheet);
                   end;
           '-backup':
                   begin
                     fSilentMode:= true;
                     DBBackup();
                   end;
        end;
        inc(i);
      end;

      if not fSilentMode then
            Init(Sender);// инициализация

      mTrayMenuFill(self); // заполнение трей меню

      __Log.Add('Main','Инициализация приложения завершена.');

  except
    on E: Exception do
      begin
        if FmWait<> nil then  FmWait.Free;

        __Log.SaveLogError(E);
        wLog('Main','Ошибка [FmMC]: "' + E.Message + '"');
        wLog('Main','Сбой инициализации приложения.');
        ShowMessage('Ошибка [FmMC]: "' + E.Message + '"');
        SetStatus(E.Message,true);
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
            wLog('Main','Ошибка [FmMC]: "' + E.Message + '"');
            wLog('Main','Сбой завершения приложения.');
            __Log.SaveLogError(E);
            ShowMessage(__Log.Text);
        end;
      end;
  finally

      // очищаем память
    FreeAndNil(Plugin);
    FreeAndNil(PluginList);

    __Log.Add('Main','Завершение приложения.');

    if __onLog then
    begin
      if __Log.SaveLog() then
      begin  // если лог-файл удалось сохранить, то удаляем краш
        if FileExists (PathLogFiles_Unsafe+'log-crash.txt') then
           DeleteFile(PathLogFiles_Unsafe+'log-crash.txt');
      end;

    end;

    if __Log <> nil then  FreeAndNil(__Log);
  end;
end;

procedure TFmMain.FormResize(Sender: TObject);
begin
  Status.Panels[0].Width:=Status.Width-200;
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
  if MessageDlg('Вы уверены, что хотите сбросить настройки внешнего вида: размер окна, ширину колонок и т.п.?',mtWarning, mbOKCancel, 0) = mrCancel then exit;
 _FileSettings:=TINIFile.Create(FileSettingsPath,true);

 try
   _FileSettings.WriteString('Others','ClearProps', '1');
    ShowMessage('Настройки внешнего вида будут сброшены при следующем запуске, перезапустите программу.');
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
fBase:= TwBase.Create(self);
fCatalog:= TCatalog.Create(self, fBase, true);

try
  fCatalog.ExportCatalogInCSV;
finally
  screen.Cursor:=crDefault;
  fCatalog.Free;
  fBase.Free;
end;
end;

procedure TFmMain.mExportCatalogInSpreadsheetClick(Sender: TObject);
begin
  fBase:= TwBase.Create(self);
  fCatalog:= TCatalog.Create(self, fBase, true);

  try
    fCatalog.ExportCatalogInSpreadsheet;
  finally
    screen.Cursor:=crDefault;
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
                       if  FmMain.WindowState<> wsNormal then
                            FmMain.WindowState:= wsNormal;
                       FmMain.Height:= 500;
                       FmMain.Width:= 900;
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
     FmMain.WindowState:= wsNormal;
     FmMain.Show;
     Application.Restore;
  end;
end;

procedure TFmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  i: LongInt;
begin
    // выгружаем подгруженные плагины
    if Plugin<> nil then
    begin
      for  i:=Plugin.Count-1 downto 0 do
         begin
              Plugin[i].Unload();
         end;
    end;
    if FmMain.WindowState in [wsMaximized, wsMinimized, wsFullScreen] then
       FmMain.WindowState:= wsNormal;
end;

procedure TFmMain.Button1Click(Sender: TObject);
begin

end;

procedure TFmMain.mCloseClick(Sender: TObject);
begin
  Close();
end;

procedure TFmMain.UpdateKurs(const aSilent:boolean = false);
var
  _DBImport: TwDBImport;
  aForm: TProgress;
begin
  _DBImport:= TwDBImport.Create(self);

  try
    _DBImport.ImportKursValut(aSilent);

    aForm:= TProgress.Create(self);
    aForm.BorderStyle:= bsNone;
    aForm.SetStatus('Идет обновление прайс-листа...');
    aForm.SetStatus('Дождитесь окончания операции');
    aForm.SetStatus('Это может занять несколько минут...');
    aForm.InitBar(pbTop,2);
    aForm.SetBar(pbTop,1);
    aForm.ShowBottom:= false;

    aForm.Show;
    try
      while not _DBImport.EndThread do
      begin
        Application.ProcessMessages;
      end;

    finally
      wLog('debug','Освобождаю форму...');
      FreeAndNil(aForm);
      wLog('debug','Освобождаю форму... OK');
    end;

    if aSilent then exit;

    if Length(_DBImport.ErrorMessage)>0 then ShowMessage(_DBImport.ErrorMessage);
    if Length(_DBImport.OKMessage)>0 then ShowMessage(_DBImport.OKMessage);

  finally
    wLog('debug','Освобождаю объект DBImport..');
    _DBImport.Destroy();
    wLog('debug','Освобождаю объект DBImport..');

    wLog('debug','Заполняю меню в трее...');
    mTrayMenuFill(self);
    wLog('debug','Заполняю меню в трее... ОК');
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

  _Form:= TFmAbout.Create(Self);

  _Form.t_ProgramName.Caption:=wProgName+' | версия: '+GetVersion;
  try
  _Form.ShowModal;

  finally
    _Form.Free;
  end;


end;


procedure TFmMain.mm_ClearDB_onStatusUpdate(Sender: TObject);
begin
 if Assigned(FmWait) then begin
   FmWait.SetStatus(fDataClearThread.Status);
   FMWait.SetBar(fDataClearThread.ProgressPosition);
 end;
end;

procedure TFmMain.mTrayClick(Sender: TObject);
begin
case TMenuItem(Sender).Name of
   'RestoreFormSize':
                     begin
                       if  FmMain.WindowState<> wsNormal then
                            FmMain.WindowState:= wsNormal;
                       FmMain.Height:= 500;
                       FmMain.Width:= 900;
                       FmMain.MoveToDefaultPosition;
                     end;
 end;
end;

procedure TFmMain.mm_ClearDB_onEnd(Sender: TObject);
var
  _Result: Boolean;
  _Status: String;
begin
    try
      //FmMain.Enabled:= true;
      FmWait.Close;
      _Result:= fDataClearThread.Result;
      _Status:= fDataClearThread.Status;

      fDataClearThread.Terminate;
      fDataClearThread:= nil;

      if _Result then
      begin
        SetStatus('Очистка БД успешно произведена.',true);
        ShowMessage('[DBClean] Очистка БД успешно произведена. После очистки рекомендуется сделать Backup/Restore через [ Утилиты->Обслуживание БД ].');
      end else
      begin
       SetStatus('Произошла ошибка при очистке БД',true);
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
  _PluginPageIndex: LongInt;
begin

    if MessageDlg('Очистить Базу данных? Это приведет к потере всех введенных данных!',mtWarning, mbOKCancel, 0) = mrCancel then exit;

    if MessageDlg('Вы уверены? Очистка БД необратима!',mtWarning, mbOKCancel, 0) = mrCancel then exit;

    fBase:= TwBase.Create(self);

    fDataClearThread:= nil;
    fDataClearThread:= TDataClearThread.Create(true);
    fDataClearThread.Base:= fBase;

    if Plugin<> nil then
    begin
      for  i:=Plugin.Count-1 downto 0 do
         begin
            _PluginPageIndex:= Plugin[i].PageIndex;
             try
               Plugin[i].Unload();
             finally
               FmMain.pcPlugins.Pages[_PluginPageIndex].Free;
             end;
         end;
    end;

       //DBase

       FmWait:= TFmWait.Create(self);

       FmWait.Height:=250;
       FmWait.Width:=540;

       FmWait.mStatus.Alignment:=taLeftJustify;
       FMWait.InitBar(8,0);
       wLog('Main','Очистка БД...');
       FmWait.SetStatus('--=== iPriceSE ===--');
       FmWait.SetStatus('Очистка БД...');
       FMWait.SetBar(2);

       fDataClearThread.ProgressPosition:=2;
       fDataClearThread.onEndThread:=@mm_ClearDB_onEnd;
       fDataClearThread.onStatusUpdate:=@mm_ClearDB_onStatusUpdate;
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
  _Form:= TFmPriceFields.Create(Self);

  try
  _Form.ShowModal;
  finally
    _Form.Free;
  end;
end;

procedure TFmMain.mKursClick(Sender:TObject);
begin
   Clipboard.AsText:= TMenuItem(Sender).Name+' = '+TMenuItem(Sender).Caption;
   ShowMessage('Строка скопирована в буфер обмена.');
end;

procedure TFmMain.tbPluginBtnCatalogClick(Sender: TObject);
begin
  wLog('Main','Запуск плагина: Каталог...');

  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(0)) else
    begin
      SetStatus('Нет соединения с БД.',false);
      SetStatus('Сбой инициализации плагина: Каталог.',true);
      wLog('Main','Сбой инициализации плагина: Каталог.');
      ShowMessage('Сбой инициализации плагина: Каталог.');
    end;
end;

procedure TFmMain.tbPluginBtnControlClick(Sender: TObject);
begin
  wLog('Main','Запуск плагина: Утилиты...');

  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(3)) else
    begin
      SetStatus('Нет соединения с БД.',false);
      SetStatus('Сбой инициализации плагина: Утилиты.',true);
      wLog('Main','Сбой инициализации плагина: Утилиты.');
      ShowMessage('Сбой инициализации плагина: Утилиты.');
    end;
end;

procedure TFmMain.tbPluginBtnFormatsClick(Sender: TObject);
begin
  wLog('Main','Запуск плагина: Форматы...');

  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(2)) else
    begin
      SetStatus('Нет соединения с БД.',false);
      SetStatus('Сбой инициализации плагина: Форматы.',true);
      wLog('Main','Сбой инициализации плагина: Форматы.');
      ShowMessage('Сбой инициализации плагина: Форматы.');
    end;
end;

procedure TFmMain.tbPluginBtnOrdersClick(Sender: TObject);
begin
 wLog('Main','Запуск плагина: Накладные...');

 if Plugin <> nil then
   Plugin.Add(TwPlugin.Create(5)) else
   begin
     SetStatus('Нет соединения с БД.',false);
     SetStatus('Сбой инициализации плагина: Накладные.',true);
     wLog('Main','Сбой инициализации плагина: Накладные.');
     ShowMessage('Сбой инициализации плагина: Накладные.');
   end;
end;

procedure TFmMain.tbPluginBtnPricesClick(Sender: TObject);
begin
  wLog('Main','Запуск плагина: Прайс-листы...');

  if Plugin <> nil then
    Plugin.Add(TwPlugin.Create(1)) else
    begin
      SetStatus('Нет соединения с БД.',false);
      SetStatus('Сбой инициализации плагина: Прайс-листы.',true);
      wLog('Main','Сбой инициализации плагина: Прайс-листы.');
      ShowMessage('Сбой инициализации плагина: Прайс-листы.');
    end;
end;

procedure TFmMain.ToolButton15Click(Sender: TObject);
begin
 wLog('Main','Запуск плагина: Аналитика...');
 SetStatus('Запуск плагина "Аналитика"... Это может занять немного времени...',true);
 Application.ProcessMessages;
 if Plugin <> nil then
   Plugin.Add(TwPlugin.Create(4)) else
   begin
     SetStatus('Нет соединения с БД.',false);
     SetStatus('Сбой инициализации плагина: Аналитика.',true);
     wLog('Main','Сбой инициализации плагина: Аналитика.');
     ShowMessage('Сбой инициализации плагина: Аналитика.');
   end;
end;

procedure TFmMain.SetStatus(_Text: string; _Log: boolean);
begin
  if _Log then
         Status.Panels.Items[0].Text:=_Text
     else
         Status.Panels.Items[1].Text:=_Text;

  wLog(wFormID,_Text);
end;

function TFmMain.GetStatus(_Log: boolean): string;
begin
  if _Log then
         result:= Status.Panels.Items[0].Text
     else
         result:= Status.Panels.Items[1].Text;
end;

end.

