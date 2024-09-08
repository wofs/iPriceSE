unit FmMatchingAddU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics,
  Dialogs, ExtCtrls, DBGrids, ComCtrls, Buttons, StdCtrls, Spin, Menus,
  XMLPropStorage, db, wLogU, UtilsU,
  wBaseU, wFuncU, wDBGridU, wTypesU,
  FmListSelectU, FmMatchingEditU, FmNomenclatureEditMassU
  ;

Type

  { TMyThread }

  TMyThread = class(TThread)
  private
    THRcnt: integer;
    THRi: integer;
    THRBase: TwBase;
    THRArrSearchPosition: ArrayOfArrayVariant;
    THRMatchLevel: integer;
    THRProgressBar: TProgressBar;
    THRGrid: TwDBGrid;
    THRLabel: TLabel;
    THRButton: TToolButton;
    THRIdMainOwner: integer;
    ThreadStarted: boolean;
    THRMode: integer; //0 - умный поиск, 1 - поиск по штрих-коду, 2 - писк по артикулу

    procedure ShowStatus;
  protected
    procedure Execute; override;
  public
    Constructor Create(CreateSuspended : boolean);
  end;

type

  { TFmMatchingAdd }

  TFmMatchingAdd = class(TForm)
    btnCancel: TBitBtn;
    btnPreventSearch: TSpeedButton;
    btnOK: TBitBtn;
    btnPriceEditSearchClear: TSpeedButton;
    edSearch: TComboBox;
    GridImageList: TImageList;
    ImageList16: TImageList;
    GridMatchingAdd: TDBGrid;
    ImageList24: TImageList;
    lbMinWCS: TLabel;
    lbMatchLevel: TLabel;
    lbPriceSearch: TLabel;
    lbStatus: TLabel;
    mAutoSearchMatching: TMenuItem;
    mAutoSearchMatching1: TMenuItem;
    mAutoSearchMatchingSelected: TMenuItem;
    mAutoSearchMatchingSelected1: TMenuItem;
    mCatalogAddItems: TMenuItem;
    mCatalogDetelMatching: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem3: TMenuItem;
    mAutoSearchMarchingScodLabel: TMenuItem;
    mAutoSearchMarchingScod: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem6: TMenuItem;
    mAutoSearchMarchingScodSelected: TMenuItem;
    mAutoSearchMarchingLabel: TMenuItem;
    mAutoSearchMarchingLabelSelected: TMenuItem;
    mBtnSearch: TPopupMenu;
    MenuItem7: TMenuItem;
    mBtnSearchMarchingScod: TMenuItem;
    mBtnSearchMarchingScodSelected: TMenuItem;
    mBtnSearchMarchingLabel: TMenuItem;
    mBtnSearchMarchingLabelSelected: TMenuItem;
    mMatchingAccept: TMenuItem;
    mMatchingEdit: TMenuItem;
    mMatchingEditList: TMenuItem;
    mSplitOther: TMenuItem;
    mSelectAll: TMenuItem;
    MenuItem2: TMenuItem;
    mClearSelect: TMenuItem;
    mClearSearchResult: TMenuItem;
    MenuItem5: TMenuItem;
    mOtherVariants: TMenuItem;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    pb1: TProgressBar;
    mGrid: TPopupMenu;
    rbSelected: TRadioButton;
    rbAll: TRadioButton;
    rbWithMatching: TRadioButton;
    rbNoMatching: TRadioButton;
    spMinWCS: TSpinEdit;
    ToolBar1: TToolBar;
    tbAutoSearchMatching: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    trbMatchLevel: TTrackBar;
    XMLPropStorage1: TXMLPropStorage;
    procedure btnCancelClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure btnPriceEditSearchClearClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GridMatchingAddDblClick(Sender: TObject);
    procedure mAutoSearchMarchingLabelClick(Sender: TObject);
    procedure mAutoSearchMarchingLabelSelectedClick(Sender: TObject);
    procedure mAutoSearchMarchingScodClick(Sender: TObject);
    procedure mAutoSearchMarchingScodSelectedClick(Sender: TObject);
    procedure mAutoSearchMatchingClick(Sender: TObject);
    procedure mAutoSearchMatchingSelectedClick(Sender: TObject);
    procedure mClearSearchResultClick(Sender: TObject);
    procedure mCatalogAddItemsClick(Sender: TObject);
    procedure mCatalogDetelMatchingClick(Sender: TObject);
    procedure mMatchingAcceptClick(Sender: TObject);
    procedure mMatchingEditClick(Sender: TObject);
    procedure mMatchingEditListClick(Sender: TObject);
    procedure tbAutoSearchMatchingClick(Sender: TObject);
    procedure XMLPropStorage1SavingProperties(Sender: TObject);
    procedure _OnSelectChange();
    procedure mSelectAllClick(Sender: TObject);
    procedure mClearSelectClick(Sender: TObject);
    procedure mGridPopup(Sender: TObject);
    procedure rbAllChange(Sender: TObject);
    procedure rbNoMatchingChange(Sender: TObject);
    procedure rbSelectedChange(Sender: TObject);
    procedure rbWithMatchingChange(Sender: TObject);
    procedure spMinWCSChange(Sender: TObject);
    procedure spMinWCSEditingDone(Sender: TObject);
    procedure spMinWCSKeyPress(Sender: TObject; var Key: char);
    procedure trbMatchLevelChange(Sender: TObject);
  private
    { private declarations }
    fFormName: string;
    fIdMainOwner: integer; // ID основного контрагента (к которому привязан каталог)
 //   _ForceClose: boolean;

    _SelectedPriceRows: ArrayOfInteger;
    _GridDataSet: TDataSet;
//   _TimeStampMaxArr: ArrayOfDateTime; // массив максимальных значений таймштамп из таблицы прайс-листы

    fBase: TwBase;
    fGridMatchingAdd: TwDBGrid;

    MinWCS: Integer;
    MinWCSEditingDone: boolean;
    _SelectedSearchRows: ArrayOfInteger;

    SearchThread: TMyThread;

    _PL_FTIMESTAMP: string;

    procedure StartSearchThread(Acnt: integer; ABase: TwBase; AArrSearchPosition: ArrayOfArrayVariant; AMatchLevel: integer; AProgressBar: TProgressBar;
      AGrid: TwDBGrid; ALabel: TLabel; AButton: TToolButton; aMode: integer = 0);
    procedure SwitchGrid();
    procedure _mOtherVariantsClick(Sender: TObject);
    procedure _mOtherVariantsListClick(Sender: TObject);
    procedure _OnSelect(Sender: TObject);

  public
    { public declarations }

    property SelectedRows: ArrayOfInteger read  _SelectedPriceRows write _SelectedPriceRows;
    property SelectedSearchRows: ArrayOfInteger read _SelectedSearchRows write _SelectedSearchRows;
    property GridDataSet: TDataSet read _GridDataSet write _GridDataSet;
//    property ForceClose: boolean read _ForceClose write _ForceClose;
    procedure SetStatus(_Text:string);
    procedure Log(aText: string);


    property Base: TwBase read fBase write fBase;

  end;

var
  FmMatchingAdd: TFmMatchingAdd;

implementation

{$R *.lfm}

{ TMyThread }


procedure TMyThread.ShowStatus;
var
  _BookMark: TBookMark;
begin
//   FmMatchingAdd.SetStatus(fStatusText);

  THRLabel.Caption:=' Обработано '+IntToStr(THRi)+' из '+IntToStr(THRcnt+1);

     if Terminated then
     begin
       _BookMark:= THRGrid.Grid.DataSource.DataSet.Bookmark;
       THRGrid.Grid.DataSource.DataSet.Close;
       THRGrid.Grid.DataSource.DataSet.Open;
       if THRGrid.Grid.DataSource.DataSet.RecordCount>0 then
       THRGrid.Grid.DataSource.DataSet.Bookmark:= _BookMark;
       THRLabel.Caption:= ' Все операции завершены.';
       THRProgressBar.Position:=THRProgressBar.Max;
       THRButton.ImageIndex:=0;
       THRProgressBar.StepIt;
     end;

     THRProgressBar.StepIt;
end;

procedure TMyThread.Execute;
var
    _indistinctmatching: string;
    i, iScod: Integer;
    _arr, _arrScodCtg: ArrayOfArrayVariant;
  begin
  //  fStatusText := 'TMyThread Starting...';
  //  Synchronize(@Showstatus);
  //  fStatusText := 'TMyThread Running...';
       i:=-1;
     _indistinctmatching:='';
    while (not Terminated) do
      begin
          for i:=0 to THRcnt do
          begin
              try
                THRBase.SQLUpdate('delete from W_TMP_TBL_NEUTRALSEARCH where IDMATCHPOSITION='+string(THRArrSearchPosition[i,0])+';',false);

                case THRMode of
                  0:
                    begin
                     case THRMatchLevel of
                       1: _indistinctmatching:= ' indistinctmatching4('+QuotedStr(string(THRArrSearchPosition[i,1]))+',CTG.NAME) ';
                       2: _indistinctmatching:= ' indistinctmatchingmanual('+QuotedStr(string(THRArrSearchPosition[i,1]))+',CTG.NAME,'+QuotedStr(IntToStr(CalcProbel(string(THRArrSearchPosition[i,1]),false)))+') ';
                     end;

                     THRBase.SQLUpdate('insert into W_TMP_TBL_NEUTRALSEARCH '
                     +' select CTG.IDOWNER,CTG.ID,CTG.NAME '
                     +' ,'+_indistinctmatching+' as WCS,'+string(THRArrSearchPosition[i,0])+' '
                     +' from "CATALOG" CTG',false);

                    end;
                  1:
                    begin
                       _arr:=nil;
                       _arr:= THRBase.SQLReadArr('PL_SCODS',['SCOD'],'IDPL_ITEMS='+string(THRArrSearchPosition[i,0]),'SCOD');

                       if Assigned(_arr) then
                       for iScod:=0 to High(_arr) do
                       begin
                           _arrScodCtg:=nil;
                           _arrScodCtg:= THRBase.SQLReadArr('SELECT IDCTG_ITEMS FROM CTG_CHECK_SCOD('+IntToStr(THRIdMainOwner)+','+QuotedStr(_arr[iScod,0])+')');
                          if Assigned(_arrScodCtg) and (_arrScodCtg[0,0]<>null) then
                             THRBase.SQLUpdate('insert into W_TMP_TBL_NEUTRALSEARCH'
                             +' select CTG.IDOWNER,CTG.ID,CTG.NAME '
                             +' ,95 as WCS,'+string(THRArrSearchPosition[i,0])+' '
                             +' from "CATALOG" CTG'
                             +' where CTG.ID='+string(_arrScodCtg[0,0])
                             ,false);
                       end;
                      end;

                  2:
                    begin
                     THRBase.SQLUpdate('insert into W_TMP_TBL_NEUTRALSEARCH '
                     +' select CTG.IDOWNER,CTG.ID,CTG.NAME '
                     +' ,95 as WCS,'+string(THRArrSearchPosition[i,0])+' '
                     +' from "CATALOG" CTG'
                     +' where CTG.LABEL='+QuotedStr(THRArrSearchPosition[i,1])+' AND CTG.LABEL <>'''' AND CTG.LABEL IS NOT NULL'
                     ,false);

                     THRBase.SQLUpdate('insert into W_TMP_TBL_NEUTRALSEARCH '
                     +' select CTG.IDOWNER,CTG.ID,CTG.NAME '
                     +' ,80 as WCS,'+string(THRArrSearchPosition[i,0])+' '
                     +' from "CATALOG" CTG'
                     +' where CTG.LABEL LIKE '+QuotedStr('%'+THRArrSearchPosition[i,1]+'%'+' AND CTG.LABEL <>'''' AND CTG.LABEL IS NOT NULL')
                     ,false);
                    end;
                end;

              finally
                THRi:=i;
                Synchronize(@Showstatus);
              end;

             if (i=THRcnt) then
             begin
               Terminate;
               Synchronize(@Showstatus);
             end;
          end;

      end;

//    Synchronize(@Showstatus);
end;

constructor TMyThread.Create(CreateSuspended: boolean);
begin
   FreeOnTerminate := true;
   inherited Create(CreateSuspended);
end;

{ TFmMatchingAdd }

procedure TFmMatchingAdd.SwitchGrid();
var
  _BookMark: TBookMark;
  _SQLString, _Where: string;
begin
  _SQLString:='';
   if rbAll.Checked then
   _SQLString:='/*rbAll*/ SELECT PL.ID,  '
     +' IIF("CATALOG".NAME IS NULL, '
     +' (SELECT BREADCRUMPS FROM GETPARENTS_GROUP_CATALOG(WTMP.CIDCTG_GROUP)), '
     +' (SELECT BREADCRUMPS FROM GETPARENTS_GROUP_CATALOG("CATALOG".IDCTG_GROUP)))  AS GBREADCRUMPS, '
     +' IIF("CATALOG".NAME IS NULL, WTMP.NAME,"CATALOG".NAME) AS CNAME, '
     +' IIF("CATALOG".NAME IS NULL, WTMP.WCS,'''') AS WCS, '
     +' PL.IDOWNER, '//MTH.VENDORCODE as MATCHINGVENDORCODE, '
     +' IIF(MTH.QUANTITYINPACKING IS NULL, IIF(WTMP.NAME IS NULL,0,1) ,MTH.QUANTITYINPACKING) AS QUANTITYINPACKING, IIF("CATALOG".ID IS NULL,WTMP.CIDCTG_GROUP, "CATALOG".IDCTG_GROUP)  as CIDCTG_GROUP, '
     +' PL.NAME AS PLNAME ,PL.VENDORCODE AS PLVENDORCODE  ,PL.PRICECALC AS PLPRICE  ,"OWNER".NAME AS OWNERNAME, IIF("CATALOG".ID IS NULL,WTMP.ID, "CATALOG".ID) AS WCATALOGID, WTMP.IDMATCHPOSITION '
     +' FROM PL_ITEMS PL '
     +' LEFT OUTER JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS = PL.ID) '
     +' LEFT OUTER JOIN "CATALOG" ON ("CATALOG".ID = MTH.IDCATALOG) '
     +' LEFT OUTER JOIN ( '
     +' SELECT T1.IDOWNER, min(T1.ID) as ID, min(T1.NAME) as NAME, T1.WCS, PLWCSMAX.IDMATCHPOSITION, min(CATALOG.IDCTG_GROUP) AS CIDCTG_GROUP '
     +' FROM W_TMP_TBL_NEUTRALSEARCH T1 '
     +' INNER JOIN   (SELECT IDMATCHPOSITION, IDOWNER, MAX(WCS) MWCS FROM W_TMP_TBL_NEUTRALSEARCH  WHERE WCS>'+IntToStr(MinWCS)+'  GROUP BY 1,2 '
     +' ) PLWCSMAX ON (T1.WCS=PLWCSMAX.MWCS AND T1.IDMATCHPOSITION = PLWCSMAX.IDMATCHPOSITION AND T1.IDOWNER=PLWCSMAX.IDOWNER) '
     +' LEFT OUTER JOIN "CATALOG" ON (CATALOG.ID = T1.ID) '
     +' GROUP BY 1,4,5 '
     +' ) WTMP ON (WTMP.IDMATCHPOSITION=PL.ID) '
     +' LEFT OUTER JOIN "CATALOG_GROUP" ON ("CATALOG_GROUP".ID = "CATALOG".IDCTG_GROUP) '
     +' LEFT OUTER JOIN "OWNER" ON ("OWNER".ID = PL.IDOWNER) '
     +' WHERE /*group_string*/ /*and_search_string*/ ';

      if rbWithMatching.Checked then
      _SQLString:= '/*rbWithMatching*/ SELECT PL.ID, (select BREADCRUMPS from GETPARENTS_GROUP_CATALOG("CATALOG".IDCTG_GROUP)) as GBREADCRUMPS, "CATALOG".name as CNAME,'
         +' MTH.quantityinpacking, PL.NAME AS PLNAME, IIF("CATALOG".ID IS NULL, 0, "CATALOG".ID) AS WCATALOGID'
         +' ,PL.VENDORCODE AS PLVENDORCODE,PL.PRICECALC AS PLPRICE, "CATALOG".IDCTG_GROUP '
         +' ,"OWNER".NAME as OWNERNAME '
         +', '''' as WCS '
         +' FROM PL_ITEMS PL '
         +' left outer join "CATALOG_MATCHING" MTH on (MTH.IDPL_ITEMS = PL.ID) '
         +' left outer join "CATALOG" on ("CATALOG".id = MTH.idcatalog) ' //right
         +' left outer join "CATALOG_GROUP" on ("CATALOG_GROUP".id = "CATALOG".IDCTG_GROUP) '
         +' left outer join "OWNER" on ("OWNER".id = PL.idowner) '
         +' WHERE '
         +' (MTH.IDPL_ITEMS = PL.ID) '
         +'  /*and_group_string*/ /*and_search_string*/ ';

      if rbNoMatching.Checked then
          _SQLString:= '/*rbNoMatching*/ SELECT PL.ID, '
         +' (SELECT BREADCRUMPS FROM GETPARENTS_GROUP_CATALOG(WTMP.CIDCTG_GROUP)) AS GBREADCRUMPS, '
         +' WTMP.NAME AS CNAME, '
         +' WTMP.WCS AS WCS, '
         +' PL.IDOWNER, '// MTH.VENDORCODE as MATCHINGVENDORCODE, '
         +' IIF(WTMP.NAME IS NULL,0, 1) AS QUANTITYINPACKING,  WTMP.CIDCTG_GROUP,'
         +' PL.NAME AS PLNAME ,PL.VENDORCODE AS PLVENDORCODE  ,PL.PRICECALC AS PLPRICE  ,"OWNER".NAME AS OWNERNAME, IIF("CATALOG".ID IS NULL, WTMP.ID, "CATALOG".ID) AS WCATALOGID, WTMP.IDMATCHPOSITION '
          +' FROM "PL_ITEMS" PL '

          +' LEFT OUTER JOIN ( '
          +' SELECT T1.IDOWNER, min(T1.ID) as ID, min(T1.NAME) as NAME, T1.WCS, PLWCSMAX.IDMATCHPOSITION, min(CATALOG.IDCTG_GROUP) AS CIDCTG_GROUP '
          +' FROM W_TMP_TBL_NEUTRALSEARCH T1 '
          +' INNER JOIN   (SELECT IDMATCHPOSITION, IDOWNER, MAX(WCS) MWCS FROM W_TMP_TBL_NEUTRALSEARCH  WHERE WCS>'+IntToStr(MinWCS)+'  GROUP BY 1,2 '
          +' ) PLWCSMAX ON (T1.WCS=PLWCSMAX.MWCS AND T1.IDMATCHPOSITION = PLWCSMAX.IDMATCHPOSITION AND T1.IDOWNER=PLWCSMAX.IDOWNER) '
          +' LEFT OUTER JOIN "CATALOG" ON (CATALOG.ID = T1.ID) '
          +' GROUP BY 1,4,5 '
          +'   ) WTMP ON (WTMP.IDMATCHPOSITION=PL.ID) '

          +' LEFT OUTER JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS = PL.ID) '
          +'  LEFT OUTER JOIN "CATALOG" ON ("CATALOG".ID = WTMP.ID)  LEFT OUTER JOIN "OWNER" ON ("OWNER".ID = PL.IDOWNER) '
          +'  WHERE MTH.IDCATALOG IS NULL /*and_group_string*/ /*and_search_string*/ ';

      if rbSelected.Checked then
      begin
        if fGridMatchingAdd.SelectedRowsCount >0 then
                _Where:= '('+fBase.PrepareWhereString('PL.ID',fGridMatchingAdd.SelectedRows)+')' else
                _Where:= 'PL.ID = 0';

        _SQLString:='/*rbSelected*/ SELECT PL.ID,  '
         +' IIF("CATALOG".NAME IS NULL, '
         +' (SELECT BREADCRUMPS FROM GETPARENTS_GROUP_CATALOG(WTMP.CIDCTG_GROUP)), '
         +' (SELECT BREADCRUMPS FROM GETPARENTS_GROUP_CATALOG("CATALOG".IDCTG_GROUP)))  AS GBREADCRUMPS, '
         +' IIF("CATALOG".NAME IS NULL, WTMP.NAME,"CATALOG".NAME) AS CNAME, '
         +' IIF("CATALOG".NAME IS NULL, WTMP.WCS,'''') AS WCS, '
         +' PL.IDOWNER, '// MTH.VENDORCODE as MATCHINGVENDORCODE, '
         +' IIF(MTH.QUANTITYINPACKING IS NULL, IIF(WTMP.NAME IS NULL,0,1) ,MTH.QUANTITYINPACKING) AS QUANTITYINPACKING, IIF("CATALOG".ID IS NULL,WTMP.CIDCTG_GROUP, "CATALOG".IDCTG_GROUP)  as CIDCTG_GROUP, '
         +' PL.NAME AS PLNAME ,PL.VENDORCODE AS PLVENDORCODE  ,PL.PRICECALC AS PLPRICE  ,"OWNER".NAME AS OWNERNAME, IIF("CATALOG".ID IS NULL,WTMP.ID, "CATALOG".ID) AS WCATALOGID, WTMP.IDMATCHPOSITION '
         +' FROM PL_ITEMS PL '
         +' LEFT OUTER JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS = PL.ID) '
         +' LEFT OUTER JOIN "CATALOG" ON ("CATALOG".ID = MTH.IDCATALOG) '
         +' LEFT OUTER JOIN ( '
         +' SELECT T1.IDOWNER, min(T1.ID) as ID, min(T1.NAME) as NAME, T1.WCS, PLWCSMAX.IDMATCHPOSITION, min(CATALOG.IDCTG_GROUP) AS CIDCTG_GROUP '
         +' FROM W_TMP_TBL_NEUTRALSEARCH T1 '
         +' INNER JOIN   (SELECT IDMATCHPOSITION, IDOWNER, MAX(WCS) MWCS FROM W_TMP_TBL_NEUTRALSEARCH  WHERE WCS>'+IntToStr(MinWCS)+'  GROUP BY 1,2 '
         +' ) PLWCSMAX ON (T1.WCS=PLWCSMAX.MWCS AND T1.IDMATCHPOSITION = PLWCSMAX.IDMATCHPOSITION AND T1.IDOWNER=PLWCSMAX.IDOWNER) '
         +' LEFT OUTER JOIN "CATALOG" ON (CATALOG.ID = T1.ID) '
         +' GROUP BY 1,4,5 '
         +' ) WTMP ON (WTMP.IDMATCHPOSITION=PL.ID) '
         +' LEFT OUTER JOIN "CATALOG_GROUP" ON ("CATALOG_GROUP".ID = "CATALOG".IDCTG_GROUP) '
         +' LEFT OUTER JOIN "OWNER" ON ("OWNER".ID = PL.IDOWNER) '
         +' WHERE '+_Where+'  /*and_group_string*/ /*and_search_string*/ ';
      end;

      if Assigned(GridMatchingAdd.DataSource) then
        begin
          _BookMark:=  GridMatchingAdd.DataSource.DataSet.BookMark;
        end;

      fGridMatchingAdd.SQL:=_SQLString;

      fGridMatchingAdd.Fill;

      if  GridMatchingAdd.DataSource.DataSet.RecordCount>0 then
       GridMatchingAdd.DataSource.DataSet.Bookmark:= _BookMark;

end;

procedure TFmMatchingAdd.FormCreate(Sender: TObject);
var
  _TimeStampMaxArr: ArrayOfDateTime;
  _SQLString: String;
begin

  ShowInTaskBar:= stAlways;

  fFormName:=Self.Name;
  fIdMainOwner:=0;
  //ForceClose:= false;
  //spMinWCS.Value:= MinWCS;
  SelectedRows:=nil;
  wLog(fFormName,'Инициализация формы... ['+fFormName+']');

  fBase:= TwBase.Create(Sender);
  try
   screen.Cursor:= crSQLWait;
   wLog(fFormName,'Инициализация формы успешно завершена.');

   except
     on E: Exception do
     begin
         screen.Cursor:= crDefault;
         __Log.SaveLogError(E);
         SetStatus('Сбой инициализации формы.');
         wLog(fFormName,'Ошибка [FmCreate]: "' + E.Message + '"');
         wLog(fFormName,'Сбой инициализации формы.');
         ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

      end;
   end;
end;

procedure TFmMatchingAdd.FormDestroy(Sender: TObject);
var
  _BookMark: TBookMark;
begin
  try
   if __Log<> nil then
   wLog('FmGroupAdditingMatching','Выгрузка формы...');

   //if Assigned(SearchThread) then SearchThread.Free;

   if Assigned(_GridDataSet) then
     begin

        try
          _BookMark:= _GridDataSet.Bookmark;
          _GridDataSet.Close;
          _GridDataSet.Open;

          if _GridDataSet.RecordCount>0 then
            _GridDataSet.GotoBookmark(_BookMark);
        except

        end;

     end;

   fGridMatchingAdd.Destroy();
   fBase.Destroy();
   if __Log<> nil then
  wLog('FmGroupAdditingMatching','Выгрузка формы успешно завершена.');

  except
    on E: Exception do
    begin
       if __Log<> nil then
        __Log.SaveLogError(E);
        SetStatus('Сбой выгрузки формы.');
       if __Log<> nil then
        wLog('FmGroupAdditingMatching','Ошибка [FmDestroy]: "' + E.Message + '"');
       if __Log<> nil then
        wLog('FmGroupAdditingMatching','Сбой выгрузки формы.');
        ShowMessage('Ошибка [FmDestroy]: "' + E.Message + '"');
     end;
  end;
end;

procedure TFmMatchingAdd.FormShow(Sender: TObject);
var
  _SQLString: String;
  __PRICE_MAX_FTIMESTAMP_ARR: ArrayOfDateTime;
begin

  fIdMainOwner:= fBase.ReadSettingByName('setDefaultOwner');


  MinWCS:= spMinWCS.Value;

  __PRICE_MAX_FTIMESTAMP_ARR:= GetMaxFTimeStampPricesArr(fBase);

  //_PL_FTIMESTAMP:= fBase.PrepareWhereStringFromDateTime('"PRICE-LISTS".FTIMESTAMP',__PRICE_MAX_FTIMESTAMP_ARR);

 ////_DBGridList
 _SQLString:='/*rbAll*/ SELECT PL.ID,  '
   +' IIF("CATALOG".NAME IS NULL, '
   +' (SELECT BREADCRUMPS FROM GETPARENTS_GROUP_CATALOG(WTMP.CIDCTG_GROUP)), '
   +' (SELECT BREADCRUMPS FROM GETPARENTS_GROUP_CATALOG("CATALOG".IDCTG_GROUP)))  AS GBREADCRUMPS, '
   +' IIF("CATALOG".NAME IS NULL, WTMP.NAME,"CATALOG".NAME) AS CNAME, '
   +' IIF("CATALOG".NAME IS NULL, WTMP.WCS,'''') AS WCS, '
   +' PL.IDOWNER, PL.VENDORCODE as PLVENDORCODE, '
   +' IIF(MTH.QUANTITYINPACKING IS NULL, IIF(WTMP.NAME IS NULL,0,1) ,MTH.QUANTITYINPACKING) AS QUANTITYINPACKING, IIF("CATALOG".ID IS NULL,WTMP.CIDCTG_GROUP, "CATALOG".IDCTG_GROUP)  as CIDCTG_GROUP, '
   +' PL.NAME AS PLNAME ,PL.VENDORCODE AS PLVENDORCODE  ,PL.PRICECALC AS PLPRICE  ,"OWNER".NAME AS OWNERNAME, IIF("CATALOG".ID IS NULL,WTMP.ID, "CATALOG".ID) AS WCATALOGID, WTMP.IDMATCHPOSITION '
   +' FROM PL_ITEMS PL '
   +' LEFT OUTER JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS = PL.ID) '
   +' LEFT OUTER JOIN "CATALOG" ON ("CATALOG".ID = MTH.IDCATALOG) '
   +' LEFT OUTER JOIN ( '
   +' SELECT T1.IDOWNER, min(T1.ID) as ID, min(T1.NAME) as NAME, T1.WCS, PLWCSMAX.IDMATCHPOSITION, min(CATALOG.IDCTG_GROUP) AS CIDCTG_GROUP '
   +' FROM W_TMP_TBL_NEUTRALSEARCH T1 '
   +' INNER JOIN   (SELECT IDMATCHPOSITION, IDOWNER, MAX(WCS) MWCS FROM W_TMP_TBL_NEUTRALSEARCH  WHERE WCS>'+IntToStr(MinWCS)+'  GROUP BY 1,2 '
   +' ) PLWCSMAX ON (T1.WCS=PLWCSMAX.MWCS AND T1.IDMATCHPOSITION = PLWCSMAX.IDMATCHPOSITION AND T1.IDOWNER=PLWCSMAX.IDOWNER) '
   +' LEFT OUTER JOIN "CATALOG" ON (CATALOG.ID = T1.ID) '
   +' GROUP BY 1,4,5 '
   +' ) WTMP ON (WTMP.IDMATCHPOSITION=PL.ID) '
   +' LEFT OUTER JOIN "CATALOG_GROUP" ON ("CATALOG_GROUP".ID = "CATALOG".IDCTG_GROUP) '
   +' LEFT OUTER JOIN "OWNER" ON ("OWNER".ID = PL.IDOWNER) '
   +' WHERE  /*group_string*/ /*and_search_string*/ ';
                           // /*and_search_string*/
 fBase.LongTransaction:= true;  //start read write transacrion

 //DBGrid.Add(TwDBGrid.Create(fFormName, GridMatchingAdd,true,nil,edSearch,btnPreventSearch,_SQLString,['$SEARCHSTRING=$TABLE.NAME','$SEARCHOTHERSTRING=$TABLE.VENDORCODE','$WHEREROWS=','$WHERETIMESTAMP=','$ORDERBY=$TABLE.NAME','$TABLE=PRICE-LISTS'],false)); // инициализация DBGrid
 fGridMatchingAdd:= TwDBGrid.Create(fBase,GridMatchingAdd,_SQLString);
 fGridMatchingAdd.MultiSelect:= true;
 fGridMatchingAdd.SearchEdit:= edSearch;
 fGridMatchingAdd.SearchPreventiveBtn:= btnPreventSearch;
 fGridMatchingAdd.SearchEntryArray:= ['PL.NAME'];
 fGridMatchingAdd.SearchParticleArray:= ['PL.VENDORCODE'];
 fGridMatchingAdd.GroupField:='PL.ID';
 fGridMatchingAdd.onSelect:=@_OnSelect;
 fGridMatchingAdd.SortTitleImagesIndex:=[1,2];
  if Length(SelectedRows)>0 then
     fGridMatchingAdd.GroupArray:= SelectedRows
     else
     fGridMatchingAdd.GroupArray:= nil;

  fGridMatchingAdd.Fill;

  ModalResult:= mrCancel;

  screen.Cursor:= crDefault;
end;

procedure TFmMatchingAdd.GridMatchingAddDblClick(Sender: TObject);
begin
     if fGridMatchingAdd.FieldName = 'CNAME' then
       begin
          if (ssShift in fGridMatchingAdd.ShiftState) then
             mMatchingEditListClick(Sender) else
             mMatchingEditClick(Sender);
       end;

     if fGridMatchingAdd.FieldName = 'QUANTITYINPACKING' then
       begin
            mMatchingEditClick(Sender);
       end;
//
end;

procedure TFmMatchingAdd.mAutoSearchMarchingLabelClick(Sender: TObject);
var
     _WhereSearchString: String;
    _arrSearchPosition: ArrayOfArrayVariant;
    cnt: integer;
begin
     if tbAutoSearchMatching.ImageIndex = 3 then
     begin
       if MessageDlg('Поиск уже запущен. Отменить?',mtConfirmation, mbOKCancel, 0) = mrOK then
            SearchThread.Terminate;
       exit;
     end;

     SetStatus('Подбор соответствий для выбранной позиции...');

     _WhereSearchString := 'ID='+GridMatchingAdd.DataSource.DataSet.FieldByName('ID').AsString;

     _arrSearchPosition:= nil;
     _arrSearchPosition := fBase.SQLReadArr('PL_ITEMS',['ID','LABEL'],' ('+_WhereSearchString+')','NAME');

      cnt:= High(_arrSearchPosition);
      pb1.Position:=0;
      pb1.Max:=cnt+1;
      pb1.Step:=1;


  // вызываем поток
     StartSearchThread(cnt, fBase, _arrSearchPosition, trbMatchLevel.Position, pb1, fGridMatchingAdd, lbStatus, tbAutoSearchMatching,2);
end;

procedure TFmMatchingAdd.mAutoSearchMarchingLabelSelectedClick(Sender: TObject);
var
 cnt: integer;
   _arrSearchPosition: ArrayOfArrayVariant;
   _WhereSearchString: string;

begin

 if tbAutoSearchMatching.ImageIndex = 3 then
   begin
     if MessageDlg('Поиск уже запущен. Отменить?',mtConfirmation, mbOKCancel, 0) = mrOK then
          SearchThread.Terminate;
     exit;
   end;


  SelectedSearchRows:= fGridMatchingAdd.SelectedRows;

  if Length(SelectedSearchRows)>1 then
        if MessageDlg('Подобрать соотвествия для выбранных позиций автоматически? Это может занять несколько минут, в зависимости от количества позиций и объема каталога.',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

   SetStatus('Подбор соответствий для выбранных позиций...');

   _WhereSearchString := fBase.PrepareWhereString('ID',SelectedSearchRows);

   _arrSearchPosition:= nil;
   _arrSearchPosition := fBase.SQLReadArr('PL_ITEMS',['ID','LABEL'],'('+_WhereSearchString+')','NAME');

    cnt:= High(_arrSearchPosition);
    pb1.Position:=0;
    pb1.Max:=cnt+1;
    pb1.Step:=1;
 //   pb1.Step:=1;


// вызываем поток
   StartSearchThread(cnt, fBase, _arrSearchPosition, trbMatchLevel.Position, pb1, fGridMatchingAdd, lbStatus, tbAutoSearchMatching,2);

   _arrSearchPosition:= nil;

end;

procedure TFmMatchingAdd.mAutoSearchMarchingScodClick(Sender: TObject);
var
  _WhereSearchString: String;
  _arrSearchPosition: ArrayOfArrayVariant;
  cnt: Integer;
begin
  if tbAutoSearchMatching.ImageIndex = 3 then
  begin
    if MessageDlg('Поиск уже запущен. Отменить?',mtConfirmation, mbOKCancel, 0) = mrOK then
         SearchThread.Terminate;
    exit;
  end;

  SetStatus('Подбор соответствий для выбранной позиции...');

  //_WhereSearchString := 'PLS.IDPL_ITEMS='+GridMatchingAdd.DataSource.DataSet.FieldByName('ID').AsString;

  _arrSearchPosition:= nil;
  //_arrSearchPosition := fBase.SQLReadArr('PL_ITEMS',['ID','NAME','IDOWNER','VENDORCODE','SCOD','LABEL'],' ('+_WhereSearchString+')','NAME');
  SetLength(_arrSearchPosition,1,1);
  _arrSearchPosition[0,0] := GridMatchingAdd.DataSource.DataSet.FieldByName('ID').AsInteger;

   cnt:= High(_arrSearchPosition);
   pb1.Position:=0;
   pb1.Max:=cnt+1;
   pb1.Step:=1;


// вызываем поток
  StartSearchThread(cnt, fBase, _arrSearchPosition, trbMatchLevel.Position, pb1, fGridMatchingAdd, lbStatus, tbAutoSearchMatching,1);
end;

procedure TFmMatchingAdd.mAutoSearchMarchingScodSelectedClick(Sender: TObject);
var
 cnt, i: integer;
   _arrSearchPosition: ArrayOfArrayVariant;
   _WhereSearchString: string;

begin

 if tbAutoSearchMatching.ImageIndex = 3 then
   begin
     if MessageDlg('Поиск уже запущен. Отменить?',mtConfirmation, mbOKCancel, 0) = mrOK then
          SearchThread.Terminate;
     exit;
   end;


  SelectedSearchRows:= fGridMatchingAdd.SelectedRows;

  if Length(SelectedSearchRows)>1 then
        if MessageDlg('Подобрать соотвествия для выбранных позиций автоматически? Это может занять несколько минут, в зависимости от количества позиций и объема каталога.',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

   SetStatus('Подбор соответствий для выбранных позиций...');

   _WhereSearchString := fBase.PrepareWhereString('PLS.IDPL_ITEMS',SelectedSearchRows);

   _arrSearchPosition:= nil;
   SetLength(_arrSearchPosition,Length(SelectedSearchRows),1);

   for i:=0 to High(SelectedSearchRows) do
       _arrSearchPosition[i,0]:= SelectedSearchRows[i];

   //_arrSearchPosition := fBase.SQLReadArr('SELECT PLS.IDPL_ITEMS,PLS.SCOD FROM PL_SCODS PLS WHERE ('+_WhereSearchString+') ORDER BY SCOD');
   //_arrSearchPosition := fBase.SQLReadArr('PL_ITEMS',['ID','NAME','IDOWNER','VENDORCODE','SCOD','LABEL'],'('+_WhereSearchString+')','NAME');

    cnt:= High(_arrSearchPosition);
    pb1.Position:=0;
    pb1.Max:=cnt+1;
    pb1.Step:=1;
 //   pb1.Step:=1;


// вызываем поток
   StartSearchThread(cnt, fBase, _arrSearchPosition, trbMatchLevel.Position, pb1, fGridMatchingAdd, lbStatus, tbAutoSearchMatching,1);

   _arrSearchPosition:= nil;
end;

procedure TFmMatchingAdd.mAutoSearchMatchingClick(Sender: TObject);
var
   _WhereSearchString: String;
  _arrSearchPosition: ArrayOfArrayVariant;
  cnt: integer;
begin
   if tbAutoSearchMatching.ImageIndex = 3 then
   begin
     if MessageDlg('Поиск уже запущен. Отменить?',mtConfirmation, mbOKCancel, 0) = mrOK then
          SearchThread.Terminate;
     exit;
   end;

   SetStatus('Подбор соответствий для выбранной позиции...');

   _WhereSearchString := 'ID='+GridMatchingAdd.DataSource.DataSet.FieldByName('ID').AsString;

   _arrSearchPosition:= nil;
   _arrSearchPosition := fBase.SQLReadArr('PL_ITEMS',['ID','NAME','IDOWNER','VENDORCODE'],' ('+_WhereSearchString+')','NAME');

    cnt:= High(_arrSearchPosition);
    pb1.Position:=0;
    pb1.Max:=cnt+1;
    pb1.Step:=1;


// вызываем поток
   StartSearchThread(cnt, fBase, _arrSearchPosition, trbMatchLevel.Position, pb1, fGridMatchingAdd, lbStatus, tbAutoSearchMatching,0);

end;

procedure TFmMatchingAdd.mAutoSearchMatchingSelectedClick(Sender: TObject);
var
 cnt: integer;
   _arrSearchPosition: ArrayOfArrayVariant;
   _WhereSearchString: string;

begin

 if tbAutoSearchMatching.ImageIndex = 3 then
   begin
     if MessageDlg('Поиск уже запущен. Отменить?',mtConfirmation, mbOKCancel, 0) = mrOK then
          SearchThread.Terminate;
     exit;
   end;


  SelectedSearchRows:= fGridMatchingAdd.SelectedRows;

  if Length(SelectedSearchRows)>1 then
        if MessageDlg('Подобрать соотвествия для выбранных позиций автоматически? Это может занять несколько минут, в зависимости от количества позиций и объема каталога.',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

   SetStatus('Подбор соответствий для выбранных позиций...');

   _WhereSearchString := fBase.PrepareWhereString('ID',SelectedSearchRows);

   _arrSearchPosition:= nil;
   _arrSearchPosition := fBase.SQLReadArr('PL_ITEMS',['ID','NAME','IDOWNER','VENDORCODE'],'('+_WhereSearchString+')','NAME');

    cnt:= High(_arrSearchPosition);
    pb1.Position:=0;
    pb1.Max:=cnt+1;
    pb1.Step:=1;
 //   pb1.Step:=1;


// вызываем поток
   StartSearchThread(cnt, fBase, _arrSearchPosition, trbMatchLevel.Position, pb1, fGridMatchingAdd, lbStatus, tbAutoSearchMatching,0);

   _arrSearchPosition:= nil;
end;

procedure TFmMatchingAdd.mClearSearchResultClick(Sender: TObject);
var
  _Where: String;
  _SelectedRowsArr: ArrayOfInteger;
  _BookMark: TBookMark;
  i, _IDPriceListIndex: Integer;
begin
  if fGridMatchingAdd.SelectedRowsCount >0 then
        if MessageDlg('Сбросить найденные соответствия для всех выделенных позиций ('+IntToStr(fGridMatchingAdd.SelectedRowsCount)+')?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

  _SelectedRowsArr:= fGridMatchingAdd.SelectedRows;

  _Where:= fBase.PrepareWhereString('IDMATCHPOSITION', _SelectedRowsArr);

  screen.Cursor:= crSQLWait;

  fBase.SQLDelete('W_TMP_TBL_NEUTRALSEARCH',_Where,false);

 fGridMatchingAdd.SelectedRowsClear; // очищаем выделение

  with GridMatchingAdd.DataSource.DataSet do
    begin
      _BookMark:= Bookmark;
      close;
      open;
      if RecordCount>0 then Bookmark:=_BookMark;
    end;
  screen.Cursor:= crDefault;
end;

procedure TFmMatchingAdd.mCatalogAddItemsClick(Sender: TObject);
var
  _FormMass: TFmNomenclatureEditMass;
  _ParentID, _SelectedRowsCount, i: Integer;
  _ParentName, _TimeStamp, _SQL_text, _Where: String;
  _arr: ArrayOfInteger;
  _Unit: TCaption;
  _Target: TComponent;
  _N, _M, _D, _C, _K, _PRICE: Double;
  _BookMark: TBookMark;
begin

  try
    _SelectedRowsCount:= Length(fGridMatchingAdd.SelectedRows);
    if _SelectedRowsCount>1 then
          if MessageDlg('Создать новые позиции каталога для всех выделенных позиций ('+IntToStr(fGridMatchingAdd.SelectedRowsCount)+')? После добавления позиций соответствия будут прописаны автоматически.',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

    _TimeStamp:= DateTimeToStr(now());

  //  if _SelectedRowsCount>1 then
  //     begin
  // множественный выбор

     screen.Cursor:= crSQLWait;

     _ParentID:= fBase.SQLReadArr('CATALOG_GROUP',['ID'],'IDPARENT=0','ID')[0,0];
     _ParentName:= fBase.SQLReadArr('select * from GETPARENTS_GROUP_CATALOG('+IntToStr(_ParentID)+');')[0,0];

     _FormMass:= TFmNomenclatureEditMass.Create(Self);
     _FormMass.Base:= fBase;
     _FormMass.gbGroup.Tag:= _ParentID;
     _FormMass.l_edGroupText.Caption:= _ParentName;

     _FormMass.cbUnit.Checked:= true;
     _FormMass.cbUnit.Enabled:= false;

     _FormMass.cbGroup.Checked:= true;
     _FormMass.cbGroup.Enabled:= false;

     _FormMass.cbPrice.Checked:= true;
     _FormMass.cbPrice.Enabled:= false;

     _FormMass.cbAll.Checked:= true;
     _FormMass.gbChange.Enabled:= false;

     _FormMass.cbMain.Caption:='Добавление в каталог новых позиций ('+IntToStr(_SelectedRowsCount)+')';

     _arr:= nil;
     _arr:= fGridMatchingAdd.SelectedRows;

     try

       _FormMass.ShowModal;
     finally

      if _FormMass.ModalResult = mrOK then
         begin

           SetStatus('Создание новых позиций Ждите...');

             // получаем ID выбранных записей
             //_arr:=nil;
             //_arr:= _DBGridCatalogPrice.SelectedRows;

             try

                     _Unit:= (_FormMass.gbUnit.Controls[1] as TComboBox).Text;

                     _ParentID:= (_FormMass.gbGroup.Controls[2] as TEdit).Tag;


    // перебор компонентов

                  _Target:= _FormMass.FindComponent('FmNomenclatureEdit');

                  for i:=0 to _Target.ComponentCount-1 do
                    if (_Target.Components[i] is TEdit) then
                     begin
                       if ((_Target.Components[i] as TEdit).Name= 'e_PRICE1') then
                         _PRICE:= EditValue(_Target.Components[i] as TEdit);

                       if ((_Target.Components[i] as TEdit).Name= 'e_PN') then
                         _N:= EditValue(_Target.Components[i] as TEdit);

                       if ((_Target.Components[i] as TEdit).Name= 'e_PM') then
                         _M:= EditValue(_Target.Components[i] as TEdit);

                       if ((_Target.Components[i] as TEdit).Name= 'e_PD') then
                         _D:= EditValue(_Target.Components[i] as TEdit);

                       if ((_Target.Components[i] as TEdit).Name= 'e_PC') then
                         _C:= EditValue(_Target.Components[i] as TEdit);

                       if ((_Target.Components[i] as TEdit).Name= 'e_PK') then
                         _K:= EditValue(_Target.Components[i] as TEdit);
                     end;
                    _Target:=nil;

                    try
                    _Where:= fBase.PrepareWhereString('ID',_arr);

                    _SQL_Text:= '';
                    _SQL_Text:= 'insert into "CATALOG" ( '
                    +' VENDORCODE,IDOWNER, NAME, UNIT, LABEL,  '
                    +' IDCTG_GROUP, PRICE, PN, PM, PD, PC, PK, IDUSER, FTIMESTAMP )'
                    +' select ID,'+IntToStr(fIdMainOwner)+', NAME, UNIT, LABEL,  '
                    +' '+IntToStr(_ParentID)+','+FloatToStr(_PRICE)+','+FloatToStr(_N)+','+FloatToStr(_M)+','+FloatToStr(_D)+','+FloatToStr(_C)+','+FloatToStr(_K)+',1,'+QuotedStr(_TimeStamp)+''
                    +' from "PL_ITEMS" '
                    +' where '
                    +' '+_Where+';';

                    fBase.SQLUpdate(_SQL_Text,false);

                    _SQL_Text:= '';
                    _SQL_Text:='insert into "CATALOG_MATCHING" (IDOWNER, IDCATALOG, IDPL_ITEMS, QUANTITYINPACKING, IDUSER, FTIMESTAMP) '
                      +' select PL.IDOWNER, CTG.ID, PL.ID, 1, 1, CTG.FTIMESTAMP from "CATALOG" CTG '
                      +' INNER JOIN "PL_ITEMS" PL ON (PL.ID=CTG.VENDORCODE) '
                      +' WHERE CTG.FTIMESTAMP='+QuotedStr(_TimeStamp)+' ;';

                    fBase.SQLUpdate(_SQL_Text,false);

                    _SQL_Text:= 'UPDATE CATALOG SET VENDORCODE='''' WHERE FTIMESTAMP='+QuotedStr(_TimeStamp)+';';

                    fBase.SQLUpdate(_SQL_Text,false);

                    _SQL_Text:='';
                    _SQL_text:='DELETE FROM W_TMP_TBL_NEUTRALSEARCH WHERE '+_Where+';';

                    fBase.SQLUpdate(_SQL_Text,false);

                    _SQL_Text:= '';

                     except
                       _arr:=nil;
                        raise;
                     end;


               _arr:=nil;

               SetStatus('Создание новых позиций каталога успешно завершено.');

    // общий except операции группового изменения
             except
              fBase.SQLTransactionEnd(false);
              SetStatus('Ошибка группового добавления записей. Операция отменена.');
             end;

           if _FormMass.cbUnselect.Checked then fGridMatchingAdd.SelectedRowsClear;

           _FormMass.Free;

         end;
     end;

  // множественный выбор
   //    end;

    with GridMatchingAdd.DataSource.DataSet do
      begin
        _BookMark:= Bookmark;
        close;
        open;
        if RecordCount>0 then Bookmark:=_BookMark;
      end;
      screen.Cursor:= crDefault;
  except
    on E: Exception do
    begin
        screen.Cursor:= crDefault;
        __Log.SaveLogError(E);
        SetStatus('Сбой добавления позиции в каталог.');
        wLog(Self.Name,'Ошибка: "' + E.Message + '"');
        wLog(Self.Name,'Сбой добавления позиции в каталог.');
        ShowMessage('Ошибка: "' + E.Message + '"');
     end;
  end;

end;

procedure TFmMatchingAdd.mCatalogDetelMatchingClick(Sender: TObject);
var
  _Where: String;
  _SelectedRowsArr: ArrayOfInteger;
  _BookMark: TBookMark;
  i: Integer;
  _arr: ArrayOfArrayVariant;
begin
  if fGridMatchingAdd.SelectedRowsCount>0 then
        if MessageDlg('Удалить соответствия для всех выделенных позиций ('+IntToStr(fGridMatchingAdd.SelectedRowsCount)+')?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

   screen.Cursor:= crSQLWait;

  _SelectedRowsArr:= fGridMatchingAdd.SelectedRows;

  _Where:= Base.PrepareWhereString('ID', _SelectedRowsArr);

  _arr:= Base.SQLReadArr('select ID, IDOWNER from "PL_ITEMS" where '+_Where);

  for i:=0 to High(_arr) do
  begin
    Base.SQLDelete('CATALOG_MATCHING','IDPL_ITEMS='+string(_arr[i,0])+' AND IDOWNER='+string(_arr[i,1]),false);
  end;

 fGridMatchingAdd.SelectedRowsClear; // очищаем выделение

  with GridMatchingAdd.DataSource.DataSet do
    begin
      _BookMark:= Bookmark;
      close;
      open;
      if RecordCount>0 then Bookmark:=_BookMark;
    end;
   screen.Cursor:= crDefault;
end;

procedure TFmMatchingAdd.mMatchingAcceptClick(Sender: TObject);
var
  _SelectedRows: ArrayOfInteger;
  i, _IDOwner, _IDCatalog, iIgnore,
  _QuantityInPacked, _IDPriceList,
  _IDPriceListIndex: Integer;
  _SelectedVendorCode, _TimeStamp: String;
  _arr: ArrayOfArrayVariant;
  _BookMark: TBookMark;
  _TimeStampMaxArr: ArrayOfDateTime;
begin
   _TimeStamp:= DateTimeToStr(Now());
   _TimeStampMaxArr:= GetMaxFTimeStampPricesArr(fBase);
    try
      _SelectedRows:=nil;
       _SelectedRows:= fGridMatchingAdd.SelectedRows;

      if _SelectedRows<> nil then
         begin

         if Length(_SelectedRows)>1 then
            begin
              if MessageDlg('Применить найденные соответствия для всех выделенных позиций ('+IntToStr(fGridMatchingAdd.SelectedRowsCount)+')?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;
            end;
                screen.Cursor:= crSQLWait;

                _arr:=nil;
                _arr:= fBase.SQLReadArr('SELECT '
                  + 'PL.ID as PLID '
                  +' ,IIF(WTMP.ID IS NULL,0,WTMP.ID) AS CATALOGID '
                  +' ,PL.IDOWNER '
                 // +' ,PL.NAME AS PLNAME '
                  //+' ,PL.VENDORCODE AS PLVENDORCODE '

                  +' FROM PL_ITEMS PL '

                  +' LEFT OUTER JOIN ( '
                  +'     SELECT T1.IDOWNER, MIN(T1.ID) AS ID, MIN(T1.NAME) AS NAME, T1.WCS, PLWCSMAX.IDMATCHPOSITION '
                  +'     , MIN(CATALOG.IDCTG_GROUP) AS CIDCTG_GROUP '
                  +'    FROM W_TMP_TBL_NEUTRALSEARCH T1 '
                  +'     INNER JOIN   (SELECT IDMATCHPOSITION, IDOWNER, MAX(WCS) MWCS FROM W_TMP_TBL_NEUTRALSEARCH  WHERE WCS>'+IntToStr(MinWCS)+'  GROUP BY 1,2 '
                  +'     ) PLWCSMAX ON '
                  +'     (T1.WCS=PLWCSMAX.MWCS AND T1.IDMATCHPOSITION = PLWCSMAX.IDMATCHPOSITION AND T1.IDOWNER=PLWCSMAX.IDOWNER) '
                  +'     LEFT OUTER JOIN "CATALOG" ON (CATALOG.ID = T1.ID)  GROUP BY 1,4,5 '
                  +' ) WTMP ON (WTMP.IDMATCHPOSITION=PL.ID) '
                  +' WHERE ('+ fBase.PrepareWhereString('PL.ID',_SelectedRows) +') '
                  +';');
            iIgnore:=0;
            for i:=0 to High(_arr) do
            begin

             if integer(_arr[i,0])>0 then
                begin
                    _IDPriceList:= integer(_arr[i,0]);
                    _IDCatalog:= integer(_arr[i,1]);

                if _IDCatalog=0 then
                  begin
                     Inc(iIgnore);
                  end else
                  begin
                     _IDOwner:= integer(_arr[i,2]);
                     //_SelectedVendorCode:= string(_arr[i,3]);
                     _QuantityInPacked:= 1;


                     if not fBase.SQLUpdate('UPDATE OR INSERT INTO "CATALOG_MATCHING" '
                             +' (IDOWNER, IDCATALOG, IDPL_ITEMS, QUANTITYINPACKING, FTIMESTAMP, IDUSER) '
                             +' VALUES ('+IntToStr(_IDOwner)+','+IntToStr(_IDCatalog)+','+IntTOStr(_IDPriceList)+','+FloatToStr(_QuantityInPacked)+','+QuotedStr(_TimeStamp)+',1) MATCHING (IDPL_ITEMS,IDCATALOG);',false) then
                    begin
                      SetStatus('Изменение соответствия завершено с ошибкой.');
                      wLog('Catalog','Изменение соответствия завершено с ошибкой.');
                    end else
                    begin
                     fbase.SQLDelete('W_TMP_TBL_NEUTRALSEARCH','IDMATCHPOSITION='+IntToStr(_SelectedRows[i])+' AND IDOWNER='+IntToStr(_IDOwner),false);

                     _IDPriceListIndex:=-1;
                     _IDPriceListIndex:= fGridMatchingAdd.SelectedRowsListIndexOf(_IDPriceList);
                     if _IDPriceListIndex>-1 then fGridMatchingAdd.SelectedRowsList.Delete(_IDPriceListIndex);

                    end;
                  end;
                end;
           end;

         if iIgnore>0 then ShowMessage('Для '+IntToStr(iIgnore)+' позиций прайс-листа не было обнаружено найденных автоматическим поиском позиций каталога. Возможно процедура была вызвана ошибочно.');

           _BookMark:= GridMatchingAdd.DataSource.DataSet.Bookmark;
           GridMatchingAdd.DataSource.DataSet.Close;
           GridMatchingAdd.DataSource.DataSet.Open;
           if GridMatchingAdd.DataSource.DataSet.RecordCount>0 then
            GridMatchingAdd.DataSource.DataSet.Bookmark:= _BookMark;

         end;
         _SelectedRows:= nil;
         _arr:= nil;
         screen.Cursor:= crDefault;
    except
      on E: Exception do
      begin
          screen.Cursor:= crDefault;
          __Log.SaveLogError(E);
          SetStatus('Сбой изменения соответствия.');
          wLog(Self.Name,'Ошибка: "' + E.Message + '"');
          wLog(Self.Name,'Сбой изменения соответствия.');
          ShowMessage('Ошибка: "' + E.Message + '"');
       end;
    end;
end;

procedure TFmMatchingAdd.mMatchingEditClick(Sender: TObject);
var
  _SelectedID, _IDOwner, _IDCatalog, _IDPrice: Integer;
  _TimeStamp, _SelectedName, _SelectedScod, _SelectedLabel,
    _SelectedVendorCode, _SelectedOwnerName,
    _SelectedOwnerNomenclatureName: String;
  _GridMatchingDataset: TDataSet;
  _SelectedIDCatalog, _SelectedIDCatalogGroup, _SelectedIDOwner,
    _IDOwnerOld, _SelectedIDPrice: LongInt;
  _SelectedQuantInPacked: Double;
  _Form: TFmMatchingEdit;
  _QuantityInPacked: Extended;
  _arr: ArrayOfArrayVariant;
  _BookMark: TBookMark;
begin
 // _SelectedID:= 0;
  _TimeStamp:= DateTimeToStr(Now());
  _SelectedQuantInPacked:=0;
  _IDCatalog:= 0;

  SetStatus('Изменение соответствия...');
  wLog('Catalog','Изменение соответствия...');

//  if TreeGroupOwner.Selected.Text = '' then exit;

  _GridMatchingDataset:= GridMatchingAdd.DataSource.DataSet;

  if _GridMatchingDataset.RecordCount>0 then
     begin
       _SelectedIDPrice:= _GridMatchingDataset.FieldByName('ID').AsInteger;
       _SelectedIDCatalog:= _GridMatchingDataset.FieldByName('WCATALOGID').AsInteger;
       _SelectedIDCatalogGroup:= _GridMatchingDataset.FieldByName('CIDCTG_GROUP').AsInteger;
       _SelectedIDOwner:= _GridMatchingDataset.FieldByName('IDOWNER').AsInteger;
       _SelectedName:= _GridMatchingDataset.FieldByName('CNAME').AsString;
       _SelectedQuantInPacked:= _GridMatchingDataset.FieldByName('QUANTITYINPACKING').AsFloat;

       _arr:= nil;
       _arr:= fBase.SQLReadArr('CATALOG',['LABEL'],'ID='+IntToStr(_SelectedIDCatalog),'');
       if _arr<>nil then
          begin
            _SelectedLabel:=string(_arr[0,0]);
          end else
          begin
           _SelectedLabel:= '';
          end;

        _arr:=nil;
        _arr:= fBase.SQLReadArr('SELECT VSCOD FROM CTG_GET_SCOD('+IntToStr(_SelectedIDCatalog)+',true)');
        if Assigned(_arr) and (_arr[0,0]<>null) then
                  _SelectedScod:= _arr[0,0] else
                  _SelectedScod:= '';

       _arr:= nil;

       _SelectedVendorCode:= _GridMatchingDataset.FieldByName('PLVENDORCODE').AsString;

       _SelectedOwnerName:= _GridMatchingDataset.FieldByName('OWNERNAME').AsString;
       _SelectedOwnerNomenclatureName:= _GridMatchingDataset.FieldByName('PLNAME').AsString;
     end;

  _IDOwnerOld:= _SelectedIDOwner;

  _Form:= TFmMatchingEdit.Create(Self);

    with _Form do begin
      Base:= fBase;
      IDPrice:= _SelectedIDPrice;
      IDOwner:= _SelectedIDOwner;
      IDCatalog:= _SelectedIDCatalog;
      IDCatalogGroup:= _SelectedIDCatalogGroup;

      //wBase:= _DBase;

      kName.Text:=_SelectedName;
      kScod.Text:=_SelectedScod;
      kLabel.Text:=_SelectedLabel;
      kVendorCode.Text:=_SelectedVendorCode;
      kOwner.Text:=_SelectedOwnerName;
      kOwnerNomenclatureName.Text:=_SelectedOwnerNomenclatureName;
      btnOpenPriceLists.Visible:=false;

      if _SelectedQuantInPacked = 0 then _SelectedQuantInPacked:= 1;

     if _SelectedQuantInPacked<1 then
        begin
          _Form.spQuantInPackLeft.Value:=1;
          _Form.spQuantInPackRight.Value:= _wRNDTO(1/_SelectedQuantInPacked,0);
        end else
        begin
           _Form.spQuantInPackLeft.Value:= _wRNDTO(1*_SelectedQuantInPacked,0);
           _Form.spQuantInPackRight.Value:= 1;
        end;

      end;
      _Form.Caption:= '[Соответствие] -= Редактирование =-';

      try
        _Form.ShowModal;
      finally
          screen.Cursor:= crSQLWait;
          if _Form.ModalResult  = mrOK then
             begin
               with _Form do
               begin

                 _SelectedVendorCode:= kVendorCode.Text;
                 _IDPrice:= IDPrice;
                 _IDOwner:= IDOwner;
                 _IDCatalog:= IDCatalog;
                 _QuantityInPacked:= _Form.spQuantInPackLeft.Value/_Form.spQuantInPackRight.Value;
               end;

             if (_IDCatalog<>0) and (_IDPrice<>0) then
                begin
                if fBase.SQLInsert('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],[_IDOwner,_IDCatalog,_IDPrice,_QuantityInPacked,_TimeStamp,integer(1)],'IDPL_ITEMS,IDCATALOG',false)=-1 then
                 begin
                   SetStatus('Изменение соответствия завершено с ошибкой.');
                   wLog('Catalog','Изменение соответствия завершено с ошибкой.');
                 end else
                 begin
                  fbase.SQLDelete('W_TMP_TBL_NEUTRALSEARCH','IDMATCHPOSITION='+IntToStr(_SelectedIDPrice)+' AND IDOWNER='+IntToStr(_IDOwner),false);

                  _BookMark:= GridMatchingAdd.DataSource.DataSet.Bookmark;
                  GridMatchingAdd.DataSource.DataSet.Close;
                  GridMatchingAdd.DataSource.DataSet.Open;
                  if GridMatchingAdd.DataSource.DataSet.RecordCount>0 then
                   GridMatchingAdd.DataSource.DataSet.Bookmark:= _BookMark;
                 end;
              end;
           _Form.Free;

         end;

      end;

      screen.Cursor:= crDefault;
end;

procedure TFmMatchingAdd.mMatchingEditListClick(Sender: TObject);
begin
  _mOtherVariantsListClick(Sender);
end;

procedure TFmMatchingAdd.tbAutoSearchMatchingClick(Sender: TObject);
begin
  mBtnSearch.PopUp;
end;

procedure TFmMatchingAdd.XMLPropStorage1SavingProperties(Sender: TObject);
begin
  DBGridClearOrderBy(GridMatchingAdd);
end;

procedure TFmMatchingAdd._OnSelect(Sender: TObject);
begin
  if rbSelected.Checked then _OnSelectChange();
end;

procedure TFmMatchingAdd._OnSelectChange();
var
  _BookMark: TBookMark;
begin
  _BookMark:= GridMatchingAdd.DataSource.DataSet.Bookmark;
  SwitchGrid();
  if GridMatchingAdd.DataSource.DataSet.RecordCount>0 then
   GridMatchingAdd.DataSource.DataSet.Bookmark:= _BookMark;
end;

procedure TFmMatchingAdd.mSelectAllClick(Sender: TObject);
begin
  fGridMatchingAdd.SelectAll:= true;
end;

procedure TFmMatchingAdd.mClearSelectClick(Sender: TObject);
begin
  fGridMatchingAdd.SelectAll:= false;
end;

procedure TFmMatchingAdd._mOtherVariantsClick(Sender: TObject);
var
  _ID: string;
  _BookMark: TBookMark;
begin
  _ID:= ReplaceStr(TMenuItem(Sender).Name,'m','');

  fBase.SQLUpdate('W_TMP_TBL_NEUTRALSEARCH',['WCS'],[Integer(99)],'IDMATCHPOSITION='+GridMatchingAdd.DataSource.DataSet.FieldByName('ID').AsString+' AND WCS=100',false);

  fBase.SQLUpdate('W_TMP_TBL_NEUTRALSEARCH',['WCS'],[Integer(100)],'ID='+_ID+' AND IDMATCHPOSITION='+GridMatchingAdd.DataSource.DataSet.FieldByName('ID').AsString,false);

  _BookMark:= GridMatchingAdd.DataSource.DataSet.Bookmark;
  GridMatchingAdd.DataSource.DataSet.Close;
  GridMatchingAdd.DataSource.DataSet.Open;
  if GridMatchingAdd.DataSource.DataSet.RecordCount>0 then
   GridMatchingAdd.DataSource.DataSet.Bookmark:= _BookMark;

end;

procedure TFmMatchingAdd._mOtherVariantsListClick(Sender: TObject);
var
  _Form: TFmListSelect;
  _BookMark: TBookMark;
  _TimeStamp: String;
  _GridMatchingDataset: TDataSet;
  _IDOwner, _SelectedIDPrice: Integer;
  _IDCatalog, _IDPrice, i: Integer;
  _QuantityInPacked: Double;
  _SelectedRows: ArrayOfInteger;
  _arr: ArrayOfArrayVariant;
  _ScodMenu: TPopupMenu;
  _ScodMenuItem: TMenuItem;
begin
  // изменение одного соответствия
  _TimeStamp:= DateTimeToStr(Now());

   try
     _GridMatchingDataset:= GridMatchingAdd.DataSource.DataSet;
     _SelectedRows:=nil;
     _Form:= TFmListSelect.Create(self);
     _Form.Caption:='Поиск аналога для позиции: ['+_GridMatchingDataset.FieldByName('PLNAME').AsString+']';
     _Form.Base:= fBase;
     _Form.wFormMode:=1; // CATALOG
     _Form.Where:= 'ID<>'+IntToStr(fIdMainOwner);
     _Form.GridList.Options:=_Form.GridList.Options - [dgMultiSelect];
     _Form.wDataSetLocateField:='ID';
     _Form.wDataSetLocateValue:=_GridMatchingDataset.FieldByName('WCATALOGID').AsInteger;
     _Form.wIDTreeItem:=_GridMatchingDataset.FieldByName('CIDCTG_GROUP').AsInteger;
     _QuantityInPacked:=  _GridMatchingDataset.FieldByName('QUANTITYINPACKING').AsFloat;
     _IDPrice:= _GridMatchingDataset.FieldByName('ID').AsInteger;
     _Form.lbQuantity.Visible:= true;
     _Form.spQuantInPackLeft.Visible:= true;
     _Form.lbK.Visible:= true;
     _Form.spQuantInPackRight.Visible:= true;

     _arr:=nil;
     _arr:= fBase.SQLReadArr('PL_ITEMS',['VENDORCODE','NAME','LABEL'],'ID='+IntToStr(_IDPrice),'');
     if Assigned(_arr) then
       begin
         _Form.ListFormInit(
           IntToStr(_IDPrice),
           VarToStr(_arr[0,0]),
           VarToStr(_arr[0,1]),
           VarToStr(_arr[0,2])
         );
       end;

     //_ScodMenu:= _Form.mQSSCode;
     //_ScodMenu.Items.Clear;
     //
     //_arr:= nil;
     //_arr:= fBase.SQLReadArr('SELECT SCOD FROM PL_SCODS WHERE IDPL_ITEMS='+IntToStr(_IDPrice)+' ORDER BY SCOD');
     //if Assigned(_arr) then
     //  for i:=0 to High(_arr) do begin
     //    _ScodMenuItem:= TMenuItem.Create(_ScodMenu);
     //    _ScodMenuItem.Caption:= _arr[i,0];
     //    _ScodMenuItem.OnClick:=@_Form._on_mScodClick;
     //    _ScodMenu.Items.Add(_ScodMenuItem);
     //  end;
     //
     //if Length(_Form.sBtnQSVendorCode.Caption)=0 then _Form.sBtnQSVendorCode.Enabled:= false;
     //if Length(_Form.sBtnQSName.Caption)=0 then _Form.sBtnQSName.Enabled:= false;
     ////if _ScodMenu.Items.Count=0 then _Form.sBtnQSSCode.Enabled:= false;
     //if Length(_Form.sBtnQSLabel.Caption)=0 then _Form.sBtnQSLabel.Enabled:= false;

     if _QuantityInPacked<1 then
        begin
         if _QuantityInPacked = 0 then _QuantityInPacked:= 1;
          _Form.spQuantInPackLeft.Value:=1;
          _Form.spQuantInPackRight.Value:= _wRNDTO(1/_QuantityInPacked,0);
        end else
        begin
           _Form.spQuantInPackLeft.Value:= _wRNDTO(1*_QuantityInPacked,0);
           _Form.spQuantInPackRight.Value:= 1;
        end;

     try
     _Form.ShowModal;
     finally
       if _Form.ModalResult <> mrCancel then
       begin
          _SelectedRows:= _Form.wSelectedRows;
          _QuantityInPacked:= _Form.wQuantityInPacked;
       end;
      _Form.Free;
     end;
     screen.Cursor:= crDefault;
     if _SelectedRows<> nil then
        begin
         _IDOwner:= _GridMatchingDataset.FieldByName('IDOWNER').AsInteger;
         _IDCatalog:= _SelectedRows[0];
         //_SelectedVendorCode:= _GridMatchingDataset.FieldByName('PLVENDORCODE').AsString;

         _SelectedIDPrice:= _GridMatchingDataset.FieldByName('ID').AsInteger;

         if _IDCatalog<>0 then
            begin
               if fBase.SQLInsert('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],[_IDOwner,_IDCatalog,_SelectedIDPrice,_QuantityInPacked,_TimeStamp,integer(1)],'IDPL_ITEMS,IDCATALOG',false)=-1 then
              begin
                SetStatus('Изменение соответствия завершено с ошибкой.');
                wLog('Catalog','Изменение соответствия завершено с ошибкой.');
              end else
              begin
               fbase.SQLDelete('W_TMP_TBL_NEUTRALSEARCH','IDMATCHPOSITION='+IntToStr(_SelectedIDPrice)+' AND IDOWNER='+IntToStr(_IDOwner),false);
              end;
            end else
                ShowMessage('Ничего не выбрано! Соответствие не было изменено.');


          _BookMark:= GridMatchingAdd.DataSource.DataSet.Bookmark;
          GridMatchingAdd.DataSource.DataSet.Close;
          GridMatchingAdd.DataSource.DataSet.Open;
          if GridMatchingAdd.DataSource.DataSet.RecordCount>0 then
           GridMatchingAdd.DataSource.DataSet.Bookmark:= _BookMark;

        end;
        _SelectedRows:= nil;
        screen.Cursor:= crDefault;
   except
     on E: Exception do
     begin
         screen.Cursor:= crSQLWait;
         __Log.SaveLogError(E);
         SetStatus('Сбой изменения соответствия.');
         wLog(Self.Name,'Ошибка: "' + E.Message + '"');
         wLog(Self.Name,'Сбой изменения соответствия.');
         ShowMessage('Ошибка: "' + E.Message + '"');
      end;
   end;
end;

procedure TFmMatchingAdd.mGridPopup(Sender: TObject);
var
  _ImageIndex: integer;
  _arr: ArrayOfArrayVariant;
  i: integer;
  _NewItem: TMenuItem;
  _CatalogID: integer;
begin
  if tbAutoSearchMatching.ImageIndex = 0 then _ImageIndex:=2 else _ImageIndex:=7;

  mAutoSearchMatching.ImageIndex:=_ImageIndex;
  mAutoSearchMatchingSelected.ImageIndex:=_ImageIndex;

  _CatalogID:= GridMatchingAdd.DataSource.DataSet.FieldByName('WCATALOGID').AsInteger;

  if _CatalogID >0 then
  begin
     mOtherVariants.Visible:= true;
    _arr:= fBase.SQLReadArr('select ID, NAME, WCS from "W_TMP_TBL_NEUTRALSEARCH" where IDMATCHPOSITION='+GridMatchingAdd.DataSource.DataSet.FieldByName('ID').AsString+' and ID<>'+IntToStr(_CatalogID)+' order by WCS DESC rows 5;');
     mOtherVariants.Clear;
     _NewItem:=nil;

     _NewItem:= NewItem(GridMatchingAdd.DataSource.DataSet.FieldByName('PLNAME').AsString, 0, False, false, nil, 0, 'm'+GridMatchingAdd.DataSource.DataSet.FieldByName('ID').AsString);
     _NewItem.ImageIndex:=15;
       mOtherVariants.Add(
           _NewItem
       );

       mOtherVariants.Add(
            NewItem('-', 0, False, True, nil, 0, 'mSplitSelectFromList')
        );

    for i:=0 to High(_arr) do
    begin
      _NewItem:= NewItem(string(_arr[i,1]), 0, False, True, @_mOtherVariantsClick, 0, 'm'+string(_arr[i,0]));
      _NewItem.ImageIndex:=-1;
        mOtherVariants.Add(
            _NewItem
        );
    end;

       mOtherVariants.Add(
            NewItem('-', 0, False, True, nil, 0, 'mSplitSelectFromList')
        );

       _NewItem:= NewItem('Выбрать из списка...', 0, False, True, @_mOtherVariantsListClick, 0, 'mSelectFromList');
       _NewItem.ImageIndex:=8;
       mOtherVariants.Add(
            _NewItem
        );

  end else
     mOtherVariants.Visible:= false;

     //mSplitOther.Visible:= mOtherVariants.Visible;

  _arr:= nil;

  if Length(GridMatchingAdd.DataSource.DataSet.FieldByName('WCS').AsString)>0 then
  begin
     mMatchingAccept.Enabled:= true;
     mClearSearchResult.Enabled:= true;

     mCatalogAddItems.Enabled:= false;
     mCatalogDetelMatching.Enabled:= false;
  end else
  begin
     mMatchingAccept.Enabled:= false;
     mClearSearchResult.Enabled:= false;

     mCatalogAddItems.Enabled:= true;
     mCatalogDetelMatching.Enabled:= true;
  end;

end;

procedure TFmMatchingAdd.rbAllChange(Sender: TObject);
begin
  if rbAll.Checked then
       SwitchGrid;
end;

procedure TFmMatchingAdd.rbNoMatchingChange(Sender: TObject);
begin
  if rbNoMatching.Checked then
       SwitchGrid;
end;

procedure TFmMatchingAdd.rbSelectedChange(Sender: TObject);
begin
  if rbSelected.Checked then
       SwitchGrid;
end;

procedure TFmMatchingAdd.rbWithMatchingChange(Sender: TObject);
begin
  if rbWithMatching.Checked then
       SwitchGrid;
end;

procedure TFmMatchingAdd.spMinWCSChange(Sender: TObject);
begin
  if MinWCSEditingDone then
  begin
    MinWCS:= spMinWCS.Value;
    SwitchGrid();
  end;
end;

procedure TFmMatchingAdd.spMinWCSEditingDone(Sender: TObject);
begin
  MinWCSEditingDone:= true;

  if MinWCSEditingDone then
  begin
    MinWCS:= spMinWCS.Value;
    SwitchGrid();
  end;

end;

procedure TFmMatchingAdd.spMinWCSKeyPress(Sender: TObject; var Key: char);
const Digit: Set of Char=['0' .. '9'];
begin
  if (Key in Digit) then MinWCSEditingDone:= false;
end;

procedure TFmMatchingAdd.StartSearchThread(Acnt: integer; ABase: TwBase; AArrSearchPosition: ArrayOfArrayVariant; AMatchLevel: integer;
  AProgressBar: TProgressBar; AGrid: TwDBGrid; ALabel: TLabel; AButton: TToolButton; aMode: integer);
begin
  //SearchThread:=nil;
  SearchThread := TMyThread.Create(True); // Таким способом он не запустится автоматически

  if Assigned(SearchThread.FatalException) then
             raise SearchThread.FatalException;

  with SearchThread do
  begin
    THRcnt:= Acnt;
    THRBase:= ABase;
    THRArrSearchPosition:= AArrSearchPosition;
    THRMatchLevel:= AMatchLevel;
    THRProgressBar:= AProgressBar;
    THRGrid:= AGrid;
    THRLabel:= ALabel;
    THRButton:= AButton;
    THRButton.ImageIndex:=3;
    THRIdMainOwner:= fIdMainOwner;
    THRMode:= aMode; // режим поиска
    Resume;
  end;

end;

procedure TFmMatchingAdd.trbMatchLevelChange(Sender: TObject);
begin

SetStatus('Изменение режима поиска повлияет только на новый цикл поиска.');

end;

procedure TFmMatchingAdd.FormCloseQuery(Sender: TObject;
  var CanClose: boolean);
begin

       if ModalResult = mrCancel then
       begin
         if MessageDlg('Закрыть без сохранения?',mtConfirmation, mbOKCancel, 0) = mrCancel
          then
            CanClose:= false
          else
          begin

            try
              if Assigned(SearchThread) and not SearchThread.Suspended then
                         SearchThread.Terminate;
            finally
               fBase.SQLTransactionEnd(false);
            end;

          end;
       end else
       begin
        try
          if Assigned(SearchThread) and not SearchThread.Suspended then
                     SearchThread.Terminate;
        finally
           fBase.SQLTransactionEnd(true);
        end;
       end;


end;

procedure TFmMatchingAdd.FormClose(Sender: TObject;
  var CloseAction: TCloseAction);
begin

  CloseAction:= caFree;
end;

procedure TFmMatchingAdd.btnOKClick(Sender: TObject);
begin
   Close();
end;

procedure TFmMatchingAdd.btnPriceEditSearchClearClick(Sender: TObject);
begin
   edSearch.Clear;
   edSearch.OnChange(edSearch);
end;

procedure TFmMatchingAdd.btnCancelClick(Sender: TObject);
begin
   Close();
end;

procedure TFmMatchingAdd.SetStatus(_Text: string);
begin
  lbStatus.Caption:=' '+_Text; //wStatus(wFormID,_Text,true);
  Log(_Text);
  Application.ProcessMessages;
end;

procedure TFmMatchingAdd.Log(aText: string);
begin
   wLog(fFormName, aText);
end;


end.

