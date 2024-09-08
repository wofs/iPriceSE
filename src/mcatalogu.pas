unit mCatalogU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, FmCatalogExportU, FmNomenclatureEditMassU, FmNomenclatureEditU,
  LCLIntf, StdCtrls, SysUtils, Controls,
  DB, ComCtrls, Forms, DBGrids, Dialogs, Menus,
  IBDatabase, IBCustomDataSet,
  wLogU, wFuncU, mUtilsU,
  FmListSelectU,
  wBaseU, wDBGridU, wDBTreeU, wReportU, wTProgressU, wTypesU;

type

  { TCatalog }

  TCatalog = class
  private
    fFormName: string;

    fGridCatalog: TwDBGrid;
    fGridCatalogFilterString: string;

    fGridMatching: TwDBGrid;
    fGridMatchingFilterString: string;

    fGridMatchingPosition: TwDBGrid;
    fProgress: TProgress;
    fReport: TwReport;

    fTreeCatalog: TwDBTree;
    FTreeInfo: TTreeView;
    fTreeMatching: TwDBTree;
    fOwnerForm: TObject;
    fIdMainOwner: string;

    fBase: TwBase;
    fUtils: TUtils;

    __PRICE_MAX_FTIMESTAMP_ARR: ArrayOfDateTime;

    procedure EventRegister();

    procedure fTreeCatalog_onSelectionChanged(Sender: TObject);
    procedure fTreeCatalog_onEditLinkPriceClick(Sender: TObject);
    procedure fGridCatalog_onDataChange(Sender: TObject; Field: TField);

    procedure fTreeMatching_onSelectionChanged(Sender: TObject);
    procedure fGridMatching_onDataChange(Sender: TObject; Field: TField);
    procedure fGridCatalog_onSelect(Sender: TObject);
    procedure fGridMatching_onSelect(Sender: TObject);

    procedure Log(aText: string);
    procedure onEndOperation(Sender: TObject);
    procedure OnEventAlert(Sender: TObject; EventName: string;
      EventCount: longint; var CancelAlerts: boolean);
    procedure onStopForce(Sender: TObject);
    procedure onStopForceWiteForEnd(Sender: TObject);
    procedure ReportOnEndThread(Sender: TObject);
    procedure ReportOnProgressInit(const aProgressBarName: TProgressBarName;
      aValue: integer);
    procedure ReportOnProgressUpdate(const aProgressBarName: TProgressBarName;
      aValue: integer);
    procedure SetStatus(aText: string; const aLog: boolean = True);
    // вывод статуса
    function VendorCodeIsExists(aCode: string): boolean;

  public
    property Base: TwBase read fBase write fBase;
    property GridCatalog: TwDBGrid read fGridCatalog write fGridCatalog;
    property GridMatching: TwDBGrid read fGridMatching write fGridMatching;
    property GridMatchingPosition: TwDBGrid
      read fGridMatchingPosition write fGridMatchingPosition;

    property IdMainOwner: string read fIdMainOwner write fIdMainOwner;
    property TreeCatalog: TwDBTree read fTreeCatalog write fTreeCatalog;
    property TreeMatching: TwDBTree read fTreeMatching write fTreeMatching;
    property TreeInfo: TTreeView read FTreeInfo write FTreeInfo;
    property PriceMaxFTImeStampArr: ArrayOfDateTime
      read __PRICE_MAX_FTIMESTAMP_ARR write __PRICE_MAX_FTIMESTAMP_ARR;

    constructor Create(Sender: TObject; aBase: TwBase;
      aGridCatalog, aGridMatching, aGridMatchingPosition: TDBGrid;
      aTreeCatalog, aTreeMatching: TTreeView);
    constructor Create(Sender: TObject; aBase: TwBase; aSilent: boolean);

    destructor Destroy();
    procedure GridCatalogFill(aGroup: ArrayOfInteger);
    procedure GridCatalogFiltered();

    procedure GridMatchingFill(aGroup: ArrayOfInteger);
    procedure GridMatchingFiltered();

    procedure GridCatalogPositionMatchingFill(aGroup: ArrayOfInteger);

    procedure TreeCatalogFill();
    procedure TreeMatchingFill();
    procedure TreeInfoFill(aGrid: TwDBGrid);
    procedure ExportData();
    procedure ItemAdd(const aCatalogTree: TwDBTree; aGridDataset: TDataSet);
    procedure ItemAdd(const aCatalogTree: TwDBTree; aGridDataset: TDataSet;
      aName, aLabel, aScod: string);
    procedure ItemCopy(const aCatalogTree: TwDBTree; awGridCatalog: TwDBGrid);
    procedure ItemEdit(const awGridCatalog: TwDBGrid; const awCatalogTree: TwDBTree);
    procedure ItemDel(const awGridCatalog: TwDBGrid);

    procedure ExportCatalogInSpreadsheet(const uFileName: string = '';
      const uStocks: ArrayOfInteger = nil; const uPrices: ArrayOfInteger = nil;
      const uSilent: boolean = False);
    procedure ExportCatalogInCSV(const aPatch: string = '';
      const aSilent: boolean = False);
  end;

implementation

uses
  pkgCatalogU, FmTreeU;

{ TCatalog }

procedure TCatalog.fTreeCatalog_onSelectionChanged(Sender: TObject);
var
  _TreeView: TTreeView;
begin
  if fTreeCatalog.MoveTheNode or TFmCatalog(fOwnerForm).tbEditMode.Down then exit;

  _TreeView := fTreeCatalog.Tree;


  if (_TreeView.SelectionCount = 0) or (fGridCatalog.Grid = nil) then exit;

  if _TreeView.Items = nil then
  begin
    SetStatus('Ошибка Tree.');
    Log('Ошибка Tree.');
    exit;
  end;
  try


    if (_TreeView.Items.Count = 0) or (_TreeView.SelectionCount = 0) then exit;

    if not fTreeCatalog.FirstFillTree then
    begin

      if (_TreeView.Selected.Level = 0) and TreeCatalog.ShowChildrenItems then
        GridCatalogFill(nil)
      else
        GridCatalogFill(fTreeCatalog.SelectedItems);
    end
    else
    begin
      fTreeCatalog.FirstFillTree := False;
    end;

  except
    on E: Exception do
    begin
      SetStatus('Сбой выбора узла дерева.');
      Log('Ошибка [' + _TreeView.Name + ']: "' + E.Message + '"');
      Log('Сбой выбора узла дерева');
    end;
  end;
end;

procedure TCatalog.fTreeCatalog_onEditLinkPriceClick(Sender: TObject);
var
  _Form: TFmTree;
  _TreeTag, _SelectedID: integer;
  _arr: ArrayOfArrayVariant;
begin
  _arr := fBase.SQLReadArr('select id from PL_GROUP where idowner=' + IdMainOwner);
  if Length(_arr) = 0 then
  begin
    ShowMessage(
      'Отсутствует прайс-лист основного контрагента! Данная операция неприменима.');
    exit;
  end;

  _Form := TFmTree.Create(TComponent(Sender));
  _Form.Base := fBase;
  _Form.Caption :=
    'Выберите группу своего прайс-листа, связанную с изменяемой группой каталога.';
  _Form.Mode := 1;
  _SelectedID := TreeCatalog.SelectedItems[0];
  _arr := fBase.SQLReadArr('CATALOG_GROUP', ['IDPL_GROUP'], 'ID=' + IntToStr(_SelectedID), '');
  if Assigned(_arr) then
    _Form.IdGroup := integer(_arr[0, 0])
  else
    _Form.IdGroup := 0;

  try
    //select id from PL_GROUP where idowner=2

    _Form.ShowModal;
    _TreeTag := _Form.IdGroup;
    if _Form.ModalResult = mrOk then
      fBase.SQLUpdate('CATALOG_GROUP', ['IDPL_GROUP'], [_TreeTag],
        'ID=' + IntToStr(_SelectedID));

  finally
    _Form.Free;
  end;

end;

procedure TCatalog.fGridCatalog_onDataChange(Sender: TObject; Field: TField);
var
  _result: string;
begin
  if fGridCatalog.FillGridNow then exit;

  _result := '';
  _result := fTreeCatalog.BreadCrumbs(
    fGridCatalog.Grid.DataSource.DataSet.FieldByName('IDCTG_GROUP').AsInteger);

  SetStatus('[' + _result + '] ' + fGridCatalog.Grid.DataSource.DataSet.FieldByName(
    'NAME').AsString);  // NAME

  if not TFmCatalog(fOwnerForm).tbCatalogBtnMatch.Marked and
    (fGridCatalog.Grid.DataSource.DataSet.RecordCount <> 0) then
    GridCatalogPositionMatchingFill(
      [fGridCatalog.Grid.DataSource.DataSet.FieldByName('ID').AsInteger]);

  TreeInfoFill(fGridCatalog);
end;

procedure TCatalog.fTreeMatching_onSelectionChanged(Sender: TObject);
var
  _TreeView: TTreeView;
begin
  //   if MoveTheNode then exit;

  _TreeView := fTreeMatching.Tree;


  if (_TreeView.SelectionCount = 0) or (fGridMatching.Grid = nil) then exit;

  if _TreeView.Items = nil then
  begin
    SetStatus('Ошибка Tree.');
    Log('Ошибка Tree.');
    exit;
  end;
  try


    if (_TreeView.Items.Count = 0) or (_TreeView.SelectionCount = 0) then exit;

    if not fTreeMatching.FirstFillTree then
    begin

      if _TreeView.Selected.Level = 0 then
        GridMatchingFill(nil)
      else
        GridMatchingFill(fTreeMatching.SelectedItems);
    end
    else
    begin
      fTreeMatching.FirstFillTree := False;
    end;

  except
    on E: Exception do
    begin
      SetStatus('Сбой выбора узла дерева.');
      Log('Ошибка [' + _TreeView.Name + ']: "' + E.Message + '"');
      Log('Сбой выбора узла дерева');
    end;
  end;
end;

procedure TCatalog.fGridMatching_onDataChange(Sender: TObject; Field: TField);
var
  _result: string;
begin
  if fGridMatching.FillGridNow then exit;

  _result := '';
  _result := fTreeMatching.BreadCrumbs(
    fGridMatching.Grid.DataSource.DataSet.FieldByName('IDOWNER').AsInteger);

  SetStatus('[' + _result + '] ' + fGridMatching.Grid.DataSource.DataSet.FieldByName(
    'CATALOGNAME').AsString);  // NAME
end;

procedure TCatalog.Log(aText: string);
begin
  if __onLog and Assigned(__Log) then
    wLog(fFormName, aText);
end;

procedure TCatalog.SetStatus(aText: string; const aLog: boolean);
begin
  wStatus(fFormName, aText, aLog);
end;

procedure TCatalog.EventRegister();
begin
    if fBase.RegisterEvents(['CATALOG_GROUP_Change', 'INVOCES_Change']) then
       fBase.EventDB.OnEventAlert := @OnEventAlert;
end;

constructor TCatalog.Create(Sender: TObject; aBase: TwBase;
  aGridCatalog, aGridMatching, aGridMatchingPosition: TDBGrid;
  aTreeCatalog, aTreeMatching: TTreeView);
var
  _TreeMenu: TPopupMenu;
  _TreeMenuSpliter, _TreeMenuEditGridItems: TMenuItem;
begin
  try
    fOwnerForm := Sender;
    fUtils := nil;

    fFormName := TFmCatalog(Sender).Name;
    fBase := aBase;
    fIdMainOwner := fBase.ReadSettingByName('setDefaultOwner');
    // считываем настройки - текущий основной прайс-лист
    fGridMatchingFilterString := '';
    // VisualComponents

    //КАТАЛОГ
    fGridCatalog := TwDBGrid.Create(Base, aGridCatalog, '');
    fGridCatalog.MultiSelect := True;
    fGridCatalog.StaticTextSelection := TFmCatalog(fOwnerForm).st_GridCatalogSelect;
    fGridCatalog.SearchEdit := TFmCatalog(fOwnerForm).edPriceSearch;
    fGridCatalog.SearchPreventiveBtn := TFmCatalog(fOwnerForm).btnPricePreventSearch;
    fGridCatalog.SearchSplitStringBtn := TFmCatalog(fOwnerForm).btnPriceSearchSplitString;
    fGridCatalog.SearchEntryArray := ['"CATALOG".NAME', '"CATALOG".LABEL'];
    fGridCatalog.SearchParticleArray :=
      ['"CATALOG".VENDORCODE', '(SELECT VRESULT FROM CTG_TRY_SCOD("CATALOG".ID,''%s'') /*=*/)'];
    fGridCatalog.GroupField := '"CATALOG".IDCTG_GROUP';
    fGridCatalog.onSelect := @fGridCatalog_onSelect;
    fGridCatalog.SortTitleImagesIndex := [2, 3];

    fTreeCatalog := TwDBTree.Create(Base, aTreeCatalog, 'CATALOG_GROUP',
      'IDPARENT, ID', ['IDOWNER', IdMainOwner]);
    fTreeCatalog.Tree.SortType := stText;

    fTreeCatalog.MultiSelect := True;
    fTreeCatalog.Tree.OnSelectionChanged := @fTreeCatalog_onSelectionChanged;
    fTreeCatalog.SetOwner := IdMainOwner;
    fGridCatalog.Tree := fTreeCatalog;

    EventRegister();

    _TreeMenu := fTreeCatalog.PopupMenu;

    _TreeMenuSpliter := TMenuItem.Create(_TreeMenu);
    _TreeMenuSpliter.Caption := '-';
    _TreeMenuSpliter.Enabled := True;
    _TreeMenu.Items.Add(_TreeMenuSpliter);

    _TreeMenuEditGridItems := TMenuItem.Create(_TreeMenu);
    _TreeMenuEditGridItems.Caption := 'Изменить товары группы';
    _TreeMenuEditGridItems.OnClick := @TFmCatalog(fOwnerForm).tbCatalogBtnEditClick;
    _TreeMenu.Items.Add(_TreeMenuEditGridItems);

    _TreeMenuSpliter := TMenuItem.Create(_TreeMenu);
    _TreeMenuSpliter.Caption := '-';
    _TreeMenuSpliter.Enabled := True;
    _TreeMenu.Items.Add(_TreeMenuSpliter);

    _TreeMenuEditGridItems := TMenuItem.Create(_TreeMenu);
    _TreeMenuEditGridItems.Caption :=
      'Изменить связь группы со своим прайс-листом';
    _TreeMenuEditGridItems.OnClick := @fTreeCatalog_onEditLinkPriceClick;
    _TreeMenu.Items.Add(_TreeMenuEditGridItems);

    with _TreeMenu do
    begin
      Images := TFmCatalog(fOwnerForm).ImageList16;
      Items[0].ImageIndex := 0;
      Items[1].ImageIndex := 1;
      Items[2].ImageIndex := 2;
      Items[4].ImageIndex := 1;
      Items[6].ImageIndex := 16;
    end;

    _TreeMenu := nil;

    //СООТВЕТСТВИЯ
    fTreeMatching := TwDBTree.Create(Base, aTreeMatching, 'OWNER', 'IDPARENT,NAME', []);
    //fTreeMatching.PopupMenu.Items.Clear;
    fTreeMatching.Tree.DragMode := dmManual;
    fTreeMatching.Expanded := True;
    fTreeMatching.Tree.OnSelectionChanged := @fTreeMatching_onSelectionChanged;

    //with fTreeMatching.Tree.PopupMenu do
    //begin
    //  Images:= TFmCatalog(fOwnerForm).ImageList16;
    //end;

    fGridMatching := TwDBGrid.Create(Base, aGridMatching, '');
    fGridMatching.MultiSelect := True;
    fGridMatching.SearchEdit := TFmCatalog(fOwnerForm).edMatchSearch;
    fGridMatching.SearchPreventiveBtn := TFmCatalog(fOwnerForm).btnMatchPreventSearch;
    fGridMatching.SearchEntryArray := ['"CATALOG".NAME', 'PL.NAME'];
    fGridMatching.SearchParticleArray :=
      ['"CATALOG".VENDORCODE', 'PL.VENDORCODE',
      '(SELECT VRESULT FROM CTG_TRY_SCOD(CATALOG_MATCHING.IDCATALOG,''%s'') /*=*/)',
      '(SELECT VRESULT FROM PL_TRY_SCOD(CATALOG_MATCHING.IDPL_ITEMS,''%s'') /*=*/)'];
    fGridMatching.GroupField := '"CATALOG_MATCHING".IDOWNER';
    fGridMatching.StaticTextSelection :=
      TFmCatalog(fOwnerForm).st_GridCatalogMarchingSelect;
    fGridMatching.onSelect := @fGridMatching_onSelect;
    fGridMatching.SortTitleImagesIndex := [2, 3];

    fGridMatchingPosition := TwDBGrid.Create(Base, aGridMatchingPosition, '');
    fGridMatchingPosition.MultiSelect := True;
    fGridMatchingPosition.StaticTextSelection :=
      TFmCatalog(fOwnerForm).st_GridCatalogMarchingPositionSelect;
    fGridMatchingPosition.GroupField := 'MTH.IDCATALOG';
    fGridMatchingPosition.SortTitleImagesIndex := [2, 3];

    TFmCatalog(fOwnerForm).pCatalogPositionMatching.Height := 0;

    __PRICE_MAX_FTIMESTAMP_ARR := GetMaxFTimeStampPricesArr(Base);
  except
    raise;
  end;
  //__PRICE_MAX_FTIMESTAMP_ARR:= Base.ReadMaxDateTimeValues('PRICE-LISTS','ID','FTIMESTAMP');
end;

constructor TCatalog.Create(Sender: TObject; aBase: TwBase; aSilent: boolean);
begin
  fBase := aBase;
  fOwnerForm := Sender;
end;

destructor TCatalog.Destroy();
begin
  fGridCatalog.Destroy();
  fTreeCatalog.Destroy();

  fGridMatching.Destroy();
  fTreeMatching.Destroy();

  fGridMatchingPosition.Destroy();
end;

procedure TCatalog.GridCatalogFill(aGroup: ArrayOfInteger);
var
  _SQLText, CatalogVendorCode: string;
begin
  if not Assigned(Base) then exit;

  GridCatalog.GroupArray := aGroup;

  if CatalogVendorCodeAsNumber then
    CatalogVendorCode:= ' CASE WHEN TRIM(CATALOG.VENDORCODE) SIMILAR TO ''[0-9]+'' THEN CAST(CATALOG.VENDORCODE AS BIGINT) ELSE 0 END AS VENDORCODE, '
  else
    CatalogVendorCode:= 'CATALOG.VENDORCODE, ';

  _SQLText := ' SELECT "CATALOG".ID, ' + wfLineEnding + ' /*formula*/ ' +
    wfLineEnding + ' "CATALOG".IDCTG_GROUP, ' + wfLineEnding + ' "CATALOG".NAME AS NAME, ' +
    wfLineEnding + ' "CATALOG".UNIT, ' +
    wfLineEnding + ' (select VSCOD from CTG_GET_SCOD("CATALOG".ID,true)) SCOD, ' +
    wfLineEnding + ' "CATALOG".LABEL, ' + wfLineEnding + ' "CATALOG".PRICE AS PRICE, ' +
    wfLineEnding + CatalogVendorCode +
    wfLineEnding + ' "CATALOG".FCOLOR, ' +
    wfLineEnding + ' "CATALOG".FTIMESTAMP, ' +
    wfLineEnding + ' "CATALOG".FTIMESTAMPCREATED, ' +
    wfLineEnding +
    ' ("PLOUR".STOCK+"PLOUR".STOCK2+"PLOUR".STOCK3+"PLOUR".STOCK4+"PLOUR".STOCK5) AS STOCK, '
    +
    wfLineEnding + ' "PLOUR".STOCK AS STOCK1, ' +
    wfLineEnding + ' "PLOUR".STOCK2 AS STOCK2, ' +
    wfLineEnding + ' "PLOUR".STOCK3 AS STOCK3, ' +
    wfLineEnding + ' "PLOUR".STOCK4 AS STOCK4, ' +
    wfLineEnding + ' "PLOUR".STOCK5 AS STOCK5, ' +
    wfLineEnding + ' "CATALOG".PN,"CATALOG".PM,"CATALOG".PD,"CATALOG".PC,"CATALOG".PK, ' +
    wfLineEnding +
    ' CASE WHEN (SELECT * FROM CATALOG_SELECT_MTHRESULT("CATALOG".ID))>0 THEN 1 ELSE 0 END MTHRESULT, '
    + wfLineEnding + '  PLFP.PRICEPL AS PRICEPL, ' +
    wfLineEnding + '  PLFP.PRICEPL2 AS PRICEPL2, ' +
    wfLineEnding + '  PLFP.PRICEPL3 AS PRICEPL3, ' +
    wfLineEnding + '  PLFP.PRICEPL4 AS PRICEPL4, ' +
    wfLineEnding + '  PLFP.PRICEPL5 AS PRICEPL5, ' +
    wfLineEnding + '  PLFP.PRICEPL6 AS PRICEPL6, ' +
    wfLineEnding + '  PLFP.PRICEPL7 AS PRICEPL7, ' +
    wfLineEnding + '  PLFP.PRICEPL8 AS PRICEPL8, ' +
    wfLineEnding + '  PLFP.PRICEPL9 AS PRICEPL9, ' +
    wfLineEnding + '  PLFP.PRICEPL10 AS PRICEPL10, ' +
    wfLineEnding + '  PLFP.PDATE AS PDATE, ' +
    wfLineEnding + ' PLOUR.PRICECALC AS PRICEOUR, ' +
    wfLineEnding + ' PLOUR.PRICECALC2 AS PRICEOUR2, ' +
    wfLineEnding + ' PLOUR.PRICECALC3 AS PRICEOUR3, ' +
    wfLineEnding + ' PLOUR.PRICECALC4 AS PRICEOUR4, ' +
    wfLineEnding + ' PLOUR.PRICECALC5 AS PRICEOUR5, ' +
    wfLineEnding + ' PLOUR.PRICECALC6 AS PRICEOUR6, ' +
    wfLineEnding + ' PLOUR.PRICECALC7 AS PRICEOUR7, ' +
    wfLineEnding + ' PLOUR.PRICECALC8 AS PRICEOUR8, ' +
    wfLineEnding + ' PLOUR.PRICECALC9 AS PRICEOUR9, ' +
    wfLineEnding + ' PLOUR.PRICECALC10 AS PRICEOUR10 ' +
    wfLineEnding + '  FROM "CATALOG" ' +
    wfLineEnding + ' LEFT JOIN CATALOG_PL_MIN_PRICE("CATALOG".ID) PLFP ON (1=1)' +
    wfLineEnding + '  LEFT OUTER JOIN "PL_ITEMS" PLOUR ON ( ' +
    wfLineEnding +
    '  "CATALOG".VENDORCODE = PLOUR.VENDORCODE AND "CATALOG".IDOWNER = PLOUR.IDOWNER) ';

  if TFmCatalog(fOwnerForm).sBtnDoublesView.Down then
    _SQLText := _SQLText +
      ' INNER JOIN (select IDOWNER,NAME from CATALOG group by IDOWNER,NAME having count(*)>1) CTGDBL ON (CTGDBL.NAME="CATALOG".NAME AND CTGDBL.IDOWNER="CATALOG".IDOWNER) ';

  _SQLText := _SQLText + '  WHERE ("CATALOG".IDOWNER=' + fIdMainOwner + ') ' +
    fGridCatalogFilterString + ' /*and_group_string*/ /*and_search_string*/ ';

  //_PL_FTIMESTAMP:= Base.PrepareWhereStringFromDateTime('PL.FTIMESTAMP',__PRICE_MAX_FTIMESTAMP_ARR);
  //_PLOUR_FTIMESTAMP:= Base.PrepareWhereStringFromDateTime('PLOUR.FTIMESTAMP',__PRICE_MAX_FTIMESTAMP_ARR);

  //if (Length(__PRICE_MAX_FTIMESTAMP_ARR)>0) and (IdMainOwner<>'0') and (Length(IdMainOwner)>0) then
  GridCatalog.Fill(_SQLText);

  if Assigned(GridCatalog.Grid.DataSource) then
    GridCatalog.Grid.DataSource.OnDataChange := @fGridCatalog_onDataChange;

  if not TFmCatalog(fOwnerForm).tbCatalogBtnMatch.Marked and
    (fGridCatalog.Grid.DataSource.DataSet.RecordCount <> 0) then
    GridCatalogPositionMatchingFill(
      [fGridCatalog.Grid.DataSource.DataSet.FieldByName('ID').AsInteger]);

  TreeInfoFill(fGridCatalog);
end;

procedure TCatalog.fGridCatalog_onSelect(Sender: TObject);
begin
  if not Assigned(fGridCatalog.Grid.DataSource) then exit;

  if TFmCatalog(fOwnerForm).sBtnSelected.Down then GridCatalogFiltered();
end;

procedure TCatalog.GridCatalogFiltered();
var
  _arr: ArrayOfInteger;
  _BookMark: TBookMark;
begin
  fGridCatalogFilterString := '';


  if TFmCatalog(fOwnerForm).sBtnStockOnly.Down then
  begin
    if Length(fGridCatalogFilterString) > 0 then
      fGridCatalogFilterString := fGridCatalogFilterString +
        ' AND ("PLOUR".STOCK>0 OR "PLOUR".STOCK2>0 OR "PLOUR".STOCK3>0 OR "PLOUR".STOCK4>0 OR "PLOUR".STOCK5>0)'
    else
      fGridCatalogFilterString :=
        fGridCatalogFilterString +
        ' ("PLOUR".STOCK>0 OR "PLOUR".STOCK2>0 OR "PLOUR".STOCK3>0 OR "PLOUR".STOCK4>0 OR "PLOUR".STOCK5>0) ';
  end;

  if TFmCatalog(fOwnerForm).sBtnWithMatching.Down then
  begin
    if Length(fGridCatalogFilterString) > 0 then
      fGridCatalogFilterString := fGridCatalogFilterString +
        ' AND (SELECT * FROM CATALOG_SELECT_MTHRESULT("CATALOG".ID))>0'
    else
      fGridCatalogFilterString :=
        fGridCatalogFilterString + ' (SELECT * FROM CATALOG_SELECT_MTHRESULT("CATALOG".ID))>0 ';
  end;

  if TFmCatalog(fOwnerForm).sBtnNoMatching.Down then
  begin
    if Length(fGridCatalogFilterString) > 0 then
      fGridCatalogFilterString := fGridCatalogFilterString +
        ' AND (SELECT * FROM CATALOG_SELECT_MTHRESULT("CATALOG".ID))=0'
    else
      fGridCatalogFilterString :=
        fGridCatalogFilterString + ' (SELECT * FROM CATALOG_SELECT_MTHRESULT("CATALOG".ID))=0 ';
  end;

  if TFmCatalog(fOwnerForm).sBtnSelected.Down then
  begin
    _arr := fGridCatalog.SelectedRows();

    if fGridCatalog.SelectedRowsCount > 0 then
    begin
      if Length(fGridCatalogFilterString) > 0 then
        fGridCatalogFilterString :=
          fGridCatalogFilterString + ' AND (' + fBase.PrepareWhereString('"CATALOG".ID', _arr) + ')'
      else
        fGridCatalogFilterString :=
          fGridCatalogFilterString + ' ' + fBase.PrepareWhereString('"CATALOG".ID', _arr);
    end
    else
    begin
      if Length(fGridCatalogFilterString) > 0 then
        fGridCatalogFilterString :=
          fGridCatalogFilterString + ' AND "CATALOG".ID=0 '
      else
        fGridCatalogFilterString :=
          fGridCatalogFilterString + ' "CATALOG".ID=0 ';
    end;

  end;

  if Length(fGridCatalogFilterString) > 0 then
    fGridCatalogFilterString := ' AND (' + fGridCatalogFilterString + ')';

  screen.Cursor := crSQLWait;
  _BookMark := fGridCatalog.Grid.DataSource.DataSet.Bookmark;
  GridCatalogFill(fGridCatalog.GroupArray);

  if fGridCatalog.Grid.DataSource.DataSet.RecordCount > 0 then
  begin
    fGridCatalog.Grid.DataSource.DataSet.Bookmark := _BookMark;
  end;

  screen.Cursor := crDefault;
end;

procedure TCatalog.GridMatchingFill(aGroup: ArrayOfInteger);
var
  _SQLText, _PL_FTIMESTAMP: string;
begin
  if not Assigned(Base) then exit;

  fGridMatching.GroupArray := aGroup;

  _SQLText := ' SELECT CATALOG_MATCHING.ID, ' + wfLineEnding +
    ' CATALOG_MATCHING.IDOWNER, ' + wfLineEnding + ' CATALOG_MATCHING.IDPL_ITEMS, ' +
    wfLineEnding + ' CAST(PL.VENDORCODE AS VARCHAR(300)) AS PLVENDORCODE, ' +
    wfLineEnding + ' OWNER.NAME AS OWNERNAME, ' +
    wfLineEnding + ' CATALOG_MATCHING.QUANTITYINPACKING AS QUANTITYINPACKING, ' +
    wfLineEnding + ' CATALOG_MATCHING.IDCATALOG, ' + wfLineEnding
    //      +' CATALOG_MATCHING.IDOWNER, '
    + ' CATALOG_MATCHING.FTIMESTAMP, ' + wfLineEnding +
    ' CATALOG.NAME AS CATALOGNAME, ' +
    wfLineEnding + ' CATALOG.VENDORCODE AS CATALOGVENDORCODE, ' +
    wfLineEnding + ' CATALOG.IDCTG_GROUP, ' + wfLineEnding
    //+' CATALOG.SCOD AS CATALOGSCOD, '
    + ' CATALOG.LABEL AS CATALOGLABEL, ' + wfLineEnding + ' PL.NAME AS PLNAME, ' +
    wfLineEnding +
    ' (PL.PRICECALC/IIF(CATALOG_MATCHING.QUANTITYINPACKING <>0,CATALOG_MATCHING.QUANTITYINPACKING,1)) AS PRICE  '
    +
    wfLineEnding + ' FROM "CATALOG_MATCHING"  ' +
    wfLineEnding + ' LEFT OUTER JOIN "PL_ITEMS" PL ON (CATALOG_MATCHING.IDPL_ITEMS=PL.ID)   '
    +
    wfLineEnding + ' LEFT JOIN OWNER ON OWNER.ID=CATALOG_MATCHING.IDOWNER   ' +
    wfLineEnding + ' LEFT JOIN CATALOG ON  CATALOG.ID=CATALOG_MATCHING.IDCATALOG  ';
  _SQLText := _SQLText + '  WHERE (0=0) ' + fGridMatchingFilterString +
    ' ' // _PL_FTIMESTAMP
    + ' /*and_group_string*/ /*and_search_string*/ ORDER BY IDCATALOG,CATALOG_MATCHING.IDOWNER ';

  //_PL_FTIMESTAMP:= Base.PrepareWhereStringFromDateTime('PL.FTIMESTAMP',__PRICE_MAX_FTIMESTAMP_ARR);

  fGridMatching.Fill(_SQLText);

  if Assigned(fGridMatching.Grid.DataSource) then
    fGridMatching.Grid.DataSource.OnDataChange := @fGridMatching_onDataChange;

end;

procedure TCatalog.fGridMatching_onSelect(Sender: TObject);
begin
  if TFmCatalog(fOwnerForm).sBtnMatchingSelected.Down then GridMatchingFiltered();
end;

procedure TCatalog.GridMatchingFiltered();
var
  _arr: ArrayOfInteger;
  _BookMark: TBookMark;
begin
  fGridMatchingFilterString := '';


  //if TFmCatalog(fOwnerForm).sBtnMatchingPriceStockOnly.Down then
  //  begin
  //    if Length(fGridCatalogFilterString)>0 then  fGridCatalogFilterString:=fGridCatalogFilterString+' AND "PL".STOCK>0' else
  //        fGridCatalogFilterString:=fGridCatalogFilterString+' "PL".STOCK>0 ';
  //  end;

  if TFmCatalog(fOwnerForm).sBtnMatchingSelected.Down then
  begin
    _arr := fGridMatching.SelectedRows();

    if fGridMatching.SelectedRowsCount > 0 then
    begin
      if Length(fGridMatchingFilterString) > 0 then
        fGridMatchingFilterString :=
          fGridMatchingFilterString + ' AND (' + fBase.PrepareWhereString(
          '"CATALOG_MATCHING".ID', _arr) + ')'
      else
        fGridMatchingFilterString :=
          fGridMatchingFilterString + ' ' + fBase.PrepareWhereString('"CATALOG_MATCHING".ID', _arr);
    end
    else
    begin
      if Length(fGridMatchingFilterString) > 0 then
        fGridMatchingFilterString :=
          fGridMatchingFilterString + ' AND "CATALOG_MATCHING".ID=0 '
      else
        fGridMatchingFilterString :=
          fGridMatchingFilterString + ' "CATALOG_MATCHING".ID=0 ';
    end;

  end;

  if Length(fGridMatchingFilterString) > 0 then
    fGridMatchingFilterString := ' AND (' + fGridMatchingFilterString + ')';

  screen.Cursor := crSQLWait;
  _BookMark := fGridMatching.Grid.DataSource.DataSet.Bookmark;
  GridMatchingFill(fGridMatching.GroupArray);

  if fGridMatching.Grid.DataSource.DataSet.RecordCount > 0 then
  begin
    fGridMatching.Grid.DataSource.DataSet.Bookmark := _BookMark;
  end;

  screen.Cursor := crDefault;
end;

procedure TCatalog.GridCatalogPositionMatchingFill(aGroup: ArrayOfInteger);
var
  _SQLText: string;
begin
  if not Assigned(Base) then exit;

  fGridMatchingPosition.GroupArray := aGroup;

  _SQLText := 'SELECT' + ' MTH.ID, ' + ' MTH.IDPL_ITEMS,' +
    ' PL.LABEL AS PLLABEL,' + ' PL.ID AS PLID,' + ' PL.IDOWNER,'
    + ' PL.VENDORCODE AS PLVENDORCODE,' + ' PL.FCOLOR,' +
    ' CAST(INV.QUANTITY AS INTEGER) QUANTITYIINVOICE, ' + ' OWN.NAME AS OWNERNAME,'
    + ' MTH.QUANTITYINPACKING AS QUANTITYINPACKING,' + ' MTH.IDCATALOG,'
    + ' PL.NAME AS PLNAME,' +
    ' ((PL.STOCK+PL.STOCK2+PL.STOCK3+PL.STOCK4+PL.STOCK5)*MTH.QUANTITYINPACKING) AS STOCK,'
    +
    ' (PL.PRICECALC/MTH.QUANTITYINPACKING) AS PRICE,' + ' PL.FTIMESTAMP,'
    + ' FMTS.STOCKONLYINFO AS STOCKONLYINFO' + ' FROM CATALOG_MATCHING MTH'
    + ' INNER JOIN PL_ITEMS PL ON (PL.ID=MTH.IDPL_ITEMS)' +
    ' LEFT JOIN INVOCES INV ON (PL.ID=INV.IDPL_ITEMS)' +
    ' INNER JOIN "FORMATS" FMTS ON (FMTS.ID= PL.IDFORMATS)' +
    ' INNER JOIN OWNER OWN ON (OWN.ID=PL.IDOWNER)' +
    ' WHERE (FMTS.FCLOSE=0) AND /*group_string*/ ORDER BY PL.PRICE, PL.IDOWNER ASC';


  fGridMatchingPosition.Fill(_SQLText);

end;

procedure TCatalog.TreeCatalogFill();
begin
  fTreeCatalog.Fill();
end;

procedure TCatalog.TreeMatchingFill();
begin
  fTreeMatching.Fill();
end;


procedure TCatalog.TreeInfoFill(aGrid: TwDBGrid);
var
  _arr, _BarcodeArr: ArrayOfArrayVariant;
begin
  if not TreeInfo.Parent.Visible then exit;

  _arr := nil;
  if aGrid.Grid.DataSource.DataSet.RecordCount > 0 then
    _arr := fBase.SQLReadArr('CATALOG', ['REMARK', 'FURL', 'FURLPICTURE'],
      'ID=' + aGrid.Grid.DataSource.DataSet.FieldByName('ID').AsString, '');

  TreeInfo.Items[0].Text := aGrid.Grid.DataSource.DataSet.FieldByName(
    'FTIMESTAMP').AsString;
  TreeInfo.Items[1].Text := aGrid.Grid.DataSource.DataSet.FieldByName(
    'VENDORCODE').AsString;
  TreeInfo.Items[2].Text := aGrid.Grid.DataSource.DataSet.FieldByName('NAME').AsString;
  TreeInfo.Items[3].Text := aGrid.Grid.DataSource.DataSet.FieldByName('UNIT').AsString;

  _BarcodeArr := nil;
  if Assigned(_arr) then
    _BarcodeArr := fBase.SQLReadArr('SELECT VSCOD FROM CTG_GET_SCOD(' +
      aGrid.Grid.DataSource.DataSet.FieldByName('ID').AsString + ',true)');
  if Assigned(_BarcodeArr) then
    TreeInfo.Items[4].Text := VarToStr(_BarcodeArr[0, 0]);

  //TreeInfo.Items[4].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('SCOD').AsString;
  TreeInfo.Items[5].Text := aGrid.Grid.DataSource.DataSet.FieldByName('LABEL').AsString;

  if Assigned(_arr) then
  begin
    TreeInfo.Items[6].Text := VarToStr(_arr[0, 0]);
    TreeInfo.Items[7].Text := VarToStr(_arr[0, 1]); //11
    TreeInfo.Items[8].Text := VarToStr(_arr[0, 2]); //12
  end
  else
  begin
    TreeInfo.Items[6].Text := '';
    TreeInfo.Items[7].Text := '';            //11
    TreeInfo.Items[8].Text := '';     //12
  end;

  TreeInfo.Items[9].Text := aGrid.Grid.DataSource.DataSet.FieldByName('FCOLOR').AsString;

end;

procedure TCatalog.onEndOperation(Sender: TObject);
begin
  fUtils.Destroy();
  fUtils := nil;
end;

procedure TCatalog.OnEventAlert(Sender: TObject; EventName: string;
  EventCount: longint; var CancelAlerts: boolean);
var
  _IdTmp: integer;
  aSelectedItems: ArrayOfInteger;
  fBookmark: TBookMark;
begin
  case EventName of
    'CATALOG_GROUP_Change':
    begin
      with TreeCatalog do
      begin
        if EventBlock then
        begin
          CancelAlerts := True;
          Exit;
        end;
        aSelectedItems := SelectedItems;
        if Assigned(aSelectedItems) then
          _IdTmp := aSelectedItems[0]
        else
          _IdTmp := 0;

        Fill();
        FindNodeWithDataInt(_IdTmp);
      end;
    end;
    'INVOCES_Change':
    begin
      if GridMatchingPosition.Grid.DataSource.DataSet.RecordCount = 0 then exit;
      with GridMatchingPosition do
      begin
        fBookmark:= GridMatchingPosition.Bookmark;
        GridMatchingPosition.Fill();
        GridMatchingPosition.Bookmark:= fBookmark;
      end;
    end;
  end;
end;

procedure TCatalog.onStopForce(Sender: TObject);
begin
  fReport.Stop();
end;

procedure TCatalog.onStopForceWiteForEnd(Sender: TObject);
begin
  ShowMessage('Дождитесь окончания операции!');
end;

procedure TCatalog.ReportOnEndThread(Sender: TObject);
begin
  try
    fProgress.NoClose := False;
    fProgress.Close;
    fProgress.Refresh;
    if not fReport.Result then
      raise Exception.Create(
        'Во время создания отчета произошла ошибка!');

  except
    on E: Exception do
    begin
      MessageDlg(E.Message, mtError, [mbOK], 0);
    end;
  end;
end;

procedure TCatalog.ReportOnProgressInit(const aProgressBarName: TProgressBarName;
  aValue: integer);
begin
  fProgress.InitBar(aProgressBarName, aValue);
end;

procedure TCatalog.ReportOnProgressUpdate(const aProgressBarName: TProgressBarName;
  aValue: integer);
begin
  fProgress.SetBar(aProgressBarName, aValue);
end;

procedure TCatalog.ExportData();
begin
  if Assigned(fUtils) then exit;

  fUtils := TUtils.Create(fOwnerForm, fBase);
  fUtils.onEndOperation := @onEndOperation;
  fUtils.WhereString := ' WHERE (' + fBase.PrepareWhereString(
    'CTGMTH.IDOWNER', TreeMatching.SelectedItems) + ')';


  fUtils.ExportData([eoMatchings], [emTemplate, emSaveFile]);
end;

function TCatalog.VendorCodeIsExists(aCode: string): boolean;
begin
  Result := (Length(aCode) > 0) and
    (fBase.GetRowsCount(Format('SELECT ID FROM CATALOG WHERE VENDORCODE = ''%s''',
    [aCode])) > 0);
end;

procedure TCatalog.ItemAdd(const aCatalogTree: TwDBTree; aGridDataset: TDataSet;
  aName, aLabel, aScod: string);
var
  _K: double;
  _C: double;
  _D: double;
  _M: double;
  _N: double;
  _P: double;
  _TimeStamp: string;
  _Edini: string;
  _Label: string;
  _Scod: string;
  _SelectedName: string;
  _SelectedID: integer;
  _ParentName: string;
  _oldParentID: integer;
  _ParentID: integer;
  _Form: TFmNomenclatureEdit;
  _IdMainOwner: integer;
  _Vendorcode: TCaption;
begin
  if aCatalogTree.Tree.Items.Count = 0 then exit;
  _SelectedID := 0;
  _IdMainOwner := fBase.ReadSettingByName('setDefaultOwner');

  if aCatalogTree.Tree.Selected.Text = '' then exit;

  _TimeStamp := DateTimeToStr(Now());
  _ParentID := aCatalogTree.SelectedItems[0];
  _oldParentID := _ParentID;

  _SelectedName := aName;
  _SelectedID := fBase.SQLInsert('CATALOG', ['IDCTG_GROUP', 'NAME', 'IDOWNER', 'FTIMESTAMP'],
    [_ParentID, _SelectedName, _IdMainOwner, _TimeStamp], False);

  _ParentName := aCatalogTree.BreadCrumbs(_ParentID);

  _Form := TFmNomenclatureEdit.Create(aCatalogTree.Tree);
  with _Form do
  begin
    Base := fBase;
    kName.Text := _SelectedName;
    edGroup.Text := _ParentName;
    edGroup.Tag := _ParentID;
    kNumber.Text := IntToStr(_SelectedID);

    if Length(aScod) > 0 then
      kScod.Text := aScod
    else
      kScod.Text := GenEAN(__SCODPREFIX, '', IntToStr(_SelectedID));

    kArticul.Text := aLabel;

    kVendorCode.Text := '';

    _Form.Caption :=
      '[Номенклатура] -= Создание на основе позиции прайс-листа =-';
    try
      ShowModal;
    finally

      if ModalResult = mrOk then
      begin
        _SelectedName := kName.Text;
        _ParentID := edGroup.Tag;
        _Edini := kEdini.Text;
        _Scod := kScod.Text;
        _Vendorcode := kVendorCode.Text;
        _Label := kArticul.Text;

        _P := EditValue(e_PRICE);
        _N := EditValue(e_PN);
        _M := EditValue(e_PM);
        _D := EditValue(e_PD);
        _C := EditValue(e_PC);
        _K := EditValue(e_PK);
        _TimeStamp := DateTimeToStr(Now());

        if VendorCodeIsExists(_Vendorcode) then
        begin
          ShowMessage(
            'Добавление новой позиции отменено: Код позиции уже существует в каталоге.');
          SetStatus(
            'Добавление новой позиции отменено: Код позиции уже существует в каталоге.');
          wLog('Catalog',
            'Добавление новой позиции отменено: Код позиции уже существует в каталоге');
          fBase.SQLTransactionEnd(False);
          exit;
        end;

        if not fBase.SQLUpdate(
          'CATALOG', ['IDCTG_GROUP', 'NAME', 'UNIT', 'PRICE', 'PN', 'PM', 'PD', 'PC',
          'PK', 'LABEL', 'VENDORCODE', 'FTIMESTAMP'], [_ParentID, _SelectedName,
          _Edini, _P, _N, _M, _D, _C, _K, _Label, _Vendorcode, _TimeStamp], 'ID=' + IntToStr(_SelectedID))
        then
        begin
          SetStatus('Добавление новой позиции завершено с ошибкой.');
          wLog('Catalog',
            'Добавление новой позиции завершено с ошибкой.');
          fBase.SQLTransactionEnd(False);
        end
        else
        begin
          fBase.SQLUpdate('EXECUTE PROCEDURE CTG_SET_SCOD(' + IntToStr(
            _IdMainOwner) + ',' + IntToStr(_SelectedID) + ',' + QuotedStr(_Scod) + ','','')');

          if _oldParentID <> _ParentID then
          begin
            try
              aCatalogTree.FindNodeWithDataInt(_ParentID);
            finally

              if aGridDataset.RecordCount > 0 then
                aGridDataset.Locate('ID', _SelectedID, []);
            end;
          end
          else
          begin
            aGridDataset.Close;
            aGridDataset.Open;
            if aGridDataset.RecordCount > 0 then
              aGridDataset.Locate('ID', _SelectedID, []);
          end;
        end;

      end
      else
      begin
        fBase.SQLTransactionEnd(False);
      end;

      _Form.Free;

    end;

  end;
end;

procedure TCatalog.ItemAdd(const aCatalogTree: TwDBTree; aGridDataset: TDataSet);
var
  _K: double;
  _C: double;
  _D: double;
  _M: double;
  _N: double;
  _P: double;
  _TimeStamp: string;
  _Edini: string;
  _Label: string;
  _Scod: string;
  _SelectedName: string;
  _SelectedID: integer;
  _ParentName: string;
  _oldParentID: integer;
  _ParentID: integer;
  _Form: TFmNomenclatureEdit;
  _IdMainOwner: integer;
  _Vendorcode: TCaption;
begin
  if aCatalogTree.Tree.Items.Count = 0 then exit;
  _SelectedID := 0;
  _IdMainOwner := fBase.ReadSettingByName('setDefaultOwner');

  if aCatalogTree.Tree.Selected.Text = '' then exit;

  _TimeStamp := DateTimeToStr(Now());
  _ParentID := aCatalogTree.SelectedItems[0];
  _oldParentID := _ParentID;

  _SelectedID := fBase.SQLInsert(
    'CATALOG', ['IDCTG_GROUP', 'NAME', 'IDOWNER', 'FTIMESTAMP'],
    [_ParentID, 'Новая позиция', _IdMainOwner, _TimeStamp], False);


  _SelectedName := 'Новая позиция';
  //aGridDataset.FieldByName('Name').AsString;

  _ParentName := aCatalogTree.BreadCrumbs(_ParentID);

  _Form := TFmNomenclatureEdit.Create(aCatalogTree.Tree);
  with _Form do
  begin
    Base := fBase;
    kName.Text := _SelectedName;
    edGroup.Text := _ParentName;
    edGroup.Tag := _ParentID;
    kNumber.Text := IntToStr(_SelectedID);
    kScod.Text := GenEAN(__SCODPREFIX, '', IntToStr(_SelectedID));
    kVendorCode.Text := '';

    _Form.Caption := '[Номенклатура] -= Создание =-';
    try
      ShowModal;
    finally

      if ModalResult = mrOk then
      begin
        _SelectedName := kName.Text;
        _ParentID := edGroup.Tag;
        _Edini := kEdini.Text;
        _Scod := kScod.Text;
        _Vendorcode := kVendorCode.Text;
        _Label := kArticul.Text;

        _P := EditValue(e_PRICE);
        _N := EditValue(e_PN);
        _M := EditValue(e_PM);
        _D := EditValue(e_PD);
        _C := EditValue(e_PC);
        _K := EditValue(e_PK);
        _TimeStamp := DateTimeToStr(Now());

        if VendorCodeIsExists(_Vendorcode) then
        begin
          ShowMessage(
            'Добавление новой позиции отменено: Код позиции уже существует в каталоге.');
          SetStatus(
            'Добавление новой позиции отменено: Код позиции уже существует в каталоге.');
          wLog('Catalog',
            'Добавление новой позиции отменено: Код позиции уже существует в каталоге');
          fBase.SQLTransactionEnd(False);
          exit;
        end;

        if not fBase.SQLUpdate(
          'CATALOG', ['IDCTG_GROUP', 'NAME', 'UNIT', 'PRICE', 'PN', 'PM', 'PD', 'PC',
          'PK', 'LABEL', 'VENDORCODE', 'FTIMESTAMP'], [_ParentID, _SelectedName,
          _Edini, _P, _N, _M, _D, _C, _K, _Label, _Vendorcode, _TimeStamp], 'ID=' + IntToStr(_SelectedID))
        then
        begin
          SetStatus(
            'Добавление новой позиции завершено с ошибкой.');
          wLog('Catalog',
            'Добавление новой позиции завершено с ошибкой.');
        end
        else
        begin
          fBase.SQLUpdate('EXECUTE PROCEDURE CTG_SET_SCOD(' +
            IntToStr(_IdMainOwner) + ',' + IntToStr(_SelectedID) + ',' + QuotedStr(_Scod) + ','','')');

          if _oldParentID <> _ParentID then
          begin
            try
              aCatalogTree.FindNodeWithDataInt(_ParentID);
            finally

              if aGridDataset.RecordCount > 0 then
                aGridDataset.Locate('ID', _SelectedID, []);
            end;
          end
          else
          begin
            aGridDataset.Close;
            aGridDataset.Open;
            if aGridDataset.RecordCount > 0 then
              aGridDataset.Locate('ID', _SelectedID, []);
          end;
        end;

      end
      else
      begin
        fBase.SQLTransactionEnd(False);
      end;

      _Form.Free;

    end;

  end;
end;

procedure TCatalog.ItemCopy(const aCatalogTree: TwDBTree; awGridCatalog: TwDBGrid);
var
  _K: double;
  _C: double;
  _D: double;
  _M: double;
  _N: double;
  _P1: double;
  _TimeStamp: string;
  _Edini: string;
  _Label: string;
  _Scod: string;
  _SelectedName: string;
  _SelectedID: integer;
  _ParentName: string;
  _oldParentID: integer;
  _ParentID, _SelectedCount: integer;
  _Form: TFmNomenclatureEdit;
  _GridDataset: TDataSet;
  _IdMainOwner: integer;
begin
  if aCatalogTree.Tree.Items.Count = 0 then exit;
  if aCatalogTree.Tree.Selected.Text = '' then exit;

  _GridDataset := awGridCatalog.Grid.DataSource.DataSet;

  _SelectedCount := Length(awGridCatalog.SelectedRows());

  _SelectedID := 0;
  _IdMainOwner := fBase.ReadSettingByName('setDefaultOwner');

  if _SelectedCount > 1 then
  begin
    ShowMessage('Для копирования выберите одну позицию!');
    exit;
  end;

  if _GridDataset.RecordCount > 0 then
  begin
    _SelectedName := _GridDataset.FieldByName('Name').AsString;
    _ParentID := _GridDataset.FieldByName('IDCTG_GROUP').AsInteger;
    _oldParentID := _ParentID;
    _ParentName := aCatalogTree.BreadCrumbs(_ParentID);
  end;

  _TimeStamp := DateTimeToStr(Now());

  _SelectedID := fBase.SQLInsert(
    'CATALOG', ['IDCTG_GROUP', 'NAME', 'IDOWNER', 'FTIMESTAMP'],
    [_ParentID, 'Новая позиция', _IdMainOwner, _TimeStamp], False);

  _Form := TFmNomenclatureEdit.Create(aCatalogTree.Tree);
  _Form.Base := fBase;

  with _Form do
  begin
    kName.Text := _SelectedName;
    edGroup.Text := _ParentName;
    edGroup.Tag := _ParentID;
    kNumber.Text := IntToStr(_SelectedID);

    kEdini.Text := _GridDataset.FieldByName('UNIT').AsString;

    _P1 := _GridDataset.FieldByName('PRICE').AsFloat;
    e_PRICE1.Text := _GridDataset.FieldByName('PRICE').AsString;
    Razdelitel(e_PRICE1, 2, False);
    //'PN','PM','PD','PC','PK'
    _N := _GridDataset.FieldByName('PN').AsFloat;
    e_PN.Text := _GridDataset.FieldByName('PN').AsString;
    Razdelitel(e_PN, 2, False);

    _M := _GridDataset.FieldByName('PM').AsFloat;
    e_PM.Text := _GridDataset.FieldByName('PM').AsString;
    Razdelitel(e_PM, 2, False);

    _D := _GridDataset.FieldByName('PD').AsFloat;
    e_PD.Text := _GridDataset.FieldByName('PD').AsString;
    Razdelitel(e_PD, 2, False);

    _C := _GridDataset.FieldByName('PC').AsFloat;
    e_PC.Text := _GridDataset.FieldByName('PC').AsString;
    Razdelitel(e_PC, 2, False);

    _K := _GridDataset.FieldByName('PK').AsFloat;
    e_PK.Text := _GridDataset.FieldByName('PK').AsString;
    Razdelitel(e_PK, 2, False);

    kScod.Text := GenEAN(__SCODPREFIX, '', IntToStr(_SelectedID));

    kArticul.Text := '';
    kVendorCode.Text := '';

    _Form.Caption := '[Номенклатура] -= Копирование =-';

    try
      ShowModal;
    finally
      if ModalResult = mrOk then
      begin
        _SelectedName := kName.Text;
        _ParentID := edGroup.Tag;
        _Edini := kEdini.Text;
        _Scod := kScod.Text;
        _Label := kArticul.Text;

        //_P:= EditValue(e_PRICE);
        _P1 := EditValue(e_PRICE1);
        _N := EditValue(e_PN);
        _M := EditValue(e_PM);
        _D := EditValue(e_PD);
        _C := EditValue(e_PC);
        _K := EditValue(e_PK);
        _TimeStamp := DateTimeToStr(Now());

        if not fBase.SQLUpdate(
          'CATALOG', ['IDCTG_GROUP', 'NAME', 'UNIT', 'PRICE', 'PN', 'PM', 'PD', 'PC',
          'PK', 'LABEL', 'FTIMESTAMP'], [_ParentID, _SelectedName, _Edini, _P1, _N,
          _M, _D, _C, _K, _Label, _TimeStamp], 'ID=' + IntToStr(_SelectedID)) then
        begin
          SetStatus(
            'Добавление новой позиции завершено с ошибкой.');
          wLog('Catalog',
            'Добавление новой позиции завершено с ошибкой.');
        end
        else
        begin
          fBase.SQLUpdate('EXECUTE PROCEDURE CTG_SET_SCOD(' +
            IntToStr(_IdMainOwner) + ',' + IntToStr(_SelectedID) + ',' + QuotedStr(_Scod) + ','','')');

          if _oldParentID <> _ParentID then
          begin
            try
              aCatalogTree.FindNodeWithDataInt(_ParentID);
            finally

              if _GridDataset.RecordCount > 0 then
                _GridDataset.Locate('ID', _SelectedID, []);
            end;
          end
          else
          begin
            _GridDataset.Close;
            _GridDataset.Open;
            if _GridDataset.RecordCount > 0 then
              _GridDataset.Locate('ID', _SelectedID, []);
          end;
        end;

      end
      else
      begin
        fBase.SQLTransactionEnd(False);
      end;

      _Form.Free;

    end;

  end;
end;

procedure TCatalog.ItemEdit(const awGridCatalog: TwDBGrid;
  const awCatalogTree: TwDBTree);
var
  _Barcodes: ArrayOfArrayVariant;
  _MassEdit: boolean;
  _VendorCode: TCaption;
  _arr: ArrayOfInteger;
  i: integer;
  _P1: double;
  _P0: double;
  _S: double;
  _K: double;
  _C: double;
  _D: double;
  _M: double;
  _N: double;
  _P: double;
  _SQL_text: string;
  _Unit: string;
  _Label: string;
  _Scod: string;
  _SelectedRowsCount: integer;
  _SelectedName: string;
  _TimeStamp: string;
  _SelectedID: integer;
  _ParentName: string;
  _oldParentID: integer;
  _ParentID: integer;
  _Target: TComponent;
  _FormMass: TFmNomenclatureEditMass;
  _Form: TFmNomenclatureEdit;
  _GridDataset: TDataSet;
  _IdMainOwner: integer;
begin
  if awCatalogTree.Tree.Items.Count = 0 then exit;
  if awCatalogTree.Tree.Selected.Text = '' then exit;

  _SelectedID := 0;
  _TimeStamp := DateTimeToStr(Now());
  _IdMainOwner := fBase.ReadSettingByName('setDefaultOwner');

  _GridDataset := awGridCatalog.Grid.DataSource.DataSet;

  if awCatalogTree.Tree.Focused then
  begin

    screen.Cursor := crSQLWait;

    awGridCatalog.SelectAll := True;

    screen.Cursor := crDefault;

    _arr := awGridCatalog.SelectedRows;
    _SelectedRowsCount := Length(_arr);

  end
  else
  begin

    _arr := awGridCatalog.SelectedRows;
    _SelectedRowsCount := Length(_arr);
  end;


  if _GridDataset.RecordCount > 0 then
  begin
    _SelectedID := _GridDataset.FieldByName('ID').AsInteger;
    _SelectedName := _GridDataset.FieldByName('Name').AsString;
    _ParentID := _GridDataset.FieldByName('IDCTG_GROUP').AsInteger;
    _oldParentID := _ParentID;
    _ParentName := awCatalogTree.BreadCrumbs(_ParentID);
  end;

  if _SelectedRowsCount > 1 then
  begin
    if MessageDlg('Выбрано ' + IntToStr(_SelectedRowsCount) +
      ' позиций. Изменить их все?', mtConfirmation,
      mbOKCancel, 0) = mrOk then
      _MassEdit := True
    else
    begin
      _MassEdit := False;
      if awCatalogTree.Tree.Focused then
      begin
        awGridCatalog.SelectAll := False;
        exit;
      end;
    end;
  end
  else
    _MassEdit := False;

  if _MassEdit then
  begin
    // множественный выбор
    screen.Cursor := crSQLWait;

    _FormMass := TFmNomenclatureEditMass.Create(awCatalogTree.Tree);
    _FormMass.Base := fBase;
    _FormMass.gbGroup.Tag := _ParentID;
    _FormMass.l_edGroupText.Caption := _ParentName;
    _FormMass.cbMain.Caption :=
      'Изменить позиций: ' + IntToStr(_SelectedRowsCount);

    try
      _FormMass.ShowModal;
      screen.Cursor := crDefault;
    finally

      if _FormMass.ModalResult = mrOk then
      begin
        if _FormMass.cbUnselect.Checked then awGridCatalog.SelectAll := False;

        screen.Cursor := crSQLWait;

        SetStatus('Принятие изменений... Ждите...');

        // получаем ID выбранных записей
        //_arr:=nil;
        //_arr:= _DBGridCatalogPrice.SelectedRows;

        try
          if _FormMass.cbUnit.Checked then
          begin
            _Unit := (_FormMass.gbUnit.Controls[1] as TComboBox).Text;

            try
              for i := 0 to Length(_arr) - 1 do
                Base.SQLUpdate('CATALOG', ['UNIT', 'FTIMESTAMP'],
                  [_Unit, _TimeStamp], 'ID=' + IntToStr(_arr[i]), False);

            except
              raise;
            end;
          end;

          if _FormMass.cbGroup.Checked then
          begin
            _ParentID := (_FormMass.gbGroup.Controls[2] as TEdit).Tag;

            try
              for i := 0 to Length(_arr) - 1 do
                Base.SQLUpdate(
                  'CATALOG', ['IDCTG_GROUP', 'FTIMESTAMP'], [_ParentID, _TimeStamp], 'ID=' +
                  IntToStr(_arr[i]), False);
            except
              raise;
            end;
          end;

          if _FormMass.cbPrice.Checked then
          begin
            // перебор компонентов

            _Target := _FormMass.FindComponent('FmNomenclatureEdit');

            for i := 0 to _Target.ComponentCount - 1 do
              if (_Target.Components[i] is TEdit) then
              begin
                if ((_Target.Components[i] as TEdit).Name = 'e_PRICE') then
                  _P := EditValue(_Target.Components[i] as TEdit);

                if ((_Target.Components[i] as TEdit).Name = 'e_PN') then
                  _N := EditValue(_Target.Components[i] as TEdit);

                if ((_Target.Components[i] as TEdit).Name = 'e_PM') then
                  _M := EditValue(_Target.Components[i] as TEdit);

                if ((_Target.Components[i] as TEdit).Name = 'e_PD') then
                  _D := EditValue(_Target.Components[i] as TEdit);

                if ((_Target.Components[i] as TEdit).Name = 'e_PC') then
                  _C := EditValue(_Target.Components[i] as TEdit);

                if ((_Target.Components[i] as TEdit).Name = 'e_PK') then
                  _K := EditValue(_Target.Components[i] as TEdit);

                if ((_Target.Components[i] as TEdit).Name = 'e_PRICE1') then
                  _P1 := EditValue(_Target.Components[i] as TEdit);
              end;
            _Target := nil;

            _SQL_text := '';

            if _FormMass.cbP1.Checked then
            begin
              if Length(_SQL_text) > 0 then
                _SQL_text := _SQL_text + ', PRICE=' + FloatToStr(_P1)
              else
                _SQL_text := ' PRICE=' + FloatToStr(_P1);
            end;
            if _FormMass.cbN.Checked then
            begin
              if Length(_SQL_text) > 0 then
                _SQL_text := _SQL_text + ', PN=' + FloatToStr(_N)
              else
                _SQL_text := ' PN=' + FloatToStr(_N);
            end;
            if _FormMass.cbM.Checked then
            begin
              if Length(_SQL_text) > 0 then
                _SQL_text := _SQL_text + ', PM=' + FloatToStr(_M)
              else
                _SQL_text := ' PM=' + FloatToStr(_M);
            end;
            if _FormMass.cbD.Checked then
            begin
              if Length(_SQL_text) > 0 then
                _SQL_text := _SQL_text + ', PD=' + FloatToStr(_D)
              else
                _SQL_text := ' PD=' + FloatToStr(_D);
            end;
            if _FormMass.cbC.Checked then
            begin
              if Length(_SQL_text) > 0 then
                _SQL_text := _SQL_text + ', PC=' + FloatToStr(_C)
              else
                _SQL_text := ' PC=' + FloatToStr(_C);
            end;
            if _FormMass.cbK.Checked then
            begin
              if Length(_SQL_text) > 0 then
                _SQL_text := _SQL_text + ', PK=' + FloatToStr(_K)
              else
                _SQL_text := ' PK=' + FloatToStr(_K);
            end;

            if Length(_SQL_text) > 0 then
              _SQL_text := _SQL_text + ', FTIMESTAMP=''' + _TimeStamp + ''''
            else
              _SQL_text := ' FTIMESTAMP=''' + _TimeStamp + '''';

            //    _SQL:= 'UPDATE "CATALOG" SET '+SQL_text+' WHERE ID='++IntToStr(_arr[i]);



            try
              for i := 0 to Length(_arr) - 1 do
                Base.SQLUpdate('UPDATE "CATALOG" SET ' +
                  _SQL_text + ' WHERE ID=' + IntToStr(_arr[i]) + ';', False);
              //  _DBase.SQLUpdate('CATALOG',['PRICE','PN','PM','PD','PC','PK','FTIMESTAMP'],[_P,_N,_M,_D,_C,_K,_TimeStamp],'ID='+IntToStr(_arr[i]),false)
            except
              _arr := nil;
              screen.Cursor := crDefault;
              raise;
            end;
          end;

          _arr := nil;
          Base.SQLTransactionEnd(True);

          SetStatus('Принятие изменений завершено.');

          // общий except операции группового изменения
        except
          Base.SQLTransactionEnd(False);
          SetStatus(
            'Ошибка группового изменения записей. Операция отменена.');
        end;

        _FormMass.Free;

        if _oldParentID <> _ParentID then
        begin
          try
            awCatalogTree.FindNodeWithDataInt(_ParentID);
          finally
            _GridDataset.DisableControls;
            if _GridDataset.RecordCount > 0 then
              _GridDataset.Locate('ID', _SelectedID, []);
            _GridDataset.EnableControls;
          end;

        end
        else
        begin
          _GridDataset.DisableControls;
          _GridDataset.Close;
          _GridDataset.Open;
          if _GridDataset.RecordCount > 0 then
            _GridDataset.Locate('ID', _SelectedID, []);
          _GridDataset.EnableControls;
          screen.Cursor := crDefault;
        end;

        screen.Cursor := crDefault;
      end
      else
      begin
        awGridCatalog.Grid.Repaint;
      end;
    end;

    // множественный выбор
  end
  else
  begin
    screen.Cursor := crSQLWait;

    _Form := TFmNomenclatureEdit.Create(awCatalogTree.Tree);
    _Form.Base := Base;
    //_Form.PriceMaxFTImeStampArr:= Catalog.PriceMaxFTImeStampArr;

    with _Form do
    begin
      kName.Text := _SelectedName;
      edGroup.Text := _ParentName;
      edGroup.Tag := _ParentID;
      kEdini.Text := _GridDataset.FieldByName('UNIT').AsString;

      _P := _GridDataset.FieldByName('PRICEPL').AsFloat;
      e_PRICE.Text := _GridDataset.FieldByName('PRICEPL').AsString;
      Razdelitel(e_PRICE, 2, False);
      //'PN','PM','PD','PC','PK'
      _N := _GridDataset.FieldByName('PN').AsFloat;
      e_PN.Text := _GridDataset.FieldByName('PN').AsString;
      Razdelitel(e_PN, 2, False);

      _M := _GridDataset.FieldByName('PM').AsFloat;
      e_PM.Text := _GridDataset.FieldByName('PM').AsString;
      Razdelitel(e_PM, 2, False);

      _D := _GridDataset.FieldByName('PD').AsFloat;
      e_PD.Text := _GridDataset.FieldByName('PD').AsString;
      Razdelitel(e_PD, 2, False);

      _C := _GridDataset.FieldByName('PC').AsFloat;
      e_PC.Text := _GridDataset.FieldByName('PC').AsString;
      Razdelitel(e_PC, 2, False);

      _K := _GridDataset.FieldByName('PK').AsFloat;
      e_PK.Text := _GridDataset.FieldByName('PK').AsString;
      Razdelitel(e_PK, 2, False);

      _P0 := _GridDataset.FieldByName('PRICEOUR').AsFloat;
      e_PRICE0.Text := _GridDataset.FieldByName('PRICEOUR').AsString;
      Razdelitel(e_PRICE0, 2, False);

      _P1 := _GridDataset.FieldByName('PRICE').AsFloat;
      e_PRICE1.Text := _GridDataset.FieldByName('PRICE').AsString;
      Razdelitel(e_PRICE1, 2, False);

      _S := _GridDataset.FieldByName('STOCK').AsFloat;
      e_STOCK.Text := FloatToStrF(_GridDataset.FieldByName('STOCK').AsFloat,
        ffNumber, 4, 0);

      _Barcodes := nil;
      _Barcodes := fBase.SQLReadArr('SELECT VSCOD FROM CTG_GET_SCOD(' +
        _GridDataset.FieldByName('ID').AsString + ',true)');
      if Assigned(_Barcodes) then
        kScod.Text := VarToStr(_Barcodes[0, 0]);

      kNumber.Text := IntToStr(_SelectedID);
      //_GridDataset.FieldByName('ID').AsString;
      kArticul.Text := _GridDataset.FieldByName('LABEL').AsString;
      kVendorCode.Text := _GridDataset.FieldByName('VENDORCODE').AsString;

    end;
    _Form.Caption :=
      '[Номенклатура] -= Редактирование =-';

    try
      _Form.ShowModal;

      screen.Cursor := crDefault;
    finally

      if _Form.ModalResult = mrOk then
      begin

        screen.Cursor := crSQLWait;

        with _Form do
        begin
          _SelectedName := kName.Text;
          _ParentID := edGroup.Tag;
          _Unit := kEdini.Text;
          _Scod := kScod.Text;
          _Label := kArticul.Text;
          _VendorCode := kVendorCode.Text;

          _P1 := EditValue(e_PRICE1);
          _N := EditValue(e_PN);
          _M := EditValue(e_PM);
          _D := EditValue(e_PD);
          _C := EditValue(e_PC);
          _K := EditValue(e_PK);
          // _S:= 0;
        end;

        if not
          Base.SQLUpdate('CATALOG', ['IDCTG_GROUP', 'NAME', 'UNIT', 'PRICE', 'PN',
          'PM', 'PD', 'PC', 'PK', 'LABEL', 'VENDORCODE', 'FTIMESTAMP'],
          [_ParentID, _SelectedName, _Unit, _P1, _N, _M, _D, _C, _K, _Label, _VendorCode, _TimeStamp],
          'ID=' + IntToStr(_SelectedID)) then
        begin
          SetStatus(
            'Изменение позиции завершено с ошибкой.');
          wLog('Catalog',
            'Изменение позиции завершено с ошибкой.');
        end
        else
        begin

          fBase.SQLUpdate('EXECUTE PROCEDURE CTG_SET_SCOD(' +
            IntToStr(_IdMainOwner) + ',' + IntToStr(_SelectedID) + ',' + QuotedStr(_Scod) + ','','')');

          if _oldParentID <> _ParentID then
          begin
            try
              awCatalogTree.FindNodeWithDataInt(_ParentID);
            finally
              if _GridDataset.RecordCount > 0 then
                _GridDataset.Locate('ID', _SelectedID, []);
            end;

          end
          else
          begin
            _GridDataset.DisableControls;
            _GridDataset.Close;
            _GridDataset.Open;
            if _GridDataset.RecordCount > 0 then
              _GridDataset.Locate('ID', _SelectedID, []);
            _GridDataset.EnableControls;
          end;
        end;

        screen.Cursor := crDefault;

      end
      else
      begin
        // если не изменено - все равно обновляем - могут быть добаслены соответствия
        _GridDataset.DisableControls;
        _GridDataset.Close;
        _GridDataset.Open;
        if _GridDataset.RecordCount > 0 then
          _GridDataset.Locate('ID', _SelectedID, []);
        _GridDataset.EnableControls;
        screen.Cursor := crDefault;
      end;

      _Form.Free;

    end;
    ////
  end;
end;

procedure TCatalog.ItemDel(const awGridCatalog: TwDBGrid);
var
  _SelCount: integer;
  _Arr: ArrayOfInteger;
  _BookMark: TBookMark;
  i: integer;
  _ID: integer;
  _GridDataset: TDataSet;
begin
  _GridDataset := awGridCatalog.Grid.DataSource.DataSet;

  _Arr := awGridCatalog.SelectedRows;

  _SelCount := Length(_Arr);

  if _SelCount > 1 then
  begin
    if MessageDlg('Удалить несколько позиций (' +
      IntToStr(_SelCount) +
      ') ? При удалении позиции так же будут удалены связанные соответствия!',
      mtConfirmation, mbOKCancel, 0) = mrOk then
    begin

      _BookMark := _GridDataset.Bookmark;

      try
        for i := 0 to Length(_Arr) - 1 do
        begin
          Base.SQLDelete('CATALOG', 'ID=' + IntToStr(_Arr[i]), False);
        end;
      finally
        Base.SQLTransactionEnd(True);
        _Arr := nil;
        with _GridDataset do
        begin
          awGridCatalog.Fill();
          if RecordCount > 0 then BookMark := _BookMark;
        end;
        wLog('Catalog', IntToStr(_SelCount) +
          ' позиций успешно удалено.');
        SetStatus(IntToStr(_SelCount) +
          ' позиций успешно удалено.');
      end;

    end;
  end
  else
  begin
    if MessageDlg('Удалить позицию "' + _GridDataset.FieldByName(
      'Name').AsString +
      '" ? При удалении позиции так же будут удалены связанные соответствия!',
      mtConfirmation, mbOKCancel, 0) = mrOk then
    begin
      _ID := _GridDataset.FieldByName('ID').AsInteger;
      _BookMark := _GridDataset.Bookmark;
      if Base.SQLDelete('CATALOG', 'ID=' + IntToStr(_ID)) then
        wLog('Catalog', 'Позиция успешно удалена');
      SetStatus('Позиция успешно удалена');
      with _GridDataset do
      begin
        awGridCatalog.Fill();
        if RecordCount > 0 then BookMark := _BookMark;
      end;
    end;
  end;
  awGridCatalog.SelectedRowsClear();
end;

procedure TCatalog.ExportCatalogInSpreadsheet(const uFileName: string;
  const uStocks: ArrayOfInteger; const uPrices: ArrayOfInteger; const uSilent: boolean);
var
  FmExport: TFmCatalogExport;
  _DS: TDataSource;
  aPrices: ArrayOfInteger;
  aStocks: ArrayOfInteger;
  aStockOnly: boolean;
  SaveDialog: TSaveDialog;
  aFileExport: string;
begin
  if Length(uFileName) = 0 then
  begin
    SaveDialog := TSaveDialog.Create(nil);
    SaveDialog.Options := [ofOverwritePrompt];

    try

      SaveDialog.Filter :=
        'OpenDocument (*.ods)|*.ods|Excel (*.xls)|*.xls|Excel (*.xlsx)|*.xlsx';
      SaveDialog.FilterIndex := 2;
      SaveDialog.FileName := 'Экспорт каталога';

      if not SaveDialog.Execute then exit;
      aFileExport := SaveDialog.FileName;
    finally
      SaveDialog.Free;
    end;
  end
  else
    aFileExport := uFileName;

  if not uSilent then
  begin
    aPrices := nil;
    aStocks := nil;

    FmExport := TFmCatalogExport.Create(nil);
    _DS := fBase.SQLReadDS('SELECT ID, NAME FROM PRICEFIELD ORDER BY PRIORITY', True);
    lbxFill(FmExport.ListPrices, _DS, ['NAME', 'ID']);

    lbxFill(FmExport.lstStocks, ['Отдел 1', 'Отдел 2',
      'Отдел 3', 'Отдел 4', 'Отдел 5']);

    _DS.DataSet.Close;

    try
      if FmExport.ShowModal = mrCancel then exit;
      aPrices := FmExport.SelectedPrices;
      aStocks := FmExport.SelectedStocks;
      aStockOnly := FmExport.StockOnly;
    finally
      FmExport.Free;
    end;
  end
  else
  begin
    aStockOnly := True;
    aPrices := uPrices;
    aStocks := uStocks;
  end;

  fProgress := TProgress.Create(TForm(fOwnerForm));
  fProgress.Caption :=
    'Экспорт каталога в файл электронной таблицы..';
  fProgress.ShowLog := False;
  fProgress.onStopForce := @onStopForceWiteForEnd;
  fProgress.NoClose := True;
  fProgress.ShowTop := False;

  fReport := TwReport.Create(True);
  fReport.Base := fBase;
  fReport.OwnerForm := TForm(fOwnerForm);
  fReport.ReportModes := rmCatalogExportSpreadSheet;
  fReport.SelectedPriceItems := aPrices;
  fReport.SelectedStockItems := aStocks;
  fReport.FlagBool := aStockOnly;
  fReport.PathToFiles := aFileExport;
  fReport.Template := PathTemplates_Unsafe + 'catalog.xls';
  fReport.WorkbookSource := nil;
  fReport.onProgressInit := @ReportOnProgressInit;
  fReport.onProgressUpdate := @ReportOnProgressUpdate;
  fReport.onEndThread := @ReportOnEndThread;

  screen.Cursor := crSQLWait;

  fReport.start;

  try

    fProgress.ShowModal;

    //fViewer.WorkbookSource:= fWorkbookSource;

    ////TFmOrders(fOwnerForm).Repaint;
    //if fReport.Result then
    //  begin
    //     fViewer.ShowModal;
    //  end;

    if fReport.Result and not uSilent then
      if MessageDlg(
        'Открыть полученный файл в программе просмотра?',
        mtConfirmation, mbOKCancel, 0) = mrOk then
        OpenDocument(aFileExport);

    //fReport.Terminate;
  finally
    screen.Cursor := crDefault;
    if Assigned(fProgress) then
      fProgress.Free;
  end;
end;

procedure TCatalog.ExportCatalogInCSV(const aPatch: string; const aSilent: boolean);
var
  OpenDialog: TSelectDirectoryDialog;
  aFileExport: string;
begin
  if Length(aPatch) = 0 then
  begin
    OpenDialog := TSelectDirectoryDialog.Create(nil);
    try
      if not OpenDialog.Execute then exit;
      aFileExport := OpenDialog.FileName;
    finally
      OpenDialog.Free;
    end;
  end
  else
    aFileExport := aPatch;


  //FmExport.ListPrices.ItemIndex:=0;

  fProgress := TProgress.Create(nil);
  fProgress.Caption := 'Экспорт каталога в CSV...';
  fProgress.ShowLog := False;
  fProgress.onStopForce := @onStopForceWiteForEnd;
  fProgress.NoClose := True;
  fProgress.ShowTop := True;
  fProgress.ShowBottom := False;

  fReport := TwReport.Create(True);

  fReport.Base := fBase;
  fReport.OwnerForm := TForm(fOwnerForm);
  fReport.ReportModes := rmCatalogExportCSV;
  fReport.PathToFiles := aFileExport;
  fReport.WorkbookSource := nil;
  fBase.onProgressInit := @ReportOnProgressInit;
  fBase.onProgressUpdate := @ReportOnProgressUpdate;
  fReport.onEndThread := @ReportOnEndThread;

  fReport.start;

  try

    fProgress.ShowModal;

    //fViewer.WorkbookSource:= fWorkbookSource;

    ////TFmOrders(fOwnerForm).Repaint;
    //if fReport.Result then
    //  begin
    //     fViewer.ShowModal;
    //  end;

    if fReport.Result and not aSilent then
      ShowMessage('Экспорт завершен!');

    //  fReport.Terminate;
  finally
    if Assigned(fProgress) then
      fProgress.Free;
  end;
end;

end.
