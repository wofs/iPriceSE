unit mPriceLists;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, FmInvoceU, fpspreadsheet, fpspreadsheetctrls, fpspreadsheetgrid, fpsTypes, mInvoceU, SysUtils, Controls,
  db, ComCtrls, Forms, DBGrids, Dialogs, Menus,DateUtils, Graphics,
  IBDatabase, IBQuery, IBCustomDataSet, IBSQL,
  wLogU, wFuncU, wReportU, wTProgressU,
  wBaseU, wDBGridU, wDBTreeU, wCustomClassThreadU, wTViewerSpreadsheetU, wTypesU
  ;

type

  { TDeletePriceThread }

  TDeletePriceThread = class(TwCustomThreadWithProgressBar)
    protected
      procedure Execute; override;
    private
      fResult: Boolean;
      fBase: TwBase;
      fStatus: String;

      fOwner: integer;

      procedure SetStatus(aText: string);
    public
      Constructor Create(CreateSuspended : boolean);

      property Status: string read fStatus;
      property Result: boolean read fResult;
      //property ProgressPosition: integer read fProgressPosition write fProgressPosition;
      property Base: TwBase write fBase;
  end;

  { TPrices }

  TPrices = class
    private
      fDeletePriceThread: TDeletePriceThread;
      fInvoce: TInvoce;
      fProgress: TProgress;
      fFormName: string;
      fGridPosition: TwDBGrid;
      fGridPrice: TwDBGrid;
      fGridInvoces: TwDBGrid;

      fGridPriceFilterString: string;

      fOwnerForm: TObject;
      fIdMainOwner: string;

      fBase: TwBase;
      fReport: TwReport;
      fTreeImportOwner: TwDBTree;
      FTreeInfo: TTreeView;
      fTreePriceGroup: TwDBTree;
      fTreePriceOwner: TwDBTree;

      __PRICE_MAX_FTIMESTAMP_ARR: ArrayOfDateTime;
      __FTIMESTAMP_KURS: string;

      procedure fGridPriceViewDataTime_and_Kurs();
      procedure fTreePriceOwner_onSelectionChanged(Sender: TObject);
      procedure fTreePriceGroup_onSelectionChanged(Sender: TObject);

      procedure fGridPrice_onDataChange(Sender: TObject; Field: TField);

      procedure fTreeMatching_onSelectionChanged(Sender: TObject);
      procedure fGridMatching_onDataChange(Sender: TObject; Field: TField);
      procedure GridPrice_onSelect(Sender: TObject);

      procedure Log(aText: string);
      procedure onCbInvoiceChange(Sender: TObject; aValue: integer);
      procedure onEndThread(Sender: TObject);
      procedure onInvoceSumCountChanged(Sender: TObject; aSum: double; aCount: integer);
      procedure onProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
      procedure onProgressStatus(const aValue: string);
      procedure onProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
      procedure onStatusUpdate(Sender: TObject);
      procedure onStopForce(Sender: TObject);
      procedure RepOnEndThread(Sender: TObject);
      procedure RepOnProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
      procedure RepOnProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
      procedure RepOnStopForce(Sender: TObject);
      procedure SetMaxFTimeStampArr(AValue: ArrayOfDateTime);
      procedure SetStatus(aText: string; const aLog: boolean = true); // вывод статуса

    public
      property Base: TwBase read fBase write fBase;

      property GridPrice: TwDBGrid read fGridPrice write fGridPrice;
      property GridPosition: TwDBGrid read fGridPosition write fGridPosition;
      property GridInvoces: TwDBGrid read fGridInvoces write fGridInvoces;

      property TreePriceOwner: TwDBTree read fTreePriceOwner write fTreePriceOwner;
      property TreePriceGroup: TwDBTree read fTreePriceGroup write fTreePriceGroup;
      property TreeImportOwner: TwDBTree read fTreeImportOwner write fTreeImportOwner;
      property TreeInfo: TTreeView read FTreeInfo write FTreeInfo;

      property IdMainOwner: string read fIdMainOwner write fIdMainOwner;

      property PriceMaxFTImeStampArr: ArrayOfDateTime read __PRICE_MAX_FTIMESTAMP_ARR write SetMaxFTimeStampArr;

      constructor Create(Sender: TObject; aBase: TwBase; aGridPrice, aGridPosition, aGridInvoce: TDBGrid; aTreePriceOwner, aTreePriceGroup,
        aTreeImportOwner: TTreeView);
      destructor Destroy();

      procedure GridPriceFill(aGroup: ArrayOfInteger);
      procedure GridPriceFiltered();
      procedure GridPositionFill(aGroup: ArrayOfInteger);
      procedure GridInvocesFill(aGroup: ArrayOfInteger);

      procedure TreePriceOwnerFill();
      procedure TreePriceGroupFill();
      procedure TreeImportOwnerFill();

      procedure TreeInfoFill(aGrid: TwDBGrid);

      procedure GridPositionModeChange();

      procedure DeletePriceList(aOwner: integer);
      procedure DeletePriceGroup;
      procedure DeletePriceItem;

      procedure InvoceAdd();
      procedure InvoceDel(aInvoces: ArrayOfInteger);
      procedure InvoceEdit(aId: integer);
      procedure InvoceChangeItem(aIdPL, aIdInvoce: integer);

      procedure mAnalogsFill(Sender: TObject; amAnalogs: TMenuItem);
      procedure PrintDateTimePrices;
end;

implementation
uses
  pkgPricesU;

{ TDeletePriceThread }

procedure TDeletePriceThread.Execute;
begin
  try
    fResult:= false;

    fBase.LongTransaction:= true;

    ProgressInit(pbBottom,5);

    try
      ProgressUpdate(pbBottom,1);

      SetStatus('Удаление архива...');
      fBase.SQLDelete('PL_VERSIONS','IDOWNER='+IntToStr(fOwner),false);
      ProgressUpdate(pbBottom,2);
      SetStatus('Удаление архива... [OK]');

      if StopForce then
        raise Exception.Create('Операция отменена!');

      SetStatus('Удаление позиций прайс-листа...');
      fBase.SQLDelete('PL_ITEMS','IDOWNER='+IntToStr(fOwner),false);
      ProgressUpdate(pbBottom,3);
      SetStatus('Удаление позиций прайс-листа... [OK]');

      if StopForce then
        raise Exception.Create('Операция отменена!');

      SetStatus('Удаление групп прайс-листа...');
      fBase.SQLDelete('PL_GROUP','IDOWNER='+IntToStr(fOwner)+' AND IDPARENT<>0',false);
      ProgressUpdate(pbBottom,4);
      SetStatus('Удаление групп прайс-листа... [OK]');

      if StopForce then
        raise Exception.Create('Операция отменена!');

      SetStatus('Обновление метаданных..');
      fBase.SQLDelete('PRICELISTS_TIMESTAMPS','IDOWNER='+IntToStr(fOwner),false);
      fBase.SQLUpdate('FORMATS',['FILEHASH'],[''],'IDOWNER='+IntToStr(fOwner),false);
      ProgressUpdate(pbBottom,5);
      SetStatus('Обновление метаданных... [OK]');

      fBase.SQLTransactionEnd(true);
    except
      fBase.SQLTransactionEnd(false);
      raise;
    end;

    fResult:= true;
    SetStatus('Удаление прайс-листа... [ЗАВЕРШЕНО]');
    onEndThread(self);
except
  on E: Exception do begin
    fResult:= false;
    SetStatus('Error: '+E.Message);
    onEndThread(self);
  end;
end;
end;

procedure TDeletePriceThread.SetStatus(aText: string);
begin
  fStatus := aText;
  ProgressStatus(aText);
  onStatusUpdate(Self);
end;

constructor TDeletePriceThread.Create(CreateSuspended: boolean);
begin
  FreeOnTerminate := true;
  StopForce:= false;

  fBase:= nil;
  inherited Create(CreateSuspended);
end;

{ TPrices }

procedure TPrices.fTreePriceOwner_onSelectionChanged(Sender: TObject);
var
  _TreeView: TTreeView;
begin
  //if fTreeCatalog.MoveTheNode then exit;

  _TreeView:= fTreePriceOwner.Tree;


  if (_TreeView.SelectionCount=0) or (fGridPrice.Grid=nil) then exit;

  if _TreeView.Items = nil then
   begin
     SetStatus('Ошибка Tree.');
     Log('Ошибка Tree.');
     exit;
   end;
  try


 if (_TreeView.Items.Count = 0) or (_TreeView.SelectionCount = 0) then exit;

     if  not fTreePriceOwner.FirstFillTree then
     begin

       __PRICE_MAX_FTIMESTAMP_ARR:= GetMaxFTimeStampPricesArr(Base);

       if _TreeView.Selected.Level=0 then
         GridPriceFill(nil)
       else
         GridPriceFill(fTreePriceOwner.SelectedItems);

         GridInvocesFill(TreePriceOwner.SelectedItems);
     end else
     begin
         fTreePriceOwner.FirstFillTree:= false;
     end;

  except
  on E: Exception do
  begin
      SetStatus('Сбой выбора узла дерева.');
      Log('Ошибка ['+_TreeView.Name+']: "' + E.Message + '"');
      Log('Сбой выбора узла дерева');
   end;
  end;
end;

procedure TPrices.fTreePriceGroup_onSelectionChanged(Sender: TObject);
var
  _TreeView: TTreeView;
begin
  _TreeView:= fTreePriceGroup.Tree;


  if (_TreeView.SelectionCount=0) or (fGridPrice.Grid=nil) then exit;

  if _TreeView.Items = nil then
   begin
     SetStatus('Ошибка Tree.');
     Log('Ошибка Tree.');
     exit;
   end;
  try


 if (_TreeView.Items.Count = 0) or (_TreeView.SelectionCount = 0) then exit;

     if  not fTreePriceGroup.FirstFillTree then
     begin

       if _TreeView.Selected.Level=0 then
        GridPriceFill(nil)
       else
        GridPriceFill(fTreePriceGroup.SelectedItems);
     end else
     begin
         fTreePriceGroup.FirstFillTree:= false;
     end;

  except
  on E: Exception do
  begin
      SetStatus('Сбой выбора узла дерева.');
      Log('Ошибка ['+_TreeView.Name+']: "' + E.Message + '"');
      Log('Сбой выбора узла дерева');
   end;
  end;
end;

procedure TPrices.fGridPriceViewDataTime_and_Kurs();
var
  aDataSet: TDataSet;
  aActualDays: LongInt;
  aTmp, aTmp2: String;

begin
  aDataSet:= fGridPrice.Grid.DataSource.DataSet;

  if not TFmPrices(fOwnerForm).tbBtnPricePosition.Marked and (aDataSet.RecordCount<>0) then
      GridPositionFill([aDataSet.FieldByName('ID').AsInteger]);

  TFmPrices(fOwnerForm).st_PriceVersion.Caption:='Дата: '+aDataSet.FieldByName('FTIMESTAMP').AsString;

  aActualDays:= aDataSet.FieldByName('ACTUALDAYS').AsInteger*-1;
  if (aActualDays = 0) or (IncDay(Now, aActualDays) < aDataSet.FieldByName('FTIMESTAMP').AsDateTime) then
    TFmPrices(fOwnerForm).st_PriceVersion.Font.Color:= clGreen
  else
    TFmPrices(fOwnerForm).st_PriceVersion.Font.Color:= clRed;

  TreeInfoFill(fGridPrice);
end;

procedure TPrices.fGridPrice_onDataChange(Sender: TObject; Field: TField);
var
  _result: string;
begin
   if fGridPrice.FillGridNow or (fGridPrice.Grid.DataSource.DataSet.RecordCount=0) then exit;

  _result:='';

  case TFmPrices(fOwnerForm).pcPriceGroup.ActivePageIndex of
    0: _result:= fTreePriceOwner.BreadCrumbs(fGridPrice.Grid.DataSource.DataSet.FieldByName('IDOWNER').AsInteger);
    1: _result:= fTreePriceGroup.BreadCrumbs(fGridPrice.Grid.DataSource.DataSet.FieldByName('IDPL_GROUP').AsInteger);
  end;

  SetStatus('['+_result+'] '+fGridPrice.Grid.DataSource.DataSet.FieldByName('PLNAME').AsString);  // NAME

  fGridPriceViewDataTime_and_Kurs();
end;

procedure TPrices.fTreeMatching_onSelectionChanged(Sender: TObject);
begin

end;

procedure TPrices.fGridMatching_onDataChange(Sender: TObject; Field: TField);
begin

end;

procedure TPrices.Log(aText: string);
begin

end;

procedure TPrices.onCbInvoiceChange(Sender: TObject; aValue: integer);
var
  _Form: TFmInvoce;
  _DS: TDataSource;
begin
  _Form:= nil;
  _Form:= TFmInvoce(Sender);

  _DS:= nil;
  _DS:= Base.SQLItemGetDS(fInvoce.Item,[aValue]);

  if Assigned(_DS) and (_DS.DataSet.RecordCount >0) then
  begin
    _Form.edQuantity.Value:= _DS.DataSet.FieldByName('QUANTITY').AsInteger;
    _Form.edRemark.Text:= _DS.DataSet.FieldByName('REMARK').AsString;
  end
  else
  begin
    _Form.edQuantity.Value:= 1;
    _Form.edRemark.Text:= '';
  end;
end;

procedure TPrices.SetMaxFTimeStampArr(AValue: ArrayOfDateTime);
begin
  __PRICE_MAX_FTIMESTAMP_ARR:=AValue;
  //fTreePriceGroup.WhereTime:=' AND ('+fBase.PrepareWhereStringFromDateTime('"GROUP-NOMENCLATURE".FTIMESTAMP',__PRICE_MAX_FTIMESTAMP_ARR)+') ';
  __FTIMESTAMP_KURS:= fBase.SQLReadArr('CURRENCY',['FTIMESTAMP'],'ID=2','')[0,0];
end;

procedure TPrices.SetStatus(aText: string; const aLog: boolean);
begin
  wStatus(fFormName, aText, aLog);
end;

constructor TPrices.Create(Sender: TObject; aBase: TwBase; aGridPrice, aGridPosition, aGridInvoce: TDBGrid; aTreePriceOwner, aTreePriceGroup,
  aTreeImportOwner: TTreeView);
begin
  fOwnerForm:=Sender;

  fFormName:= TFmPrices(Sender).Name;
  fBase:= aBase;
  fIdMainOwner:= fBase.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист

  __PRICE_MAX_FTIMESTAMP_ARR:= GetMaxFTimeStampPricesArr(Base);
  __FTIMESTAMP_KURS:= fBase.SQLReadArr('CURRENCY',['FTIMESTAMP'],'ID=2','')[0,0];

  fGridPrice:= TwDBGrid.Create(Base,aGridPrice,'');
  fGridPrice.MultiSelect:= true;
  fGridPrice.StaticTextSelection:=TFmPrices(fOwnerForm).st_GridSelect;
  fGridPrice.SearchEdit:=TFmPrices(fOwnerForm).edPriceSearch;
  fGridPrice.SearchPreventiveBtn:=TFmPrices(fOwnerForm).btnPricePreventSearch;
  fGridPrice.SearchSplitStringBtn:=TFmPrices(fOwnerForm).btnPriceSearchSplitString;

  fGridPrice.SortTitleImagesIndex:=[2,3];

  fGridPrice.SearchEntryArray:= ['PL.NAME','PL.LABEL','PL.REMARK'];
  fGridPrice.SearchParticleArray:= ['PL.VENDORCODE','(SELECT VRESULT FROM PL_TRY_SCOD(PL.ID,''%s'') /*=*/)'];
  fGridPrice.GroupField:= 'PL.IDOWNER';
  fGridPrice.onSelect:=@GridPrice_onSelect;

  fTreePriceOwner:= TwDBTree.Create(Base,aTreePriceOwner,'OWNER','IDPARENT,NAME',[]);
  fTreePriceOwner.MultiSelect:= true;
  fTreePriceOwner.Expanded:= true;
  fTreePriceOwner.Tree.OnSelectionChanged:=@fTreePriceOwner_onSelectionChanged;

  fTreePriceGroup:= TwDBTree.Create(Base,aTreePriceGroup,'PL_GROUP','IDPARENT,ID',['IDOWNER',0]);
  fTreePriceGroup.MultiSelect:= true;
  fTreePriceGroup.Expanded:= false;

  fTreePriceGroup.Tree.OnSelectionChanged:=@fTreePriceGroup_onSelectionChanged;

 if Length(__PRICE_MAX_FTIMESTAMP_ARR)>0 then
   //fTreePriceGroup.WhereTime:= ' AND ('+fBase.PrepareWhereStringFromDateTime('"GROUP-NOMENCLATURE".FTIMESTAMP',__PRICE_MAX_FTIMESTAMP_ARR)+') ';
   fTreePriceGroup.WhereRoot:=' OR (IDOWNER=%d AND IDFORMATS IS NULL)';
   fTreePriceGroup.SetOwner:= 0;

 fGridPosition:= TwDBGrid.Create(Base,aGridPosition,'');
 fGridPosition.SortTitleImagesIndex:=[2,3];
 //fGridPosition.MultiSelect:= true;

 fInvoce:= TInvoce.Create(fBase, nil);
 fInvoce.onSumCountChanged:= @onInvoceSumCountChanged;
 fGridInvoces:= TwDBGrid.Create(Base,aGridInvoce, fInvoce.List);
 fGridInvoces.GroupField:=fInvoce.GroupField;
 fGridInvoces.SearchEntryArray:= fInvoce.SearchEntryArray;
 fGridInvoces.SearchParticleArray:= fInvoce.SearchParticleArray;
 fGridInvoces.SearchEdit:= TFmPrices(fOwnerForm).edInvoceSearch;
 fGridInvoces.MultiSelect:= true;
 fGridInvoces.SortTitleImagesIndex:=[2,3];

 fInvoce.Grid:= fGridInvoces;

 fTreeImportOwner:= TwDBTree.Create(Base,aTreeImportOwner,'OWNER','IDPARENT,NAME',[]);
 fTreeImportOwner.MultiSelect:= true;
 fTreeImportOwner.Expanded:= true;

end;

destructor TPrices.Destroy();
begin
   fGridPrice.Destroy();
   fTreePriceOwner.Destroy();
   fTreePriceGroup.Destroy();
   fGridPosition.Destroy();

   fGridInvoces.Destroy();
   fInvoce.Destroy;

   fTreeImportOwner.Destroy();
end;

procedure TPrices.GridPriceFill(aGroup: ArrayOfInteger);
var
  _SQLText: string;
begin
  if not Assigned(Base) then exit;

  fGridPrice.GroupArray:= aGroup;

  _SQLText:='SELECT PL.ID,'
    +' PL.IDOWNER, '
    +' PL.NAME AS PLNAME, '
    +' PL.IDPL_GROUP, '
    +' PL.UNIT , '
    +' PL.LABEL, '
    +' PL.VENDORCODE, '
    +' (PL.STOCK+PL.STOCK2+PL.STOCK3+PL.STOCK4+PL.STOCK5) as STOCK, '
    +' PL.STOCK AS STOCK1, '
    +' PL.STOCK2, '
    +' PL.STOCK3, '
    +' PL.STOCK4, '
    +' PL.STOCK5, '
    +' PL.TRANSIT, '
    +' PL.PRICECALC*(1.0+'+StringReplace(FloatToStrF(TFmPrices(fOwnerForm).Markup.Value,ffFixed,4,2),',','.',[])+'/100.0) AS PRICE, '
    +' PL.PRICECALC2 AS PRICE2, '
    +' PL.PRICECALC3 AS PRICE3, '
    +' PL.PRICECALC4 AS PRICE4, '
    +' PL.PRICECALC5 AS PRICE5, '
    +' PL.PRICECALC6 AS PRICE6, '
    +' PL.PRICECALC7 AS PRICE7, '
    +' PL.PRICECALC8 AS PRICE8, '
    +' PL.PRICECALC9 AS PRICE9, '
    +' PL.PRICECALC10 AS PRICE10, '
    +' PL.FTIMESTAMP AS FTIMESTAMP, '
    +' PL.FCOLOR, '
    +' OWNER.NAME AS OWNERNAME, '
    +' MTH.ID AS MTHRESULT, '
    +' FMTS.STOCKONLYINFO AS STOCKONLYINFO, '
    +' IIF(PL.IDFORMATS>0, FMTS.ACTUALDAYS, 1) AS ACTUALDAYS '
    +' FROM PL_ITEMS PL  '
    +' LEFT JOIN OWNER ON PL.IDOWNER=OWNER.ID   '
    +' LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PL.IDFORMATS) '
    +' LEFT JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS=PL.ID) ';
  if TFmPrices(fOwnerForm).sBtnDoublesView.Down then
    _SQLText:= _SQLText+' INNER JOIN (select IDOWNER,NAME from PL_ITEMS group by IDOWNER,NAME having count(*)>1) PLDBL ON (PLDBL.NAME=PL.NAME AND PLDBL.IDOWNER=PL.IDOWNER) ';

    if Length(__PRICE_MAX_FTIMESTAMP_ARR)>0 then
      _SQLText:= _SQLText+'  WHERE (FMTS.FCLOSE=0 OR PL.IDFORMATS=0) '+fGridPriceFilterString
      +' /*and_where_string*/ /*and_group_string*/ /*and_search_string*/ ';

     //(SELECT VRESULT FROM PL_TRY_SCOD(PL.ID,null))
  //_PL_FTIMESTAMP:= Base.PrepareWhereStringFromDateTime('PL.FTIMESTAMP',__PRICE_MAX_FTIMESTAMP_ARR);

  fGridPrice.Fill(_SQLText);

  if Assigned(fGridPrice.Grid.DataSource) then
       fGridPrice.Grid.DataSource.OnDataChange:=@fGridPrice_onDataChange;

  fGridPriceViewDataTime_and_Kurs();

end;

procedure TPrices.GridPrice_onSelect(Sender:TObject);
begin
  if not Assigned(fGridPrice.Grid.DataSource) then exit;
  if TFmPrices(fOwnerForm).sBtnSelected.Down then
       GridPriceFiltered();


end;

procedure TPrices.GridPriceFiltered();
var
  _arr: ArrayOfInteger;
  _BookMark: TBookMark;
begin
  fGridPriceFilterString:='';


  if TFmPrices(fOwnerForm).sBtnStockOnly.Down then
    begin
      if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND (PL.STOCK>0 OR PL.STOCK2>0 OR PL.STOCK3>0 OR PL.STOCK4>0 OR PL.STOCK5>0)' else
          fGridPriceFilterString:=fGridPriceFilterString+' (PL.STOCK>0 OR PL.STOCK2>0 OR PL.STOCK3>0 OR PL.STOCK4>0 OR PL.STOCK5>0) ';
    end;

  if TFmPrices(fOwnerForm).sBtnWithMatching.Down then
    begin
      if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND MTH.ID>0' else
          fGridPriceFilterString:=fGridPriceFilterString+' MTH.ID>0 ';
    end;

  if TFmPrices(fOwnerForm).sBtnNoMatching.Down then
    begin
      if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND MTH.ID IS NULL' else
          fGridPriceFilterString:=fGridPriceFilterString+' MTH.ID IS NULL';
    end;

  if TFmPrices(fOwnerForm).sBtnSelected.Down then
    begin
      _arr:= fGridPrice.SelectedRows();

      if fGridPrice.SelectedRowsCount>0 then
        begin
              if Length(fGridPriceFilterString)>0 then
                    fGridPriceFilterString:=fGridPriceFilterString+' AND ('+fBase.PrepareWhereString('PL.ID',_arr)+')' else
                      fGridPriceFilterString:=fGridPriceFilterString+' '+fBase.PrepareWhereString('PL.ID',_arr);
        end else
        begin
              if Length(fGridPriceFilterString)>0 then
                     fGridPriceFilterString:=fGridPriceFilterString+' AND PL.ID=0 ' else
                       fGridPriceFilterString:=fGridPriceFilterString+' PL.ID=0 ';
        end;

    end;

  if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=' AND ('+fGridPriceFilterString+')';

  screen.Cursor:= crSQLWait;
  _BookMark:= fGridPrice.Grid.DataSource.DataSet.Bookmark;
  GridPriceFill(fGridPrice.GroupArray);

  if fGridPrice.Grid.DataSource.DataSet.RecordCount>0 then
      begin
        fGridPrice.Grid.DataSource.DataSet.Bookmark:= _BookMark;
      end;

  screen.Cursor:= crDefault;

end;

procedure TPrices.GridPositionFill(aGroup: ArrayOfInteger);
begin
  if not Assigned(Base) then exit;

  fGridPosition.GroupArray:= aGroup;

  GridPositionModeChange();

end;

procedure TPrices.GridInvocesFill(aGroup: ArrayOfInteger);
var
  aVisible: Boolean;
begin
  aVisible:= TFmPrices(fOwnerForm).pPricePositionVersion.Visible and  (TFmPrices(fOwnerForm).pcBottom.ActivePageIndex = 1);
  fInvoce.GridFill(aGroup, aVisible);
end;

procedure TPrices.TreePriceOwnerFill();
begin
  fTreePriceOwner.Fill();
end;

procedure TPrices.TreePriceGroupFill();
begin
  fTreePriceGroup.Fill();
end;

procedure TPrices.TreeImportOwnerFill();
begin
  fTreeImportOwner.Fill();
end;

procedure TPrices.GridPositionModeChange();
var
  _SQLText: string;
  _IdPL: LongInt;
begin

  if not Assigned(GridPrice.Grid.DataSource)
  or (GridPrice.Grid.DataSource.DataSet.RecordCount=0)
  or not TFmPrices(fOwnerForm).tbBtnPricePosition.Down
  then exit;

  //_Vendorcode:= GridPrice.Grid.DataSource.DataSet.FieldByName('VENDORCODE').AsString;
  _IdPL:= GridPrice.Grid.DataSource.DataSet.FieldByName('ID').AsInteger;

  case GridPosition.Grid.Tag of
    0:
      begin
          TFmPrices(fOwnerForm).gbPricePosition.Caption:='Архив цен на позицию';
          TFmPrices(fOwnerForm).tbPositionPriceArc.Down:= true;
          TFmPrices(fOwnerForm).tbSummaryPrice.Down:= false;

          _SQLText:='SELECT PL.ID,'
             +' PL.IDOWNER, '
             +' PL.NAME AS PLNAME, '
             +' PL.IDPL_GROUP, '
             +' PL.UNIT , '
             +' PL.LABEL, '
             +' PL.VENDORCODE, '
             +' (PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5) as STOCK, '
             +' PLV.TRANSIT, '
             +' PL.FCOLOR, '
             +' PLV.PRICECALC AS PRICE, '
             +' PLV.FTIMESTAMP AS FTIMESTAMP, '
             +' OWNER.NAME AS OWNERNAME, '
             +' MTH.ID AS MTHRESULT, '
             +' FMTS.STOCKONLYINFO AS STOCKONLYINFO, '
             + ' 1 as QUANTITYINPACKING, '
             + ' '' '' as QUANTITYINPACKINGTEXT '
             +' FROM "PL_VERSIONS" PLV  '
             +' INNER JOIN "PL_ITEMS" PL ON (PLV.IDPL_ITEMS=PL.ID) '
             +' LEFT JOIN OWNER ON PL.IDOWNER=OWNER.ID '
             +' LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PL.IDFORMATS) '
             +' LEFT JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS=PL.ID) ';

              _SQLText:= _SQLText+'  WHERE  (FMTS.FCLOSE=0) AND PLV.IDPL_ITEMS=%d ORDER BY PLV.FTIMESTAMP DESC';

              GridPosition.SetColumnCaption('PRICE','Цена');
              GridPosition.SetColumnCaption('STOCK','Остаток');

          fGridPosition.Fill(Format(_SQLText,[_IdPL]));
      end;
    1:
      begin
             TFmPrices(fOwnerForm).gbPricePosition.Caption:='Сводный прайс-лист на позицию: '+GridPrice.Grid.DataSource.DataSet.FieldByName('PLNAME').AsString;

             TFmPrices(fOwnerForm).tbPositionPriceArc.Down:= false;
             TFmPrices(fOwnerForm).tbSummaryPrice.Down:= true;

             _SQLText:='SELECT '
               +' ID, '
               +' STOCKONLYINFO, '
               +' QUANTITYINPACKING, '
               +' QUANTITYINPACKINGTEXT, '
               +' PLNAME, '
               +' VENDORCODE, '
               +' LABEL, '
               +' UNIT, '
               +' PRICE, '
               +' STOCK, '
               +' FTIMESTAMP, '
               +' FCOLOR, '
               +' OWNERNAME '
               +' FROM ANALIS_SEL_ALL_ANALOG('+IntToStr(_IdPL)+', true) ORDER BY PRICE';

             GridPosition.SetColumnCaption('PRICE','Цена *');
             GridPosition.SetColumnCaption('STOCK','Остаток *');

            fGridPosition.Fill(_SQLText);
      end;
  end;

end;

procedure TPrices.onStatusUpdate(Sender: TObject);
var
  _Status: string;
begin
  _Status:= fDeletePriceThread.Status;
  SetStatus(_Status);
  //fProgress.SetStatus(_Status);
  //fProgress.SetBar(pbBottom);
end;

procedure TPrices.onStopForce(Sender: TObject);
begin
  fDeletePriceThread.StopForce:= true;
end;

procedure TPrices.RepOnEndThread(Sender: TObject);
begin
 try
   fProgress.ForceClose;

     if not fReport.Result then
        raise Exception.Create('Во время создания отчета произошла ошибка!');

 except
   on E: Exception do begin
      MessageDlg(E.Message,mtError, [mbOK], 0);
   end;
 end
end;

procedure TPrices.RepOnProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
begin
  fProgress.InitBar(aProgressBarName, aValue);
end;

procedure TPrices.RepOnProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
begin
  fProgress.SetBar(aProgressBarName, aValue);
end;

procedure TPrices.RepOnStopForce(Sender: TObject);
begin
  fReport.Stop();
end;

procedure TPrices.onEndThread(Sender: TObject);
var
  _Result: Boolean;
begin
  _Result:= fDeletePriceThread.Result;

  fDeletePriceThread.Terminate;
  fDeletePriceThread:= nil;

  fProgress.ForceClose;

  GridPriceFill(TreePriceOwner.SelectedItems);

  if _Result then ShowMessage('Успешно завершено!') else ShowMessage('При удалении прайс-листа произошла ошибка!');

end;

procedure TPrices.onInvoceSumCountChanged(Sender: TObject; aSum: double; aCount: integer);
begin
  TFmPrices(fOwnerForm).st_InvoceItog.Caption:='Сумма: '+CurrToStrF(aSum, ffCurrency, 2)+' | Строк: '+IntToStr(aCount)+'  ';
end;

procedure TPrices.onProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
begin
   fProgress.InitBar(aProgressBarName,aValue);
end;

procedure TPrices.onProgressStatus(const aValue: string);
begin
  fProgress.SetLog(aValue);
end;

procedure TPrices.onProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
begin
   fProgress.SetBar(aProgressBarName,aValue);
end;

procedure TPrices.DeletePriceList(aOwner:integer);
begin
  fDeletePriceThread:= nil;

  fProgress:= TProgress.Create(TForm(fOwnerForm));
  fProgress.Caption:= 'Удаление...';
  fProgress.ShowTop:= false;
  fProgress.onStopForce:= @onStopForce;

  screen.Cursor:=crSQLWait;

  fDeletePriceThread:= TDeletePriceThread.Create(true);
  fDeletePriceThread.Base:= fBase;
  fDeletePriceThread.onEndThread:= @onEndThread;
  fDeletePriceThread.onStatusUpdate:= @onStatusUpdate;

  fDeletePriceThread.onProgressInit:= @onProgressInit;
  fDeletePriceThread.onProgressUpdate:= @onProgressUpdate;
  fDeletePriceThread.onProgressStatus:= @onProgressStatus;

  fDeletePriceThread.fOwner:= aOwner;
  fDeletePriceThread.Start;

  try
    fProgress.ShowModal;
  finally
    screen.Cursor:=crDefault;
    if Assigned(fProgress) then
      fProgress.Free;
  end;
end;

procedure TPrices.DeletePriceGroup;
var
  aSelectGroups: ArrayOfInteger;
begin
  aSelectGroups:= TreePriceGroup.SelectedItems;
  Base.SQLDelete('PL_GROUP','ID IN ('+Base.MakeStringFromArray(aSelectGroups)+')');
  TreePriceGroup.Fill();
end;

procedure TPrices.DeletePriceItem;
var
  aSelCount: integer;
  aArr: ArrayOfInteger;
  aBookMark: TBookMark;
  i: integer;
  aID: integer;
  aGridDataset: TDataSet;
begin
     aGridDataset:= GridPrice.Grid.DataSource.DataSet;

     aArr:= GridPrice.SelectedRows;

     aSelCount:= Length(aArr);

  if aSelCount > 1 then
    begin
    if MessageDlg('Удалить несколько позиций ('+IntToStr(aSelCount)+') ? При удалении позиции так же будут удалены связанные соответствия!',mtWarning, mbOKCancel, 0) = mrOK then
       begin

         aBookMark:= aGridDataset.Bookmark;

         try
           for i:=0 to Length(aArr)-1 do
           begin
              Base.SQLDelete('PL_ITEMS','ID='+IntToStr(aArr[i]),false);
           end;
         finally
            Base.SQLTransactionEnd(true);
            aArr:=nil;
            with aGridDataset do begin
                GridPrice.Fill();
                if RecordCount>0 then BookMark:= aBookMark;
            end;
            wLog('PL_ITEMS',IntTOStr(aSelCount)+' позиций успешно удалено.');
            SetStatus(IntTOStr(aSelCount)+' позиций успешно удалено.');
         end;

       end;
    end else
    begin
     if MessageDlg('Удалить позицию "'+aGridDataset.FieldByName('PLNAME').AsString+'" ? При удалении позиции так же будут удалены связанные соответствия!',mtWarning, mbOKCancel, 0) = mrOK then
        begin
           aID:=aGridDataset.FieldByName('ID').AsInteger;
           aBookMark:= aGridDataset.Bookmark;
           if Base.SQLDelete('PL_ITEMS','ID='+IntToStr(aID)) then
              wLog('PL_ITEMS','Позиция успешно удалена');
              SetStatus('Позиция успешно удалена');
           with aGridDataset do begin
               GridPrice.Fill();
               if RecordCount>0 then BookMark:= aBookMark;
           end;
        end;
    end;
    GridPrice.SelectedRowsClear();
end;

procedure TPrices.InvoceDel(aInvoces: ArrayOfInteger);
var
  aVisible: Boolean;
begin
  aVisible:= TFmPrices(fOwnerForm).pPricePositionVersion.Visible and  (TFmPrices(fOwnerForm).pcBottom.ActivePageIndex = 1);

  fInvoce.InvoceDel(fTreePriceOwner.SelectedItems, aVisible);
end;

procedure TPrices.InvoceAdd();
var
  aVisible: Boolean;
  aIdPricePosition: LongInt;
begin
  aVisible:= TFmPrices(fOwnerForm).pPricePositionVersion.Visible and (TFmPrices(fOwnerForm).pcBottom.ActivePageIndex = 1);
  aIdPricePosition:= fGridPrice.Grid.DataSource.DataSet.FieldByName('ID').AsInteger;

  fInvoce.InvoceAdd(aIdPricePosition, fTreePriceOwner.SelectedItems, aVisible);
end;

procedure TPrices.InvoceEdit(aId: integer);
var
  aVisible: Boolean;
begin
  aVisible:= TFmPrices(fOwnerForm).pPricePositionVersion.Visible and (TFmPrices(fOwnerForm).pcBottom.ActivePageIndex = 1);

  fInvoce.InvoceEdit(aId, TreePriceOwner.SelectedItems, aVisible);
end;

procedure TPrices.InvoceChangeItem(aIdPL, aIdInvoce: integer);
var
  _Quantity: LongInt;
  _Remark: String;
  _DS: TDataSet;
  _Result: Integer;
begin
  _Result:= -1;
  _Quantity:= GridInvoces.FieldValue['QUANTITYPL'].AsInteger;
  _Remark:= GridInvoces.FieldValue['REMARK'].AsString;

  Base.LongTransaction:= true;

  try
    _DS:= Base.SQLReadDS('PL_ITEMS',['ID','IDOWNER'],'ID='+IntToStr(aIdPL),'').DataSet;

    Base.SQLItemUpdate(fInvoce.Del,[aIdInvoce],false,false);
    //:IDOWNER, :IDPL_ITEMS, :QUANTITY, :REMARK
    _Result:= Base.SQLItemUpdate(fInvoce.New,[_DS.FieldByName('IDOWNER').AsInteger,
                      _DS.FieldByName('ID').AsInteger,
                      _Quantity,
                      _Remark],true,false);

  Base.SQLTransactionEnd(true);

  with GridInvoces.Grid.DataSource.DataSet do
  begin
    Close;
    Open;
    DisableControls;
    Locate('ID',_Result,[]);
    EnableControls;
  end;

  except
    Base.SQLTransactionEnd(false);
  end;

end;

procedure TPrices.mAnalogsFill(Sender: TObject; amAnalogs: TMenuItem);
begin
  fInvoce.mAnalogsFill(Sender, amAnalogs);
end;

procedure TPrices.PrintDateTimePrices;
var
  fViewer: TwViewer;
  fWorkbookSource: TsWorkbookSource;
begin
 ///W_TMP_ORDERS_IMPORT
 fViewer:= TwViewer.Create(nil);
 fViewer.Caption:= 'Даты импорта прайс-листов';

 fWorkbookSource:= TsWorkbookSource.Create(Application);
 fWorkbookSource.Options:= fWorkbookSource.Options+[boFileStream];

 fProgress:= TProgress.Create(nil);
 fProgress.Caption:= 'Формирование отчета...';
 fProgress.ShowLog:= false;
 fProgress.onStopForce:= @RepOnStopForce;
 fProgress.NoClose:= true;

 fReport:= TwReport.Create(true);
 fReport.Base:= fBase;
 fReport.ReportModes:= rmPriceDate;
 fReport.WorkbookSource:= fWorkbookSource;
 fReport.onProgressInit:= @RepOnProgressInit;
 fReport.onProgressUpdate:= @RepOnProgressUpdate;
 fReport.onEndThread:= @RepOnEndThread;

 screen.Cursor:= crSQLWait;

 fReport.start;

 try

   fProgress.ShowModal;

   fViewer.WorkbookSource:= fWorkbookSource;

   if fReport.Result then
     begin
        fViewer.ShowModal;
     end;

//   fReport.Terminate;
 finally
   screen.Cursor:=crDefault;
   if Assigned(fProgress) then
     fProgress.Free;
   fViewer.free;
 end;

end;

procedure TPrices.TreeInfoFill(aGrid: TwDBGrid);
var
  _arr, _BarcodeArr: ArrayOfArrayVariant;
begin
  if not TreeInfo.Parent.Visible then exit;

  _arr:=nil;
  if  aGrid.Grid.DataSource.DataSet.RecordCount>0 then
    _arr:= fBase.SQLReadArr('PL_ITEMS',['REMARK','FURL','FURLPICTURE'],'ID='+aGrid.Grid.DataSource.DataSet.FieldByName('ID').AsString,'');

  TreeInfo.Items[0].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('FTIMESTAMP').AsString;
  TreeInfo.Items[1].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('OWNERNAME').AsString;
  TreeInfo.Items[2].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('VENDORCODE').AsString;
  TreeInfo.Items[3].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('PLNAME').AsString;
  TreeInfo.Items[4].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('UNIT').AsString;


  _BarcodeArr:=nil;
  if Assigned(_arr) then
    _BarcodeArr:= fBase.SQLReadArr('SELECT VSCOD FROM PL_GET_SCOD('+aGrid.Grid.DataSource.DataSet.FieldByName('ID').AsString+',true)');
  if Assigned(_BarcodeArr) then
       TreeInfo.Items[5].Text:= VarToStr(_BarcodeArr[0,0]);

  TreeInfo.Items[6].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('LABEL').AsString;
  TreeInfo.Items[8].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[9].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE2').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[10].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE3').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[11].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE4').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[12].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE5').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[13].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE6').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[14].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE7').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[15].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE8').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[16].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE9').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[17].Text:= CurrToStrF(aGrid.Grid.DataSource.DataSet.FieldByName('PRICE10').AsCurrency, ffCurrency, 2);

  TreeInfo.Items[19].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('STOCK1').AsString;
  TreeInfo.Items[20].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('STOCK2').AsString;
  TreeInfo.Items[21].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('STOCK3').AsString;
  TreeInfo.Items[22].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('STOCK4').AsString;
  TreeInfo.Items[23].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('STOCK5').AsString;

  TreeInfo.Items[24].Text:= VarToStr(aGrid.Grid.DataSource.DataSet.FieldByName('TRANSIT').AsVariant);
  if Assigned(_arr) then
    begin
      TreeInfo.Items[25].Text:= VarToStr(_arr[0,0]);
      TreeInfo.Items[26].Text:= VarToStr(_arr[0,1]);            //11
      TreeInfo.Items[27].Text:= VarToStr(_arr[0,2]);     //12
    end else
    begin
      TreeInfo.Items[25].Text:= '';
      TreeInfo.Items[26].Text:= '';            //11
      TreeInfo.Items[27].Text:= '';     //12
    end;

  TreeInfo.Items[28].Text:= aGrid.Grid.DataSource.DataSet.FieldByName('FCOLOR').AsString;

end;

end.

