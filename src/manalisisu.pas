unit mAnalisisU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, fpspreadsheet, fpspreadsheetctrls, fpsTypes, SysUtils, Controls,
  db, ComCtrls, Forms, DBGrids, Dialogs, Menus,
  TAGraph, TAIntervalSources, TASeries, TAChartUtils, TATextElements,
  IBDatabase, IBQuery, IBCustomDataSet, IBSQL, DateUtils, Graphics, math,
  wCustomClassThreadU, wLogU, wFuncU, mUtilsU,
  wBaseU, wDBGridU, wDBTreeU, wReportU, wTProgressU, wTViewerSpreadsheetU, wTypesU;

type



   { TwAnalisis }
   TwGridPriceFiltered = record
     GroupArray: ArrayOfInteger;
     WhereString: string;
   end;


   TwAnalisis = class
    private
      fCurrentGrid: TwDBGrid;
      fFormName: string;
      fGridChangePrice: TwDBGrid;
      fGridChangeStock: TwDBGrid;

      fGridPrice: TwDBGrid;
      fChartTag: integer;
      fGridPriceFilterString: String;
      fGridPriceListNew: TwDBGrid;
      fProgress: TProgress;
      FTreeInfo: TTreeView;
      fTreePriceOwner: TwDBTree;
      fTreePriceGroup: TwDBTree;

      fOwnerForm: TObject;
      fIdMainOwner: string;

      fBase: TwBase;
      fUtils: TUtils;
      fReport: TwReport;

      __PRICE_MAX_FTIMESTAMP_ARR: ArrayOfDateTime;
      __TreePriceOwnerDataArr : TwTree_FOT_Data_Arr;

      __FTIMESTAMP_KURS: string;

      procedure fGrid_onCellClick(Sender: TObject);
      function IsOneOwner(aArr: TwTree_FOT_Data_Arr): boolean;
      procedure Log(aText: string);
      procedure onEndThread(Sender: TObject);
      procedure onProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
      procedure onProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
      procedure onStopForce(Sender: TObject);
      function PrepareGridPriceGroupArray(const aFOTArray: TwTree_FOT_Data_Arr): TwGridPriceFiltered;
      procedure SetStatus(aText: string; const aLog: boolean = true); // вывод статуса

      procedure fTreePriceOwner_onSelectionChanged(Sender: TObject);
      procedure fTreePriceGroup_onSelectionChanged(Sender: TObject);
      procedure fGridPrice_onDataChange(Sender: TObject; Field: TField);

      function ReadPricesTimeStamp():TwTree_FOT_Data_Arr;

    public
      property Base: TwBase read fBase write fBase;
      property GridPrice: TwDBGrid read fGridPrice write fGridPrice;
      property GridChangePrice: TwDBGrid read fGridChangePrice write fGridChangePrice;
      property GridPriceListNew:TwDBGrid read fGridPriceListNew write fGridPriceListNew;
      property GridChangeStock:TwDBGrid read fGridChangeStock write fGridChangeStock;
      property GridCurrent:TwDBGrid read fCurrentGrid write fCurrentGrid;

      property IdMainOwner: string read fIdMainOwner write fIdMainOwner;
      property TreePriceOwner: TwDBTree read fTreePriceOwner write fTreePriceOwner;
      property TreePriceGroup: TwDBTree read fTreePriceGroup write fTreePriceGroup;
      property TreeInfo: TTreeView read FTreeInfo write FTreeInfo;
      property PriceMaxFTimeStampArr: ArrayOfDateTime read __PRICE_MAX_FTIMESTAMP_ARR write __PRICE_MAX_FTIMESTAMP_ARR;
      property TreePriceOwnerDataArr: TwTree_FOT_Data_Arr read __TreePriceOwnerDataArr write __TreePriceOwnerDataArr;

      property ChartTag: integer read fChartTag write fChartTag;

      constructor Create(Sender: TObject; aBase: TwBase; aGridPrice, aGridChangePrice, aGridPriceNew, aGridChangeStock: TDBGrid; aTreePriceOwner, aTreePriceGroup: TTreeView);
      destructor Destroy();
      procedure GridPriceFill(aGroup: TwGridPriceFiltered);

      procedure TreePriceOwnerFill();
      procedure TreePriceGroupFill();
      procedure TreeInfoFill();
      procedure ChangedPriceData();
      procedure ChartPriceUpdate();
      procedure GridPriceFiltered(aGridOff: boolean = false);
      procedure ChangeGridSearchEdit();
      procedure TreeOwnerExpand(aLevel: integer);
      procedure ExportData();
      procedure GetPositionAnalog(aSelectedItems: ArrayOfInteger);
      procedure GetCompareHorisontal(aSelectedItems, aSelectedOwners: ArrayOfInteger; aPriceBase, aPriceCompare: TPriceType);

      function CreateGridPriceFiltered(aAray: ArrayOfInteger; const aString: string =''):TwGridPriceFiltered;
  end;
implementation

uses
  pkgAnalisisU;

{ TwAnalisis }

procedure TwAnalisis.Log(aText: string);
begin
  if __onLog and assigned (__Log) then
    wLog(fFormName, aText);
end;

procedure TwAnalisis.SetStatus(aText: string; const aLog: boolean);
begin
  wStatus(fFormName, aText, aLog);
end;


function TwAnalisis.IsOneOwner(aArr:TwTree_FOT_Data_Arr): boolean;
var
  i, _OldOwner: Integer;
begin
  Result:= true;

   for i:=0 to High(aArr) do
   begin
      if i=0 then _OldOwner:= aArr[i].IdOwner else
          if _OldOwner<> aArr[i].IdOwner then
           begin
             Result:= false;
             exit;
           end else
           Result:= true;

   end;
end;

procedure TwAnalisis.fTreePriceOwner_onSelectionChanged(Sender: TObject);
var
  _TreeView: TTreeView;
  _OwnerArr: TwTree_FOT_Data_Arr;
  _Cursor: TCursor;
begin
  _TreeView:= fTreePriceOwner.Tree;

  if _TreeView.SelectionCount>1 then
    begin
      fChartTag:=2;
    end else
      fChartTag:=TFmAnalisis(fOwnerForm).ChartPrice.Tag;

  if (_TreeView.SelectionCount=0) or (fGridPrice.Grid=nil) then exit;

  if _TreeView.Items = nil then
   begin
     SetStatus('Ошибка Tree.');
     Log('Ошибка Tree.');
     exit;
   end;
  try


 if (_TreeView.Items.Count = 0) or (_TreeView.SelectionCount = 0) then exit;

 _OwnerArr:= nil;
 _OwnerArr:= TreePriceOwner.SelectedItems(true);

 if IsOneOwner(_OwnerArr) then TFmAnalisis(fOwnerForm).TabGroup.TabVisible:= true else TFmAnalisis(fOwnerForm).TabGroup.TabVisible:= false;

 if (Length(_OwnerArr) = 2) and (_TreeView.SelectionCount = 2) and (_OwnerArr[0].IdOwner =  _OwnerArr[1].IdOwner) and _OwnerArr[0].LastItem and  _OwnerArr[1].LastItem then
  begin
    _Cursor:= _TreeView.Cursor;
    _TreeView.Cursor:= crSQLWait;
    SetStatus('Подготовка данных для анализа...');
    fBase.SQLUpdate('ALTER INDEX W_TMP_PL_VERSIONS_IDPL INACTIVE;');
    fBase.SQLUpdate('ALTER INDEX W_TMP_PL_VERSIONS_ID_PRCALC INACTIVE;');

    fBase.SQLUpdate('DELETE FROM w_tmp_pl_versions;');
    fBase.SQLUpdate('INSERT INTO w_tmp_pl_versions SELECT * FROM pl_versions WHERE pl_versions.FTIMESTAMP='+QuotedStr(DateTimeToStr(_OwnerArr[1].TimeStamp))+'');

    fBase.SQLUpdate('ALTER INDEX W_TMP_PL_VERSIONS_IDPL ACTIVE;');
    fBase.SQLUpdate('ALTER INDEX W_TMP_PL_VERSIONS_ID_PRCALC ACTIVE;');
    SetStatus('Подготовка данных для анализа завершена.');
    _TreeView.Cursor:= _Cursor;

    TFmAnalisis(fOwnerForm).btnPriceChange.Enabled:= true;
    TFmAnalisis(fOwnerForm).btnPositionChangeAssort.Enabled:= true;
    TFmAnalisis(fOwnerForm).btnPositionChangeStock.Enabled:= true;
    TFmAnalisis(fOwnerForm).btnPriceChange.Hint:= 'Показать позиции с изменением цены';
    TFmAnalisis(fOwnerForm).btnPositionChangeAssort.Hint:= 'Показать изменения в ассортименте';
    TFmAnalisis(fOwnerForm).btnPositionChangeStock.Hint:= 'Показать позиции с изменившимся остатком';
  end else
  begin

       TFmAnalisis(fOwnerForm).btnPriceChange.Enabled:= false;
       TFmAnalisis(fOwnerForm).btnPositionChangeAssort.Enabled:= false;
       TFmAnalisis(fOwnerForm).btnPositionChangeStock.Enabled:= false;
       //if fGridPrice.Grid.Tag = 1 then
       //   fGridPriceFilterString:='';

       TFmAnalisis(fOwnerForm).btnPriceChange.Hint:= 'Для сравнения выберите 2 любых прайс-листа одного контрагента';
       TFmAnalisis(fOwnerForm).btnPositionChangeAssort.Hint:= 'Для сравнения выберите 2 любых прайс-листа одного контрагента';
       TFmAnalisis(fOwnerForm).btnPositionChangeStock.Hint:= 'Для сравнения выберите 2 любых прайс-листа одного контрагента';

       TFmAnalisis(fOwnerForm).btnChangedPriceAll.Down:= true;
       TFmAnalisis(fOwnerForm).btnChangedPriceAssortAdds.Down:= true;

    fGridPrice.Grid.Tag:= TFmAnalisis(fOwnerForm).btnPriceList.Tag;
    TFmAnalisis(fOwnerForm).btnPriceList.Down:= true;

  end;

     if  not fTreePriceOwner.FirstFillTree then
     begin

       ChangeGridSearchEdit();

       if (_TreeView.Selected.Level=0) or (_TreeView.Selected.Level=1) then // переделать
       begin
        TFmAnalisis(fOwnerForm).sBtnSelected.Enabled:= false;
        TFmAnalisis(fOwnerForm).sBtnWithMatching.Enabled:= false;
        TFmAnalisis(fOwnerForm).sBtnNoMatching.Enabled:= false;
        TFmAnalisis(fOwnerForm).sBtnStockOnly.Enabled:= false;
        TFmAnalisis(fOwnerForm).tbPriceAnalisis.Enabled:= false;
        TFmAnalisis(fOwnerForm).tbPriceInfo.Enabled:= false;

        GridPriceFiltered(true);
       end
       else
        begin
          TFmAnalisis(fOwnerForm).sBtnSelected.Enabled:= true;
          TFmAnalisis(fOwnerForm).sBtnWithMatching.Enabled:= true;
          TFmAnalisis(fOwnerForm).sBtnNoMatching.Enabled:= true;
          TFmAnalisis(fOwnerForm).sBtnStockOnly.Enabled:= true;
          TFmAnalisis(fOwnerForm).tbPriceAnalisis.Enabled:= true;
          TFmAnalisis(fOwnerForm).tbPriceInfo.Enabled:= true;

          GridPriceFiltered();
        end;
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

procedure TwAnalisis.fTreePriceGroup_onSelectionChanged(Sender: TObject);
var
  _TreeView: TTreeView;
begin
  _TreeView:= fTreePriceGroup.Tree;

  fChartTag:=TFmAnalisis(fOwnerForm).ChartPrice.Tag;

  //if (_TreeView.SelectionCount=0) or (fGridPrice.Grid=nil) then exit;

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

       ChangeGridSearchEdit();
       GridPriceFiltered();
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

procedure TwAnalisis.fGrid_onCellClick(Sender:TObject);
var
  _oldTag: integer;
begin
  _oldTag:= fChartTag;
  fChartTag:= TFmAnalisis(fOwnerForm).ChartPrice.Tag;
  if _oldTag = 2 then
    if TFmAnalisis(fOwnerForm).pPriceAnalisis.Visible then
           ChartPriceUpdate();
end;

procedure TwAnalisis.fGridPrice_onDataChange(Sender: TObject; Field: TField);
  var
    _result: string;
  begin
     if GridCurrent.FillGridNow or (GridCurrent.Grid.DataSource.DataSet.RecordCount=0) then exit;

    _result:='';

    case TFmAnalisis(fOwnerForm).pcPriceGroup.ActivePageIndex of
      0: _result:= fTreePriceOwner.BreadCrumbs(GridCurrent.Grid.DataSource.DataSet.FieldByName('IDOWNER').AsInteger);
      1: _result:= fTreePriceGroup.BreadCrumbs(GridCurrent.Grid.DataSource.DataSet.FieldByName('IDPL_GROUP').AsInteger);
    end;

    SetStatus('['+_result+'] '+GridCurrent.Grid.DataSource.DataSet.FieldByName('PLNAME').AsString);  // NAME

     //if not TFmAnalisis(fOwnerForm).tbBtnPricePositionVersion.Marked then
     //    GridHistoryPositionFill([fGridPrice.Grid.DataSource.DataSet.FieldByName('ID').AsInteger]);

     ChangedPriceData();

end;

function TwAnalisis.ReadPricesTimeStamp(): TwTree_FOT_Data_Arr;
var
  _DataSet: TDataSet;
  i: Integer;
begin
  try
    _DataSet:= fBase.SQLReadDS('select IDFORMATS,IDOWNER, FTIMESTAMP from PRICELISTS_TIMESTAMPS order by 2 ASC,1 ASC,3 DESC').DataSet;
    _DataSet.Last;
    _DataSet.First;
    SetLength(result,_DataSet.RecordCount);

    for i:=0 to _DataSet.RecordCount-1 do
      begin
        result[i].IdFormat:= _DataSet.FieldByName('IDFORMATS').AsInteger;
        result[i].IdOwner:= _DataSet.FieldByName('IDOWNER').AsInteger;
        result[i].TimeStamp:= _DataSet.FieldByName('FTIMESTAMP').AsDateTime;
        _DataSet.Next;
      end;
    _DataSet.Close;
  except
    on E: Exception do
    begin
      Log('Ошибка [ReadPricesTimeStamp]: "' + E.Message + '"');
      __Log.SaveLogError(E);
      raise;
    end;
  end;
end;

procedure TwAnalisis.ChartPriceUpdate();

    procedure BringToFront(aChart:TChart; ASeries: TBasicChartSeries);
    var
      s: TBasicChartSeries;
    begin
      for s in aChart.Series do
        s.ZPosition := Ord(s = ASeries);
    end;

var
 i: integer;
 Chart1LineSeries1: TLineSeries;
 _SQLString: String;
 _arrDB: ArrayOfArrayVariant;
 _Chart: TChart;
 _WhereStr, _Name: string;
 _PriceFilter: TwGridPriceFiltered;
 _Delta: double;
 _IdPrice: LongInt;
 begin

   if not Assigned(GridCurrent.Grid.DataSource) or (GridCurrent.Grid.DataSource.DataSet.RecordCount=0) then exit;

   screen.Cursor:= crSQLWait;

   _Chart:= TFmAnalisis(fOwnerForm).ChartPrice;

   _Chart.ClearSeries;
   Chart1LineSeries1:= TLineSeries.Create(_Chart);


     Chart1LineSeries1.Marks.Style := TSeriesMarksStyle(6);

     Chart1LineSeries1.SeriesColor:= clBlue;
     Chart1LineSeries1.ShowPoints:= true;

     _Chart.Title.Visible:= true;

   // не допускаем перекрытие лейблов
   _Chart.LeftAxis.Marks.OverlapPolicy := opHideNeighbour;
   _Chart.BottomAxis.Marks.OverlapPolicy := opHideNeighbour;
   Chart1LineSeries1.Marks.OverlapPolicy := opHideNeighbour;

     _SQLString:= '';

     _IdPrice:= GridCurrent.Grid.DataSource.DataSet.FieldByName('ID').AsInteger;
     _Name:= GridCurrent.Grid.DataSource.DataSet.FieldByName('PLNAME').AsString;

     if fChartTag<>2 then
         fChartTag:= _Chart.Tag;


    case fChartTag of
         0: // изменение цены
           begin
             TFmAnalisis(fOwnerForm).tbChangedPrice.Down:= true;
             TFmAnalisis(fOwnerForm).tbChangedStock.Down:= false;
             TFmAnalisis(fOwnerForm).tbChangedPrice.Enabled:= true;
             TFmAnalisis(fOwnerForm).tbChangedStock.Enabled:= true;

             TFmAnalisis(fOwnerForm).gbChartPrice.Caption:='Изменение цены на позицию';
             _SQLString:= ' SELECT '
               +' PLV.FTIMESTAMP AS FTIMESTAMP, '
               +' PLV.PRICECALC AS PRICE '
               +' FROM "PL_VERSIONS" PLV    '
               +' WHERE  PLV.IDPL_ITEMS='+IntToStr(_IdPrice)
               +' ORDER BY FTIMESTAMP DESC ';
           end;
         1: // изменение остатка
            begin
              TFmAnalisis(fOwnerForm).tbChangedPrice.Down:= false;
              TFmAnalisis(fOwnerForm).tbChangedStock.Down:= true;
              TFmAnalisis(fOwnerForm).tbChangedPrice.Enabled:= true;
              TFmAnalisis(fOwnerForm).tbChangedStock.Enabled:= true;

                TFmAnalisis(fOwnerForm).gbChartPrice.Caption:='Изменение наличия на складе';
                _SQLString:= ' SELECT '
                  +' PLV.FTIMESTAMP AS FTIMESTAMP, '
                  +' (PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5) as STOCK '
                  +' FROM "PL_VERSIONS"  PLV   '
                  +' WHERE  PLV.IDPL_ITEMS='+IntToStr(_IdPrice)
                  +' ORDER BY FTIMESTAMP DESC ';
            end;

         2: // изменение стоимости склада
            begin
                _WhereStr:='';
                TFmAnalisis(fOwnerForm).tbChangedPrice.Down:= false;
                TFmAnalisis(fOwnerForm).tbChangedPrice.Enabled:= false;
                TFmAnalisis(fOwnerForm).tbChangedStock.Enabled:= false;
                TFmAnalisis(fOwnerForm).tbChangedStock.Down:= false;

                _PriceFilter:= PrepareGridPriceGroupArray(fTreePriceOwner.SelectedItems(true));

                if Assigned(_PriceFilter.GroupArray) then
                begin
                  for i:=0 to High(_PriceFilter.GroupArray) do
                    begin
                       if Length(_WhereStr)>0 then _WhereStr:= _WhereStr+' OR ';
                       _WhereStr:= _WhereStr+ 'PLV.IDOWNER='+ IntToStr(_PriceFilter.GroupArray[i]);
                    end;

                  if (Length(_WhereStr)>0) and (Length(_PriceFilter.WhereString)>0) then _WhereStr:= '('+_WhereStr+') AND('+_PriceFilter.WhereString+')';

                  //TFmAnalisis(fOwnerForm).gbChartPrice.Caption:='Изменение стоимости склада(ов) согласно выбранным версиям прайс-листов.';
                  _SQLString:= ' SELECT '
                    +' PLV.FTIMESTAMP, '
                    +' SUM(PLV.pricecalc*PLV.STOCK+PLV.pricecalc*PLV.STOCK2+PLV.pricecalc*PLV.STOCK3+PLV.pricecalc*PLV.STOCK4+PLV.pricecalc*PLV.STOCK5) AS STOCKPRICE '
                    +' FROM "PL_VERSIONS"  PLV '
                    +' WHERE  '+_WhereStr+''
                    +' GROUP BY 1'
                    +' ORDER BY PLV.FTIMESTAMP DESC ';
                end;
            end;
    end;

        _arrDB:= fBase.SQLReadArr(_SQLString);

    case fChartTag of
         2: // изменение стоимости склада
            begin
                if Assigned(_arrDB) then
                 begin
                   _Chart.Title.Text.Text:= '';

                   _Delta:=0;
                   for i:=0 to High(_arrDB) do
                     begin
                        Chart1LineSeries1.AddXY(_arrDB[i,0], _arrDB[i,1],'');
                        if i = 0 then _Delta:= double(_arrDB[i,1]);

                        //if _Delta>0 then _Delta:= 0-_Delta else _Delta:= _Delta+2*_Delta;
                     end;

                   _Delta:= _Delta-double(_arrDB[i,1]);

                   TFmAnalisis(fOwnerForm).gbChartPrice.Caption:= 'Изменение стоимости склада(ов) согласно выбранным версиям прайс-листов.'+' Разница: '+CurrToStrF(_Delta, ffCurrency, 2);
                 end;
            end
         else
           begin
               if Assigned(_arrDB) then
                begin
                  _Chart.Title.Text.Text:= _Name;

                  for i:=0 to High(_arrDB) do
                    Chart1LineSeries1.AddXY(_arrDB[i,0], _arrDB[i,1],'');
                end;
           end;
    end;

    _Chart.AddSeries(Chart1LineSeries1);
    _Chart.ZoomFull();

    screen.Cursor:= crDefault;
end;

function TwAnalisis.CreateGridPriceFiltered(aAray: ArrayOfInteger; const aString: string): TwGridPriceFiltered;
begin
  Result.GroupArray:= aAray;
  Result.WhereString:= aString;
end;

constructor TwAnalisis.Create(Sender: TObject; aBase: TwBase; aGridPrice, aGridChangePrice, aGridPriceNew, aGridChangeStock: TDBGrid; aTreePriceOwner,
  aTreePriceGroup: TTreeView);
begin
  fOwnerForm:=Sender;
  fUtils:= nil;

  fFormName:= TFmAnalisis(Sender).Name;
  fBase:= aBase;
  fIdMainOwner:= fBase.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист
  fGridPriceFilterString:='';
  __PRICE_MAX_FTIMESTAMP_ARR:= GetMaxFTimeStampPricesArr(Base);
  __TreePriceOwnerDataArr := ReadPricesTimeStamp();

  __FTIMESTAMP_KURS:= fBase.SQLReadArr('CURRENCY',['FTIMESTAMP'],'ID=2','')[0,0];

  fGridPrice:= TwDBGrid.Create(Base,aGridPrice,'');
  fGridPrice.MultiSelect:= true;
  fGridPrice.SearchEdit:=TFmAnalisis(fOwnerForm).edPriceSearch;
  fGridPrice.SearchPreventiveBtn:=TFmAnalisis(fOwnerForm).btnPricePreventSearch;
  fGridPrice.SearchSplitStringBtn:=TFmAnalisis(fOwnerForm).btnPriceSearchSplitString;
  fGridPrice.StaticTextSelection:=TFmAnalisis(fOwnerForm).st_GridSelect;
  fGridPrice.SearchEntryArray:= ['PL.NAME','PL.LABEL','PL.REMARK'];
  fGridPrice.SearchParticleArray:= ['PL.VENDORCODE','(SELECT VRESULT FROM PL_TRY_SCOD(PL.ID,''%s'') /*=*/)'];
  fGridPrice.GroupField:= 'PL.IDOWNER';
  fGridPrice.SortTitleImagesIndex:=[2,3];
  fGridPrice.Grid.Tag:=0;
  fGridPrice.onGridCellClick:=@fGrid_onCellClick;
  GridCurrent:= fGridPrice;

  fGridChangePrice:= TwDBGrid.Create(Base,aGridChangePrice,'');
  fGridChangePrice.MultiSelect:= true;
  fGridChangePrice.SearchEdit:= nil;//TFmAnalisis(fOwnerForm).edPriceSearch;
  fGridChangePrice.SearchPreventiveBtn:=TFmAnalisis(fOwnerForm).btnPricePreventSearch;
  fGridChangePrice.StaticTextSelection:=TFmAnalisis(fOwnerForm).st_GridSelect;
  fGridChangePrice.SearchEntryArray:= ['PL.NAME'];
  fGridChangePrice.SearchParticleArray:= ['PL.LABEL','PL.VENDORCODE','(SELECT VRESULT FROM PL_TRY_SCOD(PL.ID,''%s'') /*=*/)'];
  fGridChangePrice.GroupField:= 'PL.IDOWNER';
  fGridChangePrice.SortTitleImagesIndex:=[2,3];
  fGridChangePrice.onGridCellClick:=@fGrid_onCellClick;

  fGridPriceListNew:= TwDBGrid.Create(Base,aGridPriceNew,'');
  fGridPriceListNew.MultiSelect:= true;
  fGridPriceListNew.SearchEdit:= nil;//TFmAnalisis(fOwnerForm).edPriceSearch;
  fGridPriceListNew.SearchPreventiveBtn:=TFmAnalisis(fOwnerForm).btnPricePreventSearch;
  fGridPriceListNew.StaticTextSelection:=TFmAnalisis(fOwnerForm).st_GridSelect;
  fGridPriceListNew.SearchEntryArray:= ['PL.NAME','PL.LABEL'];
  fGridPriceListNew.SearchParticleArray:= ['PL.VENDORCODE','(SELECT VRESULT FROM PL_TRY_SCOD(PL.ID,''%s'') /*=*/)'];
  fGridPriceListNew.GroupField:= 'PL.IDOWNER';
  fGridPriceListNew.SortTitleImagesIndex:=[2,3];

  fGridChangeStock:= TwDBGrid.Create(Base,aGridChangeStock,'');
  fGridChangeStock.MultiSelect:= true;
  fGridChangeStock.SearchEdit:= nil;//TFmAnalisis(fOwnerForm).edPriceSearch;
  fGridChangeStock.SearchPreventiveBtn:=TFmAnalisis(fOwnerForm).btnPricePreventSearch;
  fGridChangeStock.StaticTextSelection:=TFmAnalisis(fOwnerForm).st_GridSelect;
  fGridChangeStock.SearchEntryArray:= ['PL.NAME','PL.LABEL'];
  fGridChangeStock.SearchParticleArray:= ['PL.VENDORCODE','(SELECT VRESULT FROM PL_TRY_SCOD(PL.ID,''%s'') /*=*/)'];
  fGridChangeStock.GroupField:= 'PL.IDOWNER';
  fGridChangeStock.SortTitleImagesIndex:=[2,3];
  fGridChangeStock.onGridCellClick:=@fGrid_onCellClick;

  fChartTag:=0;

  fTreePriceOwner:= TwDBTree.Create(Base,aTreePriceOwner,'OWNER','IDPARENT,NAME',[]);
  fTreePriceOwner.MultiSelect:= true;
  //fTreePriceOwner.Expanded:= false;
  fTreePriceOwner.Tree.OnSelectionChanged:= @fTreePriceOwner_onSelectionChanged;
  fTreePriceOwner.Tree.Tag:=0;

  fTreePriceGroup:= TwDBTree.Create(Base,aTreePriceGroup,'PL_GROUP','IDPARENT,ID',['IDOWNER',0]);
  fTreePriceGroup.MultiSelect:= true;
  //fTreePriceGroup.Expanded:= true;
  fTreePriceGroup.Tree.OnSelectionChanged:= @fTreePriceGroup_onSelectionChanged;
  //aTreePriceGroup
end;

destructor TwAnalisis.Destroy();
begin
  fGridPrice.Destroy();
  fGridChangePrice.Destroy();
  fGridPriceListNew.Destroy();
  fTreePriceOwner.Destroy();
  fGridChangeStock.Destroy();
  fTreePriceGroup.Destroy();
end;

function TwAnalisis.PrepareGridPriceGroupArray(const aFOTArray: TwTree_FOT_Data_Arr):TwGridPriceFiltered;
var
  i: Integer;
  _arr: ArrayOfInteger;
  _WhereFormat, _WhereTimeStamp: string;
begin
  SetLength(_arr,Length(aFOTArray));

  Result.GroupArray:=nil;
  Result.WhereString:= '';
  _WhereFormat:='';
  _WhereTimeStamp:='';

  for i:=0 to High(aFOTArray) do
    begin
      //Log(IntToStr(aFOTArray[i].IdOwner)+'|'+IntToStr(aFOTArray[i].IdFormat)+'|'+DateTimeToStr(aFOTArray[i].TimeStamp));

      _arr[i]:=aFOTArray[i].IdOwner;

      //if aFOTArray[i].IdFormat>0 then
      // begin
      //   if (Length(_WhereFormat)>0) then _WhereFormat:= _WhereFormat+ ' OR ';
      //   _WhereFormat:= _WhereFormat+' PL.IDFORMATS='+IntToStr(aFOTArray[i].IdFormat);
      // end;

      if (Float(aFOTArray[i].TimeStamp)>0) then
       begin
         if (Length(_WhereTimeStamp)>0) then _WhereTimeStamp:= _WhereTimeStamp+ ' OR ';
         _WhereTimeStamp:= _WhereTimeStamp+' PLV.FTIMESTAMP='+QuotedStr(DateTimeToStr(aFOTArray[i].TimeStamp));
       end;

    end;

  Result.GroupArray:= _arr;

  if Length(_WhereTimeStamp)>0 then
       Result.WhereString:= _WhereTimeStamp;

  //if (Length(_WhereTimeStamp)>0) then
  // if Length(Result.WhereString)>0 then Result.WhereString:= '('+Result.WhereString+') AND ('+ _WhereTimeStamp+')' else
  //    Result.WhereString:= _WhereTimeStamp;

end;

procedure TwAnalisis.GridPriceFill(aGroup: TwGridPriceFiltered);
var
  _SQLText: string;
  _OwnerArr: TwTree_FOT_Data_Arr;
begin
  if not Assigned(Base) then exit;

  GridCurrent.GroupArray:=aGroup.GroupArray;
  GridCurrent.Where:=aGroup.WhereString;

  if not Assigned(aGroup.GroupArray) then
    begin
      if Assigned(GridCurrent.Grid.DataSource) then GridCurrent.Grid.DataSource.DataSet.Close;
      GridCurrent.SQL:='';
      screen.Cursor:= crDefault;
      exit;
    end;

  GridCurrent.Grid.DataSource:= nil;
   _SQLText:='';
case GridPrice.Grid.Tag of
  0:
    begin
        TFmAnalisis(fOwnerForm).pcPrice.ActivePageIndex:=0;
        _SQLText:='SELECT PL.ID,'
          +' PL.IDOWNER, '
          +' PL.NAME AS PLNAME, '
          +' PL.IDPL_GROUP, '
          +' PL.UNIT , '
          //+' PL.SCOD, '
          +' PL.LABEL, '
          +' PL.VENDORCODE, '
          +' PL.FCOLOR, '
          +' (PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5) as STOCK, '
          +' PLV.STOCK as STOCK1, '
          +' PLV.STOCK2 as STOCK2, '
          +' PLV.STOCK3 as STOCK3, '
          +' PLV.STOCK4 as STOCK4, '
          +' PLV.STOCK5 as STOCK5, '
          +' PLV.TRANSIT, '
          //+' PL.REMARK, '
          //+' PL.FURL, '
          //+' PL.FURLPICTURE, '
          +' PLV.PRICECALC AS PRICE, '
          +' PLV.PRICECALC2 AS PRICE2, '
          +' PLV.PRICECALC3 AS PRICE3, '
          +' PLV.PRICECALC4 AS PRICE4, '
          +' PLV.PRICECALC5 AS PRICE5, '
          +' PLV.PRICECALC6 AS PRICE6, '
          +' PLV.PRICECALC7 AS PRICE7, '
          +' PLV.PRICECALC8 AS PRICE8, '
          +' PLV.PRICECALC9 AS PRICE9, '
          +' PLV.PRICECALC10 AS PRICE10, '
          +' PLV.FTIMESTAMP AS FTIMESTAMP, '
          +' OWNER.NAME AS OWNERNAME, '
          +' MTH.ID AS MTHRESULT, '
          +' FMTS.STOCKONLYINFO AS STOCKONLYINFO '
          +' FROM "PL_ITEMS" PL  '
          +' INNER JOIN "PL_VERSIONS" PLV ON (PL.ID=PLV.IDPL_ITEMS  /*and_where_string*/  )'
          +' LEFT JOIN OWNER ON PL.IDOWNER=OWNER.ID   '
          +' LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PLV.IDFORMATS) '
          +' LEFT JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS=PL.ID) ';
      if Assigned(aGroup.GroupArray) then
          _SQLText:= _SQLText+'  WHERE '+fGridPriceFilterString+'  /*group_string*/ /*and_search_string*/';

           _SQLText:= _SQLText+' ORDER BY PLV.FTIMESTAMP DESC';
    end;
  1:   // изменение цены
    begin
        _OwnerArr:= TreePriceOwner.SelectedItems(true);

        TFmAnalisis(fOwnerForm).pcPrice.ActivePageIndex:=1;

        _SQLText:='SELECT PL.ID as ID,'
          +' PL.IDOWNER, '
          +' PL.NAME AS PLNAME, '
          +' PL.IDPL_GROUP, '
          +' PL.UNIT , '
          //+' PL.SCOD, '
          +' PL.LABEL, '
          +' PL.VENDORCODE, '
          +' PL.FCOLOR, '
          +' (PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5) as STOCKNEW, '
          +' (PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5) as STOCKOLD, '
          +' PLV.STOCK as STOCK1, '
          +' PLV.STOCK2 as STOCK2, '
          +' PLV.STOCK3 as STOCK3, '
          +' PLV.STOCK4 as STOCK4, '
          +' PLV.STOCK5 as STOCK5, '
          +' PLV.TRANSIT, '
          //+' PL.REMARK, '
          //+' PL.FURL, '
          //+' PL.FURLPICTURE, '
          +' PLV.PRICECALC AS PRICENEW, '
          +' PLV2.PRICECALC AS PRICEOLD, '
          +' PLV.PRICECALC AS PRICE, '
          +' PLV.PRICECALC2 AS PRICE2, '
          +' PLV.PRICECALC3 AS PRICE3, '
          +' PLV.PRICECALC4 AS PRICE4, '
          +' PLV.PRICECALC5 AS PRICE5, '
          +' PLV.PRICECALC6 AS PRICE6, '
          +' PLV.PRICECALC7 AS PRICE7, '
          +' PLV.PRICECALC8 AS PRICE8, '
          +' PLV.PRICECALC9 AS PRICE9, '
          +' PLV.PRICECALC10 AS PRICE10, '
          +' PLV.FTIMESTAMP AS FTIMESTAMP, '
          +' OWNER.NAME AS OWNERNAME, '
          +' MTH.ID AS MTHRESULT, '
          +' FMTS.STOCKONLYINFO AS STOCKONLYINFO, '
          +' (PLV.PRICECALC-PLV2.PRICECALC) AS PRICEDELTA, '
          +' (CASE WHEN PLV.PRICECALC>PLV2.PRICECALC THEN 1 ELSE 0 END) AS PRICECHANHGELEVEL '
          +' FROM "PL_VERSIONS" PLV   '
          +' INNER JOIN "W_TMP_PL_VERSIONS" PLV2 ON (PLV.IDPL_ITEMS=PLV2.IDPL_ITEMS) '
          +' INNER JOIN "PL_ITEMS" PL ON (PL.ID=PLV.IDPL_ITEMS /*and_group_string*/ /*and_search_string*/)'
          +' LEFT JOIN OWNER ON PL.IDOWNER=OWNER.ID   '
          +' LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PLV.IDFORMATS) '
          +' LEFT JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS=PLV.ID) ';

          if Assigned(aGroup.GroupArray) then
            _SQLText:= _SQLText+'  WHERE '+fGridPriceFilterString
            +' (PLV.FTIMESTAMP='+QuotedStr(DateTimeToStr(_OwnerArr[0].TimeStamp))+') '
            +' AND (PLV2.PRICECALC <> PLV.PRICECALC) '
            +' /*and_where_string*/ ';
           _SQLText:= _SQLText+' ';
    end;

  2:   //изменение ассортимента
    begin
        _OwnerArr:= TreePriceOwner.SelectedItems(true);

        TFmAnalisis(fOwnerForm).pcPrice.ActivePageIndex:=2;
        if TFmAnalisis(fOwnerForm).btnChangedPriceAssortAdds.Down then
          begin
            _SQLText:='SELECT PL.ID AS ID,'
              +' PL.IDOWNER, '
              +' PL.NAME AS PLNAME, '
              +' PL.IDPL_GROUP, '
              +' PL.UNIT , '
              //+' PL.SCOD, '
              +' PL.LABEL, '
              +' PL.VENDORCODE, '
              +' PL.FCOLOR, '
              +' (PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5) as STOCK, '
              +' PLV.STOCK as STOCK1, '
              +' PLV.STOCK2 as STOCK2, '
              +' PLV.STOCK3 as STOCK3, '
              +' PLV.STOCK4 as STOCK4, '
              +' PLV.STOCK5 as STOCK5, '
              +' PLV.TRANSIT, '
              //+' PLV.REMARK, '
              //+' PLV.FURL, '
              //+' PLV.FURLPICTURE, '
              +' PLV.PRICECALC AS PRICE, '
              +' PLV.PRICECALC2 AS PRICE2, '
              +' PLV.PRICECALC3 AS PRICE3, '
              +' PLV.PRICECALC4 AS PRICE4, '
              +' PLV.PRICECALC5 AS PRICE5, '
              +' PLV.PRICECALC6 AS PRICE6, '
              +' PLV.PRICECALC7 AS PRICE7, '
              +' PLV.PRICECALC8 AS PRICE8, '
              +' PLV.PRICECALC9 AS PRICE9, '
              +' PLV.PRICECALC10 AS PRICE10, '
              +' PLV.FTIMESTAMP AS FTIMESTAMP, '
              +' OWNER.NAME AS OWNERNAME, '
              +' MTH.ID AS MTHRESULT, '
              +' FMTS.STOCKONLYINFO AS STOCKONLYINFO, '
              +'  1 AS ASSORTCHANGELEVEL '
              +' FROM "PL_VERSIONS" PLV   '
              +' LEFT JOIN "W_TMP_PL_VERSIONS" PLV2 ON (PLV.IDPL_ITEMS=PLV2.IDPL_ITEMS) '
              +' INNER JOIN "PL_ITEMS" PL ON (PL.ID=PLV.IDPL_ITEMS /*and_group_string*/ /*and_search_string*/)'
              +' LEFT JOIN OWNER ON PLV.IDOWNER=OWNER.ID   '
              +' LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PLV.IDFORMATS) '
              +' LEFT JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS=PL.ID) ';

              if Assigned(aGroup.GroupArray) then
                _SQLText:= _SQLText+'  WHERE '+fGridPriceFilterString
                +' (PLV.FTIMESTAMP='+QuotedStr(DateTimeToStr(_OwnerArr[0].TimeStamp))+') '
                +' AND (PLV2.IDPL_ITEMS IS NULL OR PLV.IDPL_ITEMS IS NULL) '
                +' /*and_where_string*/';
               _SQLText:= _SQLText+' ';
          end else
          begin
            _SQLText:='SELECT PL.ID,'
              +' PL.IDOWNER, '
              +' PL.NAME AS PLNAME, '
              +' PL.IDPL_GROUP, '
              +' PL.UNIT , '
              //+' PL.SCOD, '
              +' PL.LABEL, '
              +' PL.VENDORCODE, '
              +' PL.FCOLOR, '
              +' (PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5) as STOCK, '
              +' PLV.STOCK as STOCK1, '
              +' PLV.STOCK2 as STOCK2, '
              +' PLV.STOCK3 as STOCK3, '
              +' PLV.STOCK4 as STOCK4, '
              +' PLV.STOCK5 as STOCK5, '
              +' PLV.TRANSIT, '
              +' PLV.PRICECALC AS PRICE, '
              +' PLV.PRICECALC2 AS PRICE2, '
              +' PLV.PRICECALC3 AS PRICE3, '
              +' PLV.PRICECALC4 AS PRICE4, '
              +' PLV.PRICECALC5 AS PRICE5, '
              +' PLV.PRICECALC6 AS PRICE6, '
              +' PLV.PRICECALC7 AS PRICE7, '
              +' PLV.PRICECALC8 AS PRICE8, '
              +' PLV.PRICECALC9 AS PRICE9, '
              +' PLV.PRICECALC10 AS PRICE10, '
              +' PLV.FTIMESTAMP AS FTIMESTAMP, '
              +' OWNER.NAME AS OWNERNAME, '
              +' MTH.ID AS MTHRESULT, '
              +' FMTS.STOCKONLYINFO AS STOCKONLYINFO, '
              +'  0 AS ASSORTCHANGELEVEL '
              +' FROM W_TMP_PL_VERSIONS PLV   '
              +'  LEFT JOIN PL_VERSIONS PLV2 ON (PLV.IDPL_ITEMS=PLV2.IDPL_ITEMS AND PLV2.FTIMESTAMP='+QuotedStr(DateTimeToStr(_OwnerArr[0].TimeStamp))+') '
              +' INNER JOIN "PL_ITEMS" PL ON (PL.ID=PLV.IDPL_ITEMS /*and_group_string*/ /*and_search_string*/)'
              +' LEFT JOIN OWNER ON PLV.IDOWNER=OWNER.ID   '
              +' LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PLV.IDFORMATS) '
              +' LEFT JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS=PL.ID) ';

              if Assigned(aGroup.GroupArray) then
                _SQLText:= _SQLText+'  WHERE '+fGridPriceFilterString
                +' PLV2.IDPL_ITEMS IS NULL ';
                //+' (PLV.FTIMESTAMP='+QuotedStr(DateTimeToStr(_OwnerArr[1].TimeStamp))+') '
                //+' AND (PLV2.IDPL_ITEMS IS NOT NULL OR PLV.IDPL_ITEMS IS NULL) ';
//                +' /*and_where_string*/ ';
               _SQLText:= _SQLText+' ';
          end;
    end;
  3:   //изменение остатка
    begin
        _OwnerArr:= TreePriceOwner.SelectedItems(true);

        TFmAnalisis(fOwnerForm).pcPrice.ActivePageIndex:=3;

        _SQLText:='SELECT PL.ID as ID,'
          +' PL.IDOWNER, '
          +' PL.NAME AS PLNAME, '
          +' PL.IDPL_GROUP, '
          +' PL.UNIT , '
          //+' PL.SCOD, '
          +' PL.LABEL, '
          +' PL.VENDORCODE, '
          +' PL.FCOLOR, '
          +' (PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5) as STOCKNEW, '
          +' (PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5) as STOCKOLD, '
          +' PLV.STOCK as STOCK1, '
          +' PLV.STOCK2 as STOCK2, '
          +' PLV.STOCK3 as STOCK3, '
          +' PLV.STOCK4 as STOCK4, '
          +' PLV.STOCK5 as STOCK5, '
          +' PLV.TRANSIT, '
          +' PLV.PRICECALC AS PRICE, '
          +' PLV.PRICECALC2 AS PRICE2, '
          +' PLV.PRICECALC3 AS PRICE3, '
          +' PLV.PRICECALC4 AS PRICE4, '
          +' PLV.PRICECALC5 AS PRICE5, '
          +' PLV.PRICECALC6 AS PRICE6, '
          +' PLV.PRICECALC7 AS PRICE7, '
          +' PLV.PRICECALC8 AS PRICE8, '
          +' PLV.PRICECALC9 AS PRICE9, '
          +' PLV.PRICECALC10 AS PRICE10, '
          +' PLV.FTIMESTAMP AS FTIMESTAMP, '
          +' OWNER.NAME AS OWNERNAME, '
          +' MTH.ID AS MTHRESULT, '
          +' FMTS.STOCKONLYINFO AS STOCKONLYINFO, '
          +' ((PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5)-(PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5)) AS STOCKDELTA,'
          +' (((PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5)-(PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5))*PLV.PRICECALC) AS SUMDELTA,'
          +' (CASE WHEN (PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5)>(PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5) THEN 1 ELSE 0 END) AS STOCKCHANHGELEVEL '
          +' FROM "PL_VERSIONS" PLV   '
          +'  LEFT JOIN W_TMP_PL_VERSIONS PLV2 ON (PLV.IDPL_ITEMS=PLV2.IDPL_ITEMS) '
          +' INNER JOIN "PL_ITEMS" PL ON (PL.ID=PLV.IDPL_ITEMS /*and_group_string*/ /*and_search_string*/)'
          +' LEFT JOIN OWNER ON PLV.IDOWNER=OWNER.ID   '
          +' LEFT JOIN "FORMATS" FMTS ON (FMTS.ID= PLV.IDFORMATS) '
          +' LEFT JOIN "CATALOG_MATCHING" MTH ON (MTH.IDPL_ITEMS=PL.ID) ';

          if Assigned(aGroup.GroupArray) then
            _SQLText:= _SQLText+'  WHERE '+fGridPriceFilterString
            +' (PLV.FTIMESTAMP='+QuotedStr(DateTimeToStr(_OwnerArr[0].TimeStamp))+') '
            +' AND ((PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5) <> (PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5)) '
            +' /*and_where_string*/ ';
           _SQLText:= _SQLText+' ';
    end;

end;

  GridCurrent.Fill(_SQLText);

  if Assigned(GridCurrent.Grid.DataSource) then
       GridCurrent.Grid.DataSource.OnDataChange:=@fGridPrice_onDataChange;

  //ChangedPriceData(GridCurrent);

end;

procedure TwAnalisis.TreePriceOwnerFill();
begin
  fTreePriceOwner.Fill(__TreePriceOwnerDataArr);
end;


procedure TwAnalisis.TreePriceGroupFill();
begin
  fTreePriceGroup.Fill();
end;

procedure TwAnalisis.GridPriceFiltered(aGridOff: boolean);
var
  _arr: ArrayOfInteger;
  _BookMark: TBookMark;
  _TreeOwnerSelectedItems, _GridPriceFiltered: TwGridPriceFiltered;
begin
  fGridPriceFilterString:='';

  if aGridOff then
  begin
    if Assigned(GridCurrent.Grid.DataSource) then
        GridPriceFill(CreateGridPriceFiltered(nil));
    exit;
  end;

  if TFmAnalisis(fOwnerForm).sBtnStockOnly.Down then
    begin
      if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND (PLV.STOCK>0 OR PLV.STOCK2>0 OR PLV.STOCK3>0 OR PLV.STOCK4>0 OR PLV.STOCK5>0)' else
          fGridPriceFilterString:=fGridPriceFilterString+' (PLV.STOCK>0 OR PLV.STOCK2>0 OR PLV.STOCK3>0 OR PLV.STOCK4>0 OR PLV.STOCK5>0) ';
    end;

  if TFmAnalisis(fOwnerForm).sBtnWithMatching.Down then
    begin
      if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND MTH.ID>0' else
          fGridPriceFilterString:=fGridPriceFilterString+' MTH.ID>0 ';
    end;

  if TFmAnalisis(fOwnerForm).sBtnNoMatching.Down then
    begin
      if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND MTH.ID IS NULL' else
          fGridPriceFilterString:=fGridPriceFilterString+' MTH.ID IS NULL ';
    end;

  if TFmAnalisis(fOwnerForm).sBtnSelected.Down then
    begin
      _arr:= GridCurrent.SelectedRows();

      if GridCurrent.SelectedRowsCount>0 then
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

  if GridPrice.Grid.Tag = 1 then
    begin

      if TFmAnalisis(fOwnerForm).btnChangedPriceDown.Down then
        begin
          if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND (PLV.PRICECALC<PLV2.PRICECALC)' else
              fGridPriceFilterString:=fGridPriceFilterString+' (PLV.PRICECALC<PLV2.PRICECALC) ';
        end;

      if TFmAnalisis(fOwnerForm).btnChangedPriceUp.Down then
        begin
          if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND (PLV.PRICECALC>PLV2.PRICECALC)' else
              fGridPriceFilterString:=fGridPriceFilterString+' (PLV.PRICECALC>PLV2.PRICECALC) ';
        end;
    end;

  if GridPrice.Grid.Tag = 3 then
    begin

      if TFmAnalisis(fOwnerForm).btnChangedStockDown.Down then
        begin
          if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND ((PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5)<(PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5))' else
              fGridPriceFilterString:=fGridPriceFilterString+' ((PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5)<(PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5)) ';
        end;

      if TFmAnalisis(fOwnerForm).btnChangedStockUp.Down then
        begin
          if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=fGridPriceFilterString+' AND ((PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5)>(PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5))' else
              fGridPriceFilterString:=fGridPriceFilterString+' ((PLV.STOCK+PLV.STOCK2+PLV.STOCK3+PLV.STOCK4+PLV.STOCK5)>(PLV2.STOCK+PLV2.STOCK2+PLV2.STOCK3+PLV2.STOCK4+PLV2.STOCK5)) ';
        end;

    end;

  //(PL.STOCK+PL.STOCK2+PL.STOCK3)-(PL2.STOCK+PL2.STOCK2+PL2.STOCK3)

  //FMTS

  if Length(fGridPriceFilterString)>0 then  fGridPriceFilterString:=' ('+fGridPriceFilterString+') AND ';

  screen.Cursor:= crSQLWait;
  if Assigned(GridCurrent.Grid.DataSource) then
    _BookMark:= GridCurrent.Grid.DataSource.DataSet.Bookmark;

  if fTreePriceOwner.Tree.Tag =0 then
        GridPriceFill(PrepareGridPriceGroupArray(fTreePriceOwner.SelectedItems(true)))
     else
     begin
        _TreeOwnerSelectedItems:= PrepareGridPriceGroupArray(fTreePriceOwner.SelectedItems(true));
        _GridPriceFiltered:= CreateGridPriceFiltered(fTreePriceGroup.SelectedItems,_TreeOwnerSelectedItems.WhereString);
        GridPriceFill(_GridPriceFiltered)
     end;

  if Assigned(GridCurrent.Grid.DataSource) and (GridCurrent.Grid.DataSource.DataSet.RecordCount>0) then
      begin
        GridCurrent.Grid.DataSource.DataSet.Bookmark:= _BookMark;
      end;

  screen.Cursor:= crDefault;

end;

procedure TwAnalisis.TreeInfoFill();

var
  _arr, _BarcodeArr: ArrayOfArrayVariant;
begin
  if not TreeInfo.Parent.Visible then exit;
  _arr:=nil;
  if  GridCurrent.Grid.DataSource.DataSet.RecordCount>0 then
    _arr:= fBase.SQLReadArr('PL_ITEMS',['REMARK','FURL','FURLPICTURE'],'ID='+GridCurrent.Grid.DataSource.DataSet.FieldByName('ID').AsString,'');

  TreeInfo.Items[0].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('FTIMESTAMP').AsString;
  TreeInfo.Items[1].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('OWNERNAME').AsString;
  TreeInfo.Items[2].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('VENDORCODE').AsString;
  TreeInfo.Items[3].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('PLNAME').AsString;
  TreeInfo.Items[4].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('UNIT').AsString;

  _BarcodeArr:=nil;
  if Assigned(_arr) then
    _BarcodeArr:= fBase.SQLReadArr('SELECT VSCOD FROM PL_GET_SCOD('+GridCurrent.Grid.DataSource.DataSet.FieldByName('ID').AsString+',true)');
  if Assigned(_BarcodeArr) then
       TreeInfo.Items[5].Text:= VarToStr(_BarcodeArr[0,0]);

  TreeInfo.Items[6].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('LABEL').AsString;
  TreeInfo.Items[8].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[9].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE2').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[10].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE3').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[11].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE4').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[12].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE5').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[13].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE6').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[14].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE7').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[15].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE8').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[16].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE9').AsCurrency, ffCurrency, 2);
  TreeInfo.Items[17].Text:= CurrToStrF(GridCurrent.Grid.DataSource.DataSet.FieldByName('PRICE10').AsCurrency, ffCurrency, 2);

  TreeInfo.Items[19].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('STOCK1').AsString;
  TreeInfo.Items[20].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('STOCK2').AsString;
  TreeInfo.Items[21].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('STOCK3').AsString;
  TreeInfo.Items[22].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('STOCK4').AsString;
  TreeInfo.Items[23].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('STOCK5').AsString;
  TreeInfo.Items[24].Text:= VarToStr(GridCurrent.Grid.DataSource.DataSet.FieldByName('TRANSIT').AsVariant);
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

    TreeInfo.Items[28].Text:= GridCurrent.Grid.DataSource.DataSet.FieldByName('FCOLOR').AsString;
  //_arr:=nil;
end;

procedure TwAnalisis.ChangedPriceData();
begin

  if not Assigned(GridCurrent.Grid.DataSource) or (GridCurrent.Grid.DataSource.DataSet.RecordCount=0) then exit;

  if TFmAnalisis(fOwnerForm).pPriceAnalisis.Visible then
       ChartPriceUpdate();

  if TFmAnalisis(fOwnerForm).gbInfo.Visible then
       TreeInfoFill();

  TFmAnalisis(fOwnerForm).st_PriceVersion.Caption:='Дата импорта: '+GridCurrent.Grid.DataSource.DataSet.FieldByName('FTIMESTAMP').AsString+' / Курсы валют от '+__FTIMESTAMP_KURS+' ';
  if GridCurrent.Grid.DataSource.DataSet.FieldByName('FTIMESTAMP').AsDateTime < IncDay(Now,-1) then  TFmAnalisis(fOwnerForm).st_PriceVersion.Font.Color:= clRed else TFmAnalisis(fOwnerForm).st_PriceVersion.Font.Color:= clGreen;

end;

procedure TwAnalisis.ChangeGridSearchEdit();
begin

 case GridPrice.Grid.Tag of
     0:
       begin
         fCurrentGrid:= GridPrice;
         GridPrice.SearchEdit:= TFmAnalisis(fOwnerForm).edPriceSearch;
         GridPrice.SearchText:= GridPrice.SearchEdit.Text;
         GridChangePrice.SearchEdit:= nil;
         GridPriceListNew.SearchEdit:= nil;
         GridChangeStock.SearchEdit:= nil;
       end;
     1:
       begin
         fCurrentGrid:= GridChangePrice;
         GridPrice.SearchEdit:= nil;
         GridChangePrice.SearchEdit:= TFmAnalisis(fOwnerForm).edPriceSearch;
         GridChangePrice.SearchText:= GridChangePrice.SearchEdit.Text;
         GridPriceListNew.SearchEdit:= nil;
         GridChangeStock.SearchEdit:= nil;
       end;
     2:
       begin
         fCurrentGrid:= GridPriceListNew;
         GridPrice.SearchEdit:= nil;
         GridChangePrice.SearchEdit:= nil;
         GridPriceListNew.SearchEdit:= TFmAnalisis(fOwnerForm).edPriceSearch;
         GridPriceListNew.SearchText:= GridPriceListNew.SearchEdit.Text;
         GridChangeStock.SearchEdit:= nil;
       end;
     3:
       begin
         fCurrentGrid:= GridChangeStock;
         GridPrice.SearchEdit:= nil;
         GridChangePrice.SearchEdit:= nil;
         GridPriceListNew.SearchEdit:= nil;
         GridChangeStock.SearchEdit:= TFmAnalisis(fOwnerForm).edPriceSearch;
         GridChangeStock.SearchText:= GridChangeStock.SearchEdit.Text;
       end;

   end;
end;

procedure TwAnalisis.TreeOwnerExpand(aLevel: integer);
var
  _Noddy: TTreeNode;
begin
 _Noddy := fTreePriceOwner.Tree.Items[0];

 while Assigned(_Noddy) do
 begin
   if _Noddy.ImageIndex = aLevel then
     _Noddy.Expanded:= true;
     _Noddy := _Noddy.GetNext;
 end;

end;

procedure TwAnalisis.onEndThread(Sender: TObject);
begin
    try
      fProgress.ForceClose;

        if not fReport.Result then
           raise Exception.Create('Во время создания отчета произошла ошибка!');

    except
      on E: Exception do begin
         MessageDlg(E.Message,mtError, [mbOK], 0);
      end;
    end;
end;

procedure TwAnalisis.onProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
begin
 fProgress.InitBar(aProgressBarName, aValue);
end;

procedure TwAnalisis.onProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
begin
 fProgress.SetBar(aProgressBarName, aValue);
end;

procedure TwAnalisis.onStopForce(Sender: TObject);
begin
  fReport.Stop();
end;

procedure TwAnalisis.ExportData();
begin
  GridCurrent.ExportData();
end;

procedure TwAnalisis.GetPositionAnalog(aSelectedItems:ArrayOfInteger);
var
  fViewer: TwViewer;
  fWorkbookSource: TsWorkbookSource;
begin
 fViewer:= TwViewer.Create(TComponent(fOwnerForm));
 fViewer.Caption:= 'Аналоги позиций';

 fWorkbookSource:= TsWorkbookSource.Create(Application);
 fWorkbookSource.Options:= fWorkbookSource.Options+[boFileStream];

 fProgress:= TProgress.Create(TForm(fOwnerForm));
 fProgress.Caption:= 'Формирование отчета...';
 fProgress.ShowLog:= false;
 fProgress.onStopForce:= @onStopForce;
 fProgress.NoClose:= true;

 fReport:= TwReport.Create(true);
 fReport.SelectedPriceItems:= aSelectedItems;
 fReport.Base:= fBase;
 fReport.ReportModes:= rmAnalogs;
 fReport.WorkbookSource:= fWorkbookSource;
 fReport.onProgressInit:= @onProgressInit;
 fReport.onProgressUpdate:= @onProgressUpdate;
 fReport.onEndThread:= @onEndThread;

 screen.Cursor:= crSQLWait;

 fReport.start;

 try

   fProgress.ShowModal;

   fViewer.WorkbookSource:= fWorkbookSource;

   TFmAnalisis(fOwnerForm).Repaint;
   if fReport.Result then
     begin
        fViewer.ShowModal;
     end;

   //fReport.Terminate;
 finally
   screen.Cursor:=crDefault;
   if Assigned(fProgress) then
     fProgress.Free;
   fViewer.free;
 end;

end;

procedure TwAnalisis.GetCompareHorisontal(aSelectedItems, aSelectedOwners:ArrayOfInteger; aPriceBase, aPriceCompare: TPriceType);
var
  fViewer: TwViewer;
  fWorkbookSource: TsWorkbookSource;
begin
 fViewer:= TwViewer.Create(TComponent(fOwnerForm));
 fViewer.Caption:= 'Матрица цен на позиции';

 fWorkbookSource:= TsWorkbookSource.Create(Application);
 fWorkbookSource.Options:= fWorkbookSource.Options+[boFileStream];

 fProgress:= TProgress.Create(TForm(fOwnerForm));
 fProgress.Caption:= 'Формирование отчета...';
 fProgress.ShowLog:= false;
 fProgress.onStopForce:= @onStopForce;
 fProgress.NoClose:= true;

 fReport:= TwReport.Create(true);
 fReport.SelectedPriceItems:= aSelectedItems;
 fReport.SelectedOwners:= aSelectedOwners;
 fReport.PriceBase:= aPriceBase;
 fReport.PriceCompare:= aPriceCompare;
 fReport.Base:= fBase;
 fReport.ReportModes:= rmCompareHorisontal;
 fReport.WorkbookSource:= fWorkbookSource;
 fReport.onProgressInit:= @onProgressInit;
 fReport.onProgressUpdate:= @onProgressUpdate;
 fReport.onEndThread:= @onEndThread;

 screen.Cursor:= crSQLWait;

 fReport.start;

 try

   fProgress.ShowModal;

   fViewer.WorkbookSource:= fWorkbookSource;

   TFmAnalisis(fOwnerForm).Repaint;
   if fReport.Result then
     begin
        fViewer.ShowModal;
     end;

   //fReport.Terminate;
 finally
   screen.Cursor:=crDefault;
   if Assigned(fProgress) then
     fProgress.Free;
   fViewer.free;
 end;

end;

end.

