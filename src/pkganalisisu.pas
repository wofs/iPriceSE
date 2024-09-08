unit pkgAnalisisU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, FmReportU, fpsexport, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls, ComCtrls, StdCtrls, Buttons, DBGrids, Menus,
  TAGraph, TAIntervalSources, TASeries, TAChartCombos, TATools, TATransformations, TANavigation, TADataTools,
  wBaseU, wLogU, mAnalisisU, UtilsU, wFuncU, wTViewerSpreadsheetU, wTypesU,
  FmMatchingAddU,
  Grids, LCLIntf, Clipbrd, db, DateUtils,
  FmWaitU, XMLPropStorage
  ;

type

  { TFmAnalisis }

  TFmAnalisis = class(TForm)
    btnChangedPriceAll: TSpeedButton;
    btnChangedStockAll: TSpeedButton;
    btnChangedPriceAssortAdds: TSpeedButton;
    btnChangedStockUp: TSpeedButton;
    btnChangedPriceUp: TSpeedButton;
    btnChangedPriceDown: TSpeedButton;
    btnChangedPriceAssortDel: TSpeedButton;
    btnChangedStockDown: TSpeedButton;
    btnPriceEditSearchClear: TSpeedButton;
    btnPricePreventSearch: TSpeedButton;
    btnPriceSearchSplitString: TSpeedButton;
    ChartPrice: TChart;
    ChartToolset1: TChartToolset;
    ChartToolset1PanAny: TPanDragTool;
    ChartToolset1ZoomMouseWheelTool1: TZoomMouseWheelTool;
    edPriceSearch: TComboBox;
    DateTimeIntervalChartSource1: TDateTimeIntervalChartSource;
    GridImageList: TImageList;
    GridPriceList: TDBGrid;
    gbChartPrice: TGroupBox;
    gbInfo: TGroupBox;
    GridPriceListChangeStock: TDBGrid;
    GridPriceListNew: TDBGrid;
    GridPriceListChangePrice: TDBGrid;
    ImageListTree: TImageList;
    ImageListTreeGroup: TImageList;
    Images16: TImageList;
    ImagesTreeInfo: TImageList;
    lbPriceFind: TLabel;
    mChangedPrice: TMenuItem;
    mChangedStock: TMenuItem;
    mClipboardVNLP: TMenuItem;
    mViewAnalisSuperPrice: TMenuItem;
    mViewAnalogs: TMenuItem;
    mExportInSpreadsheet: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem6: TMenuItem;
    mMatchingAdd: TMenuItem;
    MenuItem2: TMenuItem;
    mClipboardVNL: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem5: TMenuItem;
    MenuTreeInfo: TPopupMenu;
    mGoToGroup: TMenuItem;
    mGridPriceList: TPopupMenu;
    mInfoPrice: TMenuItem;
    mKurs: TPopupMenu;
    mKursImages: TImageList;
    mPriceAnalisis: TMenuItem;
    mSelectAll: TMenuItem;
    mSelectAllPositionInGroup: TMenuItem;
    mSelectionClear: TMenuItem;
    mTGPAddToCatalog: TMenuItem;
    mTreeGroupPrice: TPopupMenu;
    mTreeInfoCopyOne: TMenuItem;
    pPriceChangeBtns: TPanel;
    pcPrice: TPageControl;
    pAnalisis: TPanel;
    pInfo: TPanel;
    pGridPrice: TPanel;
    pPriceBtnLeft: TPanel;
    pPriceBtns: TPanel;
    pcPriceArc: TPageControl;
    Panel1: TPanel;
    pcPriceGroup: TPageControl;
    mChartPrice: TPopupMenu;
    pPrice: TPanel;
    pPriceChangeBtns1: TPanel;
    pPriceChangeBtns2: TPanel;
    pPriceDateTime_Kurs: TPanel;
    pPriceFind: TPanel;
    pPriceGroup: TPanel;
    pPriceAnalisis: TPanel;
    sBtnKurs: TSpeedButton;
    sBtnSelected: TSpeedButton;
    sBtnStockOnly: TSpeedButton;
    sBtnWithMatching: TSpeedButton;
    btnPriceList: TSpeedButton;
    btnPriceChange: TSpeedButton;
    btnPositionChangeAssort: TSpeedButton;
    btnPositionChangeStock: TSpeedButton;
    sBtnNoMatching: TSpeedButton;
    Splitter2: TSplitter;
    SplitterInfo: TSplitter;
    SplitterPriceAnalisis: TSplitter;
    st_GridSelect: TStaticText;
    st_PriceVersion: TStaticText;
    TabGroup: TTabSheet;
    TabKontr: TTabSheet;
    tbPriceAnalisis: TToolButton;
    TabSheet1: TTabSheet;
    TabSheet3: TTabSheet;
    tbPriceList: TTabSheet;
    TabSheet2: TTabSheet;
    tbPosition: TToolBar;
    tbChangedPrice: TToolButton;
    tbPrice1: TToolBar;
    tbPriceInfo: TToolButton;
    tbPricesArc: TTabSheet;
    tbChangedStock: TToolButton;
    tbTree: TToolBar;
    tbTree1: TToolBar;
    tbTreeBtnExpand: TToolButton;
    tbTreeBtnSort: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    tbLevelOwner: TToolButton;
    ToolButton3: TToolButton;
    tbLevelFormat: TToolButton;
    ToolButton4: TToolButton;
    tbTreeBtnShowChild: TToolButton;
    ToolButton6: TToolButton;
    tbLevelMonth: TToolButton;
    TreeGroupOwner: TTreeView;
    TreeGroupPrice: TTreeView;
    TreeViewInfo: TTreeView;
    XMLPropStorage1: TXMLPropStorage;
    procedure btnChangedPriceAllClick(Sender: TObject);
    procedure btnPriceEditSearchClearClick(Sender: TObject);
    procedure btnPriceListClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure GridPriceListChangePriceDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure GridPriceListChangeStockDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure GridPriceListDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure GridPriceListNewDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure mChartPricePopup(Sender: TObject);
    procedure mClipboardVNLClick(Sender: TObject);
    procedure mClipboardVNLPClick(Sender: TObject);
    procedure mExportInSpreadsheetClick(Sender: TObject);
    procedure mGoToGroupClick(Sender: TObject);
    procedure mGridPriceListPopup(Sender: TObject);
    procedure mMatchingAddClick(Sender: TObject);
    procedure mSelectAllClick(Sender: TObject);
    procedure mSelectAllPositionInGroupClick(Sender: TObject);
    procedure mSelectionClearClick(Sender: TObject);
    procedure mTGPAddToCatalogClick(Sender: TObject);
    procedure mTreeInfoCopyOneClick(Sender: TObject);
    procedure mViewAnalisSuperPriceClick(Sender: TObject);
    procedure mViewAnalogsClick(Sender: TObject);
    procedure pcPriceGroupChange(Sender: TObject);
    procedure pcPriceGroupResize(Sender: TObject);
    procedure pGridPriceResize(Sender: TObject);
    procedure pInfoResize(Sender: TObject);
    procedure pPriceAnalisisResize(Sender: TObject);
    procedure pPriceResize(Sender: TObject);
    procedure sBtnKursClick(Sender: TObject);
    procedure sBtnNoMatchingClick(Sender: TObject);
    procedure sBtnSelectedClick(Sender: TObject);
    procedure sBtnStockOnlyClick(Sender: TObject);
    procedure sBtnWithMatchingClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure tbLevelFormatClick(Sender: TObject);
    procedure tbLevelMonthClick(Sender: TObject);
    procedure tbLevelOwnerClick(Sender: TObject);
    procedure tbPriceAnalisisClick(Sender: TObject);
    procedure tbChangedPriceClick(Sender: TObject);
    procedure tbChangedStockClick(Sender: TObject);
    procedure tbPriceInfoClick(Sender: TObject);
    procedure tbTreeBtnExpandClick(Sender: TObject);
    procedure tbTreeBtnShowChildClick(Sender: TObject);
    procedure tbTreeBtnSortClick(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupPriceGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupPriceGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeViewInfoAdvancedCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; Stage: TCustomDrawStage; var PaintImages,
      DefaultDraw: Boolean);
    procedure TreeViewInfoDblClick(Sender: TObject);
    procedure TreeViewInfoGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure XMLPropStorage1SavingProperties(Sender: TObject);
  private
    fBase: TwBase;
    fFormID:string;
    Analisis: TwAnalisis;
    fFormReport: TFmReport;
    procedure mKursClick(Sender: TObject);
  public
    procedure SetStatus(_Text:string);
  end;

var
  FmAnalisis: TFmAnalisis;
  FmMatchingAdd: TFmMatchingAdd;
implementation

{$R *.lfm}

{ TFmAnalisis }

procedure TFmAnalisis.FormCreate(Sender: TObject);
begin
   try
     screen.Cursor:= crSQLWait;
     fFormID:= Self.Name;
     fBase:= TwBase.Create(self);
     Analisis:= TwAnalisis.Create(self,fBase,GridPriceList,GridPriceListChangePrice,GridPriceListNew,GridPriceListChangeStock, TreeGroupOwner,TreeGroupPrice);
     Analisis.TreeInfo:= TreeViewInfo;
     Analisis.GridPriceFill(Analisis.CreateGridPriceFiltered(nil));
     Analisis.TreePriceOwnerFill();
     screen.Cursor:= crDefault;
     ChartPrice.Tag:= 0;
     pInfo.Width:= 0;
     pcPrice.ShowTabs:= false;
     pcPrice.ActivePageIndex:=0;

     SetStatus('Плагин загружен успешно.');
   except
     on E: Exception do
     begin
         screen.Cursor:= crDefault;
         SetStatus('Сбой инициализации плагина.');
         wLog(fFormID,'Ошибка [FmCreate]: "' + E.Message + '"');
         wLog(fFormID,'Сбой инициализации плагина.');
         ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');
      end;
   end;
   //fAnalisis.TreePriceOwnerWithDatesFill();
end;

procedure TFmAnalisis.btnPriceEditSearchClearClick(Sender: TObject);
begin
  edPriceSearch.Clear;
  edPriceSearch.OnChange(edPriceSearch);
end;

procedure TFmAnalisis.btnChangedPriceAllClick(Sender: TObject);
begin
   Analisis.GridPriceFiltered();
end;

procedure TFmAnalisis.btnPriceListClick(Sender: TObject);
begin
  Analisis.GridPrice.Grid.Tag:= TSpeedButton(Sender).Tag;
  Analisis.ChangeGridSearchEdit();
  pcPriceGroupChange(self);
  Analisis.GridPriceFiltered();
end;

procedure TFmAnalisis.FormDestroy(Sender: TObject);
begin
  // выгружаем форму добавлени соответствий
  if Assigned(FmMatchingAdd) then
      FmMatchingAdd.GridDataSet:= nil;

   fBase.Destroy();
   Analisis.Destroy();
end;

procedure TFmAnalisis.GridPriceListChangePriceDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  _ColumnText: string;
  _ImageIndex: integer;
begin
 if (gdFocused in State) then // если строка не выделена, то
   begin
     TDBGrid(Sender).Canvas.Brush.Color:= TDBGrid(Sender).FixedHotColor;
     TDBGrid(Sender).Canvas.Font.Color:= clBlack;

     TDBGrid(Sender).Canvas.FillRect(Rect);
     TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
   end;

 if Column.FieldName = 'ID' then
  begin
    if Assigned(TDBGrid(Sender).TitleImageList) then
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);
        TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

        if Analisis.GridChangePrice.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         begin
                  _ImageIndex:= 0;
                  // А теперь пусть ImageList нарисует ее на канве DBGrid'а
                  TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
         end;
     end else
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);

        if Analisis.GridChangePrice.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         _ColumnText:='+'
            else
         _ColumnText:=' ';

         TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
     end;
  end;


    if Column.FieldName = 'PRICECHANHGELEVEL' then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

       if TDBGrid(Sender).DataSource.DataSet.FieldByName('PRICECHANHGELEVEL').AsInteger=1 then
          TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, 4)
       else
          TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, 5);
    end;

    if Column.FieldName = 'PRICEDELTA' then
     begin
       with TDBGrid(Sender).Canvas do
          if TDBGrid(Sender).DataSource.DataSet.FieldByName('PRICEDELTA').AsFloat<0 then
             begin
               FillRect(Rect);
               font.Color:=clGreen;
               TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
             end else
             begin
               FillRect(Rect);
               font.Color:=clRed;
               TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
             end;
     end;


   if Column.FieldName = 'MTHRESULT' then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

       if TDBGrid(Sender).DataSource.DataSet.FieldByName('MTHRESULT').AsInteger>0 then
        begin
     	    _ImageIndex:= 1;
     	    // А теперь пусть ImageList нарисует ее на канве DBGrid'а
          TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
        end;
    end;

   if Column.FieldName = 'STOCK' then
    begin
      with TDBGrid(Sender).Canvas do
         if TDBGrid(Sender).DataSource.DataSet.FieldByName('STOCKONLYINFO').AsInteger>0 then
            begin
              FillRect(Rect);
              if gdSelected in State
                 then font.Color:=clRed
                 else font.Color:=clBlue;
              TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
            end;
    end;
   Analisis.GridCurrent.HighLightText(Sender,'PLNAME', Rect,DataCol,Column,State); // подсветить часть текста;
end;

procedure TFmAnalisis.GridPriceListChangeStockDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  _ColumnText: string;
  _ImageIndex: integer;
begin
 if (gdFocused in State) then // если строка не выделена, то
   begin
     TDBGrid(Sender).Canvas.Brush.Color:= TDBGrid(Sender).FixedHotColor;
     TDBGrid(Sender).Canvas.Font.Color:= clBlack;

     TDBGrid(Sender).Canvas.FillRect(Rect);
     TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
   end;

 if Column.FieldName = 'ID' then
  begin
    if Assigned(TDBGrid(Sender).TitleImageList) then
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);
        TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

        if Analisis.GridChangeStock.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         begin
                  _ImageIndex:= 0;
                  // А теперь пусть ImageList нарисует ее на канве DBGrid'а
                  TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
         end;
     end else
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);

        if Analisis.GridChangeStock.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         _ColumnText:='+'
            else
         _ColumnText:=' ';

         TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
     end;
  end;


    if Column.FieldName = 'STOCKCHANHGELEVEL' then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

       if TDBGrid(Sender).DataSource.DataSet.FieldByName('STOCKCHANHGELEVEL').AsInteger=1 then
          TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, 6)
       else
          TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, 7);
    end;

    if Column.FieldName = 'STOCKDELTA' then
     begin
       with TDBGrid(Sender).Canvas do
          if TDBGrid(Sender).DataSource.DataSet.FieldByName('STOCKDELTA').AsFloat>0 then
             begin
               FillRect(Rect);
               font.Color:=clGreen;
               TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
             end else
             begin
               FillRect(Rect);
               font.Color:=clRed;
               TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
             end;
     end;


   if Column.FieldName = 'MTHRESULT' then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

       if TDBGrid(Sender).DataSource.DataSet.FieldByName('MTHRESULT').AsInteger>0 then
        begin
     	    _ImageIndex:= 1;
     	    // А теперь пусть ImageList нарисует ее на канве DBGrid'а
          TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
        end;
    end;

   if Column.FieldName = 'STOCK' then
    begin
      with TDBGrid(Sender).Canvas do
         if TDBGrid(Sender).DataSource.DataSet.FieldByName('STOCKONLYINFO').AsInteger>0 then
            begin
              FillRect(Rect);
              if gdSelected in State
                 then font.Color:=clRed
                 else font.Color:=clBlue;
              TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
            end;
    end;

   Analisis.GridCurrent.HighLightText(Sender,'PLNAME', Rect,DataCol,Column,State); // подсветить часть текста;
end;

procedure TFmAnalisis.GridPriceListDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  _ColumnText: string;
  _ImageIndex: integer;
begin
 if (gdFocused in State) then // если строка не выделена, то
   begin
     TDBGrid(Sender).Canvas.Brush.Color:= TDBGrid(Sender).FixedHotColor;
     TDBGrid(Sender).Canvas.Font.Color:= clBlack;

     TDBGrid(Sender).Canvas.FillRect(Rect);
     TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
   end;

 if Column.FieldName = 'ID' then
  begin
    if Assigned(TDBGrid(Sender).TitleImageList) then
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);
        TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

        if Analisis.GridPrice.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         begin
                  _ImageIndex:= 0;
                  // А теперь пусть ImageList нарисует ее на канве DBGrid'а
                  TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
         end;
     end else
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);

        if Analisis.GridPrice.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         _ColumnText:='+'
            else
         _ColumnText:=' ';

         TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
     end;
  end;

   if Column.FieldName = 'MTHRESULT' then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

       if TDBGrid(Sender).DataSource.DataSet.FieldByName('MTHRESULT').AsInteger>0 then
        begin
     	    _ImageIndex:= 1;
     	    // А теперь пусть ImageList нарисует ее на канве DBGrid'а
          TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
        end;
    end;

   if Column.FieldName = 'STOCK' then
    begin
      with TDBGrid(Sender).Canvas do
         if TDBGrid(Sender).DataSource.DataSet.FieldByName('STOCKONLYINFO').AsInteger>0 then
            begin
              FillRect(Rect);
              if gdSelected in State
                 then font.Color:=clRed
                 else font.Color:=clBlue;
              TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
            end;
    end;

   if Column.FieldName = 'FTIMESTAMP' then
    begin
      with TDBGrid(Sender).Canvas do
         begin
            FillRect(Rect);
         if TDBGrid(Sender).DataSource.DataSet.FieldByName('FTIMESTAMP').AsDateTime < IncDay(Now,-1) then
            font.Color:=clRed else
            font.Color:=clGreen;
         end;
         TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
    end;
   Analisis.GridCurrent.HighLightText(Sender,'PLNAME', Rect,DataCol,Column,State); // подсветить часть текста;
   Analisis.GridCurrent.HighLightText(Sender,'LABEL', Rect,DataCol,Column,State); // подсветить часть текста
end;

procedure TFmAnalisis.GridPriceListNewDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  _ColumnText: string;
  _ImageIndex: integer;
begin
 if (gdFocused in State) then // если строка не выделена, то
   begin
     TDBGrid(Sender).Canvas.Brush.Color:= TDBGrid(Sender).FixedHotColor;
     TDBGrid(Sender).Canvas.Font.Color:= clBlack;

     TDBGrid(Sender).Canvas.FillRect(Rect);
     TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
   end;

 if Column.FieldName = 'ID' then
  begin
    if Assigned(TDBGrid(Sender).TitleImageList) then
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);
        TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

        if Analisis.GridPriceListNew.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         begin
                  _ImageIndex:= 0;
                  // А теперь пусть ImageList нарисует ее на канве DBGrid'а
                  TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
         end;
     end else
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);

        if Analisis.GridPriceListNew.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         _ColumnText:='+'
            else
         _ColumnText:=' ';

         TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
     end;
  end;


   if Column.FieldName = 'MTHRESULT' then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

       if TDBGrid(Sender).DataSource.DataSet.FieldByName('MTHRESULT').AsInteger>0 then
        begin
     	    _ImageIndex:= 1;
     	    // А теперь пусть ImageList нарисует ее на канве DBGrid'а
          TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
        end;
    end;

   if Column.FieldName = 'FTIMESTAMP' then
    begin
      with TDBGrid(Sender).Canvas do
         if TDBGrid(Sender).DataSource.DataSet.FieldByName('ASSORTCHANGELEVEL').AsFloat=1 then
            begin
              FillRect(Rect);
              font.Color:=clGreen;
              TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
            end else
            begin
              FillRect(Rect);
              font.Color:=clRed;
              TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
            end;
    end;

   Analisis.GridCurrent.HighLightText(Sender,'PLNAME', Rect,DataCol,Column,State); // подсветить часть текста;
   Analisis.GridCurrent.HighLightText(Sender,'LABEL', Rect,DataCol,Column,State); // подсветить часть текста
end;

procedure TFmAnalisis.mChartPricePopup(Sender: TObject);
begin


   mChartPrice.Items[0].Enabled:= tbChangedPrice.Enabled;
   mChartPrice.Items[1].Enabled:= tbChangedStock.Enabled;

   mChartPrice.Items[0].Checked:= tbChangedPrice.Down;
   mChartPrice.Items[1].Checked:= tbChangedStock.Down;

end;

procedure TFmAnalisis.mClipboardVNLClick(Sender: TObject);
begin
   Analisis.GridCurrent.CopyToClipboard(['VENDORCODE', 'PLNAME', 'UNIT', 'LABEL'], [''],'PL.ID');
end;

procedure TFmAnalisis.mClipboardVNLPClick(Sender: TObject);
begin
   Analisis.GridCurrent.CopyToClipboard(['VENDORCODE', 'PLNAME', 'UNIT', 'LABEL', 'PRICE'], ['PRICE'],'PL.ID');
end;

procedure TFmAnalisis.mExportInSpreadsheetClick(Sender: TObject);
begin
  Analisis.ExportData();
end;

procedure TFmAnalisis.mGoToGroupClick(Sender: TObject);
var
  _GridDataset: TDataSet;
  _ParentID: integer;
  _ID: integer;
begin


  if TreeGroupOwner.Selected.Text = '' then exit;

  _GridDataset:= Analisis.GridCurrent.Grid.DataSource.DataSet;
  _ID:= _GridDataset.FieldByName('ID').AsInteger;
  try
    case pcPriceGroup.ActivePageIndex of
        0:
          begin
             _ParentID:= _GridDataset.FieldByName('IDOWNER').AsInteger ;
             Analisis.TreePriceOwner.FindNodeWithDataInt(_ParentID);
          end;
        1:
          begin
             _ParentID:= _GridDataset.FieldByName('IDPL_GROUP').AsInteger ;
             Analisis.TreePriceGroup.FindNodeWithDataInt(_ParentID);
          end;
    end;
  finally

  if _GridDataset.RecordCount>0 then
    _GridDataset.Locate('ID',_ID,[]);
  end;
end;

procedure TFmAnalisis.mGridPriceListPopup(Sender: TObject);
begin
  if not Assigned(Analisis.GridCurrent.Grid.DataSource) then
    begin
       mSelectAll.Enabled:= false;
       mSelectionClear.Enabled:= false;
       mGoToGroup.Enabled:= false;
       mPriceAnalisis.Enabled:= false;
       mInfoPrice.Enabled:= false;
       mMatchingAdd.Enabled:= false;

    end else
    begin
      mSelectAll.Enabled:= true;
      mSelectionClear.Enabled:= true;
      mGoToGroup.Enabled:= true;
      mPriceAnalisis.Enabled:= true;
      mInfoPrice.Enabled:= true;
      mMatchingAdd.Enabled:= true;
    end;

    mPriceAnalisis.Checked:= tbPriceAnalisis.Down;
    mInfoPrice.Checked:= tbPriceInfo.Down;
end;

procedure TFmAnalisis.mMatchingAddClick(Sender: TObject);
begin
 FmMatchingAdd:= TFmMatchingAdd.Create(Application);
 FmMatchingAdd.SelectedRows:= Analisis.GridCurrent.SelectedRows;
 FmMatchingAdd.GridDataSet:=Analisis.GridCurrent.Grid.DataSource.DataSet;
 FmMatchingAdd.Show;
end;

procedure TFmAnalisis.mSelectAllClick(Sender: TObject);
begin
 Analisis.GridCurrent.Grid.Cursor:= crSQLWait;
 Application.ProcessMessages;
 Analisis.GridCurrent.SelectAll:= true;
 Analisis.GridCurrent.Grid.Cursor:= crDefault;
end;

procedure TFmAnalisis.mSelectAllPositionInGroupClick(Sender: TObject);
begin
 Analisis.GridCurrent.Grid.Cursor:= crSQLWait;
 Application.ProcessMessages;
 Analisis.GridCurrent.SelectAll:= true;
 Analisis.GridCurrent.Grid.Cursor:= crDefault;
end;

procedure TFmAnalisis.mSelectionClearClick(Sender: TObject);
var
  _SelCount: Integer;
begin

   _SelCount:=  Analisis.GridCurrent.SelectedRowsCount;

   if _SelCount>0 then
     begin
       if MessageDlg('Сбросить выделение ('+IntToStr(_SelCount)+') позиций?',mtConfirmation, mbOKCancel, 0) = mrOK then
        begin
           Analisis.GridCurrent.SelectAll:= false;
        end;
     end;
end;

procedure TFmAnalisis.mTGPAddToCatalogClick(Sender: TObject);
begin
   if MessageDlg('Добавить выделенные категории с подкатегориями и товарами в каталог?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

   PricesLoadNomenclatureEditMassForm(Sender,fBase,Analisis.TreePriceGroup);
end;

procedure TFmAnalisis.mTreeInfoCopyOneClick(Sender: TObject);
begin
  Clipboard.AsText:= TreeViewInfo.Selected.Text;
end;

procedure TFmAnalisis.mViewAnalisSuperPriceClick(Sender: TObject);
begin
  fFormReport:= TFmReport.Create(self);
  fFormReport.Base:= fBase;

  try
    fFormReport.ShowModal;
    if fFormReport.ModalResult = mrOK then
       Analisis.GetCompareHorisontal(Analisis.GridCurrent.SelectedRows(),fFormReport.SelectedItems,fFormReport.PriceBase,fFormReport.PriceCompare);

  finally
    fFormReport.Free;
  end;
end;

procedure TFmAnalisis.mViewAnalogsClick(Sender: TObject);
begin
  Analisis.GetPositionAnalog(Analisis.GridCurrent.SelectedRows());
end;

procedure TFmAnalisis.pcPriceGroupChange(Sender: TObject);
var
  _IdOwner: Integer;
begin
  _IdOwner:= Analisis.TreePriceOwner.SelectedItems(true)[0].IdOwner;
case pcPriceGroup.ActivePageIndex of
    0:
    begin
      if (Sender is TPageControl) then
            Analisis.TreePriceOwner.Tree.Tag:= 0;

      Analisis.GridCurrent.GroupField:= 'PL.IDOWNER';
      Analisis.GridCurrent.GroupArray:= nil;
      Analisis.GridCurrent.Where:='';
      //Analisis.GridPriceFiltered();
      //Analisis.TreePriceGroup.SetOwner:= Analisis.TreePriceOwner.SelectedItems(true)[0].IdOwner;
      //Analisis.Tr();
    end;
    1:
      begin
        Analisis.GridCurrent.GroupField:= 'PL.IDPL_GROUP';
        Analisis.GridCurrent.GroupArray:= nil;
        Analisis.GridCurrent.Where:='PL.IDOWNER='+IntToStr(_IdOwner);
        Analisis.TreePriceGroup.SetOwner:= _IdOwner;
        if (Sender is TPageControl) then
         begin
           Analisis.TreePriceOwner.Tree.Tag:= 1;
           Analisis.TreePriceGroupFill();
         end;
      end;
  end;
end;

procedure TFmAnalisis.pcPriceGroupResize(Sender: TObject);
begin
   pcPriceGroup.Repaint;
end;

procedure TFmAnalisis.pGridPriceResize(Sender: TObject);
begin
    pGridPrice.Repaint;
end;

procedure TFmAnalisis.pInfoResize(Sender: TObject);
begin
  pInfo.Repaint;
end;

procedure TFmAnalisis.pPriceAnalisisResize(Sender: TObject);
begin
   pPriceAnalisis.Repaint;
end;

procedure TFmAnalisis.pPriceResize(Sender: TObject);
begin
   pPrice.Repaint;
end;

procedure TFmAnalisis.sBtnKursClick(Sender: TObject);
var
  _arr: ArrayOfArrayVariant;
  _PopupMenu: TPopupMenu;
  _PopupMenuKurs_Data,
    _PopupMenuKurs_USD,
    _PopupMenuKurs_EUR,
    _PopupMenuKurs_KZT,
    _PopupMenuKurs_UAH,
    _PopupMenuKurs_Spliter: TMenuItem;

  i: Integer;
begin
_arr:= fBase.SQLReadArr('CURRENCY',['ID','KURS','FTIMESTAMP'],'ID<>1','ID');

if Assigned(_arr) and Assigned(sBtnKurs.PopupMenu) then
begin
   _PopupMenu:= sBtnKurs.PopupMenu;
   _PopupMenu.Items.Clear;

   _PopupMenuKurs_Data:= TMenuItem.Create(_PopupMenu);
   _PopupMenuKurs_Data.Name:='Data';
   _PopupMenuKurs_Data.Caption:= _arr[0,2];
   _PopupMenuKurs_Data.ImageIndex:=0;
   _PopupMenuKurs_Data.OnClick:=@mKursClick;
   _PopupMenu.Items.Add(_PopupMenuKurs_Data);

   _PopupMenuKurs_Spliter:= TMenuItem.Create(_PopupMenu);
   _PopupMenuKurs_Spliter.Caption:= '-';
   _PopupMenu.Items.Add(_PopupMenuKurs_Spliter);

   for i:=0 to High(_arr) do
       begin
          case i of
            0:
              begin
                _PopupMenuKurs_USD:= TMenuItem.Create(_PopupMenu);
                _PopupMenuKurs_USD.Name:='USD';
                _PopupMenuKurs_USD.Caption:= _arr[i,1];
                _PopupMenuKurs_USD.ImageIndex:= 1;
                _PopupMenuKurs_USD.OnClick:=@mKursClick;
                _PopupMenu.Items.Add(_PopupMenuKurs_USD);
              end;
            1:
              begin
                _PopupMenuKurs_EUR:= TMenuItem.Create(_PopupMenu);
                _PopupMenuKurs_EUR.Name:='EUR';
                _PopupMenuKurs_EUR.Caption:= _arr[i,1];
                _PopupMenuKurs_EUR.ImageIndex:= 2;
                _PopupMenuKurs_EUR.OnClick:=@mKursClick;
                _PopupMenu.Items.Add(_PopupMenuKurs_EUR);
              end;
            2:
              begin
                 _PopupMenuKurs_KZT:= TMenuItem.Create(_PopupMenu);
                 _PopupMenuKurs_KZT.Name:='KZT';
                 _PopupMenuKurs_KZT.Hint:= 'KZT';
                 _PopupMenuKurs_KZT.Caption:= _arr[i,1];
                 _PopupMenuKurs_KZT.ImageIndex:= 3;
                 _PopupMenuKurs_KZT.OnClick:=@mKursClick;
                 _PopupMenu.Items.Add(_PopupMenuKurs_KZT);
              end;
            3:
              begin
                 _PopupMenuKurs_UAH:= TMenuItem.Create(_PopupMenu);
                 _PopupMenuKurs_UAH.Name:='UAH';
                 _PopupMenuKurs_UAH.Caption:= _arr[i,1];
                 _PopupMenuKurs_UAH.ImageIndex:= 3;
                 _PopupMenuKurs_UAH.OnClick:=@mKursClick;
                 _PopupMenu.Items.Add(_PopupMenuKurs_UAH);
              end;
          end;
       end;

   _PopupMenu.PopUp;
end;
end;

procedure TFmAnalisis.sBtnNoMatchingClick(Sender: TObject);
begin
  Analisis.GridPriceFiltered();
end;

procedure TFmAnalisis.sBtnSelectedClick(Sender: TObject);
begin
 Analisis.GridPriceFiltered();
end;

procedure TFmAnalisis.sBtnStockOnlyClick(Sender: TObject);
begin
 Analisis.GridPriceFiltered();
end;

procedure TFmAnalisis.sBtnWithMatchingClick(Sender: TObject);
begin
 Analisis.GridPriceFiltered();
end;

procedure TFmAnalisis.SpeedButton1Click(Sender: TObject);
begin
  SplitterInfo.Visible:= false;
  gbInfo.Visible:= false;
end;

procedure TFmAnalisis.tbLevelFormatClick(Sender: TObject);
begin
  Analisis.TreePriceOwner.Tree.FullCollapse;
  Analisis.TreePriceOwner.Tree.Items[0].Expanded:= true;
  Analisis.TreeOwnerExpand(0);
  Analisis.TreeOwnerExpand(1);
end;

procedure TFmAnalisis.tbLevelMonthClick(Sender: TObject);
begin
  Analisis.TreePriceOwner.Tree.FullCollapse;
  Analisis.TreePriceOwner.Tree.Items[0].Expanded:= true;
  Analisis.TreeOwnerExpand(0);
  Analisis.TreeOwnerExpand(1);
  Analisis.TreeOwnerExpand(2);
end;

procedure TFmAnalisis.tbLevelOwnerClick(Sender: TObject);
begin
  Analisis.TreePriceOwner.Tree.FullCollapse;
  Analisis.TreePriceOwner.Tree.Items[0].Expanded:= true;
  Analisis.TreeOwnerExpand(0);
end;

procedure TFmAnalisis.tbPriceAnalisisClick(Sender: TObject);
begin
if tbPriceAnalisis.Marked then
   begin
       tbPriceAnalisis.Marked:=false;
       tbPriceAnalisis.Down:=true;

        pPriceAnalisis.Visible:=true;
        SplitterPriceAnalisis.Visible:=true;
        //tbBtnPricePositionVersion1.Marked:=false;
        //tbBtnPricePositionVersion1.Down:=true;
     if Assigned(GridPriceList.DataSource) then
               Analisis.ChangedPriceData();
   end
     else
   begin
      pPriceAnalisis.Visible:=false;
      SplitterPriceAnalisis.Visible:=false;
      tbPriceAnalisis.Marked:=true;
      tbPriceAnalisis.Down:=false;

      //tbBtnPricePositionVersion1.Marked:=true;
      //tbBtnPricePositionVersion1.Down:=false;
   end;
end;

procedure TFmAnalisis.tbChangedPriceClick(Sender: TObject);
begin
 ChartPrice.Tag:=0;
 Analisis.ChangedPriceData();
end;

procedure TFmAnalisis.tbChangedStockClick(Sender: TObject);
begin
 ChartPrice.Tag:=1;
 Analisis.ChangedPriceData();
end;

procedure TFmAnalisis.tbPriceInfoClick(Sender: TObject);
begin

if tbPriceInfo.Marked then
   begin

     tbPriceInfo.Marked:=false;
     tbPriceInfo.Down:=true;
     SplitterInfo.Visible:= true;
     pInfo.Width:= 170;
     gbInfo.Visible:= true;

   if Assigned(Analisis.GridPrice.Grid.DataSource) then
        Analisis.ChangedPriceData();
   end
     else
   begin
      tbPriceInfo.Marked:=true;
      tbPriceInfo.Down:=false;
      gbInfo.Visible:= false;
      pInfo.Width:= 0;
      SplitterInfo.Visible:= false;
   end;
end;

procedure TFmAnalisis.tbTreeBtnExpandClick(Sender: TObject);
begin
if tbTreeBtnExpand.Marked then
   begin
     Analisis.TreePriceGroup.Expanded:=false;
     Analisis.TreePriceGroup.Tree.FullCollapse;
     Analisis.TreePriceGroup.Tree.Items[0].Expanded:= true;
     tbTreeBtnExpand.Marked:=false;
   end
     else
   begin
      Analisis.TreePriceGroup.Expanded:=true;
      Analisis.TreePriceGroup.Tree.FullExpand;
      tbTreeBtnExpand.Marked:=true;
   end;
end;

procedure TFmAnalisis.tbTreeBtnShowChildClick(Sender: TObject);
begin
  if Analisis.TreePriceGroup.Tree.Items.Count=0 then exit;

  if tbTreeBtnShowChild.Marked then
     begin
       Analisis.TreePriceGroup.ShowChildrenItems:=true;
       Analisis.GridPriceFiltered();
       tbTreeBtnShowChild.Marked:=false;
     end
       else
     begin
        Analisis.TreePriceGroup.ShowChildrenItems:=false;
        Analisis.GridPriceFiltered();
        tbTreeBtnShowChild.Marked:=true;
     end;
end;

procedure TFmAnalisis.tbTreeBtnSortClick(Sender: TObject);
begin
if tbTreeBtnSort.Marked then
   begin
     Analisis.TreePriceGroup.OrderBy:='IDPARENT, ID';
     Analisis.TreePriceGroup.Fill();
     tbTreeBtnSort.Marked:=false;
   end
     else
   begin
      Analisis.TreePriceGroup.OrderBy:='IDPARENT, NAME';
      Analisis.TreePriceGroup.Fill();
      tbTreeBtnSort.Marked:=true;
   end;
end;

procedure TFmAnalisis.ToolButton3Click(Sender: TObject);
begin
 Analisis.TreePriceOwner.Tree.FullCollapse;
 Analisis.TreePriceOwner.Tree.Items[0].Expanded:= true;
end;

procedure TFmAnalisis.TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmAnalisis.TreeGroupPriceGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmAnalisis.TreeGroupPriceGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmAnalisis.TreeViewInfoAdvancedCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; Stage: TCustomDrawStage;
  var PaintImages, DefaultDraw: Boolean);
var
  NodeRect: TRect;
begin
NodeRect := Node.DisplayRect(True);
 with (Sender as TTreeView).Canvas do
 begin
    DefaultDraw := True;

    if cdsSelected in State then // Выбранный пользователем элемент?
    begin
       Brush.Color := clHighlight;
       FillRect(NodeRect);
       Font.Color := clWhite
    end
    else // Обычный, не выбранный
    begin
       if (Node.Index = 11) or (Node.Index = 12) then
          Font.Color := clBlue
       else
          Font.Color := clBlack;
    end;
    TextOut(NodeRect.Left + 2, NodeRect.Top + 1, Node.Text);
 end;
end;

procedure TFmAnalisis.TreeViewInfoDblClick(Sender: TObject);
var
  _FormWait: TFmWait;
begin
  case TTreeView(sender).Selected.Index of
      10:
        begin
          if Length(TTreeView(sender).Selected.Text)=0 then exit;
          _FormWait:= TFmWait.Create(self);
          _FormWait.Caption:='Примечание';
          _FormWait.pbStatus.Visible:= false;
          _FormWait.mStatus.Append(TTreeView(sender).Selected.Text);

          _FormWait.BorderStyle:= bsSizeable;
          _FormWait.Height:=250;
          _FormWait.Width:=600;
          _FormWait.Memo.Alignment:= taLeftJustify;
          _FormWait.Memo.Font.Style:=_FormWait.Memo.Font.Style-[fsBold];

          try
           _FormWait.ShowModal;
          finally
            _FormWait.Free;
          end;

        end;
      11,12: if Length(TTreeView(sender).Selected.Text)>0 then OpenURL(TTreeView(sender).Selected.Text);
    end;
end;

procedure TFmAnalisis.TreeViewInfoGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmAnalisis.XMLPropStorage1SavingProperties(Sender: TObject);
begin
  DBGridClearOrderBy(GridPriceList);
  DBGridClearOrderBy(GridPriceListChangePrice);
  DBGridClearOrderBy(GridPriceListChangeStock);
  DBGridClearOrderBy(GridPriceListNew);
end;

procedure TFmAnalisis.mKursClick(Sender: TObject);
begin
 Clipboard.AsText:= TMenuItem(Sender).Name+' = '+TMenuItem(Sender).Caption;
 ShowMessage('Строка скопирована в буфер обмена.');
end;

procedure TFmAnalisis.SetStatus(_Text: string);
begin
  wStatus(fFormID,_Text,true);
end;

end.

