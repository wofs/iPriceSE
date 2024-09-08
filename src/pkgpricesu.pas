unit pkgPricesU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Spin, SysUtils, Dialogs,  Forms, Controls, ComCtrls, ExtCtrls, StdCtrls, DBGrids, Classes, Graphics, db,
  LCLIntf, Clipbrd, LazFileUtils,
  FmMatchingAddU, FmArcViewU, FmWaitU, FmTreeU,
  wZipperU,
  wLogU, wFuncU,
  wBaseU, wDBTreeU, wDBGridU, wTypesU,
  wDBImportU,
  mPriceLists,
  UtilsU

  , Grids, Buttons, ExtDlgs, DbCtrls, Menus, XMLPropStorage, DBDateTimePicker;
type

  TPriceBtmMode = (pbmAnalis, pbmInvoce);

  { TFmPrices }

  TFmPrices = class(TForm)
    btnedInvoceSearchClear: TSpeedButton;
    btnPriceEditSearchClear: TSpeedButton;
    btnPricePreventSearch: TSpeedButton;
    btnPriceSearchSplitString: TSpeedButton;
    edInvoceSearch: TComboBox;
    edPriceSearch: TComboBox;
    Markup: TFloatSpinEdit;
    gbInfo: TGroupBox;
    gbInvoce: TGroupBox;
    GridInvoce: TDBGrid;
    GridPriceList: TDBGrid;
    gbExportLog: TGroupBox;
    gbKontr: TGroupBox;
    GridPosition: TDBGrid;
    gbPricePosition: TGroupBox;
    GridImageList: TImageList;
    ImageListTreeGroup: TImageList;
    ImageListTree: TImageList;
    ImagesTreeInfo: TImageList;
    lbInvoceSearch: TLabel;
    mClipboardVNLP: TMenuItem;
    mClipboardVNL: TMenuItem;
    mClipboardAll: TMenuItem;
    mClipboardAnalVNLP: TMenuItem;
    mClipboardAnalVNL: TMenuItem;
    MenuItem10: TMenuItem;
    mAddIntoInvoce: TMenuItem;
    MenuItem11: TMenuItem;
    mAnalogs: TMenuItem;
    MenuItem12: TMenuItem;
    mClpbCopyOwnerIDs: TMenuItem;
    mDeletePosition: TMenuItem;
    mClipboardVNLPK: TMenuItem;
    mTGDeleteGroup: TMenuItem;
    mPricesDateTime: TMenuItem;
    MenuItem15: TMenuItem;
    mInvoceSelectAllClear: TMenuItem;
    mInvoceSelectAll: TMenuItem;
    MenuItem16: TMenuItem;
    mInvoceFindPanelSplit: TMenuItem;
    mInvoceFindPanel: TMenuItem;
    MenuItem4: TMenuItem;
    mPrintDateTimePrices: TMenuItem;
    mVCodeNameLabelQuantity: TMenuItem;
    MenuItem13: TMenuItem;
    mInvoceEdit: TMenuItem;
    mInvoceDel: TMenuItem;
    mVCodeNameLabelPrice: TMenuItem;
    mInvoce: TMenuItem;
    mExportInSpreadsheetPosition: TMenuItem;
    mExportInSpreadsheet: TMenuItem;
    MenuItem9: TMenuItem;
    mPriceDelete: TMenuItem;
    MenuItem8: TMenuItem;
    mSelectAllPositionInGroup: TMenuItem;
    MenuItem5: TMenuItem;
    mTGPAddToCatalog: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    mTreeInfoCopyOne: TMenuItem;
    mInfoPrice: TMenuItem;
    mKursImages: TImageList;
    Images24: TImageList;
    Images16: TImageList;
    ImageListTreeMenu: TImageList;
    lbPriceFind: TLabel;
    mAddMatching: TMenuItem;
    MenuItem1: TMenuItem;
    mSummaryPrice: TMenuItem;
    mPositionPriceArc: TMenuItem;
    mSelectAll: TMenuItem;
    mGoToGroup: TMenuItem;
    MenuItem3: TMenuItem;
    mSelectionClear: TMenuItem;
    MenuItem2: TMenuItem;
    mPricePosition: TMenuItem;
    mExportLog: TMemo;
    mGridPriceList: TPopupMenu;
    mTreeGroupOwner: TPopupMenu;
    mGridPricePositionVersion: TPopupMenu;
    mKurs: TPopupMenu;
    Panel1: TPanel;
    Panel2: TPanel;
    pPriceBtmLeft: TPanel;
    pPriceBtmRight: TPanel;
    pPriceBtm: TPanel;
    pGridPriceBtmRight: TPanel;
    pGridPriceBtmLeft: TPanel;
    pGridPriceBtm: TPanel;
    pcBottom: TPageControl;
    pGridPrice: TPanel;
    MenuTreeInfo: TPopupMenu;
    mTreeGroupPrice: TPopupMenu;
    pInvoceSearch: TPanel;
    mGridInvoces: TPopupMenu;
    pPriceDateTime_Kurs: TPanel;
    pPricePositionVersion: TPanel;
    pcPrices: TPageControl;
    pcPriceGroup: TPageControl;
    pKontrPrice: TPanel;
    pExportLog: TPanel;
    pPriceGroup: TPanel;
    pPrice: TPanel;
    pPriceFind: TPanel;
    pMainPrices: TPanel;
    sBtnDoublesView: TSpeedButton;
    sBtnKurs: TSpeedButton;
    sBtnStockOnly: TSpeedButton;
    sBtnSelected: TSpeedButton;
    sBtnWithMatching: TSpeedButton;
    sBtnNoMatching: TSpeedButton;
    splitPricePositionVersion: TSplitter;
    SplitterInfo: TSplitter;
    SpltView: TSplitter;
    SpltExport: TSplitter;
    st_PriceVersion: TStaticText;
    st_GridSelect: TStaticText;
    st_InvoceItog: TStaticText;
    TabPriceView: TTabSheet;
    TabPriceImportPage: TTabSheet;
    TabKontr: TTabSheet;
    TabGroup: TTabSheet;
    rbImportExport: TToolBar;
    tbBtnInvoce: TToolButton;
    tbInvoce: TToolBar;
    tbInvoceBtnEdit: TToolButton;
    tbInvoceBtnDel: TToolButton;
    ToolButton8: TToolButton;
    tsInvoce: TTabSheet;
    tsPricePosition: TTabSheet;
    tbBtnPricePosition: TToolButton;
    tbPriceInfo: TToolButton;
    tbImport: TToolButton;
    tbLogClear: TToolButton;
    tbPrice: TToolBar;
    tbTree: TToolBar;
    tbTree1: TToolBar;
    tbTreeBtnExpand: TToolButton;
    tbTreeOwnerBtnExpand: TToolButton;
    tbTreeBtnSort: TToolButton;
    tbTreeGroupBtnShowChild: TToolButton;
    tbTreeOwnerBtnSort: TToolButton;
    tbPosition: TToolBar;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    tbIgnoreVersion: TToolButton;
    tbTreeOwnerBtnShowChild: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    tbPositionPriceArc: TToolButton;
    tbSummaryPrice: TToolButton;
    ToolButton6: TToolButton;
    TreeGroupOwner: TTreeView;
    TreeGroupPrice: TTreeView;
    TreeGroupOwnerExport: TTreeView;
    TreeViewInfo: TTreeView;
    XMLPropStorage1: TXMLPropStorage;
    procedure btnedInvoceSearchClearClick(Sender: TObject);
    procedure btnPriceEditSearchClearClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure GridPriceListDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure mAddIntoInvoceClick(Sender: TObject);
    procedure mAddMatchingClick(Sender: TObject);
    procedure MarkupChange(Sender: TObject);
    procedure mClipboardAllClick(Sender: TObject);
    procedure mClipboardAnalVNLClick(Sender: TObject);
    procedure mClipboardAnalVNLPClick(Sender: TObject);
    procedure mClipboardVNLClick(Sender: TObject);
    procedure mClipboardVNLPClick(Sender: TObject);
    procedure mClpbCopyOwnerIDsClick(Sender: TObject);
    procedure mDeletePositionClick(Sender: TObject);
    procedure mClipboardVNLPKClick(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure mExportInSpreadsheetClick(Sender: TObject);
    procedure mExportInSpreadsheetPositionClick(Sender: TObject);
    procedure mGridInvocesPopup(Sender: TObject);
    procedure mInvoceDelClick(Sender: TObject);
    procedure mInvoceEditClick(Sender: TObject);
    procedure mInvoceFindPanelClick(Sender: TObject);
    procedure mInvoceSelectAllClearClick(Sender: TObject);
    procedure mInvoceSelectAllClick(Sender: TObject);
    procedure mPriceDeleteClick(Sender: TObject);
    procedure mPrintDateTimePricesClick(Sender: TObject);
    procedure mSelectAllPositionInGroupClick(Sender: TObject);
    procedure mTGDeleteGroupClick(Sender: TObject);
    procedure mTGPAddToCatalogClick(Sender: TObject);
    procedure mTreeInfoCopyOneClick(Sender: TObject);
    procedure mGoToGroupClick(Sender: TObject);
    procedure mGridPricePositionVersionPopup(Sender: TObject);
    procedure mPositionPriceArcClick(Sender: TObject);
    procedure mSelectAllClick(Sender: TObject);
    procedure mSelectionClearClick(Sender: TObject);
    procedure mGridPriceListPopup(Sender: TObject);
    procedure mSummaryPriceClick(Sender: TObject);
    procedure mVCodeNameLabelPriceClick(Sender: TObject);
    procedure mVCodeNameLabelQuantityClick(Sender: TObject);
    procedure pcPriceGroupChange(Sender: TObject);
    procedure pcPricesChange(Sender: TObject);
    procedure sBtnKursClick(Sender: TObject);
    procedure sBtnNoMatchingClick(Sender: TObject);
    procedure sBtnStockOnlyClick(Sender: TObject);
    procedure sBtnSelectedClick(Sender: TObject);
    procedure sBtnWithMatchingClick(Sender: TObject);
    procedure tbBtnInvoceClick(Sender: TObject);
    procedure tbBtnPricePositionClick(Sender: TObject);
    procedure tbPriceInfoClick(Sender: TObject);
    procedure tbImportClick(Sender: TObject);
    procedure tbLogClearClick(Sender: TObject);
    procedure tbTreeBtnExpandClick(Sender: TObject);
    procedure tbTreeGroupBtnShowChildClick(Sender: TObject);
    procedure tbTreeOwnerBtnShowChildClick(Sender: TObject);
    procedure tbTreeBtnSortClick(Sender: TObject);
    procedure tbTreeOwnerBtnExpandClick(Sender: TObject);
    procedure tbTreeOwnerBtnSortClick(Sender: TObject);
    procedure TreeGroupOwnerChange(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupOwnerExportGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupOwnerExportGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupPriceGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupPriceGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeViewInfoAdvancedCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; Stage: TCustomDrawStage; var PaintImages,
      DefaultDraw: Boolean);
    procedure TreeViewInfoDblClick(Sender: TObject);
    procedure TreeViewInfoGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure XMLPropStorage1SavingProperties(Sender: TObject);
  private
    { private declarations }
    FormIDent: string;
    OwnerID: integer;     // ID выбранного контрагента
    IdMainOwner: string; // ID основного контрагента (к которому привязан каталог)
   _TimeStampMaxArr: ArrayOfDateTime; // массив максимальных значений таймштамп из таблицы прайс-листы

    fBase: TwBase;
    fDBImport: TwDBImport;
    Prices: TPrices;



    _Form: TFmMatchingAdd;

    procedure mKursClick(Sender: TObject);
    procedure pcBottomVisible(aVisible: boolean; PriceBtmMode: TPriceBtmMode);
    property wFormID: string read FormIDent write FormIDent;

  protected
    procedure OpenPrice;
  public
    { public declarations }
    procedure SetStatus(_Text:string);
  end;

var
  FmPrices: TFmPrices;

implementation

{$R *.lfm}

{ TFmPrices }

procedure TFmPrices.SetStatus(_Text: string);
begin
     wStatus(wFormID,_Text,true);
end;

procedure TFmPrices.FormCreate(Sender: TObject);
begin
  wFormID:=Self.Name;
  screen.Cursor:= crSQLWait;

  OwnerID:=0; // инициализируем.
  wLog('Prices','Инициализация плагина... ['+wFormID+']');

  fBase:= TwBase.Create(Sender);
  IdMainOwner:= fBase.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист

  try
  Prices:= TPrices.Create(Sender,fBase,GridPriceList,GridPosition,GridInvoce,TreeGroupOwner,TreeGroupPrice,TreeGroupOwnerExport);
  Prices.TreeInfo:= TreeViewInfo;

  Prices.GridPriceFill(nil);
  Prices.TreePriceOwnerFill();
  Prices.GridPosition.Grid.Tag:=1; // Указываем, что при открытии аналитики показывать сводный прайс-лист

  Prices.GridInvocesFill(Prices.TreePriceOwner.SelectedItems);

  fDBImport:= TwDBImport.Create(Self,mExportLog);
    wLog('Prices','Инициализация плагина успешно завершена.');

  if TreeGroupOwner.Selected.Text = '' then
    begin
      ShowMessage('Перед работой с прайс-листами добавьте хотя бы одного контрагента в модуле "Форматы".');
      SetStatus('Перед работой с прайс-листами добавьте хотя бы одного контрагента в модуле "Форматы".');
    end;
   screen.Cursor:= crDefault;
  except
    on E: Exception do
    begin
        screen.Cursor:= crDefault;
        SetStatus('Сбой инициализации плагина.');
        wLog('Prices','Ошибка [FmCreate]: "' + E.Message + '"');
        wLog('Prices','Сбой инициализации плагина.');
        ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

     end;
  end;
//
  //wLicenseOK(Licensed); // [wLicense]

end;

procedure TFmPrices.btnPriceEditSearchClearClick(Sender: TObject);
begin
  edPriceSearch.Text:='';
  edPriceSearch.OnChange(edPriceSearch);
end;

procedure TFmPrices.btnedInvoceSearchClearClick(Sender: TObject);
begin
  edInvoceSearch.Clear;
  edInvoceSearch.OnChange(edInvoceSearch);
end;

procedure TFmPrices.FormDestroy(Sender: TObject);
var
  i: integer;
begin

   try
   wLog('Prices','Выгрузка плагина...');

      try
      // выгружаем форму добавлени соответствий
        if Assigned(_Form) then
         _Form.GridDataSet:= nil;

      finally
        Prices.Destroy();
        fBase.Destroy();
        fDBImport.Destroy();
        //wLicense_ReadKey(_DBase); // считываем ключ из БД и проверяем [wLicense]
      end;


      wLog('Prices','Выгрузка плагина успешно завершена.');

      except
        on E: Exception do
        begin
            SetStatus('Сбой выгрузки плагина: Каталог.');
            wLog('Prices','Ошибка [FmDestroy]: "' + E.Message + '"');
            wLog('Prices','Сбой выгрузки плагина.');
            ShowMessage('Ошибка [FmDestroy]: "' + E.Message + '"');
         end;
      end;
end;

procedure TFmPrices.GridPriceListDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  _ColumnText: string;
  _ImageIndex: integer;
  _FieldValue: Double;
begin
  if Prices.GridPrice.Grid.DataSource.DataSet.RecordCount = 0 then exit;

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

        if Prices.GridPrice.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         begin
                  _ImageIndex:= 0;
                  // А теперь пусть ImageList нарисует ее на канве DBGrid'а
                  TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
         end;
     end else
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);

        if Prices.GridPrice.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         _ColumnText:='+'
            else
         _ColumnText:=' ';

         TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
     end;
  end;

 if Column.FieldName = 'QUANTITYINPACKING' then
  begin
    with TDBGrid(Sender).Canvas do
    begin
      _FieldValue:= TDBGrid(Sender).DataSource.DataSet.FieldByName('QUANTITYINPACKING').AsFloat;
      if _FieldValue <> 0 then
        begin
          FillRect(Rect);
          if _FieldValue<1 then
            _ColumnText:= '1 к '+FloatToStr(_wRND(1/_FieldValue))
          else
            _ColumnText:= FloatToStr(1*_FieldValue)+' к 1';

          TextOut(Rect.Right - 2 - TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
        end else
        begin
           FillRect(Rect);
           _ColumnText:='';
           TextOut(Rect.Right - 2 - TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
        end;
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
      begin
      FillRect(Rect);

         if TDBGrid(Sender).DataSource.DataSet.FieldByName('STOCKONLYINFO').AsInteger>0 then
            begin
              if gdSelected in State
                 then font.Color:=clRed
                 else font.Color:=clBlue;


            end;

         if (TDBGrid(Sender).Name = GridPosition.Name) and (TDBGrid(Sender).Tag = 1) then
           if TDBGrid(Sender).DataSource.DataSet.FieldByName('QUANTITYINPACKING').AsFloat<>1 then
                  Font.Style := [fsBold];

      TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
      end;
    end;

   Prices.GridPrice.HighLightText(Sender,'PLNAME', Rect,DataCol,Column,State); // подсветить часть текста;
   Prices.GridPrice.HighLightText(Sender,'LABEL', Rect,DataCol,Column,State); // подсветить часть текста
end;

procedure TFmPrices.mAddIntoInvoceClick(Sender: TObject);
begin
  Prices.InvoceAdd();
end;

procedure TFmPrices.mAddMatchingClick(Sender: TObject);
begin

 _Form:= TFmMatchingAdd.Create(Application);
 _Form.SelectedRows:= Prices.GridPrice.SelectedRows;
 _Form.GridDataSet:=GridPriceList.DataSource.DataSet;
 _Form.Show;
end;

procedure TFmPrices.MarkupChange(Sender: TObject);
begin
   Prices.GridPriceFill(Prices.GridPrice.GroupArray);
   if TFloatSpinEdit(Sender).Value>0 then
      TFloatSpinEdit(Sender).Color:=clSkyBlue
   else
      TFloatSpinEdit(Sender).Color:=clDefault;
end;

procedure TFmPrices.mClipboardAllClick(Sender: TObject);
var
  _DS: TDataSet;
  _SQLText: String;
begin

if not Assigned(Prices.GridPrice.Grid.DataSource) then exit;
_SQLText:='SELECT '
  +' PL.VENDORCODE, '
  +' PL.NAME AS PLNAME, '
  +' PL.UNIT , '
  +' PL.PRICECALC AS PRICE, '
  +' (PL.STOCK+PL.STOCK2+PL.STOCK3+PL.STOCK4+PL.STOCK5) as STOCK, '
  +' (select VSCOD from PL_GET_SCOD(PL.ID,true)) SCOD, '
  +' PL.LABEL, '
  +' PL.TRANSIT '
  +'from PL_ITEMS PL where ID='+Prices.GridPrice.Grid.DataSource.DataSet.FieldByName('ID').AsString;

_DS:= fBase.SQLReadDS(_SQLText).DataSet;
 Prices.GridPrice.CopyToClipboard(_DS, nil, ['PRICE']);
end;

procedure TFmPrices.mClipboardAnalVNLClick(Sender: TObject);
begin
  Prices.GridPosition.CopyToClipboard(['VENDORCODE', 'PLNAME', 'UNIT', 'LABEL'], [''],'PL.ID');
end;

procedure TFmPrices.mClipboardAnalVNLPClick(Sender: TObject);
begin
  Prices.GridPosition.CopyToClipboard(['VENDORCODE', 'PLNAME', 'UNIT', 'LABEL', 'PRICE'], ['PRICE'],'PL.ID');
end;

procedure TFmPrices.mClipboardVNLClick(Sender: TObject);
begin
  Prices.GridPrice.CopyToClipboard(['VENDORCODE', 'PLNAME', 'UNIT', 'LABEL'], [''],'PL.ID');
end;

procedure TFmPrices.mClipboardVNLPClick(Sender: TObject);
begin
  Prices.GridPrice.CopyToClipboard(['VENDORCODE', 'PLNAME', 'UNIT', 'LABEL', 'PRICE'], ['PRICE'],'PL.ID');
end;

procedure TFmPrices.mClpbCopyOwnerIDsClick(Sender: TObject);
begin
  Clipboard.AsText:= fBase.MakeStringFromArray(Prices.TreePriceOwner.SelectedItems);
end;

procedure TFmPrices.mDeletePositionClick(Sender: TObject);
begin
  Prices.DeletePriceItem;
end;

procedure TFmPrices.mClipboardVNLPKClick(Sender: TObject);
begin
  Prices.GridPrice.CopyToClipboard(['VENDORCODE', 'PLNAME', 'UNIT', 'LABEL', 'PRICE', 'OWNERNAME'], ['PRICE'],'PL.ID');
end;

procedure TFmPrices.MenuItem1Click(Sender: TObject);
begin
OpenPrice;

end;

procedure TFmPrices.MenuItem4Click(Sender: TObject);
begin
  Prices.GridInvoces.ExportData();
end;

procedure TFmPrices.mExportInSpreadsheetClick(Sender: TObject);
begin
  Prices.GridPrice.ExportData();
end;

procedure TFmPrices.mExportInSpreadsheetPositionClick(Sender: TObject);
begin
  Prices.GridPosition.ExportData();
end;

procedure TFmPrices.mGridInvocesPopup(Sender: TObject);
begin
  mInvoceFindPanel.Checked:= pInvoceSearch.Visible;
  Prices.mAnalogsFill(Sender, mAnalogs);
end;

procedure TFmPrices.mInvoceDelClick(Sender: TObject);
begin
  Prices.InvoceDel(Prices.GridInvoces.SelectedRows);
end;

procedure TFmPrices.mInvoceEditClick(Sender: TObject);
begin
if Prices.GridInvoces.Grid.DataSource.DataSet.RecordCount = 0 then exit;

  Prices.InvoceEdit(Prices.GridInvoces.SelectedRows[0]);
end;

procedure TFmPrices.mInvoceFindPanelClick(Sender: TObject);
begin
  pInvoceSearch.Visible:= not pInvoceSearch.Visible;
  if not pInvoceSearch.Visible then
     btnedInvoceSearchClearClick(self)
  else
  edInvoceSearch.SetFocus;
end;

procedure TFmPrices.mInvoceSelectAllClearClick(Sender: TObject);
begin
  Prices.GridInvoces.SelectAll:= false;
end;

procedure TFmPrices.mInvoceSelectAllClick(Sender: TObject);
begin
  Prices.GridInvoces.SelectAll:= true;
end;

procedure TFmPrices.mPriceDeleteClick(Sender: TObject);
var
  _Owner: integer;
  _OwnerName: String;
begin
  _Owner:= Prices.TreePriceOwner.SelectedItems[0];
  _OwnerName:= Prices.TreePriceOwner.Tree.Selected.Text;

  if MessageDlg('Удалить прайс-лист '+_OwnerName+'? Будут удалены: Группы, прайс-лист и архив.',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;
  if MessageDlg('Внимание! Удаление прайс-листа приведет к удалению СООТВЕТСТВИЙ прайс-листа! Продолжить?',mtWarning, mbOKCancel, 0) = mrCancel then exit;

  try
      SetStatus('Удаление прайс-листа...');
      Prices.DeletePriceList(_Owner);
  except
    on E: Exception do
    begin
      __Log.SaveLogError(E);
      ShowMessage(E.Message);
      SetStatus(E.Message);
    end;
  end;
end;

procedure TFmPrices.mPrintDateTimePricesClick(Sender: TObject);
begin
   Prices.PrintDateTimePrices;
end;

procedure TFmPrices.mSelectAllPositionInGroupClick(Sender: TObject);
begin
  GridPriceList.Cursor:= crSQLWait;
  Application.ProcessMessages;
  Prices.GridPrice.SelectAll:= true;
  GridPriceList.Cursor:= crDefault;
end;

procedure TFmPrices.mTGDeleteGroupClick(Sender: TObject);
begin
  if MessageDlg('Удалить выбранные группы? Внимание! Буду удалены так же подгруппы и товары, содержащиеся в группах и их соответствия!',mtWarning, mbOKCancel, 0) = mrCancel then exit;

  try
    Prices.DeletePriceGroup;

    ShowMessage('Удаление успешно завершено');
  except
    ShowMessage('Удаление завершено с ОШИБКОЙ!');
    raise;
  end;
end;

procedure TFmPrices.mTGPAddToCatalogClick(Sender: TObject);
begin
  if MessageDlg('Добавить выделенные категории с подкатегориями и товарами в каталог?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

  PricesLoadNomenclatureEditMassForm(Sender,fBase,Prices.TreePriceGroup);

end;

procedure TFmPrices.mTreeInfoCopyOneClick(Sender: TObject);
begin
  Clipboard.AsText:= TreeViewInfo.Selected.Text;
end;

procedure TFmPrices.mGoToGroupClick(Sender: TObject);
var
  _GridDataset: TDataSet;
  _ParentID: integer;
  _ID: integer;
begin


  if TreeGroupOwner.Selected.Text = '' then exit;

  _GridDataset:= GridPriceList.DataSource.DataSet;
  _ID:= _GridDataset.FieldByName('ID').AsInteger;
  try
    case pcPriceGroup.ActivePageIndex of
        0:
          begin
             _ParentID:= _GridDataset.FieldByName('IDOWNER').AsInteger ;
             Prices.TreePriceOwner.FindNodeWithDataInt(_ParentID);
          end;
        1:
          begin
             _ParentID:= _GridDataset.FieldByName('IDPL_GROUP').AsInteger ;
             Prices.TreePriceGroup.FindNodeWithDataInt(_ParentID);
          end;
    end;
  finally

  if _GridDataset.RecordCount>0 then
    _GridDataset.Locate('ID',_ID,[]);
  end;
//
end;

procedure TFmPrices.mGridPricePositionVersionPopup(Sender: TObject);
begin
  mPositionPriceArc.Checked:= tbPositionPriceArc.Down;
  mSummaryPrice.Checked := tbSummaryPrice.Down;
end;

procedure TFmPrices.mPositionPriceArcClick(Sender: TObject);
begin
  GridPosition.Tag:= 0;
  Prices.GridPositionModeChange();
end;

procedure TFmPrices.mSelectAllClick(Sender: TObject);
begin
     GridPriceList.Cursor:= crSQLWait;
     Application.ProcessMessages;
     Prices.GridPrice.SelectAll:= true;
     GridPriceList.Cursor:= crDefault;
end;

procedure TFmPrices.mSelectionClearClick(Sender: TObject);
var
  _SelCount: Integer;
begin

   _SelCount:=  Prices.GridPrice.SelectedRowsCount;

   if _SelCount>0 then
     begin
       if MessageDlg('Сбросить выделение ('+IntToStr(_SelCount)+') позиций?',mtConfirmation, mbOKCancel, 0) = mrOK then
        begin
           Prices.GridPrice.SelectAll:= false;
        end;
     end;
end;

procedure TFmPrices.mGridPriceListPopup(Sender: TObject);
begin
  mPricePosition.Checked:= tbBtnPricePosition.Down;
  mInvoce.Checked:= tbBtnInvoce.Down;

  mInfoPrice.Checked:= gbInfo.Visible;
end;

procedure TFmPrices.mSummaryPriceClick(Sender: TObject);
begin
GridPosition.Tag:= 1;
Prices.GridPositionModeChange();
end;

procedure TFmPrices.mVCodeNameLabelPriceClick(Sender: TObject);
begin
  Prices.GridInvoces.CopyToClipboard(['VENDORCODE', 'NAME', 'UNIT', 'LABEL', 'PRICEPL'], ['PRICEPL'],'INV.ID');
end;

procedure TFmPrices.mVCodeNameLabelQuantityClick(Sender: TObject);
begin
  Prices.GridInvoces.CopyToClipboard(['VENDORCODE', 'NAME', 'UNIT', 'LABEL', 'QUANTITY'], [''],'INV.ID');
end;

procedure TFmPrices.pcPriceGroupChange(Sender: TObject);
var
  _Tree:TTreeview;
  _PageControl: TPageControl;
  _arr: ArrayOfInteger;
begin
  _Tree:= TreeGroupOwner;
  _PageControl:= (Sender as TPageControl);
  fBase.SQLTransactionEnd();
  _arr:= Prices.TreePriceOwner.SelectedItems;
case _PageControl.ActivePageIndex of
    0:begin
       Prices.GridPrice.GroupField:= 'PL.IDOWNER';
       Prices.GridPrice.GroupArray:= nil;
       Prices.GridPrice.Where:='';
    end;
    1:begin
        if (_Tree.SelectionCount >1) or (_Tree.Selected.Count>0) or not Assigned(_arr) then
           begin
             _PageControl.ActivePageIndex:=0;
             ShowMessage('Для доступа к группам номенклатуры выберите одного из контрагентов.');
           end else
           begin
              Prices.GridPrice.GroupField:= 'PL.IDPL_GROUP';
              Prices.GridPrice.GroupArray:= nil;
              Prices.GridPrice.Where:='PL.IDOWNER='+IntToStr(_arr[0]);
              Prices.TreePriceGroup.SetOwner:= OwnerID;
              Prices.TreePriceGroupFill();
           end;
    end;

  end;
end;

procedure TFmPrices.pcPricesChange(Sender: TObject);
var
  _PageControl: TPageControl;
begin

try
_PageControl:= (Sender as TPageControl);


case _PageControl.ActivePageIndex of
     0:
       begin
          pcPriceGroup.ActivePageIndex:=0;
          _TimeStampMaxArr:=nil;
          _TimeStampMaxArr:= GetMaxFTimeStampPricesArr(fBase);
          Prices.PriceMaxFTImeStampArr:= _TimeStampMaxArr;
          Prices.TreePriceOwnerFill();
          btnPriceEditSearchClear.OnClick(btnPriceEditSearchClear);
          //Prices.GridPriceFill(nil);
       end;
     1:
       begin
          Prices.TreeImportOwnerFill();
          Prices.TreeImportOwner.FindNodeWithDataInt(OwnerID);
       end;
end;

  except
        on E: Exception do
    begin
       wLog('Prices','Ошибка [pcPricesChange]: "' + E.Message + '"');
       ShowMessage('Ошибка [pcPricesChange]: "' + E.Message + '"');
    end;
  end;
end;

procedure TFmPrices.mKursClick(Sender: TObject);
begin
    Clipboard.AsText:= TMenuItem(Sender).Name+' = '+TMenuItem(Sender).Caption;
    ShowMessage('Строка скопирована в буфер обмена.');
end;

procedure TFmPrices.sBtnKursClick(Sender: TObject);
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

procedure TFmPrices.sBtnNoMatchingClick(Sender: TObject);
begin
  Prices.GridPriceFiltered();
end;

procedure TFmPrices.sBtnStockOnlyClick(Sender: TObject);
begin
   Prices.GridPriceFiltered();
end;

procedure TFmPrices.sBtnSelectedClick(Sender: TObject);
begin
   Prices.GridPriceFiltered();
end;

procedure TFmPrices.sBtnWithMatchingClick(Sender: TObject);
begin
   Prices.GridPriceFiltered();
end;

procedure TFmPrices.tbBtnInvoceClick(Sender: TObject);
begin
if tbBtnInvoce.Marked then
   begin
     if GridPriceList.DataSource.DataSet.RecordCount<>0 then
        begin
          try
             tbBtnInvoce.Marked:=false;
             tbBtnInvoce.Down:=true;
          finally
            pcBottomVisible(true,pbmInvoce);
            Prices.GridInvocesFill(Prices.TreePriceOwner.SelectedItems);
          end;
        end else
        begin
           ShowMessage('Прайс-лист пуст!');
        end;

   end
     else
   begin
      pcBottomVisible(false,pbmInvoce);
      tbBtnInvoce.Marked:=true;
      tbBtnInvoce.Down:=false;
   end;
end;

procedure TFmPrices.pcBottomVisible(aVisible: boolean; PriceBtmMode: TPriceBtmMode);
begin
case PriceBtmMode of
  pbmAnalis:
      begin
        tbBtnPricePosition.Marked:=not aVisible;
        tbBtnPricePosition.Down:=aVisible;

        if aVisible then
          begin
            tbBtnInvoce.Marked:= aVisible;
            tbBtnInvoce.Down:= not aVisible;
          end;

        pcBottom.ActivePage:= tsPricePosition;
      end;
  pbmInvoce:
      begin
        tbBtnInvoce.Marked:=not aVisible;
        tbBtnInvoce.Down:=aVisible;

        if aVisible then
          begin
            tbBtnPricePosition.Marked:= aVisible;
            tbBtnPricePosition.Down:= not aVisible;
          end;

        pcBottom.ActivePage:= tsInvoce;
      end;
end;

  pPricePositionVersion.Visible:=aVisible;
  splitPricePositionVersion.Visible:=aVisible;

end;

procedure TFmPrices.OpenPrice;
var
  i: integer;
  _FormArc: TFmArcView;
  _FilesArr: ArrayOfArrayVariant;
  _arr: ArrayOfString;
  _UnPackPath: string;
  _FileName: string;
  _Zipper: TwZipper;
begin
  _FilesArr:= fBase.SQLReadArr('FORMATS', ['FILE'], 'IDOWNER='+IntToStr(OwnerID)+' AND IDFMTS_CATEGORY=1', '');

  if not Assigned(_FilesArr) then exit;

  if High(_FilesArr)>0 then
   begin
     _FormArc:= TFmArcView.Create(self);
     _FormArc.Caption:='Список прайс-листов | Выберите прайс-лист';

     for i:=0 to High(_FilesArr) do
         _FormArc.ListFiles.AddItem(_FilesArr[i, 0], nil);

     try
       _FormArc.ShowModal;
     finally
       _FileName:= _FormArc.SelectedFileName;
       _FormArc.Free;
     end;

   end else
       _FileName:= _FilesArr[0, 0];

  if Length(_FileName)=0 then Exit;

   _Zipper:= TwZipper.Create();

   try
     _arr:= _Zipper.ParseComboFileName(_FileName);
     if Assigned(_arr) then
       begin
         if Length(_arr[1])>0 then
           begin
             _UnPackPath:= includeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
             _UnPackPath:= _UnPackPath+'tmp';

             if not DirectoryExistsUTF8(_UnPackPath) then ForceDirectoriesUTF8(_UnPackPath);
             _UnPackPath:=_UnPackPath+DirectorySeparator+IntTOStr(OwnerID);
             if not DirectoryExistsUTF8(_UnPackPath) then ForceDirectoriesUTF8(_UnPackPath);

             //sgFormat.Cells[1,12]:= aFileName+'|'+_FileExtract;
             _Zipper.ExtractOneFile(_arr[0], _arr[1], _UnPackPath);
             _FileName:=includeTrailingPathDelimiter(_UnPackPath)+_arr[1];
           end else
           _FileName:= _arr[0];
         //Length(
       end;
       if FileExists(_FileName) then
            OpenDocument(_FileName) else
            ShowMessage('Сохраненный локально файл не найден!');
   finally
      _arr:=nil;
      _Zipper.Destroy();
   end;
end;

procedure TFmPrices.tbBtnPricePositionClick(Sender: TObject);
begin
if tbBtnPricePosition.Marked then
   begin
     if GridPriceList.DataSource.DataSet.RecordCount<>0 then
        begin
           pcBottomVisible(true,pbmAnalis);
           Prices.GridPositionFill([GridPriceList.DataSource.DataSet.FieldByName('ID').AsInteger]);
        end else
        begin
           ShowMessage('Прайс-лист пуст!');
        end;

   end
     else
   begin
      pcBottomVisible(false,pbmAnalis);
      //tbBtnPricePosition.Marked:=true;
      //tbBtnPricePosition.Down:=false;
   end;
end;

procedure TFmPrices.tbPriceInfoClick(Sender: TObject);
begin
  if tbPriceInfo.Marked then
     begin
       if GridPriceList.DataSource.DataSet.RecordCount<>0 then
          begin
             tbPriceInfo.Marked:=false;
             tbPriceInfo.Down:=true;
             gbInfo.Visible:= true;
             SplitterInfo.Visible:= true;
             Prices.TreeInfoFill(Prices.GridPrice);
          end else
          begin
             ShowMessage('Прайс-лист пуст!');
          end;

     end
       else
     begin
        tbPriceInfo.Marked:=true;
        tbPriceInfo.Down:=false;
        gbInfo.Visible:= false;
        SplitterInfo.Visible:= false;
     end;
end;

procedure TFmPrices.tbImportClick(Sender: TObject);
var
  i: integer;
  _DataSet: TDataSet;
  _FieldsArray: ArrayOfString;
  _Where: string;
begin
   _FieldsArray:= nil;
  // _arrSettings:= nil;
   //if tbImport.ImageIndex = 3 then
   //  begin
   //    ShowMessage('Импорт уже запущен! Дождитесь окончания операции!');
   //  end;

   if TreeGroupOwnerExport.Selected.Level = 0 then
      begin
        mExportLog.Lines.Add('Выберите контрагента или группу контрагентов для импорта...');

        exit;
      end;

   //tbImport.ImageIndex:=3;

   _FieldsArray:=fBase.MakeArrayFromString(FormatImportFields);

   _Where:= fBase.PrepareWhereString('IDOWNER',Prices.TreeImportOwner.SelectedItems);

   _Where:= '('+_Where+') AND IDFMTS_CATEGORY=1 AND FCLOSE=0';// только прайс-листы

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


 fDBImport.IgnoreVersion:=tbIgnoreVersion.Down;

  try
    try
    fDBImport.Import();
    finally
      _DataSet:= nil;
      _FieldsArray:= nil;
      //SetStatus('Импорт завершен.');
      //tbImport.ImageIndex:=0;
    end;

  except
    on E: Exception do
      begin
         tbImport.ImageIndex:= 0;
         wLog('Prices','Ошибка [Import] "' + E.Message + '"');
         SetStatus('Ошибка [Import]: "' + E.Message + '"');
      end;
  end;


end;

procedure TFmPrices.tbLogClearClick(Sender: TObject);
begin
  mExportLog.Clear;
end;

procedure TFmPrices.tbTreeBtnExpandClick(Sender: TObject);
begin
  if tbTreeBtnExpand.Marked then
     begin
       Prices.TreePriceGroup.Expanded:=false;
       Prices.TreePriceGroup.Tree.FullCollapse;
       Prices.TreePriceGroup.Tree.Items[0].Expanded:= true;
       tbTreeBtnExpand.Marked:=false;
     end
       else
     begin
        Prices.TreePriceGroup.Expanded:=true;
        Prices.TreePriceGroup.Tree.FullExpand;
        tbTreeBtnExpand.Marked:=true;
     end;
end;

procedure TFmPrices.tbTreeGroupBtnShowChildClick(Sender: TObject);
begin
  if Prices.TreePriceGroup.Tree.Items.Count=0 then exit;

  if tbTreeGroupBtnShowChild.Marked then
     begin
       Prices.TreePriceGroup.ShowChildrenItems:=true;
       Prices.GridPriceFill(Prices.TreePriceGroup.SelectedItems);
       tbTreeGroupBtnShowChild.Marked:=false;
     end
       else
     begin
        Prices.TreePriceGroup.ShowChildrenItems:=false;
        Prices.GridPriceFill(Prices.TreePriceGroup.SelectedItems);
        tbTreeGroupBtnShowChild.Marked:=true;
     end;
end;

procedure TFmPrices.tbTreeOwnerBtnShowChildClick(Sender: TObject);
begin
  if Prices.TreePriceOwner.Tree.Items.Count=0 then exit;

  if tbTreeOwnerBtnShowChild.Marked then
     begin
       Prices.TreePriceOwner.ShowChildrenItems:=true;
       Prices.GridPriceFill(Prices.TreePriceOwner.SelectedItems);
       tbTreeOwnerBtnShowChild.Marked:=false;
     end
       else
     begin
        Prices.TreePriceOwner.ShowChildrenItems:=false;
        Prices.GridPriceFill(Prices.TreePriceOwner.SelectedItems);
        tbTreeOwnerBtnShowChild.Marked:=true;
     end;
end;

procedure TFmPrices.tbTreeBtnSortClick(Sender: TObject);
begin
  if tbTreeBtnSort.Marked then
     begin
       Prices.TreePriceGroup.OrderBy:='IDPARENT, ID';
       Prices.TreePriceGroup.Fill();
       tbTreeBtnSort.Marked:=false;
     end
       else
     begin
        Prices.TreePriceGroup.OrderBy:='IDPARENT, NAME';
        Prices.TreePriceGroup.Fill();
        tbTreeBtnSort.Marked:=true;
     end;
end;

procedure TFmPrices.tbTreeOwnerBtnExpandClick(Sender: TObject);
begin
  if tbTreeOwnerBtnExpand.Marked then
     begin
       Prices.TreePriceOwner.Expanded:=false;
       Prices.TreePriceOwner.Tree.FullCollapse;
       Prices.TreePriceOwner.Tree.Items[0].Expanded:= true;
       tbTreeOwnerBtnExpand.Marked:=false;
     end
       else
     begin
        Prices.TreePriceOwner.Expanded:=true;
        Prices.TreePriceOwner.Tree.FullExpand;
        tbTreeOwnerBtnExpand.Marked:=true;
     end;
end;

procedure TFmPrices.tbTreeOwnerBtnSortClick(Sender: TObject);
begin
  if tbTreeOwnerBtnSort.Marked then
     begin
       Prices.TreePriceOwner.OrderBy:='IDPARENT, ID';
       Prices.TreePriceOwner.Fill();
       tbTreeOwnerBtnSort.Marked:=false;
     end
       else
     begin
        Prices.TreePriceOwner.OrderBy:='IDPARENT, NAME';
        Prices.TreePriceOwner.Fill();
        tbTreeOwnerBtnSort.Marked:=true;
     end;
end;

procedure TFmPrices.TreeGroupOwnerChange(Sender: TObject; Node: TTreeNode);
var
  _Tree: TTreeView;
begin
     _Tree:= (Sender as TTreeVIew);
     OwnerID:= TTreeData(_Tree.Selected.Data).Value;
end;

procedure TFmPrices.TreeGroupOwnerExportGetImageIndex(Sender: TObject; Node: TTreeNode);
var
  _idOwner: integer;
begin
  TryStrToInt(IdMainOwner,_idOwner);
  if TTreeData(Node.Data).Value = _idOwner then
  begin
    Node.ImageIndex:=2;
    exit;
  end;

  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmPrices.TreeGroupOwnerExportGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmPrices.TreeGroupOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
var
  _idOwner: integer;
begin
  TryStrToInt(IdMainOwner,_idOwner);
  if TTreeData(Node.Data).Value = _idOwner then
  begin
    Node.ImageIndex:=2;
    exit;
  end;

  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmPrices.TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmPrices.TreeGroupPriceGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmPrices.TreeGroupPriceGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmPrices.TreeViewInfoAdvancedCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; Stage: TCustomDrawStage;
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

procedure TFmPrices.TreeViewInfoDblClick(Sender: TObject);
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

procedure TFmPrices.TreeViewInfoGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmPrices.XMLPropStorage1SavingProperties(Sender: TObject);
begin
  DBGridClearOrderBy(GridPriceList);
  DBGridClearOrderBy(GridPosition);
end;

end.

