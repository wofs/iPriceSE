unit pkgCatalogU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  SysUtils, Dialogs, Forms, Controls, ComCtrls, ExtCtrls, StdCtrls, DBGrids,
  Classes, Graphics, Grids, Menus, Buttons, wDBGridU, XMLPropStorage, IDEWindowIntf,
  FmNomenclatureEditU, FmNomenclatureEditMassU, FmMatchingEditU, db, Clipbrd,
  LCLIntf,
  wLogU, wDBTreeU, wFuncU,  wFormulaU,
  wBaseU, mUtilsU,
  FmListSelectU, FmQuantityInPackingU, FmWaitU,
  mCatalogU, wTypesU,
  UtilsU, mInvoceU
  ;

type

  { TFmCatalog }

  TFmCatalog = class(TForm)
    btnMatchPreventSearch: TSpeedButton;
    btnPriceEditSearchClear: TSpeedButton;
    btnedMatchSearchClear: TSpeedButton;
    btnPriceSearchSplitString: TSpeedButton;
    cbPriceField: TComboBox;
    edMatchSearch: TComboBox;
    DBGrid1: TDBGrid;
    edPriceSearch: TComboBox;
    gbInfo: TGroupBox;
    GridImageList: TImageList;
    GridCatalogPrice: TDBGrid;
    gbGroupPrice: TGroupBox;
    gbGroupOwner: TGroupBox;
    GridCatalogPositionMatching: TDBGrid;
    GridMatching: TDBGrid;
    ILtabs: TImageList;
    ImageList16: TImageList;
    ImageListTree: TImageList;
    ImageListTreeGroup: TImageList;
    ImagesTreeInfo: TImageList;
    lbPriceSearch: TLabel;
    lbMatchSearch: TLabel;
    mCatalogMatchingPositionSelectClear: TMenuItem;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    mCatalogMatchingPositionSelectAll: TMenuItem;
    mClipboardVNL: TMenuItem;
    mClipboardAll: TMenuItem;
    mClipboardVNLP: TMenuItem;
    mClipboardPosVNL: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    mClipboardPosVNLP: TMenuItem;
    mCatalogInfo: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    mClipboardPosVNLPK: TMenuItem;
    mAddInvoice: TMenuItem;
    mExportInSpreadsheetMPG: TMenuItem;
    mExportInSpreadsheetMG: TMenuItem;
    mExportInSpreadsheet: TMenuItem;
    mTmExport: TMenuItem;
    MenuTreeInfo: TPopupMenu;
    mMatchingPositionAdd: TMenuItem;
    mMatchingPositionDelete: TMenuItem;
    mMatchingPositionEdit: TMenuItem;
    mMatchingPositionEditQuantInPack: TMenuItem;
    MenuItem1: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    mSelectAll: TMenuItem;
    mNomClearSelect: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    mNomCopy: TMenuItem;
    mNomDelete1: TMenuItem;
    mNomEdit1: TMenuItem;
    mNomGoToGroup1: TMenuItem;
    mNomGoToGroup: TMenuItem;
    mNomEdit: TMenuItem;
    mNomDelete: TMenuItem;
    mNomSootv: TMenuItem;
    mNomAdd: TMenuItem;
    mMatchingGrid: TPopupMenu;
    mTreeInfoCopyOne: TMenuItem;
    m_MatchingPosition: TPopupMenu;
    pGridPrice: TPanel;
    pMain: TPanel;
    pcCatalog: TPageControl;
    mPriceGrid: TPopupMenu;
    pCatalogPositionMatching: TPanel;
    mTreeMatching: TPopupMenu;
    pPriceDateTime_Kurs: TPanel;
    pPriceDateTime_Kurs1: TPanel;
    pPriceDateTime_Kurs2: TPanel;
    pPriceSearch: TPanel;
    pMatchGroup: TPanel;
    pMatchList: TPanel;
    pMarchSearch: TPanel;
    pPriceGroup: TPanel;
    pPrice: TPanel;
    btnPricePreventSearch: TSpeedButton;
    SaveDialog1: TSaveDialog;
    sBtnSelected: TSpeedButton;
    sBtnMatchingSelected: TSpeedButton;
    sBtnStockOnly: TSpeedButton;
    sBtnDoublesView: TSpeedButton;
    sBtnWithMatching: TSpeedButton;
    sBtnNoMatching: TSpeedButton;
    Separator1: TMenuItem;
    splitCatalogPositionMatching: TSplitter;
    SplitterInfo: TSplitter;
    SpltPrice: TSplitter;
    SpltMatch: TSplitter;
    st_GridCatalogMarchingSelect: TStaticText;
    st_GridCatalogSelect: TStaticText;
    st_GridCatalogMarchingPositionSelect: TStaticText;
    TabNomenclature: TTabSheet;
    TabMatching: TTabSheet;
    tbCatalogBtnAdd: TToolButton;
    tbCatalogPosirionMatchingAdd: TToolButton;
    tbCatalogBtnCopy: TToolButton;
    tbCatalogBtnDelete: TToolButton;
    tbCatalogPosirionMatchingDelete: TToolButton;
    tbCatalogBtnEdit: TToolButton;
    tbCatalogPosirionMatchingEdit: TToolButton;
    tbCatalogPosirionMatchingQuantity: TToolButton;
    tbCatalogBtnGoToGroup: TToolButton;
    tbCatalogBtnMatch: TToolButton;
    tbMatch: TToolBar;
    tbMatchBtnEdit: TToolButton;
    tbMatchBtnDelete: TToolButton;
    tbMatchBtnGoToGroup: TToolButton;
    tbPrice: TToolBar;
    tbCatalogPosirionMatching: TToolBar;
    tbCatalogInfo: TToolButton;
    tbTree: TToolBar;
    tbTreeBtnExpand: TToolButton;
    tbTreeBtnSort: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    tbTreeBtnShowChild: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    tbEditMode: TToolButton;
    tbAddInvoce: TToolButton;
    TreeCatalog: TTreeView;
    TreeGroupOwner: TTreeView;
    TreeViewInfo: TTreeView;
    XMLSession: TXMLPropStorage;

    procedure btnedMatchSearchClearClick(Sender: TObject);
    procedure btnPriceEditSearchClearClick(Sender: TObject);
    procedure cbPriceFieldChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure GridCatalogPositionMatchingDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure mCatalogMatchingPositionSelectAllClick(Sender: TObject);
    procedure mCatalogMatchingPositionSelectClearClick(Sender: TObject);
    procedure mClipboardAllClick(Sender: TObject);
    procedure mClipboardPosVNLClick(Sender: TObject);
    procedure mClipboardPosVNLPClick(Sender: TObject);
    procedure mClipboardVNLClick(Sender: TObject);
    procedure mClipboardVNLPClick(Sender: TObject);
    procedure mClipboardPosVNLPKClick(Sender: TObject);
    procedure mExportInSpreadsheetClick(Sender: TObject);
    procedure mExportInSpreadsheetMGClick(Sender: TObject);
    procedure mExportInSpreadsheetMPGClick(Sender: TObject);
    procedure mMatchingPositionAddClick(Sender: TObject);
    procedure mMatchingPositionDeleteClick(Sender: TObject);
    procedure mMatchingPositionEditClick(Sender: TObject);
    procedure mMatchingPositionEditQuantInPackClick(Sender: TObject);
    procedure MenuItem6Click(Sender: TObject);
    procedure MenuItem7Click(Sender: TObject);
    procedure mNomClearSelectClick(Sender: TObject);
    procedure mPriceGridPopup(Sender: TObject);
    procedure mSelectAllClick(Sender: TObject);
    procedure mTmExportClick(Sender: TObject);
    procedure mTreeInfoCopyOneClick(Sender: TObject);
    procedure pcCatalogChange(Sender: TObject);
    procedure sBtnMatchingPriceStockOnlyClick(Sender: TObject);
    procedure sBtnMatchingSelectedClick(Sender: TObject);
    procedure sBtnNoMatchingClick(Sender: TObject);
    procedure sBtnSelectedClick(Sender: TObject);
    procedure sBtnStockOnlyClick(Sender: TObject);
    procedure sBtnWithMatchingClick(Sender: TObject);
    procedure tbCatalogBtnAddClick(Sender: TObject);
    procedure tbCatalogBtnDeleteClick(Sender: TObject);
    procedure tbCatalogBtnEditClick(Sender: TObject);
    procedure tbCatalogBtnGoToGroupClick(Sender: TObject);
    procedure tbCatalogBtnCopyClick(Sender: TObject);
    procedure tbCatalogBtnMatchClick(Sender: TObject);
    procedure tbEditModeClick(Sender: TObject);
    procedure tbMatchBtnDeleteClick(Sender: TObject);
    procedure tbMatchBtnEditClick(Sender: TObject);
    procedure tbMatchBtnGoToGroupClick(Sender: TObject);
    procedure tbCatalogInfoClick(Sender: TObject);
    procedure tbTreeBtnExpandClick(Sender: TObject);
    procedure tbTreeBtnSortClick(Sender: TObject);
    procedure tbTreeBtnShowChildClick(Sender: TObject);
    procedure tbAddInvoceClick(Sender: TObject);
    procedure TreeCatalogGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeCatalogGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeViewInfoAdvancedCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; Stage: TCustomDrawStage; var PaintImages,
      DefaultDraw: Boolean);
    procedure TreeViewInfoDblClick(Sender: TObject);
    procedure TreeViewInfoGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure XMLSessionRestoreProperties(Sender: TObject);
    procedure XMLSessionSavingProperties(Sender: TObject);
  private
    FormIDent: string;
    OwnerID: integer;     // ID выбранного контрагента
    Catalog: TCatalog;
    fBase: TwBase;
    IdMainOwner: string; // ID основного контрагента (к которому привязан каталог)
    fInvoce: TInvoce;

    property wFormID: string read FormIDent write FormIDent;

  public
    property Base: TwBase read fBase;
    procedure SetStatus(_Text:string);

    { public declarations }
  end;

var
  FmCatalog: TFmCatalog;
  IDint:integer;

implementation

{$R *.lfm}

{ TFmCatalog }

procedure TFmCatalog.SetStatus(_Text: string);
begin
     wStatus(wFormID,_Text,true);
     Application.ProcessMessages;
end;

procedure TFmCatalog.FormCreate(Sender: TObject);
//var
//  _TreeMenu: TPopupMenu;
//  TreeMenuSpliter, TreeMenuEditGridItems: TMenuItem;
//  _SQL_string, _SQL_PositionMAtchingString: String;
begin
      wFormID:=Self.Name;
      screen.Cursor:= crSQLWait;

      OwnerID:=0; // инициализируем.
      wLog('Catalog','Инициализация плагина... ['+wFormID+']');

      fBase:= TwBase.Create(Sender);
      IdMainOwner:= Base.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист

      try
      Catalog:= TCatalog.Create(Sender,Base,GridCatalogPrice,GridMatching,GridCatalogPositionMatching,TreeCatalog,TreeGroupOwner);
      Catalog.TreeInfo:= TreeViewInfo;

      // заполняем список цен
      cmbxFill(cbPriceField,Base.SQLReadDS('PRICEFIELD',['NAME','ID'],'FCLOSE=0','PRIORITY'),['NAME','ID']);
      // Инициализируем формулу
      Catalog.GridCatalog.Formula:= TFormula.Create(Sender as TComponent);
      Catalog.GridCatalog.Formula.CalculateField:= 'CPRICE';
      Catalog.GridCatalog.Formula.CurrencyArray:= Base.GetCurrencyArray();
      Catalog.GridCatalog.FormulaText:=Catalog.GridCatalog.Formula.Prepare(string(Base.SQLReadArr('PRICEFIELD',['FORMULA'],'ID='+IntToStr(cmbxSelectID(cbPriceField)),'')[0,0]));

      Catalog.GridCatalogFill(nil);
      Catalog.TreeCatalogFill();
      fInvoce:= TInvoce.Create(Base, nil);

      screen.Cursor:= crDefault;
      except
        on E: Exception do
        begin
            screen.Cursor:= crDefault;
            SetStatus('Сбой инициализации плагина.');
            wLog('Catalog','Ошибка [FmCreate]: "' + E.Message + '"');
            wLog('Catalog','Сбой инициализации плагина.');
            ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

         end;
      end;
end;

procedure TFmCatalog.cbPriceFieldChange(Sender: TObject);
var
  _BookMark: TBookMark;
begin
  if not Assigned(GridCatalogPrice.DataSource) then exit;

  Catalog.GridCatalog.Formula.CurrencyArray:= fBase.GetCurrencyArray();
  Catalog.GridCatalog.FormulaText:=Catalog.GridCatalog.Formula.Prepare(string(fBase.SQLReadArr('PRICEFIELD',['FORMULA'],'ID='+IntToStr(cmbxSelectID(cbPriceField)),'')[0,0]));


  with GridCatalogPrice.DataSource.DataSet do begin
     _BookMark:= Bookmark;
     Catalog.GridCatalog.Fill();
     Bookmark:= _BookMark;
  end;


end;

procedure TFmCatalog.btnPriceEditSearchClearClick(Sender: TObject);
begin
  edPriceSearch.Text:='';
  edPriceSearch.OnChange(edPriceSearch);
end;

procedure TFmCatalog.btnedMatchSearchClearClick(Sender: TObject);
begin
  edMatchSearch.Text:='';
  edMatchSearch.OnChange(edMatchSearch);
end;

procedure TFmCatalog.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  CloseAction := caFree;
end;

procedure TFmCatalog.FormDestroy(Sender: TObject);
begin
      try
       wLog('Catalog','Выгрузка плагина...');

    Catalog.Destroy();
    Base.Destroy();
    fInvoce.Destroy;
    cmbxClearData(cbPriceField);  // очищаем объекты комбобокса


      wLog('Catalog','Выгрузка плагина успешно завершена.');

      except
        on E: Exception do
        begin
            SetStatus('Сбой выгрузки плагина: Каталог.');
            wLog('Catalog','Ошибка [FmDestroy]: "' + E.Message + '"');
            wLog('Catalog','Сбой выгрузки плагина.');
            ShowMessage('Ошибка [FmDestroy]: "' + E.Message + '"');
         end;
      end;
end;

procedure TFmCatalog.GridCatalogPositionMatchingDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  _ColumnText: string;
  _ImageIndex: integer;
  _FieldValue: Double;
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

        if Catalog.GridMatchingPosition.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         begin
                  _ImageIndex:= 0;
                  // А теперь пусть ImageList нарисует ее на канве DBGrid'а
                  TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
         end;
     end else
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);

        if Catalog.GridMatchingPosition.SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
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
         if TDBGrid(Sender).DataSource.DataSet.FieldByName('STOCKONLYINFO').AsInteger>0 then
            begin
              FillRect(Rect);
              if gdSelected in State
                 then
                    Font.Color:=clRed
                 else
                    Font.Color:=clBlue;

                if TDBGrid(Sender).DataSource.DataSet.FieldByName('QUANTITYINPACKING').AsFloat<>1 then
                   Font.Style := [fsBold];

               TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
            end else
            begin
              if TDBGrid(Sender).DataSource.DataSet.FieldByName('QUANTITYINPACKING').AsFloat<>1 then
                 begin
                   FillRect(Rect);
                   Font.Style := [fsBold];
                   TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
                 end;
            end;
    end;
end;

procedure TFmCatalog.mCatalogMatchingPositionSelectAllClick(Sender: TObject);
begin
  Catalog.GridMatchingPosition.SelectAll:= true;
end;

procedure TFmCatalog.mCatalogMatchingPositionSelectClearClick(Sender: TObject);
begin
  Catalog.GridMatchingPosition.SelectAll:= false;
end;

procedure TFmCatalog.mClipboardAllClick(Sender: TObject);
begin
  Catalog.GridCatalog.CopyToClipboard(nil,['CPRICE']);
end;

procedure TFmCatalog.mClipboardPosVNLClick(Sender: TObject);
begin
  Catalog.GridMatchingPosition.CopyToClipboard(['PLVENDORCODE', 'PLNAME', 'PLLABEL'],[''],'MTH.ID');
end;

procedure TFmCatalog.mClipboardPosVNLPClick(Sender: TObject);
begin
  Catalog.GridMatchingPosition.CopyToClipboard(['PLVENDORCODE', 'PLNAME', 'PLLABEL', 'PRICE'],['PRICE'],'MTH.ID');
end;

procedure TFmCatalog.mClipboardVNLClick(Sender: TObject);
begin
  Catalog.GridCatalog.CopyToClipboard(['SCOD', 'NAME', 'UNIT', 'LABEL'],[],'CATALOG.ID');
end;

procedure TFmCatalog.mClipboardVNLPClick(Sender: TObject);
begin
  Catalog.GridCatalog.CopyToClipboard(['SCOD', 'NAME', 'UNIT', 'LABEL', 'CPRICE'],['CPRICE'],'CATALOG.ID');
end;

procedure TFmCatalog.mClipboardPosVNLPKClick(Sender: TObject);
begin
  Catalog.GridMatchingPosition.CopyToClipboard(['PLVENDORCODE', 'PLNAME', 'PLLABEL', 'PRICE', 'OWNERNAME'], ['PRICE'],'MTH.ID');
end;

procedure TFmCatalog.mExportInSpreadsheetClick(Sender: TObject);
begin
  Catalog.GridCatalog.ExportData();
end;

procedure TFmCatalog.mExportInSpreadsheetMGClick(Sender: TObject);
begin
  Catalog.GridMatching.ExportData();
end;

procedure TFmCatalog.mExportInSpreadsheetMPGClick(Sender: TObject);
begin
  Catalog.GridMatchingPosition.ExportData();
end;

procedure TFmCatalog.mMatchingPositionAddClick(Sender: TObject);
var
  _Form: TFmListSelect;
  _SelectedRows: ArrayOfInteger;
  _TimeStamp: String;
  _IDCatalog: LongInt;
  _QuantityInPacked: Double;
  i, _IDMatching: Integer;
  _arr: ArrayOfArrayVariant;
  _GridDataset: TDataSet;
  _ScodMenu: TPopupMenu;
  _ScodMenuItem: TMenuItem;
begin
  _GridDataset:= GridCatalogPrice.DataSource.DataSet;
if _GridDataset.RecordCount = 0 then exit;

screen.Cursor:= crSQLWait;
_Form:= TFmListSelect.Create(self);
_Form.Base:= fBase;
_Form.MultiSelectGrid:= true;
_Form.wFormMode:= 0; // PriceLists
_Form.Where:= 'ID<>'+IdMainOwner;

_Form.ListFormInit(
    _GridDataSet.FieldByName('ID').AsString,
    _GridDataSet.FieldByName('VENDORCODE').AsString,
    _GridDataSet.FieldByName('NAME').AsString,
    _GridDataSet.FieldByName('LABEL').AsString
);

_SelectedRows:= nil;
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
if _SelectedRows<> nil then
  begin
    _TimeStamp:= DateTimeToStr(Now());

    _IDCatalog:= _GridDataset.FieldByName('ID').AsInteger;

    try
    for i:=0 to High(_SelectedRows) do
       begin
         _arr:= nil;
          _arr:= fBase.SQLReadArr('PL_ITEMS',['IDOWNER','ID'],'ID='+IntToStr(_SelectedRows[i]),'ID');
          if Length(_arr)>0 then
              //_IDMatching:= fBase.SQLInsert('INSERT INTO "CATALOG_MATCHING" (IDOWNER,IDCATALOG,IDPL_ITEMS,QUANTITYINPACKING,FTIMESTAMP,IDUSER) VALUES ('+IntTOStr(_arr[0,0])+','+IntToStr(_IDCatalog)+','+QuotedStr(_arr[0,1])+','+FloatToStr(_QuantityInPacked)+','+QuotedStr(_TimeStamp)+',1)');
               _IDMatching:= fBase.SQLInsert('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],[_arr[0,0],_IDCatalog,_arr[0,1],_QuantityInPacked,_TimeStamp,integer(1)],'IDOWNER, IDPL_ITEMS',false);
       end;
    fBase.SQLTransactionEnd(true);
    except
      on E: Exception do
      begin
          fBase.SQLTransactionEnd(false);
          __Log.SaveLogError(E);
          SetStatus('Сбой добавления соответствий.');
          wLog('FmNomenclatureEdit','Ошибка: "' + E.Message + '"');
          wLog('FmNomenclatureEdit','Сбой добавления соответствий.');
          ShowMessage('Ошибка: "' + E.Message + '"');
       end;
    end;

       try
        Catalog.GridCatalog.Fill();
      if _GridDataset.RecordCount>0 then
         _GridDataset.Locate('ID',_IDCatalog,[]);
       finally
        if GridCatalogPositionMatching.DataSource.DataSet.RecordCount>0 then
                 GridCatalogPositionMatching.DataSource.DataSet.Locate('ID',IntTOStr(_IDMatching),[loCaseInsensitive]);
       end;

  end;
  _arr:= nil;
  _SelectedRows:= nil;
  screen.Cursor:= crDefault;
end;

procedure TFmCatalog.mMatchingPositionDeleteClick(Sender: TObject);
var
  i: Integer;
  _BookMark: TBookMark;
  _SelectedRows: ArrayOfInteger;
begin
  if TreeCatalog.Items.Count = 0 then exit;
  _SelectedRows:=nil;
  _SelectedRows:= Catalog.GridMatchingPosition.SelectedRows;

  if (Length(_SelectedRows)>0) and (Catalog.GridMatchingPosition.Grid.DataSource.DataSet.RecordCount>0) then
     begin
       if MessageDlg('Удалить выделенные соответствия ('+IntTOStr(Length(_SelectedRows))+') ?',mtConfirmation, mbOKCancel, 0) = mrOK
        then
         begin

             try
             _BookMark:= Catalog.GridMatchingPosition.Grid.DataSource.DataSet.Bookmark;

             for i:=0 to High(_SelectedRows) do
                fBase.SQLDelete('CATALOG_MATCHING','ID='+IntToStr(_SelectedRows[i]),false);
              fBase.SQLTransactionEnd(true);
              Catalog.GridMatchingPosition.Fill();
              if Catalog.GridMatchingPosition.Grid.DataSource.DataSet.RecordCount>0 then Catalog.GridMatchingPosition.Grid.DataSource.DataSet.Bookmark:= _BookMark;

              _BookMark:= Catalog.GridCatalog.Grid.DataSource.DataSet.Bookmark;

              Catalog.GridCatalogFiltered();
              if Catalog.GridCatalog.Grid.DataSource.DataSet.RecordCount>0 then
                       Catalog.GridCatalog.Grid.DataSource.DataSet.Bookmark:= _BookMark;

             except
               on E: Exception do
               begin
                   __Log.SaveLogError(E);
                   fBase.SQLTransactionEnd(false);
                   SetStatus('Сбой удаления соответствий.');
                   wLog(wFormID,'Ошибка: "' + E.Message + '"');
                   wLog(wFormID,'Сбой удаления соответствий.');
                   ShowMessage('Ошибка: "' + E.Message + '"');
                end;
             end;

         end;
     end;

  _SelectedRows:=nil;
end;

procedure TFmCatalog.mMatchingPositionEditClick(Sender: TObject);
var
  _Form: TFmListSelect;
  _IDMatching, i: Integer;
  _TimeStamp, _SelectedID: String;
  _QuantityInPackedGrid, _QuantityInPacked: Double;
  _SelectedRows: ArrayOfInteger;
  _IDCatalog: LongInt;
  _arr: ArrayOfArrayVariant;
  _GridDataset: TDataSet;
begin
// изменение одного соответствия
  if GridCatalogPositionMatching.DataSource.DataSet.RecordCount=0 then exit;
  _GridDataset:= GridCatalogPrice.DataSource.DataSet;
  try
    _IDMatching:= Catalog.GridMatchingPosition.SelectedRows[0];
    _SelectedID:= GridCatalogPositionMatching.DataSource.DataSet.FieldByName('IDPL_ITEMS').AsString;
    _QuantityInPackedGrid:= GridCatalogPositionMatching.DataSource.DataSet.FieldByName('QUANTITYINPACKING').AsFloat;
    screen.Cursor:= crSQLWait;

    _Form:= TFmListSelect.Create(self);
    _Form.Base:= fBase;
    _Form.MultiSelectGrid:= true;
    _Form.wFormMode:= 0; // PriceLists
    _Form.Where:= 'ID<>'+IdMainOwner;

    _Form.ListFormInit(
        _GridDataSet.FieldByName('ID').AsString,
        _GridDataSet.FieldByName('VENDORCODE').AsString,
        _GridDataSet.FieldByName('NAME').AsString,
        _GridDataSet.FieldByName('LABEL').AsString
    );

    _SelectedRows:= nil;

    _Form.GridList.Options:=_Form.GridList.Options - [dgMultiSelect];
    _Form.wDataSetLocateField:='ID';
    _Form.wDataSetLocateValue:=_SelectedID;
    _Form.wIDTreeItem:=GridCatalogPositionMatching.DataSource.DataSet.FieldByName('IDOWNER').AsInteger;

    if _QuantityInPackedGrid<1 then
       begin
         _Form.spQuantInPackLeft.Value:=1;
         _Form.spQuantInPackRight.Value:= _wRNDTO(1/_QuantityInPackedGrid,0);
       end else
       begin
          _Form.spQuantInPackLeft.Value:= _wRNDTO(1*_QuantityInPackedGrid,0);
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
    if _SelectedRows<> nil then
       begin
         _TimeStamp:= DateTimeToStr(Now());
         _IDCatalog:= GridCatalogPrice.DataSource.DataSet.FieldByName('ID').AsInteger;

         try
           _arr:= nil;
           _arr:= fBase.SQLReadArr('PL_ITEMS',['IDOWNER','ID'],'ID='+IntToStr(_SelectedRows[0]),'ID');
           fBase.SQLUpdate('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],[_arr[0,0],_IDCatalog,_arr[0,1],_QuantityInPacked,_TimeStamp,integer(1)],'ID='+IntToStr(_IDMatching),false);

         fBase.SQLTransactionEnd(true);
         except
           fBase.SQLTransactionEnd(false);
           raise;
         end;
         Catalog.GridMatchingPosition.Fill;
         GridCatalogPositionMatching.DataSource.DataSet.Locate('ID',IntTOStr(_IDMatching),[loCaseInsensitive]);
       end;
       _arr:= nil;
       _SelectedRows:= nil;
       screen.Cursor:= crDefault;
  except
    on E: Exception do
    begin
        __Log.SaveLogError(E);
        SetStatus('Сбой изменения соответствия.');
        wLog('FmNomenclatureEdit','Ошибка: "' + E.Message + '"');
        wLog('FmNomenclatureEdit','Сбой изменения соответствия.');
        ShowMessage('Ошибка: "' + E.Message + '"');
     end;
  end;
   screen.Cursor:= crDefault;
end;

procedure TFmCatalog.mMatchingPositionEditQuantInPackClick(Sender: TObject);
var
  _SelectedRows: ArrayOfInteger;
  _Form: TFmQuantityInPacking;
  _QuantityInPacked: Extended;
  _TimeStamp: String;
  i: Integer;
begin
// изменение одного соответствия
  try
    _SelectedRows:= Catalog.GridMatchingPosition.SelectedRows;

    _Form:= TFmQuantityInPacking.Create(self);

    try
    _Form.ShowModal;
    finally
      if _Form.ModalResult <> mrCancel then
      begin
         _QuantityInPacked:= _Form.spQuantInPackLeft.Value/_Form.spQuantInPackRight.Value;
      end else _QuantityInPacked:=0;
     _Form.Free;
    end;
    if _QuantityInPacked <> 0 then
       begin
         _TimeStamp:= DateTimeToStr(Now());

         try
           for i:=0 to High(_SelectedRows) do
                fBase.SQLUpdate('CATALOG_MATCHING',['QUANTITYINPACKING','FTIMESTAMP','IDUSER'],[_QuantityInPacked,_TimeStamp,integer(1)],'ID='+IntToStr(_SelectedRows[i]),false);

         fBase.SQLTransactionEnd(true);
         except
           fBase.SQLTransactionEnd(false);
           raise;
         end;
         Catalog.GridMatchingPosition.Fill;
         GridCatalogPositionMatching.DataSource.DataSet.Locate('ID',IntTOStr(_SelectedRows[i]),[loCaseInsensitive]);
       end;
       _SelectedRows:= nil;
  except
    on E: Exception do
    begin
        __Log.SaveLogError(E);
        SetStatus('Сбой изменения фасовки.');
        wLog('FmNomenclatureEdit','Ошибка: "' + E.Message + '"');
        wLog('FmNomenclatureEdit','Сбой изменения фасовки.');
        ShowMessage('Ошибка: "' + E.Message + '"');
     end;
  end;
end;

procedure TFmCatalog.MenuItem6Click(Sender: TObject);
begin
  Catalog.GridMatching.SelectAll:= true;
end;

procedure TFmCatalog.MenuItem7Click(Sender: TObject);
var
  _SelCount: Integer;
begin
  _SelCount:= Catalog.GridMatching.SelectedRowsCount;

  if _SelCount>0 then
    begin
      if MessageDlg('Сбросить выделение ('+IntToStr(_SelCount)+') позиций?',mtConfirmation, mbOKCancel, 0) = mrOK then
       begin
          Catalog.GridMatching.SelectAll:= false;
       end;
    end;
end;

procedure TFmCatalog.mNomClearSelectClick(Sender: TObject);
var
  _SelCount: Integer;
begin

   _SelCount:= Catalog.GridCatalog.SelectedRowsCount;

   if _SelCount>0 then
     begin
       if MessageDlg('Сбросить выделение ('+IntToStr(_SelCount)+') позиций?',mtConfirmation, mbOKCancel, 0) = mrOK then
        begin
           Catalog.GridCatalog.SelectAll:= false;
        end;
     end;
end;

procedure TFmCatalog.mPriceGridPopup(Sender: TObject);
begin
  mNomSootv.Checked:= tbCatalogBtnMatch.Down;
  mCatalogInfo.Checked:= gbInfo.Visible;
end;

procedure TFmCatalog.mSelectAllClick(Sender: TObject);
begin
  Catalog.GridCatalog.SelectAll:= true;
end;

procedure TFmCatalog.mTmExportClick(Sender: TObject);
begin
  Catalog.ExportData();
end;

procedure TFmCatalog.mTreeInfoCopyOneClick(Sender: TObject);
begin
  Clipboard.AsText:= TreeViewInfo.Selected.Text;
end;

procedure TFmCatalog.pcCatalogChange(Sender: TObject);
begin
try
  if TreeGroupOwner.Showing then
    begin
        Catalog.GridMatchingFill(nil);
        Catalog.TreeMatchingFill();
    end
  else
  begin
      Catalog.GridCatalogFill(nil);
      Catalog.TreeCatalogFill();
  end;
except
      on E: Exception do
  begin
     wLog('Catalog','Ошибка [pcCatalogChange]: "' + E.Message + '"');
     ShowMessage('Ошибка [pcCatalogChange]: "' + E.Message + '"');
  end;
end;
end;

procedure TFmCatalog.sBtnMatchingPriceStockOnlyClick(Sender: TObject);
begin
  Catalog.GridMatchingFiltered();
end;

procedure TFmCatalog.sBtnMatchingSelectedClick(Sender: TObject);
begin
  Catalog.GridMatchingFiltered();
end;

procedure TFmCatalog.sBtnNoMatchingClick(Sender: TObject);
begin
  Catalog.GridCatalogFiltered();
end;

procedure TFmCatalog.sBtnSelectedClick(Sender: TObject);
begin
  Catalog.GridCatalogFiltered();
end;

procedure TFmCatalog.sBtnStockOnlyClick(Sender: TObject);
begin
  Catalog.GridCatalogFiltered();
end;

procedure TFmCatalog.sBtnWithMatchingClick(Sender: TObject);
begin
  Catalog.GridCatalogFiltered();
end;

procedure TFmCatalog.tbCatalogBtnAddClick(Sender: TObject);
begin
  Catalog.ItemAdd(Catalog.TreeCatalog, GridCatalogPrice.DataSource.DataSet);
end;

procedure TFmCatalog.tbCatalogBtnDeleteClick(Sender: TObject);
var
  awGrid: TwDBGrid;
begin
      if TreeCatalog.Items.Count=0 then exit;
      if TreeCatalog.Selected.Text = '' then exit;
      awGrid:= Catalog.GridCatalog;

     Catalog.ItemDel(awGrid);
end;

procedure TFmCatalog.tbCatalogBtnEditClick(Sender: TObject);
var
  _GridDataset: TDataSet;
  _CatalogTree: TwDBTree;
  _GridCatalog: TwDBGrid;
begin
  _CatalogTree:= Catalog.TreeCatalog;
  _GridDataset:= GridCatalogPrice.DataSource.DataSet;
  _GridCatalog:= Catalog.GridCatalog;

  Catalog.ItemEdit(_GridCatalog, _CatalogTree);
end;

procedure TFmCatalog.tbCatalogBtnGoToGroupClick(Sender: TObject);
var
  _GridDataset: TDataSet;
  _ParentID: integer;
  _ID: integer;
begin
  if TreeCatalog.Items.Count=0 then exit;

  if (TreeCatalog.Selected.Text = '') then exit;

  _GridDataset:= GridCatalogPrice.DataSource.DataSet;
  _ParentID:= _GridDataset.FieldByName('IDCTG_GROUP').AsInteger;
  _ID:= _GridDataset.FieldByName('ID').AsInteger;
  try
  Catalog.TreeCatalog.FindNodeWithDataInt(_ParentID);
  finally

  if _GridDataset.RecordCount>0 then
    _GridDataset.Locate('ID',_ID,[]);
  end;
end;

procedure TFmCatalog.tbCatalogBtnCopyClick(Sender: TObject);
begin
  Catalog.ItemCopy(Catalog.TreeCatalog, Catalog.GridCatalog);
end;

procedure TFmCatalog.tbCatalogBtnMatchClick(Sender: TObject);
var
  _GridDataset: TDataSet;
begin

if not Assigned(GridCatalogPrice.DataSource) then exit;

if tbCatalogBtnMatch.Marked then
   begin
     if GridCatalogPrice.DataSource.DataSet.RecordCount<>0 then
        begin
          try
             Catalog.GridCatalogPositionMatchingFill([GridCatalogPrice.DataSource.DataSet.FieldByName('ID').AsInteger]);// FillDBGridCatalogPositionMatching();
          finally
            pCatalogPositionMatching.Height:=200;
            GridCatalogPositionMatching.Visible:=true;
            splitCatalogPositionMatching.Visible:=true;
            tbCatalogBtnMatch.Marked:=false;
            tbCatalogBtnMatch.Down:=true;
          end;
        end else
        begin
           ShowMessage('Каталог пуст!');
        end;

   end
     else
   begin
      pCatalogPositionMatching.Height:=0;//tbPrice.Height;
      GridCatalogPositionMatching.Visible:=false;
      splitCatalogPositionMatching.Visible:=false;
      tbCatalogBtnMatch.Marked:=true;
      tbCatalogBtnMatch.Down:=false;
   end;

_GridDataset:= GridCatalogPrice.DataSource.DataSet;

  if (TreeCatalog.Selected.Text = '') or (_GridDataset.RecordCount = 0) then exit;

  screen.Cursor:= crHourGlass;

  screen.Cursor:= crDefault;
end;

procedure TFmCatalog.tbEditModeClick(Sender: TObject);
begin
  if TToolButton(Sender).Down then
     begin
      Catalog.GridCatalog.Grid.Enabled:= false;
      Catalog.GridCatalog.Grid.Color:= clForm;
     end
    else
    begin
      Catalog.GridCatalog.Grid.Enabled:= true;
      Catalog.GridCatalog.Grid.Color:= clWindow;
    end;
end;

procedure TFmCatalog.tbMatchBtnDeleteClick(Sender: TObject);
var
  i: Integer;
  _BookMark: TBookMark;
  _SelectedRows: ArrayOfInteger;
begin
  if TreeGroupOwner.Items.Count = 0 then exit;
  _SelectedRows:=nil;
  _SelectedRows:= Catalog.GridMatching.SelectedRows;

  if (Length(_SelectedRows)>0) and (Catalog.GridMatching.Grid.DataSource.DataSet.RecordCount>0) then
     begin
       if MessageDlg('Удалить выделенные соответствия ('+IntTOStr(Length(_SelectedRows))+') ?',mtConfirmation, mbOKCancel, 0) = mrOK
        then
         begin

             try
             _BookMark:= Catalog.GridMatching.Grid.DataSource.DataSet.Bookmark;

             for i:=0 to High(_SelectedRows) do
                fBase.SQLDelete('CATALOG_MATCHING','ID='+IntToStr(_SelectedRows[i]),false);
              fBase.SQLTransactionEnd(true);
              Catalog.GridMatching.Fill();
              if Catalog.GridMatching.Grid.DataSource.DataSet.RecordCount>0 then Catalog.GridMatching.Grid.DataSource.DataSet.Bookmark:= _BookMark;
             except
               on E: Exception do
               begin
                   __Log.SaveLogError(E);
                   fBase.SQLTransactionEnd(false);
                   SetStatus('Сбой удаления соответствий.');
                   wLog('FmNomenclatureEdit','Ошибка: "' + E.Message + '"');
                   wLog('FmNomenclatureEdit','Сбой удаления соответствий.');
                   ShowMessage('Ошибка: "' + E.Message + '"');
                end;
             end;

         end;
     end;

  _SelectedRows:=nil;
end;

procedure TFmCatalog.tbMatchBtnEditClick(Sender: TObject);
var
  _SelectedID, _IDOwner, _IDOwnerOld, _IDCatalog, _IDPrice: Integer;
  _TimeStamp: string;
  _GridDataset: TDataSet;
  _SelectedName, _SelectedScod, _SelectedLabel,
    _SelectedVendorCode, _SelectedOwnerName,
    _SelectedOwnerNomenclatureName: String;
  _Form: TFmMatchingEdit;
  _SelectedQuantInPacked: Double;
  _SelectedIDCatalog, _SelectedIDOwner, _SelectedIDCatalogGroup: LongInt;
  _QuantityInPacked: Extended;
  _Barcodes: ArrayOfArrayVariant;
begin

_SelectedID:= 0;
_TimeStamp:= DateTimeToStr(Now());

SetStatus('Изменение соответствия...');
wLog('Catalog','Изменение соответствия...');

if TreeGroupOwner.Selected.Text = '' then exit;

_GridDataset:= GridMatching.DataSource.DataSet;

if _GridDataset.RecordCount>0 then
   begin
     _SelectedID:= _GridDataset.FieldByName('ID').AsInteger;
     _SelectedIDCatalog:= _GridDataset.FieldByName('IDCATALOG').AsInteger;
     _SelectedIDCatalogGroup:= _GridDataset.FieldByName('IDCTG_GROUP').AsInteger;
     _SelectedIDOwner:= _GridDataset.FieldByName('IDOWNER').AsInteger;
     _SelectedName:= _GridDataset.FieldByName('CATALOGNAME').AsString;
     _SelectedQuantInPacked:= _GridDataset.FieldByName('QUANTITYINPACKING').AsFloat;

     _Barcodes:=nil;
     _Barcodes:= fBase.SQLReadArr('SELECT VSCOD FROM CTG_GET_SCOD('+IntToStr(_SelectedIDCatalog)+',true)');
     if Assigned(_Barcodes) and (_Barcodes[0,0]<>null) then
               _SelectedScod:= _Barcodes[0,0] else
               _SelectedScod:= '';

     //_SelectedScod:= _GridDataset.FieldByName('CATALOGSCOD').AsString;
     _SelectedLabel:= _GridDataset.FieldByName('CATALOGLABEL').AsString;
     _SelectedVendorCode:= _GridDataset.FieldByName('PLVENDORCODE').AsString;
     _SelectedOwnerName:= _GridDataset.FieldByName('OWNERNAME').AsString;
     _SelectedOwnerNomenclatureName:= _GridDataset.FieldByName('PLNAME').AsString;
   end
   else exit;

_IDOwnerOld:= _SelectedIDOwner;

_Form:= TFmMatchingEdit.Create(Self);

  with _Form do begin

    IDPrice:= _GridDataset.FieldByName('IDPL_ITEMS').AsInteger;
    IDOwner:= _SelectedIDOwner;
    IDCatalog:= _SelectedIDCatalog;
    IDCatalogGroup:= _SelectedIDCatalogGroup;

    Base:= fBase;

    kName.Text:=_SelectedName;
    kScod.Text:=_SelectedScod;
    kLabel.Text:=_SelectedLabel;
    kVendorCode.Text:=_SelectedVendorCode;
    kOwner.Text:=_SelectedOwnerName;
    kOwnerNomenclatureName.Text:=_SelectedOwnerNomenclatureName;

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


           if not fBase.SQLUpdate('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],[_IDOwner,_IDCatalog,_IDPrice,_QuantityInPacked,_TimeStamp,integer(1)],'ID='+IntToStr(_SelectedID))
           then
             begin
               SetStatus('Изменение соответствия завершено с ошибкой.');
               wLog('Catalog','Изменение соответствия завершено с ошибкой.');
             end else
             begin
              if _SelectedIDOwner <> _IDOwner then
                begin
                  try
                    Catalog.TreeMatching.FindNodeWithDataInt(_IDOwner);
                  finally
                    _GridDataset:= GridMatching.DataSource.DataSet;
                    if _GridDataset.RecordCount>0 then
                       _GridDataset.Locate('ID',_SelectedID,[]);
                  end;
                end else
                begin
                     _GridDataset:= GridMatching.DataSource.DataSet;
                     _GridDataset.Close;
                     _GridDataset.Open;
                     if _GridDataset.RecordCount>0 then
                        _GridDataset.Locate('ID',_SelectedID,[]);
                end;
             end;

         _Form.Free;

       end;

    end;

end;

procedure TFmCatalog.tbMatchBtnGoToGroupClick(Sender: TObject);
var
  _GridDataset: TDataSet;
  _ParentID: integer;
  _ID: integer;
begin
  if TreeGroupOwner.Items.Count = 0 then exit;
  if TreeGroupOwner.Selected.Text = '' then exit;

  _GridDataset:= GridMatching.DataSource.DataSet;
  _ParentID:= _GridDataset.FieldByName('IDOWNER').AsInteger;
  _ID:= _GridDataset.FieldByName('ID').AsInteger;
  try
    Catalog.TreeMatching.FindNodeWithDataInt(_ParentID);
  finally

  if _GridDataset.RecordCount>0 then
    _GridDataset.Locate('ID',_ID,[]);
  end;
end;

procedure TFmCatalog.tbCatalogInfoClick(Sender: TObject);
begin
  if tbCatalogInfo.Marked then
     begin
       if GridCatalogPrice.DataSource.DataSet.RecordCount<>0 then
          begin
             tbCatalogInfo.Marked:=false;
             tbCatalogInfo.Down:=true;
             gbInfo.Visible:= true;
             SplitterInfo.Visible:= true;
             Catalog.TreeInfoFill(Catalog.GridCatalog);
          end else
          begin
             ShowMessage('Прайс-лист пуст!');
          end;

     end
       else
     begin
        tbCatalogInfo.Marked:=true;
        tbCatalogInfo.Down:=false;
        gbInfo.Visible:= false;
        SplitterInfo.Visible:= false;
     end;
end;

procedure TFmCatalog.tbTreeBtnExpandClick(Sender: TObject);
begin
  if TreeCatalog.Items.Count=0 then exit;
  if tbTreeBtnExpand.Marked then
     begin
       Catalog.TreeCatalog.Expanded:=false;
       Catalog.TreeCatalog.Tree.FullCollapse;
       Catalog.TreeCatalog.Tree.Items[0].Expanded:= true;
       tbTreeBtnExpand.Marked:=false;
     end
       else
     begin
        Catalog.TreeCatalog.Expanded:=true;
        Catalog.TreeCatalog.Tree.FullExpand;
        tbTreeBtnExpand.Marked:=true;
     end;
end;

procedure TFmCatalog.tbTreeBtnSortClick(Sender: TObject);
begin
  if TreeCatalog.Items.Count=0 then exit;

  if tbTreeBtnSort.Marked then
     begin
       Catalog.TreeCatalog.OrderBy:='IDPARENT, ID';
       Catalog.TreeCatalog.Tree.SortType:= stNone;
       Catalog.TreeCatalog.Fill();
       //Catalog.TreeCatalog.Tree.AlphaSort;
       tbTreeBtnSort.Marked:=false;
     end
       else
     begin
        Catalog.TreeCatalog.OrderBy:='IDPARENT, NAME';
        Catalog.TreeCatalog.Tree.SortType:= stText;
        Catalog.TreeCatalog.Fill();
        //Catalog.TreeCatalog.Tree.AlphaSort;
        tbTreeBtnSort.Marked:=true;
     end;
end;

procedure TFmCatalog.tbTreeBtnShowChildClick(Sender: TObject);
begin
  if TreeCatalog.Items.Count=0 then exit;

  if tbTreeBtnShowChild.Marked then
     begin
       Catalog.TreeCatalog.ShowChildrenItems:=true;
       Catalog.GridCatalogFill(Catalog.TreeCatalog.SelectedItems);
       tbTreeBtnShowChild.Marked:=false;
     end
       else
     begin
        Catalog.TreeCatalog.ShowChildrenItems:=false;
        Catalog.GridCatalogFill(Catalog.TreeCatalog.SelectedItems);
        tbTreeBtnShowChild.Marked:=true;
     end;
end;

procedure TFmCatalog.tbAddInvoceClick(Sender: TObject);
var
  aIdPricePosition: LongInt;
  _BookMark: TBookMark;
  fGridMatch: TwDBGrid;
begin
  fGridMatch:= Catalog.GridMatchingPosition;
  if fGridMatch.Grid.DataSource.DataSet.RecordCount = 0 then exit;

  aIdPricePosition:= fGridMatch.Grid.DataSource.DataSet.FieldByName('PLID').AsInteger;
 // fInvoce.EventBlock:= false;
  fInvoce.InvoceAdd(aIdPricePosition);
 // fInvoce.EventBlock:= true;
  try
     _BookMark:= fGridMatch.Grid.DataSource.DataSet.Bookmark;
    fGridMatch.Grid.DataSource.DataSet.DisableControls;
    fGridMatch.Fill();
  finally
    fGridMatch.Grid.DataSource.DataSet.Bookmark:= _BookMark;
    fGridMatch.Grid.DataSource.DataSet.EnableControls;
  end;

end;

procedure TFmCatalog.TreeCatalogGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmCatalog.TreeCatalogGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmCatalog.TreeGroupOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmCatalog.TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
    if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
    exit;
    Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmCatalog.TreeViewInfoAdvancedCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; Stage: TCustomDrawStage;
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
       if (Node.Index = 7) or (Node.Index = 8) then
          Font.Color := clBlue
       else
          Font.Color := clBlack;
    end;
    TextOut(NodeRect.Left + 2, NodeRect.Top + 1, Node.Text);
 end;

end;

procedure TFmCatalog.TreeViewInfoDblClick(Sender: TObject);
var
 _FormWait: TFmWait;
begin
case TTreeView(sender).Selected.Index of
    6:
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
    7,8: if Length(TTreeView(sender).Selected.Text)>0 then OpenURL(TTreeView(sender).Selected.Text);
  end;

end;

procedure TFmCatalog.TreeViewInfoGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
    if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
    exit;
    Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmCatalog.XMLSessionRestoreProperties(Sender: TObject);
begin

end;

procedure TFmCatalog.XMLSessionSavingProperties(Sender: TObject);
begin
  DBGridClearOrderBy(GridCatalogPrice);
  DBGridClearOrderBy(GridCatalogPositionMatching);
  DBGridClearOrderBy(GridMatching);
end;

end.

