unit pkgOrdersU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, db, FmTreeU, fpspreadsheet, fpspreadsheetctrls, SysUtils, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  ComCtrls, StdCtrls, EditBtn, Buttons, DBGrids, Menus, LazUTF8,
  LCLIntf, Clipbrd, wTViewerSpreadsheetU,
  mOrdersU, wLogU, wBaseU, wDBTreeU, wFuncU,wTypesU, fpsexport
  , Grids, XMLPropStorage;

type

  { TFmOrders }

  TFmOrders = class(TForm)
    btnedInvoceSearchClear: TSpeedButton;
    btnInvoceAddNewItem: TToolButton;
    btnRemark: TBitBtn;
    btnLoad: TBitBtn;
    cbx_Format: TComboBox;
    edInvoceSearch: TComboBox;
    gbInvoce: TGroupBox;
    GridImageList: TImageList;
    GridInvoces: TDBGrid;
    GridOrderFinded: TDBGrid;
    GridOrderNoFinded: TDBGrid;
    FileNameEdit1: TFileNameEdit;
    GridInvoceNoFinded: TDBGrid;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox4: TGroupBox;
    ImageListTree: TImageList;
    Images16: TImageList;
    lbInvoceSearch: TLabel;
    lFormat: TLabel;
    mAnalogs: TMenuItem;
    MenuItem1: TMenuItem;
    mAnalisDopCreateOrderIgnoreStock: TMenuItem;
    mAnalisDopCreateOrder: TMenuItem;
    MenuItem10: TMenuItem;
    mAnalisDopExportWithVendorcode: TMenuItem;
    mClearRemarks: TMenuItem;
    mExportToPrice: TMenuItem;
    mInvoceFindedViewAnalogs: TMenuItem;
    mInvoceNotFindSelectAllClear: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    mInvoceNotFindSelectAll: TMenuItem;
    MenuItem9: TMenuItem;
    mInvocesNoFindedViewAlanogs: TMenuItem;
    mExportToOwnerFiles: TMenuItem;
    mInvoceAddNewItem: TMenuItem;
    mGridFindedEdit: TMenuItem;
    MenuItem5: TMenuItem;
    mGridFindedDelete: TMenuItem;
    mGridInvoces: TPopupMenu;
    mInvoceDel: TMenuItem;
    mInvoceEdit: TMenuItem;
    mInvoceFindPanel: TMenuItem;
    mInvoceFindPanelSplit: TMenuItem;
    mInvoceSelectAll: TMenuItem;
    mInvoceSelectAllClear: TMenuItem;
    mNoFindedEditMatching: TMenuItem;
    migBtnFindScod: TMenuItem;
    migBtnFindLabel: TMenuItem;
    mResetMatching: TMenuItem;
    mPassedMatching: TMenuItem;
    mNoFindedSelectAll: TMenuItem;
    mNoFindedClearSelection: TMenuItem;
    mNoFindedLabel: TMenuItem;
    mFindedName: TMenuItem;
    mNoFindedName: TMenuItem;
    mNoFindedScod: TMenuItem;
    mFindedVendorCode: TMenuItem;
    mFindedLabel: TMenuItem;
    mFindedScod: TMenuItem;
    mNoFindedVendorCode: TMenuItem;
    mGridNoFinded: TPopupMenu;
    miBtnFindLabel: TMenuItem;
    miBtnFindScod: TMenuItem;
    mVCodeNameLabelPrice: TMenuItem;
    mVCodeNameLabelQuantity: TMenuItem;
    PageControl1: TPageControl;
    mGridFinded: TPopupMenu;
    Panel3: TPanel;
    pcOrders: TPageControl;
    pBtm1: TPanel;
    pCenter1: TPanel;
    pInvoceSearch: TPanel;
    pInvocesBottom: TPanel;
    mGridInvoceNoFinded: TPopupMenu;
    mAnalisDop: TPopupMenu;
    pOrderBottomSum: TPanel;
    pOrderBottom: TPanel;
    pMainOrders: TPanel;
    Panel2: TPanel;
    pCenter: TPanel;
    pBtm: TPanel;
    mBtnFindMatching: TPopupMenu;
    pOrderBottomSum1: TPanel;
    pRight: TPanel;
    pTop: TPanel;
    pTree: TPanel;
    btnOpenWithOuterProgram: TSpeedButton;
    SaveDialog: TSaveDialog;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    Splitter3: TSplitter;
    stOrderSum: TStaticText;
    stInvoceSum: TStaticText;
    tbInvoce: TToolBar;
    tbInvoceBtnEdit: TToolButton;
    tbInvoceBtnDel: TToolButton;
    ToolButton5: TToolButton;
    tsNakl: TTabSheet;
    tsInvoce: TTabSheet;
    tbImport: TTabSheet;
    tbNoFindedInv: TToolBar;
    tbTree1: TToolBar;
    tbTreeOwnerBtnExpand: TToolButton;
    tbTreeOwnerBtnSort: TToolButton;
    tbFinded: TToolBar;
    tbNoFinded: TToolBar;
    btnEditMatching: TToolButton;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    btnPassed: TToolButton;
    ToolButton3: TToolButton;
    btnExportNewPosition: TToolButton;
    btnExportOrderToExcel: TToolButton;
    btnResetMatching: TToolButton;
    btnFindMatching: TToolButton;
    ToolButton4: TToolButton;
    btnMatchingOnly: TToolButton;
    TreeGroupOwner: TTreeView;
    XMLPropStorage1: TXMLPropStorage;
    procedure btnedInvoceSearchClearClick(Sender: TObject);
    procedure btnEditCatalogPosition(Sender: TObject);
    procedure btnExportNewPositionClick(Sender: TObject);
    procedure btnExportOrderToExcelClick(Sender: TObject);
    procedure btnFindMatchingClick(Sender: TObject);
    procedure btnInvoceAddNewItemClick(Sender: TObject);
    procedure btnLoadClick(Sender: TObject);
    procedure btnMatchingOnlyClick(Sender: TObject);
    procedure btnOpenWithOuterProgramClick(Sender: TObject);
    procedure btnPassedClick(Sender: TObject);
    procedure btnRemarkClick(Sender: TObject);
    procedure btnResetMatchingClick(Sender: TObject);
    procedure FileNameEdit1Change(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure GridInvoceNoFindedDblClick(Sender: TObject);
    procedure GridOrderFindedDblClick(Sender: TObject);
    procedure GridOrderFindedDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure mAnalisDopCreateOrderClick(Sender: TObject);
    procedure mAnalisDopCreateOrderIgnoreStockClick(Sender: TObject);
    procedure mAnalisDopExportWithVendorcodeClick(Sender: TObject);
    procedure mClearRemarksClick(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure MenuItem7Click(Sender: TObject);
    procedure mExportToPriceClick(Sender: TObject);
    procedure mExportToOwnerFilesClick(Sender: TObject);
    procedure mInvoceAddNewItemClick(Sender: TObject);
    procedure mInvoceFindedViewAnalogsClick(Sender: TObject);
    procedure mInvoceNotFindSelectAllClearClick(Sender: TObject);
    procedure mInvoceNotFindSelectAllClick(Sender: TObject);
    procedure mSummaryInvoce(Sender: TObject);
    procedure mFindedVendorCodeClick(Sender: TObject);
    procedure mGridFindedDeleteClick(Sender: TObject);
    procedure mGridInvocesPopup(Sender: TObject);
    procedure miBtnFindLabelClick(Sender: TObject);
    procedure miBtnFindScodClick(Sender: TObject);
    procedure mInvoceFindPanelClick(Sender: TObject);
    procedure mInvoceSelectAllClearClick(Sender: TObject);
    procedure mInvoceSelectAllClick(Sender: TObject);
    procedure mNoFindedClearSelectionClick(Sender: TObject);
    procedure mNoFindedSelectAllClick(Sender: TObject);
    procedure mVCodeNameLabelPriceClick(Sender: TObject);
    procedure mVCodeNameLabelQuantityClick(Sender: TObject);
    procedure mInvocesNoFindedViewAlanogsClick(Sender: TObject);
    procedure tbInvoceBtnDelClick(Sender: TObject);
    procedure tbInvoceBtnEditClick(Sender: TObject);
    procedure tbTreeOwnerBtnExpandClick(Sender: TObject);
    procedure tbTreeOwnerBtnSortClick(Sender: TObject);
    procedure TreeGroupOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure XMLPropStorage1SavingProperties(Sender: TObject);
  private
    fBase: TwBase;
    fOrders: TOrders;
    wFormID: string;
    IdMainOwner: integer;
  public
    procedure SetStatus(_Text:string);
  end;

var
  FmOrders: TFmOrders;

implementation

{$R *.lfm}

{ TFmOrders }

procedure TFmOrders.FormCreate(Sender: TObject);
begin
  try
    wFormID:=Self.Name;
    pcOrders.ShowTabs:= false;
    fBase:= TwBase.Create(self);

    fOrders:= TOrders.Create(Sender,fBase,TreeGroupOwner,GridOrderFinded,GridOrderNoFinded, GridInvoces, GridInvoceNoFinded, cbx_Format);
    fOrders.GridOrderFindedFill();
    fOrders.GridOrderNoFindedFill();
    fOrders.TreeOwnersFill();
    cbx_Format.Clear;
    cbx_Format.Enabled:= false;
    FileNameEdit1.FileName:='';
    //fBase.Memo:= mLog;
    TryStrToInt(fBase.ReadSettingByName('setDefaultOwner'),IdMainOwner); // считываем настройки - текущий основной прайс-лист
  except
    on E: Exception do
    begin
        SetStatus('Сбой инициализации плагина.');
        wLog('Orders','Ошибка [FmCreate]: "' + E.Message + '"');
        wLog('Orders','Сбой инициализации плагина.');
       // ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

     end;
  end;
end;

procedure TFmOrders.btnRemarkClick(Sender: TObject);
begin
  Showmessage(fOrders.FormatOrder.Remark);
end;

procedure TFmOrders.btnResetMatchingClick(Sender: TObject);
begin
  if MessageDlg('Удалить соответствия для выделенных позиций?',mtWarning, mbOKCancel, 0) = mrCancel then exit;

   fOrders.ResetMatching(fOrders.GridOrderNoFinded);
end;

procedure TFmOrders.FileNameEdit1Change(Sender: TObject);
begin
  TFileNameEdit(Sender).InitialDir:= ExtractFileDir(TFileNameEdit(Sender).FileName);
end;

procedure TFmOrders.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
     FmOrders.WindowState:= wsNormal;
end;

procedure TFmOrders.btnLoadClick(Sender: TObject);
begin

  try
    if (Length(FileNameEdit1.FileName)=0) or not FileExists(UTF8ToSys(FileNameEdit1.FileName)) then
        begin
          ShowMessage('Файл не найден!');
          exit;
        end;

    if TBitBtn(Sender).Tag = 0 then
      mAnalisDop.PopUp
    else
    begin
      fBase.DataBase.Connected:= false;
      fBase.DataBase.Connected:= true;

      fOrders.LoadFromFile(FileNameEdit1.FileName, TBitBtn(Sender).Tag);
    end;

  except
    on E: Exception do
    begin
        //SetStatus('Файл не найден!');
        wLog('Orders','Ошибка [Load]: "' + E.Message + '"');
        ShowMessage('Ошибка [Load]: "' + E.Message + '"');
     end;
  end;
end;

procedure TFmOrders.btnMatchingOnlyClick(Sender: TObject);
begin
    if btnMatchingOnly.Marked then
     begin
       fOrders.GridOrderNoFinded.Where:='CTG.ID IS NOT NULL';
       fOrders.GridOrderNoFindedFill();
       btnMatchingOnly.Marked:=false;
     end
       else
     begin
        fOrders.GridOrderNoFinded.Where:='';
        fOrders.GridOrderNoFindedFill();
        btnMatchingOnly.Marked:=true;
     end;
end;

procedure TFmOrders.btnOpenWithOuterProgramClick(Sender: TObject);
begin
  if MessageDlg('Открыть накладную во внешней программе просмотра?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;
  if Length(FileNameEdit1.FileName)=0 then exit;

  OpenDocument(FileNameEdit1.FileName);
end;

procedure TFmOrders.btnPassedClick(Sender: TObject);
begin
  if MessageDlg('Применить соответствия для всех отображаемых позиций?',mtWarning, mbOKCancel, 0) = mrCancel then exit;

  fOrders.PassedMatching();

  btnPassed.Enabled:= false;
  btnResetMatching.Enabled:= false;

end;

procedure TFmOrders.btnEditCatalogPosition(Sender: TObject);
begin
  fOrders.EditMatching(fOrders.GridOrderNoFinded);
end;

procedure TFmOrders.btnedInvoceSearchClearClick(Sender: TObject);
begin
  edInvoceSearch.Clear;
  edInvoceSearch.OnChange(edInvoceSearch);
end;

procedure TFmOrders.btnExportNewPositionClick(Sender: TObject);
begin
  fOrders.OrderExportNewPosition('new_items_'+ExtractFileName(FileNameEdit1.FileName));
end;

procedure TFmOrders.btnExportOrderToExcelClick(Sender: TObject);
begin
  fOrders.OrderExportResult('convert_'+ExtractFileName(FileNameEdit1.FileName));
end;

procedure TFmOrders.btnFindMatchingClick(Sender: TObject);
begin
  mBtnFindMatching.PopUp;
end;

procedure TFmOrders.btnInvoceAddNewItemClick(Sender: TObject);
begin
  if fOrders.GridInvoceNoFinded.Grid.DataSource.DataSet.RecordCount = 0 then exit;

  fOrders.Invoce.InvoceAddNewItem(fOrders.GridInvoceNoFinded, fOrders.TreeOwners.SelectedItems, true);
end;

procedure TFmOrders.FormDestroy(Sender: TObject);
begin
  if fBase.LongTransaction then fBase.SQLTransactionEnd(false);
  fOrders.Destroy();
  fBase.Destroy();
end;

procedure TFmOrders.GridInvoceNoFindedDblClick(Sender: TObject);
begin
  btnInvoceAddNewItemClick(self);
end;

procedure TFmOrders.GridOrderFindedDblClick(Sender: TObject);
begin
  if fOrders.GridOrderFinded.FieldName = 'CTGNAME' then
       fOrders.EditMatching(fOrders.GridOrderFinded);
end;

procedure TFmOrders.GridOrderFindedDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if (gdFocused in State) then // если строка не выделена, то
    begin
      TDBGrid(Sender).Canvas.Brush.Color:= TDBGrid(Sender).FixedHotColor;
      TDBGrid(Sender).Canvas.Font.Color:= clBlack;

      TDBGrid(Sender).Canvas.FillRect(Rect);
      TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
    end;

  if Column.FieldName = 'ORDQUANTITY' then
   begin
     with TDBGrid(Sender).Canvas do
        if TDBGrid(Sender).DataSource.DataSet.FieldByName('ORDQUANTITYCALCULATED').AsInteger>0 then
           begin
             FillRect(Rect);
             font.Color:=clBlue;
             font.Style:= [fsBold];
             TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
           end;
   end;

  if Column.FieldName = 'ORDPRICE' then
   begin
     with TDBGrid(Sender).Canvas do
        if TDBGrid(Sender).DataSource.DataSet.FieldByName('ORDQUANTITYCALCULATED').AsInteger>0 then
           begin
             FillRect(Rect);
             font.Color:=clBlue;
             font.Style:= [fsBold];
             TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
           end;
   end;
end;

procedure TFmOrders.mAnalisDopCreateOrderClick(Sender: TObject);
begin
  try

    fBase.DataBase.Connected:= false;
    fBase.DataBase.Connected:= true;
    fOrders.LoadFromFile(FileNameEdit1.FileName, 0);

  except
    on E: Exception do
    begin
        //SetStatus('Файл не найден!');
        wLog('Orders','Ошибка [Load]: "' + E.Message + '"');
        ShowMessage('Ошибка [Load]: "' + E.Message + '"');
     end;
  end;
end;

procedure TFmOrders.mAnalisDopCreateOrderIgnoreStockClick(Sender: TObject);
begin
  try

    fBase.DataBase.Connected:= false;
    fBase.DataBase.Connected:= true;

    fOrders.LoadFromFile(FileNameEdit1.FileName, 2);

  except
    on E: Exception do
    begin
        //SetStatus('Файл не найден!');
        wLog('Orders','Ошибка [Load]: "' + E.Message + '"');
        ShowMessage('Ошибка [Load]: "' + E.Message + '"');
     end;
  end;
end;

procedure TFmOrders.mAnalisDopExportWithVendorcodeClick(Sender: TObject);
var
  aFmTree: TFmTree;
  aSelected: ArrayOfInteger;
begin

  fBase.DataBase.Connected:= false;
  fBase.DataBase.Connected:= true;

  aFmTree:= TFmTree.Create(self);
  aFmTree.Mode:= 2;
  aFmTree.PanelBtn.Visible:= true;
  aFmTree.Base:= fBase;

  try
    aFmTree.ShowModal;
    if aFmTree.ModalResult = mrOK then
     begin
       aFmTree.Tree.ShowChildrenItems:= true;
       aSelected:= aFmTree.Tree.SelectedItems;
       fOrders.LoadFromFile(FileNameEdit1.FileName, 3, aSelected);
     end;

  finally
    FreeAndNil(aFmTree);
  end;
end;

procedure TFmOrders.mClearRemarksClick(Sender: TObject);
begin
  fOrders.InvoceClearRemark(fOrders.GridInvoces.SelectedRows);
end;

procedure TFmOrders.MenuItem4Click(Sender: TObject);
begin
  fOrders.GridInvoces.ExportData();
end;

procedure TFmOrders.MenuItem7Click(Sender: TObject);
begin
  fOrders.GridInvoceNoFinded.ExportData();
end;

procedure TFmOrders.mExportToPriceClick(Sender: TObject);
var
  aSelected: ArrayOfInteger;
  aDataSet: TDataSet;
begin
  aDataSet:= fOrders.GridInvoces.Grid.DataSource.DataSet;

  if MessageDlg('Отметить позиции в прайс-листе контрагента "'+aDataSet.FieldByName('OWNERNAME').AsString+'" ?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

   aSelected:= [aDataSet.FieldByName('IDOWNER').AsInteger];

   SetStatus('Выгрузка ...');
   fOrders.ExportInvoceInOwnerPrice(aSelected);
   fOrders.GridInvocesFill();
end;

procedure TFmOrders.mExportToOwnerFilesClick(Sender: TObject);
begin
  fOrders.ExportInvoceToOwnerFiles;
end;

procedure TFmOrders.mInvoceAddNewItemClick(Sender: TObject);
begin
  fOrders.Invoce.InvoceAddNewItem(fOrders.GridInvoces, true);
end;

procedure TFmOrders.mInvoceFindedViewAnalogsClick(Sender: TObject);
begin
  fOrders.Invoce.GetPositionAnalog(fOrders.GridInvoces.SelectedRows, true);
end;

procedure TFmOrders.mInvoceNotFindSelectAllClearClick(Sender: TObject);
begin
  fOrders.GridInvoceNoFinded.SelectAll:= false;
end;

procedure TFmOrders.mInvoceNotFindSelectAllClick(Sender: TObject);
begin
  fOrders.GridInvoceNoFinded.SelectAll:= true;
end;

procedure TFmOrders.mSummaryInvoce(Sender: TObject);
begin
  fOrders.SummaryInvoce;
end;

procedure TFmOrders.mFindedVendorCodeClick(Sender: TObject);
begin
  fOrders.CopyData(Sender);
end;

procedure TFmOrders.mGridFindedDeleteClick(Sender: TObject);
begin
  if MessageDlg('Удалить соответствия для выделенных позиций?',mtWarning, mbOKCancel, 0) = mrCancel then exit;

   fOrders.ResetMatching(fOrders.GridOrderFinded);
end;

procedure TFmOrders.mGridInvocesPopup(Sender: TObject);
begin
  mInvoceFindPanel.Checked:= pInvoceSearch.Visible;

  if fOrders.GridInvoces.Grid.DataSource.DataSet.RecordCount = 0 then exit;
  fOrders.mAnalogsFill(Sender, mAnalogs);
end;

procedure TFmOrders.miBtnFindLabelClick(Sender: TObject);
begin
  if MessageDlg('Найти соответствия по артикулу для выделенных позиций?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;
  fOrders.FindMatching(2);
end;

procedure TFmOrders.miBtnFindScodClick(Sender: TObject);
begin
  if MessageDlg('Найти соответствия по штрих-коду для выделенных позиций?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;
  fOrders.FindMatching(1);
end;

procedure TFmOrders.mInvoceFindPanelClick(Sender: TObject);
begin
  pInvoceSearch.Visible:= not pInvoceSearch.Visible;
  if not pInvoceSearch.Visible then
     btnedInvoceSearchClearClick(self)
     else
     edInvoceSearch.SetFocus;
end;

procedure TFmOrders.mInvoceSelectAllClearClick(Sender: TObject);
begin
  fOrders.GridInvoces.SelectAll:= false;
end;

procedure TFmOrders.mInvoceSelectAllClick(Sender: TObject);
begin
  fOrders.GridInvoces.SelectAll:= true;
end;

procedure TFmOrders.mNoFindedClearSelectionClick(Sender: TObject);
begin
  fOrders.GridOrderNoFinded.SelectAll:= false;
end;

procedure TFmOrders.mNoFindedSelectAllClick(Sender: TObject);
begin
  fOrders.GridOrderNoFinded.SelectAll:= true;
end;

procedure TFmOrders.mVCodeNameLabelPriceClick(Sender: TObject);
begin
  fOrders.GridInvoces.CopyToClipboard(['VENDORCODE', 'NAME', 'UNIT', 'LABEL', 'PRICEPL'], ['PRICEPL'],'INV.ID');
end;

procedure TFmOrders.mVCodeNameLabelQuantityClick(Sender: TObject);
begin
  fOrders.GridInvoces.CopyToClipboard(['VENDORCODE', 'NAME', 'UNIT', 'LABEL', 'QUANTITY'], [''],'INV.ID');
end;

procedure TFmOrders.mInvocesNoFindedViewAlanogsClick(Sender: TObject);
begin
  fOrders.Invoce.GetPositionAnalog(fOrders.GridInvoceNoFinded.SelectedRows, false);
end;

procedure TFmOrders.tbInvoceBtnDelClick(Sender: TObject);
begin
  fOrders.InvoceDel(fOrders.GridInvoces.SelectedRows);
end;

procedure TFmOrders.tbInvoceBtnEditClick(Sender: TObject);
begin
  if fOrders.GridInvoces.Grid.DataSource.DataSet.RecordCount = 0 then exit;

  fOrders.InvoceEdit(fOrders.GridInvoces.SelectedRows[0]);
end;

procedure TFmOrders.tbTreeOwnerBtnExpandClick(Sender: TObject);
begin
  if TreeGroupOwner.Items.Count=0 then exit;
  if tbTreeOwnerBtnExpand.Marked then
     begin
       fOrders.TreeOwners.Expanded:=false;
       fOrders.TreeOwners.Tree.FullCollapse;
       fOrders.TreeOwners.Tree.Items[0].Expanded:= true;
       tbTreeOwnerBtnExpand.Marked:=false;
     end
       else
     begin
        fOrders.TreeOwners.Expanded:=true;
        fOrders.TreeOwners.Tree.FullExpand;
        tbTreeOwnerBtnExpand.Marked:=true;
     end;
end;

procedure TFmOrders.tbTreeOwnerBtnSortClick(Sender: TObject);
begin
  if tbTreeOwnerBtnSort.Marked then
     begin
       fOrders.TreeOwners.OrderBy:='IDPARENT, ID';
       fOrders.TreeOwners.Fill();
       tbTreeOwnerBtnSort.Marked:=false;
     end
       else
     begin
        fOrders.TreeOwners.OrderBy:='IDPARENT, NAME';
        fOrders.TreeOwners.Fill();
        tbTreeOwnerBtnSort.Marked:=true;
     end;
end;

procedure TFmOrders.TreeGroupOwnerGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
  if TTreeData(Node.Data).Value = IdMainOwner then
  begin
    Node.ImageIndex:=2;
    exit;
  end;

  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmOrders.TreeGroupOwnerGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmOrders.XMLPropStorage1SavingProperties(Sender: TObject);
begin
  DBGridClearOrderBy(GridOrderFinded);
  DBGridClearOrderBy(GridOrderNoFinded);
end;

procedure TFmOrders.SetStatus(_Text: string);
begin
  wStatus(self.Name,_Text,true);
end;

end.

