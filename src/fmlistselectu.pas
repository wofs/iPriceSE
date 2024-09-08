unit FmListSelectU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Spin, Forms, Controls, Graphics, Dialogs, db,LazUTF8,
  StdCtrls, ExtCtrls, ComCtrls, DBGrids, Buttons, XMLPropStorage, Menus, wLogU, wFuncU,
  //wDBaseU, wDBTreeU, wDBGridU,
  wBaseU, wDBGridU, wDBTreeU, wTypesU
  ;

type

  { TFmListSelect }

  TFmListSelect = class(TForm)
    btnCancel: TBitBtn;
    btnOK: TBitBtn;
    btnListPreventSearch: TSpeedButton;
    btnListEditSearchClear: TSpeedButton;
    btnListSearchSplitString: TSpeedButton;
    eQSName: TEdit;
    edSearch: TComboBox;
    GridImageList: TImageList;
    GridList: TDBGrid;
    gbLeft: TGroupBox;
    menuImageList: TImageList;
    ImageListTreeOwner: TImageList;
    ImageListTreeGroup: TImageList;
    lbK: TLabel;
    lbQuantity: TLabel;
    lbPriceSearch: TLabel;
    mNomAddFromPriceItem: TMenuItem;
    MenuItem3: TMenuItem;
    mNomAdd: TMenuItem;
    mNomCopy: TMenuItem;
    mNomDelete: TMenuItem;
    mNomEdit: TMenuItem;
    mNomGoToGroup: TMenuItem;
    mPriceGrid: TPopupMenu;
    mQSSCode: TPopupMenu;
    pSearch: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    pBtnQuickSearch: TPanel;
    pRight: TPanel;
    sBtnQSVendorCode: TSpeedButton;
    sBtnQSSCode: TSpeedButton;
    sBtnQSLabel: TSpeedButton;
    Splitter1: TSplitter;
    spQuantInPackLeft: TSpinEdit;
    spQuantInPackRight: TSpinEdit;
    TreeGroup: TTreeView;
    XMLSession: TXMLPropStorage;
    procedure btnListEditSearchClearClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GridListDblClick(Sender: TObject);
    procedure mNomAddClick(Sender: TObject);
    procedure mNomAddFromPriceItemClick(Sender: TObject);
    procedure mNomCopyClick(Sender: TObject);
    procedure mNomDeleteClick(Sender: TObject);
    procedure mNomEditClick(Sender: TObject);
    procedure mNomGoToGroupClick(Sender: TObject);
    procedure sBtnQSSCodeClick(Sender: TObject);
    procedure sBtnQSVendorCodeClick(Sender: TObject);
    procedure TreeGroupGetImageIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupGetSelectedIndex(Sender: TObject; Node: TTreeNode);
    procedure TreeGroupSelectionChanged(Sender: TObject);
    procedure XMLSessionSavingProperties(Sender: TObject);
  private
    fCatalog: TObject;
    FormMode: integer;
    SelectedRows: ArrayOfInteger;
    wFormLoaded: boolean;
    fFormName: string;
    procedure EventRegister();
    procedure OnEventAlert(Sender: TObject; EventName: string; EventCount: longint; var CancelAlerts: Boolean);
  private
    DataSetLocateField: string;
    DataSetLocateValue: variant;
    fMultiSelectGrid: boolean;
    IDTreeItem: integer;
    QuantityInPacked: Double;
    fWhere: string;
    fBase: TwBase;
    fTreeGroup: TwDBTree;
    fGridList: TwDBGrid;
    fIdMainOwner: integer;
    procedure SetFormMode(aValue: integer);
  public
    procedure SetStatus(_Text:string);
    procedure _on_mScodClick(Sender: TObject);
    procedure ListFormInit(aIDPL_or_Catalog, aVendorcode, aName, aLabel: string; const aScod: string = '-1');

    property wFormMode: integer read FormMode write SetFormMode;
    property Where: string read fWhere write fWhere;
    property wSelectedRows: ArrayOfInteger read SelectedRows write SelectedRows;
    property wQuantityInPacked: Double read QuantityInPacked write QuantityInPacked;
    property wDataSetLocateField: string read DataSetLocateField write DataSetLocateField;
    property wDataSetLocateValue: variant read DataSetLocateValue write DataSetLocateValue;
    property wIDTreeItem: integer read IDTreeItem write IDTreeItem;
    property MultiSelectGrid: boolean read fMultiSelectGrid write fMultiSelectGrid;
    property Base: TwBase read fBase write fBase;

  end;

var
  FmListSelect: TFmListSelect;

implementation
uses
  mCatalogU;

{$R *.lfm}

{ TFmListSelect }

procedure TFmListSelect._on_mScodClick(Sender: TObject);
var
  _BtnDown: Boolean;
begin
   fGridList.SearchComplete:= false;
   edSearch.Text:= TMenuItem(Sender).Caption;
   fGridList.SearchText:= edSearch.Text;
   fGridList.Fill();
   fGridList.SearchComplete:=true;
end;

procedure TFmListSelect.FormCreate(Sender: TObject);
begin
  SelectedRows:= nil;
  wIDTreeItem:= 0;
  wFormLoaded:= false;
  fFormName:= Self.Name;
  fMultiSelectGrid:= false;
  wLog('FmListSelect','Инициализация формы... ['+fFormName+']');
  screen.Cursor:= crSQLWait;

  try

    Base:=nil;
    fCatalog:= nil;

    // заполнение формы в SHOW

     wLog('FmListSelect','Инициализация формы успешно завершена.');

   except
     on E: Exception do
     begin
         Screen.Cursor := crDefault;
         __Log.SaveLogError(E);
         SetStatus('Сбой инициализации формы.');
         wLog('FmListSelect','Ошибка [FmCreate]: "' + E.Message + '"');
         wLog('FmListSelect','Сбой инициализации формы.');
         ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

      end;
   end;
end;

procedure TFmListSelect.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if ModalResult = mrCancel then
    if MessageDlg('Закрыть без сохранения?',mtConfirmation, mbOKCancel, 0) = mrCancel
     then
       CanClose:= false
     else
       ModalResult:= mrCancel;
end;

procedure TFmListSelect.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  wSelectedRows:= fGridList.SelectedRows;
  wQuantityInPacked:= spQuantInPackLeft.Value/spQuantInPackRight.Value;
end;

procedure TFmListSelect.btnListEditSearchClearClick(Sender: TObject);
begin
  edSearch.Text:='';
  edSearch.OnChange(edSearch);
end;

procedure TFmListSelect.FormDestroy(Sender: TObject);
begin
   try
    wLog('FmListSelect','Выгрузка формы...');

     // выгружаем подгруженные DBTree
     fTreeGroup.Destroy();

         // выгружаем подгруженные DBGrid
     fGridList.Destroy();

     if Assigned(fCatalog) then fCatalog.Destroy();

   wLog('FmListSelect','Выгрузка формы успешно завершена.');

   except
     on E: Exception do
     begin
         __Log.SaveLogError(E);
         SetStatus('Сбой выгрузки формы.');
         wLog('FmListSelect','Ошибка [FmDestroy]: "' + E.Message + '"');
         wLog('FmListSelect','Сбой выгрузки формы.');
         ShowMessage('Ошибка [FmDestroy]: "' + E.Message + '"');
      end;
   end;
end;

procedure TFmListSelect.EventRegister();
begin
  if fBase.EventDB.Events.IndexOf('CATALOG_GROUP_Change')>-1 then exit;

  fBase.EventDB.Events.Add('CATALOG_GROUP_Change');
  fBase.EventDB.RegisterEvents;
  fBase.EventDB.OnEventAlert:= @OnEventAlert;
end;

procedure TFmListSelect.OnEventAlert(Sender: TObject; EventName: string; EventCount: longint; var CancelAlerts: Boolean);
var
  _IdTmp: Integer;
begin
   case EventName of
       'CATALOG_GROUP_Change':
             begin
               with fTreeGroup do begin
                 if EventBlock then exit;
                 _IdTmp:= SelectedItems[0];
                 Fill();
                 FindNodeWithDataInt(_IdTmp);
               end;
             end;
     end;
end;

procedure TFmListSelect.SetFormMode(aValue: integer);
var
  _SQL_text: String;
begin
  //if FormMode=aValue then Exit;
  FormMode:=aValue;

  fIdMainOwner:= fBase.ReadSettingByName('setDefaultOwner');

  case wFormMode of
    0:   // PriceLists
      begin
          //DBTree
           //DBTree.Add(TwDBTreeView.Create(wFormID, TreeGroup, GridList, true,'OWNER','IDPARENT,NAME',[])); // инициализация дерева
           fTreeGroup:= TwDBTree.Create(fBase,TreeGroup,'OWNER','IDPARENT,NAME',[]);

           fTreeGroup.Tree.Images:= ImageListTreeOwner;

           _SQL_text:='select '
              +' PL.ID, '
              +' PL.IDOWNER, '
              +' PL.NAME, '
              +' PL.IDPL_GROUP, '
              +' PL.UNIT, '
              +' PL.LABEL, '
              +' PL.VENDORCODE, '
              +' (PL.STOCK+PL.STOCK2+PL.STOCK3+PL.STOCK4+PL.STOCK5) STOCK, '
              +' PL.PRICECALC as PRICE, '
              +' OWN.NAME as OWNERNAME '
              +' from "CATALOG_MATCHING" MTH '
              +' right join "PL_ITEMS" PL ON (MTH.idpl_items = PL.id) '
              +' LEFT JOIN OWNER OWN ON (PL.IDOWNER=OWN.ID) '
              +' WHERE '
              +' (MTH.id IS NULL) AND'
              +' (PL.IDOWNER<>'+IntToStr(fIdMainOwner)+')  ';
              _SQL_text:= _SQL_text +' /*and_group_string*/ /*and_search_string*/ ';

           fGridList:= TwDBGrid.Create(fBase,GridList,_SQL_text);
           fGridList.MultiSelect:= fMultiSelectGrid;
           fGridList.SearchEdit:= edSearch;
           fGridList.SearchPreventiveBtn:= btnListPreventSearch;
           fGridList.SearchSplitStringBtn:= btnListSearchSplitString;
           fGridList.SearchEntryArray:= ['PL.NAME','PL.LABEL'];
           fGridList.SearchParticleArray:= ['PL.VENDORCODE','(SELECT VRESULT FROM PL_TRY_SCOD(PL.ID,''%s'') /*=*/)'];
           fGridList.GroupField:= 'PL.IDOWNER';
           fGridList.SortTitleImagesIndex:=[2,3];

           GridList.Columns[1].FieldName:= 'OWNERNAME';
           GridList.Columns[2].FieldName:= 'NAME';
           GridList.Columns[3].FieldName:= 'UNIT';
           GridList.Columns[4].FieldName:= 'STOCK';
           GridList.Columns[5].FieldName:= 'VENDORCODE';
           GridList.Columns[6].FieldName:= 'LABEL';
           GridList.Columns[7].FieldName:= 'PRICE';

           GridList.Columns[1].Width:=70;
           GridList.Columns[2].Width:=250;
           GridList.Columns[3].Width:=40;
           GridList.Columns[4].Width:=50;
           GridList.Columns[5].Width:=70;
           GridList.Columns[6].Width:=70;
           GridList.Columns[7].Width:=70;
      end;
    1:   // CATALOG
      begin
          //DBTree
           fCatalog:= TCatalog.Create(self,fBase,true);
           fTreeGroup:= TwDBTree.Create(fBase,TreeGroup,'CATALOG_GROUP','IDPARENT,NAME',['IDOWNER',fIdMainOwner]);
           fTreeGroup.Tree.Images:= ImageListTreeGroup;

           EventRegister();

           with fTreeGroup.Tree.PopupMenu do
           begin
             Images:= fTreeGroup.Tree.Images;
             Items[0].ImageIndex:= 2;
             Items[1].ImageIndex:= 3;
             Items[2].ImageIndex:= 4;
           end;

           _SQL_text:=' SELECT CTG.ID, '
               +' CTG.IDCTG_GROUP, '
               +' CTG.NAME AS NAME, '
               +' CTG.UNIT, '
               +' (select VSCOD from CTG_GET_SCOD(CTG.ID,true)) SCOD, '
               +' CTG.LABEL, '
               +' CTG.PRICE AS PRICE, '
               +' CTG.VENDORCODE, '
               +' CTG.FCOLOR, '
               +' CTG.FTIMESTAMP, '
               +' (PLOUR.STOCK+PLOUR.STOCK2+PLOUR.STOCK3+PLOUR.STOCK4+PLOUR.STOCK5) AS STOCK, '
               +' PLOUR.STOCK AS STOCK1, '
               +' PLOUR.STOCK2 AS STOCK2, '
               +' PLOUR.STOCK3 AS STOCK3, '
               +' PLOUR.STOCK4 AS STOCK4, '
               +' PLOUR.STOCK5 AS STOCK5, '
               +' CTG.PN,CTG.PM,CTG.PD,CTG.PC,CTG.PK, '
               +' CASE WHEN (SELECT * FROM CATALOG_SELECT_MTHRESULT(CTG.ID))>0 THEN 1 ELSE 0 END MTHRESULT, '
               +'  PLFP.PRICEPL AS PRICEPL, '
               +'  PLFP.PRICEPL2 AS PRICEPL2, '
               +'  PLFP.PRICEPL3 AS PRICEPL3, '
               +'  PLFP.PRICEPL4 AS PRICEPL4, '
               +'  PLFP.PRICEPL5 AS PRICEPL5, '
               +' PLOUR.PRICECALC AS PRICEOUR, '
               +' PLOUR.PRICECALC2 AS PRICEOUR2, '
               +' PLOUR.PRICECALC3 AS PRICEOUR3, '
               +' PLOUR.PRICECALC4 AS PRICEOUR4, '
               +' PLOUR.PRICECALC5 AS PRICEOUR5 '
               +'  FROM "CATALOG" CTG '
               +' LEFT JOIN CATALOG_PL_MIN_PRICE(CTG.ID) PLFP ON (1=1)'
               +'  LEFT OUTER JOIN "PL_ITEMS" PLOUR ON ( '
               +'  CTG.VENDORCODE = PLOUR.VENDORCODE AND CTG.IDOWNER = PLOUR.IDOWNER) '
               +' where (CTG.IDOWNER='+IntToStr(fIdMainOwner)+') /*and_group_string*/ /*and_search_string*/  ';


           fGridList:= TwDBGrid.Create(fBase,GridList,_SQL_text);
           fGridList.MultiSelect:= fMultiSelectGrid;
           fGridList.SearchEdit:= edSearch;
           fGridList.SearchPreventiveBtn:= btnListPreventSearch;
           fGridList.SearchEntryArray:= ['CTG.NAME','CTG.LABEL'];
           fGridList.SearchParticleArray:= ['CTG.VENDORCODE','(SELECT VRESULT FROM CTG_TRY_SCOD(CTG.ID,''%s'') /*=*/)'];
           fGridList.GroupField:= 'CTG.IDCTG_GROUP';
           fGridList.Grid.PopupMenu:= mPriceGrid;

           GridList.Columns[2].FieldName:= 'NAME';
           GridList.Columns[3].FieldName:= 'UNIT';
           GridList.Columns[4].FieldName:= 'STOCK';
           GridList.Columns[5].FieldName:= 'VENDORCODE';
           GridList.Columns[6].FieldName:= 'LABEL';
           GridList.Columns[7].FieldName:= 'PRICEOUR';

           GridList.Columns[1].Width:=0;
           GridList.Columns[2].Width:=250;
           GridList.Columns[3].Width:=40;
           GridList.Columns[4].Width:=50;
           GridList.Columns[5].Width:=70;
           GridList.Columns[6].Width:=70;
           GridList.Columns[7].Width:=70;
      end;
  end;

end;

procedure TFmListSelect.FormShow(Sender: TObject);
var
  _SQL_text: String;
begin

  screen.Cursor:= crSQLWait;
  try
// заполнение формы

  //Base


// заполнение формы
//
    fTreeGroup.Where:=fWhere;
    fTreeGroup.Fill();
    try
      if wIDTreeItem<>0 then
        fTreeGroup.FindNodeWithDataInt(wIDTreeItem)
      else
        fGridList.Fill;

      if  not assigned (GridList.DataSource) then fGridList.Fill;

    finally
      if Assigned(GridList.DataSource) and (DataSetLocateValue<>null) and (VarToStr(DataSetLocateValue)<>'0') then
        GridList.DataSource.DataSet.Locate(DataSetLocateField,DataSetLocateValue,[]);
      //ShowMessage(DataSetLocateField+'|'+string(DataSetLocateValue)) ;

      wFormLoaded:= true;
      screen.Cursor:= crDefault;
    end;
  except
     on E: Exception do
     begin
         __Log.SaveLogError(E);
         SetStatus('Сбой отображения формы.');
         wLog('FmListSelect','Ошибка [FormShow]: "' + E.Message + '"');
         wLog('FmListSelect','Сбой выгрузки формы.');
         ShowMessage('Ошибка [FormShow]: "' + E.Message + '"');
     end;
  end;
end;

procedure TFmListSelect.GridListDblClick(Sender: TObject);
begin
  btnOK.Click;
end;

procedure TFmListSelect.mNomAddClick(Sender: TObject);
begin
  TCatalog(fCatalog).ItemAdd(fTreeGroup,fGridList.Grid.DataSource.DataSet);
end;

procedure TFmListSelect.mNomAddFromPriceItemClick(Sender: TObject);
var
  aScod: String;
begin
  aScod:= EmptyStr;
  if mQSSCode.Items.Count>0 then aScod:= mQSSCode.Items[0].Caption;
  TCatalog(fCatalog).ItemAdd(fTreeGroup,fGridList.Grid.DataSource.DataSet, eQSName.Text, sBtnQSLabel.Caption, aScod);
end;


procedure TFmListSelect.mNomCopyClick(Sender: TObject);
begin
  TCatalog(fCatalog).ItemCopy(fTreeGroup,fGridList);
end;

procedure TFmListSelect.mNomDeleteClick(Sender: TObject);
begin
  TCatalog(fCatalog).ItemDel(fGridList);
end;

procedure TFmListSelect.mNomEditClick(Sender: TObject);
begin
  TCatalog(fCatalog).ItemEdit(fGridList,fTreeGroup);
end;

procedure TFmListSelect.mNomGoToGroupClick(Sender: TObject);
begin
  fTreeGroup.FindNodeWithDataInt(fGridList.Grid.DataSource.DataSet.FieldByName('IDCTG_GROUP').AsInteger);
end;

procedure TFmListSelect.sBtnQSSCodeClick(Sender: TObject);
begin
  if mQSSCode.Items.Count = 1 then
    begin
      fGridList.SearchComplete:= false;
      edSearch.Text:= mQSSCode.Items[0].Caption;
      fGridList.SearchText:= edSearch.Text;
      fGridList.Fill();
      fGridList.SearchComplete:=true;
    end else
     mQSSCode.PopUp;
end;

procedure TFmListSelect.sBtnQSVendorCodeClick(Sender: TObject);
begin
  fGridList.SearchComplete:= false;
  edSearch.Text:= TSpeedButton(Sender).Caption;
  fGridList.SearchText:= edSearch.Text;
  fGridList.Fill();
  fGridList.SearchComplete:=true;
end;

procedure TFmListSelect.TreeGroupGetImageIndex(Sender: TObject; Node: TTreeNode);
begin
  if Node.Expanded then
  Node.ImageIndex:=1 else
  Node.ImageIndex:=0;
end;

procedure TFmListSelect.TreeGroupGetSelectedIndex(Sender: TObject; Node: TTreeNode);
begin
  if ((TTreeView(Sender).Selected=nil) or (Node=nil)) then
  exit;
  Node.SelectedIndex:=Node.ImageIndex;
end;

procedure TFmListSelect.TreeGroupSelectionChanged(Sender: TObject);
var
  _TreeView: TTreeView;
begin

    _TreeView:= fTreeGroup.Tree;


    if (_TreeView.SelectionCount=0) or (fGridList.Grid=nil) then exit;

    if _TreeView.Items = nil then
     begin
       SetStatus('Ошибка Tree.');
//       Log('Ошибка Tree.');
       exit;
     end;
    try


   if (_TreeView.Items.Count = 0) or (_TreeView.SelectionCount = 0) then exit;

       if  not fTreeGroup.FirstFillTree then
       begin

         if _TreeView.Selected.Level=0 then
          fGridList.GroupArray:= nil
         else
          fGridList.GroupArray:= fTreeGroup.SelectedItems;

         fGridList.Fill();
       end else
       begin
           fTreeGroup.FirstFillTree:= false;
       end;

    except
    on E: Exception do
    begin
        SetStatus('Сбой выбора узла дерева.');
     end;
    end;
end;

procedure TFmListSelect.XMLSessionSavingProperties(Sender: TObject);
begin
  DBGridClearOrderBy(GridList);
end;

procedure TFmListSelect.SetStatus(_Text: string);
begin
   wStatus(fFormName,_Text,true);
end;

procedure TFmListSelect.ListFormInit(aIDPL_or_Catalog, aVendorcode, aName, aLabel: string; const aScod: string);
var
  aArr: ArrayOfArrayVariant;
  aScodMenu: TPopupMenu;
  aScodMenuItem: TMenuItem;
  i: Integer;
begin

    sBtnQSVendorCode.Caption:= aVendorcode;
    eQSName.Text:= aName;
    sBtnQSLabel.Caption:= aLabel;

    aScodMenu:= mQSSCode;
    aScodMenu.Items.Clear;

    aArr:= nil;
    if aScod = '-1' then
    begin
      if wFormMode = 0 then
          aArr:= fBase.SQLReadArr('SELECT SCOD FROM CATALOG_SCODS WHERE IDCTG_ITEMS='+aIDPL_or_Catalog+' ORDER BY SCOD')
      else
          aArr:= fBase.SQLReadArr('SELECT SCOD FROM PL_SCODS WHERE IDPL_ITEMS='+aIDPL_or_Catalog+' ORDER BY SCOD');
    end else
    begin
      if Length(aScod)>0 then
      begin
        SetLength(aArr,1,1);
        aArr[0,0]:= aScod;
      end;
    end;

    if Assigned(aArr) then
      for i:=0 to High(aArr) do begin
        aScodMenuItem:= TMenuItem.Create(aScodMenu);
        aScodMenuItem.Caption:= VarToStr(aArr[i,0]);
        aScodMenuItem.OnClick:=@_on_mScodClick;
        aScodMenu.Items.Add(aScodMenuItem);
      end;

    if Length(sBtnQSVendorCode.Caption)=0 then sBtnQSVendorCode.Enabled:= false;
    //if Length(TFmListSelect(aForm).eQSName.Text)=0 then TFmListSelect(aForm).sBtnQSName.Enabled:= false;
    if Length(sBtnQSLabel.Caption)=0 then sBtnQSLabel.Enabled:= false;
end;
end.

