unit FmNomenclatureEditU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  //wDBTreeU, wDBaseU, wDBGridU,
  wDBTreeU, wDBGridU, wBaseU, wFormulaU, wTypesU, UtilsU,
  ComCtrls, StdCtrls, Buttons, ExtDlgs, Spin, DBGrids, Menus, FmListSelectU, FmQuantityInPackingU, db, Grids;

type

  { TFmNomenclatureEdit }

  TFmNomenclatureEdit = class(TForm)
    btnCancel: TBitBtn;
    btnOK: TBitBtn;
    c_PRICE0: TSpeedButton;
    c_PRICE1: TSpeedButton;
    DBGridPriceFields: TDBGrid;
    dCalc: TCalculatorDialog;
    edGroup: TEdit;
    e_PRICE0: TEdit;
    e_PRICE1: TEdit;
    e_STOCK: TEdit;
    gbPriceFiels: TGroupBox;
    e_PD: TEdit;
    e_PK: TEdit;
    e_PC: TEdit;
    gbName: TGroupBox;
    gbGroup: TGroupBox;
    gb_Unit: TGroupBox;
    GridImageList: TImageList;
    DBGridMatching: TDBGrid;
    ImageList16: TImageList;
    kVendorCode: TLabeledEdit;
    lbC: TLabel;
    lbP1: TLabel;
    lbS: TLabel;
    lbP0: TLabel;
    l_P: TLabel;
    lbD: TLabel;
    lbK: TLabel;
    l_M: TLabel;
    l_C: TLabel;
    l_N: TLabel;
    l_D: TLabel;
    l_K: TLabel;
    gbIdent: TGroupBox;
    kArticul: TLabeledEdit;
    kEdini: TComboBox;
    kName: TMemo;
    kNumber: TLabeledEdit;
    e_PN: TEdit;
    e_PRICE: TEdit;
    kScod: TLabeledEdit;
    e_PM: TEdit;
    Label3: TLabel;
    lbP: TLabel;
    lbN: TLabel;
    lbM: TLabel;
    l_P0: TLabel;
    l_P1: TLabel;
    l_S: TLabel;
    mAdd: TMenuItem;
    mEditMatch: TMenuItem;
    mDelete: TMenuItem;
    mEditQuantInPack: TMenuItem;
    m_EditMatch: TMenuItem;
    m_EditQuantInPack: TMenuItem;
    Panel4: TPanel;
    m_Matching: TPopupMenu;
    m_EditMatching: TPopupMenu;
    pPrices: TPanel;
    pc: TPageControl;
    Panel1: TPanel;
    pTop: TPanel;
    Panel3: TPanel;
    c_PRICE: TSpeedButton;
    btnChangeGroup: TSpeedButton;
    c_PN: TSpeedButton;
    c_PM: TSpeedButton;
    c_PD: TSpeedButton;
    c_PK: TSpeedButton;
    c_PC: TSpeedButton;
    SpeedButton1: TSpeedButton;
    pcTbPrice: TTabSheet;
    pcTbMatching: TTabSheet;
    tbMatch: TToolBar;
    tbMatchBtnAdd: TToolButton;
    tbMatchBtnDelete: TToolButton;
    tbMatchBtnEdit: TToolButton;
    procedure btnChangeGroupClick(Sender: TObject);
    procedure c_PCClick(Sender: TObject);
    procedure c_PDClick(Sender: TObject);
    procedure c_PKClick(Sender: TObject);
    procedure c_PMClick(Sender: TObject);
    procedure c_PNClick(Sender: TObject);
    procedure c_PRICE0Click(Sender: TObject);
    procedure c_PRICE1Click(Sender: TObject);
    procedure DBGridPriceFieldsTitleClick(Column: TColumn);
    procedure e_PCKeyPress(Sender: TObject; var Key: char);
    procedure e_PNChange(Sender: TObject);
    procedure e_PNEnter(Sender: TObject);
    procedure e_PNExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure e_PRICEChange(Sender: TObject);
    procedure e_PRICEEnter(Sender: TObject);
    procedure e_PRICEExit(Sender: TObject);
    procedure e_PRICEKeyPress(Sender: TObject; var Key: char);
    procedure c_PRICEClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure kArticulEnter(Sender: TObject);
    procedure kArticulExit(Sender: TObject);
    procedure kNameEnter(Sender: TObject);
    procedure kNameExit(Sender: TObject);
    procedure kScodEnter(Sender: TObject);
    procedure kScodExit(Sender: TObject);
    procedure mEditMatchClick(Sender: TObject);
    procedure mEditQuantInPackClick(Sender: TObject);
    procedure m_MatchingPopup(Sender: TObject);
    procedure m_EditMatchingPopup(Sender: TObject);
    procedure pcChange(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure tbMatchBtnAddClick(Sender: TObject);
    procedure tbMatchBtnDeleteClick(Sender: TObject);
    procedure tbMatchBtnEditClick(Sender: TObject);
  private
    fOwnerID: integer;     // ID выбранного контрагента
    fIdMainOwner: string; // ID основного контрагента (к которому привязан каталог)
    dbPriceArr: ArrayOfArrayVariant;

    fGridPriceFields: TwDBGrid;
    fGridMatching: TwDBGrid;
    fBase: TwBase;

    __PRICE_MAX_FTIMESTAMP_ARR: ArrayOfDateTime; // массив максимальных значений таймштамп из таблицы прайс-листы

    procedure SetStatus(_Text: string);
  public
    { public declarations }
    procedure CreateGridPriceFields();
    property Base: TwBase read fBase write fBase;
    property GridPriceFields: TwDBGrid read fGridPriceFields write fGridPriceFields;
    property GridMatching: TwDBGrid read fGridMatching write fGridMatching;
    //property PriceMaxFTImeStampArr: ArrayOfDateTime read __PRICE_MAX_FTIMESTAMP_ARR write __PRICE_MAX_FTIMESTAMP_ARR;
    procedure GridPriceFieldsRefrash();
  end;

var
  FmNomenclatureEdit: TFmNomenclatureEdit;

implementation

uses
  wLogU, wFuncU, FmTreeU;

{$R *.lfm}

{ TFmNomenclatureEdit }

procedure TFmNomenclatureEdit.SetStatus(_Text: string);
begin
     wStatus(FmNomenclatureEdit.Name,_Text,true);
end;

procedure TFmNomenclatureEdit.c_PRICEClick(Sender: TObject);
begin
  CalcOpen(e_PRICE,dCalc);
end;

procedure TFmNomenclatureEdit.FormDestroy(Sender: TObject);
begin
  try
   wLog('FmNomenclatureEdit','Выгрузка формы...');

    // выгружаем подгруженные DBGrid
    if Assigned(fGridPriceFields) then fGridPriceFields.Destroy();
    if Assigned(fGridMatching) then fGridMatching.Destroy();


  wLog('FmNomenclatureEdit','Выгрузка формы успешно завершена.');

  except
    on E: Exception do
    begin
        __Log.SaveLogError(E);
        SetStatus('Сбой выгрузки формы: Каталог.');
        wLog('FmNomenclatureEdit','Ошибка [FmDestroy]: "' + E.Message + '"');
        wLog('FmNomenclatureEdit','Сбой выгрузки плагина.');
        ShowMessage('Ошибка [FmDestroy]: "' + E.Message + '"');
     end;
  end;
end;

procedure TFmNomenclatureEdit.GridPriceFieldsRefrash();
var
  _P, _N, _M, _D, _C, _K, _S, _P0, _P1, _P2, _P3, _P4, _P5, _P6, _P7,
    _P8, _P9, _P10, _P02, _P03, _P04, _P05, _P06, _P07, _P08, _P09,
    _P010, _S1, _S2, _S3, _S4, _S5, _PDATEDIFF: Double;
begin
   if not Assigned(fGridPriceFields) then CreateGridPriceFields;

   _P:= EditValue(e_PRICE); // цена поставщика

   _N:= EditValue(e_PN);

   _N:= _P*(1+_N/100)-_P;

   _M:= EditValue(e_PM);
   _D:= EditValue(e_PD);
   _C:= EditValue(e_PC);
   _C:= _C/100;
   _K:= EditValue(e_PK);

   _S:= EditValue(e_STOCK);

   _P0:= EditValue(e_PRICE0);  // закуп (базовая)

   _P1:= EditValue(e_PRICE1); // каталог

   if self.Showing and Assigned(dbPriceArr) then
   begin
     _P2:= VarToDouble(dbPriceArr[0,7]);
     _P3:= VarToDouble(dbPriceArr[0,8]);
     _P4:= VarToDouble(dbPriceArr[0,9]);
     _P5:= VarToDouble(dbPriceArr[0,10]);
     _P6:= VarToDouble(dbPriceArr[0,11]);
     _P7:= VarToDouble(dbPriceArr[0,12]);
     _P8:= VarToDouble(dbPriceArr[0,13]);
     _P9:= VarToDouble(dbPriceArr[0,14]);
     _P10:= VarToDouble(dbPriceArr[0,15]);

     _S1:= VarToDouble(dbPriceArr[0,1]);
     _S2:= VarToDouble(dbPriceArr[0,2]);
     _S3:= VarToDouble(dbPriceArr[0,3]);
     _S4:= VarToDouble(dbPriceArr[0,4]);
     _S5:= VarToDouble(dbPriceArr[0,5]);


     _P02:= VarToDouble(dbPriceArr[0,17]);
     _P03:= VarToDouble(dbPriceArr[0,18]);
     _P04:= VarToDouble(dbPriceArr[0,19]);
     _P05:= VarToDouble(dbPriceArr[0,20]);
     _P06:= VarToDouble(dbPriceArr[0,21]);
     _P07:= VarToDouble(dbPriceArr[0,22]);
     _P08:= VarToDouble(dbPriceArr[0,23]);
     _P09:= VarToDouble(dbPriceArr[0,24]);
     _P010:= VarToDouble(dbPriceArr[0,25]);
     _PDATEDIFF:= VarToDouble(dbPriceArr[0,26]);
   end else
   begin
     _P2:=0;
     _P3:=0;
     _P4:=0;
     _P5:=0;
     _P6:=0;
     _P7:=0;
     _P8:=0;
     _P9:=0;
     _P10:=0;

     _S1:=0;
     _S2:=0;
     _S3:=0;
     _S4:=0;
     _S5:=0;


     _P02:=0;
     _P03:=0;
     _P04:=0;
     _P05:=0;
     _P06:=0;
     _P07:=0;
     _P08:=0;
     _P09:=0;
     _P010:=0;
     _PDATEDIFF:=0;
   end;

   fGridPriceFields.Formula.VarArray:= [
               {P1}	_P1,  // каталог
               {P}	_P,   // цена поставщика
               {P2}	_P2,
               {P3}	_P3,
               {P4}	_P4,
               {P5}	_P5,
               {P6}	_P6,
               {P7}	_P7,
               {P8}	_P8,
               {P9}	_P9,
               {P10}	_P10,
               {P0}	_P0,  // закуп (базовая)
               {P02}	_P02,
               {P03}	_P03,
               {P04}	_P04,
               {P05}	_P05,
               {P06}	_P06,
               {P07}	_P07,
               {P08}	_P08,
               {P09}	_P09,
               {P010}	_P010,
               {S}	_S,
               {S1}	_S1,
               {S2}	_S2,
               {S3}	_S3,
               {S4}	_S4,
               {S5}	_S5,
               {N}	_N,
               {M}	_M,
               {D}	_D,
               {C}	_C,
               {K}	_K,
               {PDATEDIFF} _PDATEDIFF
                         ];
   fGridPriceFields.Fill(); // заполнение грида
   //_DBase.PriceField:=false;
end;

procedure TFmNomenclatureEdit.CreateGridPriceFields();
var
  _SQL_Text: String;
begin
  fIdMainOwner:= fBase.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист

  _SQL_Text:='select ID, NAME, FORMULA FROM PRICEFIELD WHERE FCLOSE=0 ORDER BY PRIORITY';

  fGridPriceFields:= TwDBGrid.Create(Base,DBGridPriceFields,_SQL_Text);
  fGridPriceFields.Formula:= TFormula.Create(fGridPriceFields.Grid);
  fGridPriceFields.Formula.CalculateField:= 'CPRICE';
  fGridPriceFields.Formula.CurrencyArray:= fBase.GetCurrencyArray();
  fGridPriceFields.SortON:= false;
end;

procedure TFmNomenclatureEdit.FormShow(Sender: TObject);
var
  _SQL_Text: String;
begin

 dbPriceArr:= fBase.SQLReadArr('SELECT "CATALOG".ID,  ' + wfLineEnding //0
     +'  PLOUR.STOCK AS STOCK1, ' + wfLineEnding                     //1
     +'  PLOUR.STOCK2 AS STOCK2, ' + wfLineEnding                    //2
     +'  PLOUR.STOCK3 AS STOCK3, ' + wfLineEnding                    //3
     +'  PLOUR.STOCK4 AS STOCK4, ' + wfLineEnding                    //4
     +'  PLOUR.STOCK5 AS STOCK5, ' + wfLineEnding                    //5
     +'   PLFP.PRICEPL AS PRICEPL, ' + wfLineEnding                  //6
     +'   PLFP.PRICEPL2 AS PRICEPL2,  ' + wfLineEnding               //7
     +'   PLFP.PRICEPL3 AS PRICEPL3,  ' + wfLineEnding               //8
     +'   PLFP.PRICEPL4 AS PRICEPL4,  ' + wfLineEnding               //9
     +'   PLFP.PRICEPL5 AS PRICEPL5,  ' + wfLineEnding               //10
     +'   PLFP.PRICEPL6 AS PRICEPL6,  ' + wfLineEnding               //11
     +'   PLFP.PRICEPL7 AS PRICEPL7,  ' + wfLineEnding               //12
     +'   PLFP.PRICEPL8 AS PRICEPL8,  ' + wfLineEnding               //13
     +'   PLFP.PRICEPL9 AS PRICEPL9,  ' + wfLineEnding               //14
     +'   PLFP.PRICEPL10 AS PRICEPL10,  ' + wfLineEnding             //15
     +'  PLOUR.PRICECALC AS PRICEOUR,  ' + wfLineEnding              //16
     +'  PLOUR.PRICECALC2 AS PRICEOUR2,  ' + wfLineEnding            //17
     +'  PLOUR.PRICECALC3 AS PRICEOUR3,  ' + wfLineEnding            //18
     +'  PLOUR.PRICECALC4 AS PRICEOUR4,  ' + wfLineEnding            //19
     +'  PLOUR.PRICECALC5 AS PRICEOUR5,  ' + wfLineEnding            //20
     +'  PLOUR.PRICECALC6 AS PRICEOUR6,  ' + wfLineEnding            //21
     +'  PLOUR.PRICECALC7 AS PRICEOUR7,  ' + wfLineEnding            //22
     +'  PLOUR.PRICECALC8 AS PRICEOUR8,  ' + wfLineEnding            //23
     +'  PLOUR.PRICECALC9 AS PRICEOUR9,  ' + wfLineEnding            //24
     +'  PLOUR.PRICECALC10 AS PRICEOUR10,  ' + wfLineEnding           //25
     +'  PLFP.PDATE AS PDATE  ' + wfLineEnding                        //26
     +'   FROM "CATALOG"  ' + wfLineEnding
     +'  LEFT JOIN CATALOG_PL_MIN_PRICE("CATALOG".ID) PLFP ON (1=1) ' + wfLineEnding
     +'   LEFT OUTER JOIN "PL_ITEMS" PLOUR ON (  ' + wfLineEnding
     +'   "CATALOG".VENDORCODE = PLOUR.VENDORCODE AND "CATALOG".IDOWNER = PLOUR.IDOWNER) ' + wfLineEnding
     +'   WHERE ("CATALOG".ID='+kNumber.Text+') ');

GridPriceFieldsRefrash();

//_PL_FTIMESTAMP:= Base.PrepareWhereStringFromDateTime('PL.FTIMESTAMP',__PRICE_MAX_FTIMESTAMP_ARR);

_SQL_Text:='select MTH.ID, '
          +' MTH.IDOWNER, '
          +' MTH.IDPL_ITEMS, '
          +' PL.VENDORCODE AS PLVENDORCODE, '
          +' OWNER.NAME as OWNERNAME, '
          +' PL.NAME as PLNAME, '
          +' MTH.QUANTITYINPACKING as QUANTITYINPACKING, '
          +' MTH.IDOWNER as IDOWNER  '
          +' from "CATALOG_MATCHING" MTH '
          +' left outer join "PL_ITEMS" PL ON (MTH.IDPL_ITEMS=PL.ID)    '
          +' left join "OWNER" on "OWNER".ID=MTH.IDOWNER '
          +' WHERE /*group_string*/';

fGridMatching:= TwDBGrid.Create(Base,DBGridMatching,_SQL_Text);
fGridMatching.GroupField:= 'MTH.IDCATALOG';
fGridMatching.GroupArray:= [StrToInt(kNumber.Text)];
fGridMatching.MultiSelect:= true;
fGridMatching.Fill();

end;

procedure TFmNomenclatureEdit.kArticulEnter(Sender: TObject);
begin
  (Sender as TLabeledEdit).Color:=clSkyBlue; //clDefault    clSkyBlue
end;

procedure TFmNomenclatureEdit.kArticulExit(Sender: TObject);
begin
  (Sender as TLabeledEdit).Color:=clDefault; //clDefault    clSkyBlue
end;

procedure TFmNomenclatureEdit.kNameEnter(Sender: TObject);

begin
 (Sender as TMemo).Color:=clSkyBlue; //clDefault
end;

procedure TFmNomenclatureEdit.kNameExit(Sender: TObject);
begin
  (Sender as TMemo).Color:=clDefault; //clDefault    clSkyBlue
end;

procedure TFmNomenclatureEdit.kScodEnter(Sender: TObject);
begin
  (Sender as TLabeledEdit).Color:=clSkyBlue; //clDefault    clSkyBlue
end;

procedure TFmNomenclatureEdit.kScodExit(Sender: TObject);
begin
  (Sender as TLabeledEdit).Color:=clDefault; //clDefault    clSkyBlue
end;

procedure TFmNomenclatureEdit.mEditMatchClick(Sender: TObject);
var
  _IDMatching: integer;
  //_VendorcodeMatching: string;

  _Form: TFmListSelect;
  _SelectedRows: ArrayOfInteger;
  _TimeStamp: string;
  _arr: ArrayOfArrayVariant;
  _IDCatalog: Integer;
  _QuantityInPacked: Double;
  _QuantityInPackedGrid: Double;
  _SelectedID: LongInt;
begin
  // изменение одного соответствия
  if DBGridMatching.DataSource.DataSet.RecordCount=0 then exit;
  try
    _IDMatching:= fGridMatching.SelectedRows[0];
    _SelectedID:= DBGridMatching.DataSource.DataSet.FieldByName('IDPL_ITEMS').AsInteger;
    //_VendorcodeMatching:= DBGridMatching.DataSource.DataSet.FieldByName('ID').AsString;
    _QuantityInPackedGrid:= DBGridMatching.DataSource.DataSet.FieldByName('QUANTITYINPACKING').AsFloat;
    screen.Cursor:= crSQLWait;

    _Form:= TFmListSelect.Create(self);
    _Form.Base:= fBase;
    _Form.MultiSelectGrid:= true;
    _Form.wFormMode:= 0; // PriceLists
    _Form.Where:= 'ID<>'+fIdMainOwner;

    _Form.ListFormInit(
      kNumber.Text,
      kVendorCode.Text,
      kName.Text,
      kArticul.Text
    );

    _SelectedRows:= nil;

    _Form.GridList.Options:=_Form.GridList.Options - [dgMultiSelect];
    _Form.wDataSetLocateField:='ID';
    _Form.wDataSetLocateValue:=_SelectedID;
    _Form.wIDTreeItem:=DBGridMatching.DataSource.DataSet.FieldByName('IDOWNER').AsInteger;

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
         _IDCatalog:= StrToInt(kNumber.Text);

         try
           _arr:= nil;
           _arr:= fBase.SQLReadArr('PL_ITEMS',['IDOWNER','ID'],'ID='+IntToStr(_SelectedRows[0]),'ID');
           fBase.SQLUpdate('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],[_arr[0,0],_IDCatalog,_arr[0,1],_QuantityInPacked,_TimeStamp,integer(1)],'ID='+IntToStr(_IDMatching),false);

         fBase.SQLTransactionEnd(true);
         except
           fBase.SQLTransactionEnd(false);
           raise;
         end;
         fGridMatching.Fill;
         DBGridMatching.DataSource.DataSet.Locate('ID',IntTOStr(_IDMatching),[loCaseInsensitive]);
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

procedure TFmNomenclatureEdit.mEditQuantInPackClick(Sender: TObject);
var
  _Form: TFmQuantityInPacking;
  _SelectedRows: ArrayOfInteger;
  i: integer;
  _TimeStamp: string;
  _QuantityInPacked: Double;
begin
   //изменение одного соответствия
  try
    _SelectedRows:= fGridMatching.SelectedRows;

    _Form:= TFmQuantityInPacking.Create(self);

    //if _QuantityInPackedGrid<1 then
    //   begin
    //     _Form.spQuantInPackLeft.Value:=1;
    //     _Form.spQuantInPackRight.Value:= 1/_QuantityInPackedGrid;
    //   end else
    //   begin
    //      _Form.spQuantInPackLeft.Value:= 1*_QuantityInPackedGrid;
    //      _Form.spQuantInPackRight.Value:= 1;
    //   end;

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
         fGridMatching.Fill;
         DBGridMatching.DataSource.DataSet.Locate('ID',IntTOStr(_SelectedRows[i]),[loCaseInsensitive]);
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

procedure TFmNomenclatureEdit.m_MatchingPopup(Sender: TObject);
begin
  // if DBGridMatching.DataSource.DataSet.RecordCount=0 then
  //    begin
  //       mEditMatch.Enabled:=false;
  //       mEditQuantInPack.Enabled:=false;
  //    end else
  //    begin
  //       mEditMatch.Enabled:=true;
  //       mEditQuantInPack.Enabled:=true;
  //    end;
  //
  //if DBGridMatching.SelectedRows.Count>1 then
  //   mEditMatch.Enabled:=false else mEditMatch.Enabled:=true;
end;

procedure TFmNomenclatureEdit.m_EditMatchingPopup(Sender: TObject);
begin

  //if DBGridMatching.DataSource.DataSet.RecordCount=0 then
  //   begin
  //      m_EditMatch.Enabled:=false;
  //      m_EditQuantInPack.Enabled:=false;
  //   end else
  //   begin
  //      m_EditMatch.Enabled:=true;
  //      m_EditQuantInPack.Enabled:=true;
  //   end;
  //
  //  if DBGridMatching.SelectedRows.Count>1 then
  //   m_EditMatch.Enabled:=false else m_EditMatch.Enabled:=true;
end;

procedure TFmNomenclatureEdit.pcChange(Sender: TObject);
begin
//if (DBaseCatalog <> nil) and DBaseCatalog.TransactionUpdate.Active then
// begin
//   if MessageDlg('Перед добавлением соответствий новую позицию необходимо сохранить. Сохранить "'+kName.Text+'" ?',mtConfirmation, mbOKCancel, 0) = mrOK then
//    begin
//       DBaseCatalog.SQLTransactionEnd(true);
//    end else
//    begin
//       TPageControl(Sender).ActivePageIndex:=0;
//       exit;
//    end;
// end;
//
//case TPageControl(Sender).ActivePageIndex of
//   0 : GridPriceFieldsRefrash() ;
//   1 :
//     begin
//       _DBGridMatching.SetParam('$WHEREID','MATCHING.IDCATALOG='+kNumber.Text);
//       _DBGridMatching.FillGrid;
//     end;
//  end;
end;

procedure TFmNomenclatureEdit.SpeedButton1Click(Sender: TObject);
begin
  if MessageDlg('Сгенерировать новый штрих-код для "'+kName.Text+'" ?',mtConfirmation, mbOKCancel, 0) = mrOK
     then
       begin
         kScod.Text:=GenEAN(__SCODPREFIX, '', kNumber.Text);
       end;
end;

procedure TFmNomenclatureEdit.tbMatchBtnAddClick(Sender: TObject);
var
  _Form: TFmListSelect;
  _SelectedRows: ArrayOfInteger;
  _QuantityInPacked: Double;
  i: integer;
  _TimeStamp: string;
  _arr: ArrayOfArrayVariant;
  _IDCatalog: Integer;
  _IDMatching: integer;
begin
  screen.Cursor:= crSQLWait;
 _Form:= TFmListSelect.Create(self);
 _Form.Base:= fBase;
 _Form.MultiSelectGrid:= true;
 _Form.wFormMode:= 0; // PriceLists
 _Form.Where:= 'ID<>'+fIdMainOwner;

 _Form.ListFormInit(
      kNumber.Text,
      kVendorCode.Text,
      kName.Text,
      kArticul.Text
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
      _IDCatalog:= StrToInt(kNumber.Text);

      try
      for i:=0 to High(_SelectedRows) do
         begin
           _arr:= nil;
            _arr:= fBase.SQLReadArr('PL_ITEMS',['IDOWNER','ID'],'ID='+IntToStr(_SelectedRows[i]),'ID');
            if Length(_arr)>0 then
                //_IDMatching:= fBase.SQLInsert('INSERT INTO "CATALOG_MATCHING" (IDOWNER,IDCATALOG,IDPL_ITEMS,QUANTITYINPACKING,FTIMESTAMP,IDUSER) VALUES ('+IntTOStr(_arr[0,0])+','+IntToStr(_IDCatalog)+','+QuotedStr(_arr[0,1])+','+FloatToStr(_QuantityInPacked)+','+QuotedStr(_TimeStamp)+',1)',false);
                _IDMatching:= fBase.SQLInsert('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],
                                              [_arr[0,0],_IDCatalog,_arr[0,1],_QuantityInPacked,_TimeStamp,integer(1)],'IDOWNER, IDPL_ITEMS',false);

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
      fGridMatching.Fill;
      finally
        DBGridMatching.DataSource.DataSet.Locate('ID',IntTOStr(_IDMatching),[loCaseInsensitive]);
      end;
    end;
    _arr:= nil;
    _SelectedRows:= nil;
    screen.Cursor:= crDefault;
end;

procedure TFmNomenclatureEdit.tbMatchBtnDeleteClick(Sender: TObject);
var
  _SelectedRows:ArrayOfInteger;
  i: integer;
  _BookMark: TBookMark;
begin
  _SelectedRows:=nil;
  _SelectedRows:= fGridMatching.SelectedRows;

  if Length(_SelectedRows)>0 then
     begin
       if MessageDlg('Удалить выделенные соответствия ('+IntTOStr(Length(_SelectedRows))+') ?',mtConfirmation, mbOKCancel, 0) = mrOK
        then
         begin

             try
             _BookMark:= DBGridMatching.DataSource.DataSet.Bookmark;

             for i:=0 to High(_SelectedRows) do
                fBase.SQLDelete('CATALOG_MATCHING','ID='+IntToStr(_SelectedRows[i]),false);
              fBase.SQLTransactionEnd(true);
              fGridMatching.Fill;
              if DBGridMatching.DataSource.DataSet.RecordCount>0 then DBGridMatching.DataSource.DataSet.Bookmark:= _BookMark;
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

procedure TFmNomenclatureEdit.tbMatchBtnEditClick(Sender: TObject);
var
  pnt: TPoint;
begin
  pnt := Mouse.CursorPos;
  m_EditMatching.PopUp(pnt.x,pnt.y);
end;

procedure TFmNomenclatureEdit.e_PRICEEnter(Sender: TObject);
begin
 //(Sender as TEdit).Color:=clSkyBlue; //clDefault    clSkyBlue
  Razdelitel(Sender,2,true);
end;

procedure TFmNomenclatureEdit.e_PRICEExit(Sender: TObject);
begin
  //(Sender as TEdit).Color:=clDefault; //clDefault    clSkyBlue
  Razdelitel(Sender,2,false);
  GridPriceFieldsRefrash();
end;

procedure TFmNomenclatureEdit.e_PRICEKeyPress(Sender: TObject; var Key: char
  );
begin
  key:=FilterSimvol(Sender,Key,false);
end;

procedure TFmNomenclatureEdit.e_PRICEChange(Sender: TObject);
begin
  with Sender as TEdit do
    begin
      CheckClear(Sender,2,Text);
    end;
end;

procedure TFmNomenclatureEdit.FormCreate(Sender: TObject);
begin

  wLog('FmNomenclatureEdit','Инициализация формы... [FmNomenclatureEdit]');

  try
    Razdelitel(e_PRICE,2,false);
    Razdelitel(e_PM,2,false);
    Razdelitel(e_PC,2,false);
    Razdelitel(e_PN,2,false);
    Razdelitel(e_PD,2,false);
    Razdelitel(e_PK,2,false);
    Razdelitel(e_PRICE0,2,false);
    Razdelitel(e_PRICE1,2,false);
    Razdelitel(e_STOCK,0,false);


//
//     DBGrid.Add(TwDBGrid.Create(wFormID, DBGridPriceFields,false,nil,nil,nil,'select ID, NAME, FORMULA FROM PRICEFIELD WHERE FCLOSE=0 ORDER BY $ORDERBY',['$ORDERBY=PRIORITY'],false)); // инициализация DBGrid
//     _DBGridPriceFields:= __DBGrid(wFormID,DBGridPriceFields);
//
//     _TimeStampMaxArr:= _DBase.ReadMaxDateTimeValues('PRICE-LISTS','FTIMESTAMP','IDOWNER');
//
     //DBGrid.Add(TwDBGrid.Create(wFormID, DBGridMatching,true,nil,nil,nil,'select $FIELDS,MATCHING.VENDORCODE AS MATCHINGVENDORCODE, OWNER.NAME as OWNERNAME, PL.NAME as PLNAME, MATCHING.QUANTITYINPACKING as QUANTITYINPACKING, MATCHING.IDOWNER as IDOWNER from $TABLE left outer join "PRICE-LISTS" PL ON ($WHERETIMESTAMP MATCHING.VENDORCODE=PL.VENDORCODE and MATCHING.IDOWNER=PL.IDOWNER)   left join OWNER on OWNER.ID=MATCHING.IDOWNER where $WHEREID order by $ORDERBY',['$FIELDS=MATCHING.ID,MATCHING.IDOWNER','$WHEREID=MATCHING.IDCATALOG','$WHERETIMESTAMP=','$ORDERBY=MATCHING.IDOWNER','$TABLE=MATCHING'] ,true)); // инициализация DBGrid
//     // заполняем при отображении вкладки с гридом
//     _DBGridMatching:= __DBGrid(wFormID,DBGridMatching);
//
//     _DBGridMatching.SetParam('$WHERETIMESTAMP',_DBase.PrepareWhereStringFromDateTime('PL.FTIMESTAMP',_TimeStampMaxArr,true,2));
//
    wLog('FmNomenclatureEdit','Инициализация формы успешно завершена.');
  except
    on E: Exception do
    begin
        __Log.SaveLogError(E);
        SetStatus('Сбой инициализации формы.');
        wLog('FmNomenclatureEdit','Ошибка [FmCreate]: "' + E.Message + '"');
        wLog('FmNomenclatureEdit','Сбой инициализации формы.');
        ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

     end;
  end;


end;

procedure TFmNomenclatureEdit.FormClose(Sender: TObject;
  var CloseAction: TCloseAction);
begin
 // CloseAction := caFree;
end;

procedure TFmNomenclatureEdit.FormCloseQuery(Sender: TObject;
  var CanClose: boolean);
begin
    if ModalResult = mrCancel then
      if MessageDlg('Закрыть без сохранения?',mtConfirmation, mbOKCancel, 0) = mrCancel
       then
         CanClose:= false
       else
         ModalResult:= mrCancel;
end;

procedure TFmNomenclatureEdit.btnChangeGroupClick(Sender: TObject);
var
  _Form: TFmTree;
  _TreeTag: integer;
begin

  _Form := TFmTree.Create(Self);
  _Form.Base:= fBase;
  _Form.IdGroup:= edGroup.Tag;
  try

    _Form.ShowModal;
    _TreeTag:= _Form.IdGroup;
    if _Form.ModalResult = mrOK then
      if _TreeTag>0 then
      begin
           edGroup.Tag:=_TreeTag;
           edGroup.Text:=_Form.Tree.BreadCrumbs(_TreeTag);
      end;


  finally
    _Form.Free;
  end;
end;

procedure TFmNomenclatureEdit.c_PCClick(Sender: TObject);
begin
  CalcOpen(e_PC,dCalc);
end;

procedure TFmNomenclatureEdit.c_PDClick(Sender: TObject);
begin
  CalcOpen(e_PD,dCalc);
end;

procedure TFmNomenclatureEdit.c_PKClick(Sender: TObject);
begin
  CalcOpen(e_PK,dCalc);
end;

procedure TFmNomenclatureEdit.c_PMClick(Sender: TObject);
begin
  CalcOpen(e_PM,dCalc);
end;

procedure TFmNomenclatureEdit.c_PNClick(Sender: TObject);
begin
  CalcOpen(e_PN,dCalc);
end;

procedure TFmNomenclatureEdit.c_PRICE0Click(Sender: TObject);
begin
  CalcOpen(e_PRICE0,dCalc);
end;

procedure TFmNomenclatureEdit.c_PRICE1Click(Sender: TObject);
begin
  CalcOpen(e_PRICE1,dCalc);
end;

procedure TFmNomenclatureEdit.DBGridPriceFieldsTitleClick(Column: TColumn);
begin
  if Column <> nil then exit;
end;

procedure TFmNomenclatureEdit.e_PCKeyPress(Sender: TObject; var Key: char);
begin
  key:=FilterSimvol(Sender,Key,true);
end;

procedure TFmNomenclatureEdit.e_PNChange(Sender: TObject);
begin
with Sender as TEdit do
  begin
    CheckClear(Sender,2,Text);
  end;
end;

procedure TFmNomenclatureEdit.e_PNEnter(Sender: TObject);
begin
(Sender as TEdit).Color:=clSkyBlue; //clDefault    clSkyBlue
 Razdelitel(Sender,2,true);
end;

procedure TFmNomenclatureEdit.e_PNExit(Sender: TObject);
begin
(Sender as TEdit).Color:=clDefault; //clDefault    clSkyBlue
Razdelitel(Sender,2,false);
GridPriceFieldsRefrash();
end;

end.

