unit FmPriceFieldsU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  SysUtils, Forms, Dialogs, ExtCtrls, StdCtrls, ComCtrls, Buttons, DBGrids,Spin, db, Graphics
  ,wLogU, wBaseU, wDBGridU, wFuncU, wFormulaU, wTypesU,
  Classes, Controls, ExtDlgs, MaskEdit, Menus;

type

  { TFmPriceFields }

  TFmPriceFields = class(TForm)
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    cb_Closed: TCheckBox;
    c_PC: TSpeedButton;
    c_PD: TSpeedButton;
    c_PK: TSpeedButton;
    c_PM: TSpeedButton;
    c_PN: TSpeedButton;
    c_PRICE: TSpeedButton;
    c_PRICE0: TSpeedButton;
    c_PRICE1: TSpeedButton;
    DBGridPriceFields: TDBGrid;
    dCalc: TCalculatorDialog;
    e_PRICE0: TEdit;
    e_PRICE1: TEdit;
    e_result: TEdit;
    e_Stock: TSpinEdit;
    e_Name: TEdit;
    e_PC: TEdit;
    e_PD: TEdit;
    e_PK: TEdit;
    e_PM: TEdit;
    e_PN: TEdit;
    e_PRICE: TEdit;
    gbPrice: TGroupBox;
    gbEdit: TGroupBox;
    gbResult: TGroupBox;
    gbMain: TGroupBox;
    gbFormula: TGroupBox;
    gbHelp: TGroupBox;
    ILtoolbars: TImageList;
    ImageList16: TImageList;
    Label1: TLabel;
    lbC: TLabel;
    lbP1: TLabel;
    lbS: TLabel;
    lbP: TLabel;
    lbP0: TLabel;
    lbN: TLabel;
    lbM: TLabel;
    lbD: TLabel;
    lbK: TLabel;
    l_P1: TLabel;
    l_S: TLabel;
    l_P: TLabel;
    l_M: TLabel;
    l_C: TLabel;
    l_N: TLabel;
    l_D: TLabel;
    l_K: TLabel;
    l_P0: TLabel;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    m_Help: TMemo;
    m_Formula: TMemo;
    Panel1: TPanel;
    Panel2: TPanel;
    m_PriceField: TPopupMenu;
    pRight: TPanel;
    gbPrices: TGroupBox;
    e_Priority: TSpinEdit;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    tbPriceFieldBtnAdd: TToolButton;
    tbPriceFieldBtnCopy: TToolButton;
    tbPriceFieldBtnDelete: TToolButton;
    tbPriceField: TToolBar;
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure cb_ClosedChange(Sender: TObject);
    procedure c_PCClick(Sender: TObject);
    procedure c_PDClick(Sender: TObject);
    procedure c_PKClick(Sender: TObject);
    procedure c_PMClick(Sender: TObject);
    procedure c_PNClick(Sender: TObject);
    procedure c_PRICE0Click(Sender: TObject);
    procedure c_PRICE1Click(Sender: TObject);
    procedure c_PRICEClick(Sender: TObject);
    procedure e_NameChange(Sender: TObject);
    procedure e_NameEnter(Sender: TObject);
    procedure e_NameExit(Sender: TObject);
    procedure e_PCKeyPress(Sender: TObject; var Key: char);
    procedure e_PRICEChange(Sender: TObject);
    procedure e_PRICEEnter(Sender: TObject);
    procedure e_PRICEExit(Sender: TObject);
    procedure e_PRICEKeyPress(Sender: TObject; var Key: char);
    procedure e_PriorityChange(Sender: TObject);
    procedure e_PriorityEnter(Sender: TObject);
    procedure e_PriorityExit(Sender: TObject);
    procedure e_StockChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure gbMainClick(Sender: TObject);
    procedure lbP0Click(Sender: TObject);
    procedure m_FormulaChange(Sender: TObject);
    procedure m_FormulaExit(Sender: TObject);
    procedure tbPriceFieldBtnAddClick(Sender: TObject);
    procedure tbPriceFieldBtnCopyClick(Sender: TObject);
    procedure tbPriceFieldBtnDeleteClick(Sender: TObject);
  private
    { private declarations }
    FormIDent: string;
    fFormula: TFormula;
    fGrid: TwDBGrid;
    fBase: TwBase;

    _FormShow: boolean;     // флаг того, что форма отображена
    _Changed: boolean;      // есть несохраненные изменения
    _FormulaFiled: boolean; // ошибка в формуле             _

    procedure SetStatus(_Text: string);
    procedure _onDataChange(Sender: TObject; Field: TField); // изменение датасета в гриде
    procedure _resultRefrash();// обновление результата
    procedure _formulaChange(); // изменение формулы
    procedure _SetChanged(const AValue: boolean = true); // установка флага - изменен

    property wFormID: string read FormIDent write FormIDent;
    property Changed: boolean read _Changed write _Changed;
  public
    { public declarations }
  end;

var
  FmPriceFields: TFmPriceFields;

implementation


{$R *.lfm}

{ TFmPriceFields }

procedure TFmPriceFields.FormCreate(Sender: TObject);
begin
  wFormID:=Self.Name;
  _FormulaFiled:= false;
  _Changed:= false;

   wLog('FmPriceFields','Инициализация формы... ['+wFormID+']');

   Razdelitel(e_PRICE,2,false);
   Razdelitel(e_PM,2,false);
   Razdelitel(e_PC,2,false);
   Razdelitel(e_PN,2,false);
   Razdelitel(e_PD,2,false);
   Razdelitel(e_PK,2,false);
   Razdelitel(e_PRICE0,2,false);
   Razdelitel(e_PRICE1,2,false);
   try

   //DBase

   fBase:= TwBase.Create(Sender);
    //DBGrid
   // DBGrid.Add(TwDBGrid.Create(wFormID, DBGridPriceFields,nil,'PRICEFIELD',['ID','NAME','FORMULA','PRIORITY','FCLOSE'],[],[])); // инициализация DBGrid

    fGrid:= TwDBGrid.Create(fBase,DBGridPriceFields,'select ID, NAME, FORMULA, PRIORITY, FCLOSE FROM PRICEFIELD WHERE FCLOSE=0 ORDER BY PRIORITY');

    //_DBGrid.OrderBy:= 'PRIORITY';
    fGrid.Fill();

    DBGridPriceFields.DataSource.OnDataChange:=@_onDataChange;// присваем свой обработчик изменений датасету
    DBGridPriceFields.DataSource.DataSet.First;

    // создаем формулу для расчета цен

    fFormula:= TFormula.Create(self);

   wLog('FmPriceFields','Инициализация формы успешно завершена.');
   except
     on E: Exception do
     begin
         SetStatus('Сбой инициализации формы.');
         wLog('FmPriceFields','Ошибка [FmCreate]: "' + E.Message + '"');
         wLog('FmPriceFields','Сбой инициализации формы.');
         ShowMessage('FmNomenclatureEdit [FmCreate]: "' + E.Message + '"');

      end;
   end;
end;

procedure TFmPriceFields.e_NameChange(Sender: TObject);
begin
  _SetChanged();// устанавливаем флаг наличия изменений
end;

procedure TFmPriceFields.e_NameEnter(Sender: TObject);
begin
  (Sender as TEdit).Color:=clSkyBlue; //clDefault    clSkyBlue
end;

procedure TFmPriceFields.e_NameExit(Sender: TObject);
begin
  (Sender as TEdit).Color:=clDefault; //clDefault    clSkyBlue
end;

procedure TFmPriceFields.e_PCKeyPress(Sender: TObject; var Key: char);
begin
  key:=FilterSimvol(Sender,Key,true);
end;

procedure TFmPriceFields.c_PRICEClick(Sender: TObject);
begin
  CalcOpen(e_PRICE,dCalc);
end;

procedure TFmPriceFields.c_PNClick(Sender: TObject);
begin
  CalcOpen(e_PN,dCalc);
end;

procedure TFmPriceFields.c_PRICE0Click(Sender: TObject);
begin
  CalcOpen(e_PRICE0,dCalc);
end;

procedure TFmPriceFields.c_PRICE1Click(Sender: TObject);
begin
  CalcOpen(e_PRICE1,dCalc);
end;

procedure TFmPriceFields.c_PMClick(Sender: TObject);
begin
  CalcOpen(e_PM,dCalc);
end;

procedure TFmPriceFields.c_PDClick(Sender: TObject);
begin
  CalcOpen(e_PD,dCalc);
end;

procedure TFmPriceFields.c_PKClick(Sender: TObject);
begin
  CalcOpen(e_PK,dCalc);
end;

procedure TFmPriceFields.c_PCClick(Sender: TObject);
begin
  CalcOpen(e_PC,dCalc);
end;

procedure TFmPriceFields.btnSaveClick(Sender: TObject);
var
  _GridDataset: TDataSet;
  _Name, _Formula :string;
  _Priority, _Closed: integer;
  _BookMark: TBookMark;
  _ID: integer;
begin
  if _FormulaFiled then
  begin
    ShowMessage('Формула содержит ошибки! Сохранение отменено.');
   exit;
  end;
  if not Changed then
  begin
     ShowMessage('Изменений не было. Сохранение отменено.');
     exit;
  end;

   _GridDataset:= DBGridPriceFields.DataSource.DataSet;
   _ID:= _GridDataset.FieldByName('ID').AsInteger;
   _Name:= e_Name.Text;
   _Formula:= m_Formula.Text;
   _Priority:=e_Priority.Value;
   if cb_Closed.Checked then _Closed:= 1 else _Closed:= 0;
   try
    fBase.SQLUpdate('PRICEFIELD',['NAME','FORMULA','PRIORITY','FCLOSE'],[_Name,_Formula,_Priority,_Closed],'ID='+IntTOStr(_ID));
   finally
    _BookMark:= _GridDataset.Bookmark;
    _GridDataset.Close;
    _GridDataset.Open;
    if _GridDataset.RecordCount>0 then
        _GridDataset.Bookmark:= _BookMark;
    _SetChanged(false);
   end;

end;

procedure TFmPriceFields.btnCancelClick(Sender: TObject);
begin

end;

procedure TFmPriceFields.cb_ClosedChange(Sender: TObject);
begin
  _SetChanged();// устанавливаем флаг наличия изменений
end;

procedure TFmPriceFields.e_PRICEChange(Sender: TObject);
begin
    with Sender as TEdit do
    begin
      CheckClear(Sender,2,Text);
    end;
end;

procedure TFmPriceFields.e_PRICEEnter(Sender: TObject);
begin
    (Sender as TEdit).Color:=clSkyBlue; //clDefault    clSkyBlue
  Razdelitel(Sender,2,true);
end;

procedure TFmPriceFields.e_PRICEExit(Sender: TObject);
begin
    (Sender as TEdit).Color:=clDefault; //clDefault    clSkyBlue
  Razdelitel(Sender,2,false);
  _resultRefrash();
end;

procedure TFmPriceFields.e_PRICEKeyPress(Sender: TObject; var Key: char);
begin
  key:=FilterSimvol(Sender,Key,false);
end;

procedure TFmPriceFields.e_PriorityChange(Sender: TObject);
begin
  _SetChanged();// устанавливаем флаг наличия изменений
end;

procedure TFmPriceFields.e_PriorityEnter(Sender: TObject);
begin
  (Sender as TSpinEdit).Color:=clSkyBlue; //clDefault    clSkyBlue
end;

procedure TFmPriceFields.e_PriorityExit(Sender: TObject);
begin
  (Sender as TSpinEdit).Color:=clDefault; //clDefault    clSkyBlue
end;

procedure TFmPriceFields.e_StockChange(Sender: TObject);
begin
  _resultRefrash();
end;

procedure TFmPriceFields.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
      if Changed then
      if MessageDlg('Закрыть без сохранения?',mtConfirmation, mbOKCancel, 0) = mrOK
       then
         ModalResult:= mrCancel
       else
         CanClose:= false

end;

procedure TFmPriceFields.FormDestroy(Sender: TObject);
begin
  try
   wLog('FmPriceFields','Выгрузка формы...');

   fGrid.Destroy();
   fBase.Destroy();
   fFormula.Destroy();


  wLog('FmPriceFields','Выгрузка формы успешно завершена.');

  except
    on E: Exception do
    begin
        SetStatus('Сбой выгрузки формы: Каталог.');
        wLog('FmPriceFields','Ошибка [FmDestroy]: "' + E.Message + '"');
        wLog('FmPriceFields','Сбой выгрузки плагина.');
        ShowMessage('Ошибка [FmDestroy]: "' + E.Message + '"');
     end;
  end;
end;

procedure TFmPriceFields.FormShow(Sender: TObject);
begin
  _resultRefrash();
  _formulaChange();
  _FormShow:= true;
end;

procedure TFmPriceFields.gbMainClick(Sender: TObject);
begin

end;

procedure TFmPriceFields.lbP0Click(Sender: TObject);
begin

end;

procedure TFmPriceFields.m_FormulaChange(Sender: TObject);
begin
     _formulaChange();
     _SetChanged();// устанавливаем флаг наличия изменений
end;

procedure TFmPriceFields.m_FormulaExit(Sender: TObject);
begin
  _resultRefrash();
end;

procedure TFmPriceFields.tbPriceFieldBtnAddClick(Sender: TObject);
var
  _ID: integer;
  _GridDataset: TDataSet;
begin
  _GridDataset:= DBGridPriceFields.DataSource.DataSet;
  try
    _ID:= fBase.SQLInsert('PRICEFIELD',['NAME','FORMULA','PRIORITY','FCLOSE'],['Новая цена','P+N',_GridDataset.RecordCount+1,0]);
  finally
    _GridDataset.Close;
    _GridDataset.Open;
    _GridDataset.Locate('ID',_ID,[]);
  end;
end;

procedure TFmPriceFields.tbPriceFieldBtnCopyClick(Sender: TObject);
var
  _ID: integer;
  _GridDataset: TDataSet;
  _Name, _Formula: string;
begin
  _GridDataset:= DBGridPriceFields.DataSource.DataSet;
  _Name:= _GridDataset.FieldByName('NAME').AsString;
  _Formula:= _GridDataset.FieldByName('FORMULA').AsString;
  try
    _ID:= fBase.SQLInsert('PRICEFIELD',['NAME','FORMULA','PRIORITY','FCLOSE'],[_Name,_Formula,_GridDataset.RecordCount+1,0]);
  finally
    _GridDataset.Close;
    _GridDataset.Open;
    _GridDataset.Locate('ID',_ID,[]);
  end;
end;

procedure TFmPriceFields.tbPriceFieldBtnDeleteClick(Sender: TObject);
var
  _GridDataset: TDataSet;
  _ID: integer;
  _BookMark: TBookMark;
begin

     _GridDataset:= DBGridPriceFields.DataSource.DataSet;

      if _GridDataset.RecordCount=1 then
      begin
        ShowMessage('Вы не можете удалить единственный вид цены!');
        exit;
      end;
      if MessageDlg('Удалить цену "'+_GridDataset.FieldByName('NAME').AsString+'"?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

     _BookMark:= _GridDataset.Bookmark;
     _ID:= _GridDataset.FieldByName('ID').AsInteger;
     try
       fBase.SQLDelete('PRICEFIELD','ID='+IntToStr(_ID));
       _GridDataset.Close;
       _GridDataset.Open;
       if _GridDataset.RecordCount>0 then _GridDataset.Bookmark:= _BookMark;
     except
       on E: Exception do
       begin
           SetStatus('Сбой удаления цены.');
           wLog('FmPriceFields','Ошибка [PriceFieldDelete]: "' + E.Message + '"');
           wLog('FmPriceFields','Сбой удаления.');
           ShowMessage('Ошибка [PriceFieldDelete]: "' + E.Message + '"');
        end;
     end;
end;

procedure TFmPriceFields.SetStatus(_Text: string);
begin
  wStatus(wFormID,_Text,true);
end;

procedure TFmPriceFields._onDataChange(Sender: TObject; Field: TField);
var
  _GridDataset: TDataSet;
begin
  if fGrid.FillGridNOW then exit;

  _GridDataset:= DBGridPriceFields.DataSource.DataSet;

  with _GridDataset do
  begin
       e_Name.Text:= FieldByName('NAME').AsString;
       e_Priority.Value:= FieldByName('PRIORITY').AsInteger;
       if FieldByName('FCLOSE').AsInteger = 1 then cb_Closed.Checked:=true else cb_Closed.Checked:=false;
       m_Formula.Text:= FieldByName('FORMULA').AsString;
       _SetChanged(false);
  end;

  if _FormShow then _resultRefrash();

end;

procedure TFmPriceFields._resultRefrash;
var
  _Formula: string;
  _P, _N, _M, _D, _C, _K, _S, _P0, _P1: Double;
  _arr: ArrayOfArrayVariant;
  i: Integer;
begin
   _Formula:= m_Formula.Text;
   _P:= EditValue(e_PRICE);
   //P*(1+N/100)-P
   _N:= EditValue(e_PN);
   _N:= _P*(1+_N/100)-_P;

   _M:= EditValue(e_PM);
   _D:= EditValue(e_PD);
   _C:= EditValue(e_PC);
   _C:= EditValue(e_PC);
   _C:= _C/100;
   _K:= EditValue(e_PK);
   _S:= e_Stock.Value;
   _P0:= EditValue(e_PRICE0);
   _P1:= EditValue(e_PRICE1);

   fFormula.CurrencyArray:= fBase.GetCurrencyArray();

   if Length(_Formula)>0 then
   begin
     fFormula.VarArray:=[
               {P1}	_P1,
               {P}	_P,
               {P2}	 0,
               {P3}	 0,
               {P4}	 0,
               {P5}	 0,
               {P6}	 0,
               {P7}	 0,
               {P8}	 0,
               {P9}	 0,
               {P10}	 0,
               {P0}	_P0,
               {P02}	 0,
               {P03}	 0,
               {P04}	 0,
               {P05}	 0,
               {P06}	 0,
               {P07}	 0,
               {P08}	 0,
               {P09}	 0,
               {P010}	 0,
               {S}	_S,
               {S1}	 0,
               {S2}	 0,
               {S3}	 0,
               {S4}	 0,
               {S5}	 0,
               {N}	_N,
               {M}	_M,
               {D}	_D,
               {C}	_C,
               {K}	_K
                         ];
     try
     e_result.Text:=FloatToStr(fFormula.Calc(_Formula));
     Razdelitel(e_result,2,true);
     Razdelitel(e_result,2,false);
     _FormulaFiled:= false;
     except
       on E: Exception do
       begin
            _FormulaFiled:= true;
            e_result.Text:=E.Message;
        end;
     end;
   end;

end;

procedure TFmPriceFields._formulaChange;
var
  _cnt: integer;
begin
  _cnt:= Length(m_Formula.Text);
  if _cnt>1024 then m_Formula.Color:=clRed else m_Formula.Color:=clDefault;
  gbFormula.Caption:='Формула. Введено: '+IntToStr(_cnt)+' символов из 1024';
end;

procedure TFmPriceFields._SetChanged(const AValue: boolean);
begin
  if not _FormShow then exit;
  if AValue then
     begin
       Changed:= true;
       gbEdit.Caption:='Редактор * изменен *';
     end else
     begin
      Changed:= false;
      gbEdit.Caption:='Редактор';
     end;
end;


end.

