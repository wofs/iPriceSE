unit FmMatchingEditU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Spin, Forms, Controls, Graphics, Dialogs, DBGrids,
  FmListSelectU,wFuncU, UtilsU,
  //wDBaseU, wFuncU,
  wLogU,
  wBaseU, wDBGridU, wDBTreeU, wTypesU,
  ExtCtrls, StdCtrls, Buttons, Menus;

type

  { TFmMatchingEdit }

  TFmMatchingEdit = class(TForm)
    btnCancel: TBitBtn;
    btnOpenCatalog: TSpeedButton;
    btnOK: TBitBtn;
    gbIdent: TGroupBox;
    gbName: TGroupBox;
    gbName1: TGroupBox;
    GroupBox1: TGroupBox;
    kLabel: TLabeledEdit;
    kName: TMemo;
    kOwnerNomenclatureName: TMemo;
    kScod: TLabeledEdit;
    kVendorCode: TLabeledEdit;
    kOwner: TLabeledEdit;
    Label1: TLabel;
    Panel1: TPanel;
    Panel3: TPanel;
    btnOpenPriceLists: TSpeedButton;
    spQuantInPackLeft: TSpinEdit;
    spQuantInPackRight: TSpinEdit;
    procedure btnOpenCatalogClick(Sender: TObject);
    procedure btnOpenPriceListsClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
     fFormName: string;
     fIDCatalog: integer;
     fIDCatalogGroup: integer;

     fIdMainOwner: string;
     fBase: TwBase;
     fIDOwner: integer;
     fIDPrice: integer;

  public
 //    property wIDCatalog: integer read  IDCatalog write IDCatalog;
 //    property wIDMatching: integer read  IDMatching write IDMatching;
     property IDPrice: integer read fIDPrice write fIDPrice;
     property IDCatalog: integer read fIDCatalog write fIDCatalog;
     property IDCatalogGroup: integer read  fIDCatalogGroup write fIDCatalogGroup;
     property IDOwner: integer read  fIDOwner write fIDOwner;
     property Base: TwBase read fBase write fBase;

    procedure SetStatus(_Text:string);

  end;

var
  FmMatchingEdit: TFmMatchingEdit;

implementation

{$R *.lfm}

{ TFmMatchingEdit }

procedure TFmMatchingEdit.btnOpenPriceListsClick(Sender: TObject);
var
  _IDMatching, _IDCatalog, i: Integer;
  _Form: TFmListSelect;
  _SelectedRows: ArrayOfInteger;
  _TimeStamp: Int64;
  _arr: ArrayOfArrayVariant;
  _ScodMenu: TPopupMenu;
  _ScodMenuItem: TMenuItem;
begin
  // изменение одного соответствия
  // if GridMatching.DataSource.DataSet.RecordCount=0 then exit;
   try


     _Form:= TFmListSelect.Create(self);
     _Form.Base:= fBase;
     _Form.wFormMode:=0; // PriceLists
     //_SelectedRows:= nil;
     //_Form.Base:= fBase;
     _Form.GridList.Options:=_Form.GridList.Options - [dgMultiSelect];
     _Form.wDataSetLocateField:='ID';
     _Form.wDataSetLocateValue:=fIDPrice;
     _Form.wIDTreeItem:=IDOwner;
     _Form.lbQuantity.Visible:= false;
     _Form.spQuantInPackLeft.Visible:= false;
     _Form.lbK.Visible:= false;
     _Form.spQuantInPackRight.Visible:= false;

     _arr:=nil;
     _arr:= fBase.SQLReadArr('CATALOG',['VENDORCODE','NAME','LABEL'],'ID='+IntToStr(fIDCatalog),'');
     if Assigned(_arr) then
       begin
         _Form.ListFormInit(
           IntToStr(fIDCatalog),
           VarToStr(_arr[0,0]),
           VarToStr(_arr[0,1]),
           VarToStr(_arr[0,2])
         );
       end;

     try
     _Form.ShowModal;
     finally
       if _Form.ModalResult <> mrCancel then
       begin
          _SelectedRows:= _Form.wSelectedRows;
      //    _QuantityInPacked:= _Form.wQuantityInPacked;
       end;
      _Form.Free;
     end;
     if _SelectedRows<> nil then
        begin


            _arr:= nil;
            _arr:= fBase.SQLReadArr('PL_ITEMS',['IDOWNER','ID','NAME','VENDORCODE'],'ID='+IntToStr(_SelectedRows[0]),'ID');

            //fBase.SQLUpdate('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],[_arr[0,0],_IDCatalog,_arr[0,1],_QuantityInPacked,_TimeStamp,1],'ID='+IntToStr(_IDMatching),false);

            if Length(_arr)>0 then
               begin
                IDOwner:= integer(_arr[0,0]);
                IDPrice:= integer(_arr[0,1]);
                kVendorCode.Text:=string(_arr[0,3]);
                kOwner.Text:=fBase.SQLReadArr('OWNER',['NAME'],'ID='+IntToStr(IDOwner),'ID')[0,0];
                kOwnerNomenclatureName.Text:=_arr[0,2];
               end else
                   ShowMessage('Ничего не выбрано! Соответствие не было изменено.');

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

end;

procedure TFmMatchingEdit.FormCloseQuery(Sender: TObject; var CanClose: boolean
  );
begin
  if ModalResult = mrCancel then
    if MessageDlg('Закрыть без сохранения?',mtConfirmation, mbOKCancel, 0) = mrCancel
     then
       CanClose:= false
     else
       ModalResult:= mrCancel;
end;

procedure TFmMatchingEdit.btnOpenCatalogClick(Sender: TObject);
var
  _Form: TFmListSelect;
  _SelectedRows: ArrayOfInteger;
  _arr, _Barcodes: ArrayOfArrayVariant;

begin
  // изменение одного соответствия
   try

     _Form:= TFmListSelect.Create(self);
     _Form.Base:= fBase;
     _Form.wFormMode:=1; // CATALOG
     //_Form.Base:= fBase;
     _Form.Where:= 'ID<>'+fIdMainOwner;
     _Form.GridList.Options:=_Form.GridList.Options - [dgMultiSelect];
     _Form.wDataSetLocateField:='ID';
     _Form.wDataSetLocateValue:=IDCatalog;
     _Form.wIDTreeItem:=IDCatalogGroup;
     _Form.lbQuantity.Visible:= false;
     _Form.spQuantInPackLeft.Visible:= false;
     _Form.lbK.Visible:= false;
     _Form.spQuantInPackRight.Visible:= false;

     _arr:=nil;
     _arr:= fBase.SQLReadArr('PL_ITEMS',['VENDORCODE','NAME','LABEL'],'ID='+IntToStr(fIDPrice),'');
     if Assigned(_arr) then
       begin

         _Form.ListFormInit(
           IntToStr(fIDPrice),
           VarToStr(_arr[0,0]),
           VarToStr(_arr[0,1]),
           VarToStr(_arr[0,2])
         );

       end;

    try
     _Form.ShowModal;
     finally
       if _Form.ModalResult <> mrCancel then
       begin
          _SelectedRows:= _Form.wSelectedRows;
       end;
      _Form.Free;
     end;
     if _SelectedRows<> nil then
        begin

            _arr:= nil;
            _arr:= fBase.SQLReadArr('CATALOG',['ID','LABEL','NAME'],'ID='+IntToStr(_SelectedRows[0]),'ID');
            if Length(_arr)>0 then
              begin
                IDCatalog:= integer(_arr[0,0]);

                _Barcodes:=nil;
                _Barcodes:= fBase.SQLReadArr('SELECT VSCOD FROM CTG_GET_SCOD('+IntToStr(_SelectedRows[0])+',true)');
                if Assigned(_Barcodes) then
                            kScod.Text:= VarToStr(_Barcodes[0,0]);
                kLabel.Text:= string(_arr[0,1]);
                kName.Text:= string(_arr[0,2]);
              end else
                  ShowMessage('Ничего не выбрано! Соответствие не было изменено.');

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
end;

procedure TFmMatchingEdit.FormCreate(Sender: TObject);
begin
    IDCatalogGroup:=0;
    IDCatalog:=0;
    IDPrice:=0;
    fFormName:= Self.Name;
//  SelectedRows:= nil;
 try
  wLog('FmMatchingEdit','Инициализация формы... ['+fFormName+']');

  wLog('FmMatchingEdit','Инициализация формы успешно завершена.');

   except
     on E: Exception do
     begin
         Screen.Cursor := crDefault;
         __Log.SaveLogError(E);
         SetStatus('Сбой инициализации формы.');
         wLog('FmMatchingEdit','Ошибка [FmCreate]: "' + E.Message + '"');
         wLog('FmMatchingEdit','Сбой инициализации формы.');
         ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

      end;
   end;
end;

procedure TFmMatchingEdit.FormShow(Sender: TObject);
begin
  fIdMainOwner:= fBase.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист
end;

procedure TFmMatchingEdit.SetStatus(_Text: string);
begin
  wStatus(fFormName,_Text,true);
end;

end.

