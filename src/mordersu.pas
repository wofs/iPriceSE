unit mOrdersU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, fpspreadsheet, fpspreadsheetctrls, fpsTypes, mInvoceU, mUtilsU, SysUtils, Controls, StdCtrls,
  db, ComCtrls, Forms, DBGrids, Dialogs, Menus, LaZUTF8, Clipbrd,
  IBDatabase, IBCustomDataSet, FPSExport,
  wLogU, wFuncU, wDBImportU, UtilsU, wReportU, wTProgressU,
  wBaseU, wDBGridU, wDBTreeU, wTViewerSpreadsheetU, wTypesU;

type

  { TOrders }

  TOrders = class
  private
    FFormatComboBox: TComboBox;
    fFormName: string;
    fGridInvoceNoFinded: TwDBGrid;
    fGridInvoces: TwDBGrid;
    fInvoce: TInvoce;
    fOutStringArr: ArrayOfString;
    fProgress: TProgress;
    fReport: TwReport;
    fUtils: TUtils;
    IdMainOwner: integer;
    fOwnerForm: TObject;

    fBase: TwBase;
    fImport: TwDBImport;

    fGridOrderFinded: TwDBGrid;
    fGridOrderNoFinded: TwDBGrid;

    fTreeOwners: TwDBTree;

    fFormat: TRecOrderFormat;

    procedure ChangeFileFilter;
    procedure DeleteMaching(aIDTMP: integer);
    procedure FFormatComboBox_onChange(Sender: TObject);
    procedure fTreeOwners_onSelectionChanged(Sender: Tobject);
    procedure onAddNewItem(Sender: TObject; aValue: integer);
    procedure onEndOperation(Sender: TObject);
    procedure onInvoceSumCountChanged(Sender: TObject; aSum: double; aCount: integer);
    procedure SetStatus(aText: string);
    procedure WorkBookAddOuterString(const aExportFileName: string);

  public
    constructor Create(Sender: TObject; aBase: TwBase; aTreeOwners: TTreeView; aGridOrderFinded, aGridOrderNoFinded, aGridInvoces, aGridInvoceNoFinded: TDBGrid;
      aCBXFormat: TComboBox);
    destructor Destroy();

    procedure GridOrderFindedFill();
    procedure GridOrderNoFindedFill();
    procedure GridInvoceNoFindedFill();
    procedure GridInvocesFill();
    procedure TreeOwnersFill();
    procedure UpdateMatchingsLinks(aBtnLoadTag: integer);

    procedure LoadFromFile(aFileName: string; aBtnLoadTag: integer; const aSelected: ArrayOfInteger = nil);

    procedure FindMatching(aMode: integer);
    procedure EditMatching(awGrid: TwDBGrid);
    procedure ResetMatching(awGrid: TwDBGrid);
    procedure PassedMatching();

    procedure CopyData(Sender: TObject);

    procedure OrderExportResult(aFileName: string);
    procedure OrderExportNewPosition(aFileName: string);

    procedure InvoceEdit(aId: integer);
    procedure InvoceDel(aInvoces: ArrayOfInteger);
    procedure InvoceClearRemark(aInvoces: ArrayOfInteger);
    procedure mAnalogsFill(Sender: TObject; amAnalogs: TMenuItem);
    procedure SummaryInvoce;
    procedure ExportInvoceWithSelectOwnerCode(aSelected: ArrayOfInteger);
    procedure ExportInvoceInOwnerPrice(aSelected: ArrayOfInteger);
    procedure ExportInvoceToOwnerFiles;

    property GridOrderFinded: TwDBGrid read fGridOrderFinded write fGridOrderFinded;
    property GridOrderNoFinded: TwDBGrid read fGridOrderNoFinded write fGridOrderNoFinded;
    property GridInvoceNoFinded: TwDBGrid read fGridInvoceNoFinded;
    property GridInvoces: TwDBGrid read fGridInvoces write fGridInvoces;
    property Invoce: TInvoce read fInvoce;
    property TreeOwners: TwDBTree read fTreeOwners write fTreeOwners;
    property FormatOrder: TRecOrderFormat read fFormat write fFormat;
    property FormatComboBox: TComboBox read FFormatComboBox write FFormatComboBox;
  end;

implementation
uses
  pkgOrdersU, FmListSelectU;
{ TOrders }

procedure TOrders.SetStatus(aText: string);
begin
  wStatus(fFormName, aText, true);
end;

procedure TOrders.WorkBookAddOuterString(const aExportFileName: string);
var
  _Worksheet: TsWorksheet;
  _WorkBook: TsWorkbook;
  i: integer;
begin
  _WorkBook := TsWorkbook.Create();
  _WorkBook.ReadFromFile(UTF8ToSys(aExportFileName), sfExcel8);

  try
    _WorkBook.Options:= [boBufStream];
    _Worksheet := _WorkBook.GetWorksheetByIndex(0);

    for i:= High(fOutStringArr) downto 0  do begin
       _Worksheet.InsertRow(0);
       _Worksheet.WriteText(0, 0, fOutStringArr[i]);
    end;

    _Worksheet.InsertRow(High(fOutStringArr)+1);
   _WorkBook.WriteToFile(aExportFileName, true);
  finally
    _WorkBook.Free;
  end;
end;

constructor TOrders.Create(Sender: TObject; aBase: TwBase; aTreeOwners: TTreeView; aGridOrderFinded, aGridOrderNoFinded, aGridInvoces, aGridInvoceNoFinded: TDBGrid;
  aCBXFormat: TComboBox);
var
  _SQLText: String;
begin
  fOwnerForm:= Sender;
  fFormName:= TFmOrders(fOwnerForm).Name;

  FFormatComboBox:= aCBXFormat;

  fBase:= aBase;
  TryStrToInt(fBase.ReadSettingByName('setDefaultOwner'),IdMainOwner); // считываем настройки - текущий основной прайс-лист
  fUtils:= nil;

  fImport:= TwDBImport.Create(fOwnerForm);

  _SQLText:='SELECT '
      //+' COUNT(1) as RN,'
      +' WTOI.ID,'
      +' WTOI.ORDVENDORCODE,'
      +' WTOI.ORDNAME,'
      +' WTOI.ORDLABEL,'
      +' WTOI.ORDSCOD,'
      +' WTOI.ORDUNIT, '
      +' WTOI.ORDSUM, '
      +' WTOI.ORDOWNER, '
      +' CTG.NAME CTGNAME,'
      +' CTG.ID CTGID,'
      +' CASE WHEN CTGMTH.QUANTITYINPACKING IS NOT NULL AND CTGMTH.QUANTITYINPACKING<>1 THEN 1 ELSE 0 END AS ORDQUANTITYCALCULATED,'
      // расчитыаем количество и цену, в зависимости от фасовки
      +' CASE WHEN CTGMTH.QUANTITYINPACKING IS NOT NULL THEN '
      +'   CASE WHEN CTGMTH.QUANTITYINPACKING<1 THEN TRUNC(WTOI.ORDQUANTITY*CTGMTH.QUANTITYINPACKING,0) ELSE  WTOI.ORDQUANTITY*CTGMTH.QUANTITYINPACKING END'
      +' ELSE WTOI.ORDQUANTITY END AS ORDQUANTITY,'
      +' CASE WHEN CTGMTH.QUANTITYINPACKING IS NOT NULL THEN '
      +'   WTOI.ORDSUM/(CASE WHEN CTGMTH.QUANTITYINPACKING<1 THEN TRUNC(WTOI.ORDQUANTITY*CTGMTH.QUANTITYINPACKING,0) ELSE  WTOI.ORDQUANTITY*CTGMTH.QUANTITYINPACKING END) '
      +' ELSE WTOI.ORDPRICE END AS ORDPRICE '
      //End расчитыаем количество и цену, в зависимости от фасовки
      +' FROM W_TMP_ORDERS_IMPORT WTOI '
      +' join CATALOG_MATCHING CTGMTH ON (CTGMTH.ID=WTOI.MTHID) '
      +' join CATALOG CTG ON (CTGMTH.IDCATALOG=CTG.ID) '
      +' WHERE WTOI.FPASSED=1 /*and_group_string*/';
      //+' GROUP BY 2,3,4,5,6,7';

  fGridOrderFinded:= TwDBGrid.Create(fBase,aGridOrderFinded,_SQLText);
  fGridOrderFinded.MultiSelect:= false;
  fGridOrderFinded.SortON:= false;
  fGridOrderFinded.GroupField:='WTOI.ORDOWNER';
  fGridOrderFinded.GroupArray:=nil;

  _SQLText:='SELECT '
      //+' COUNT(1) as RN,'
      +' CTG.ID CTGID, '
      +' CTG.NAME CTGNAME, '
      +' CTG.UNIT CTGUNIT, '
      +' PL.PRICECALC as CTGPRICE, '
      +' WTOI.ID,'
      +' WTOI.ORDOWNER,'
      +' WTOI.ORDVENDORCODE,'
      +' WTOI.ORDLABEL,'
      +' WTOI.ORDSCOD,'
      +' WTOI.ORDNAME,'
      +' WTOI.ORDUNIT, '
      +' WTOI.MTHID, '
      +' CTGMTH.QUANTITYINPACKING, '
      +' WTOI.ORDPRICE, '
      +' WTOI.ORDSUM '
      +' FROM W_TMP_ORDERS_IMPORT WTOI '
      +' LEFT JOIN CATALOG_MATCHING CTGMTH ON (WTOI.MTHID=CTGMTH.ID)'
      +' LEFT JOIN CATALOG CTG ON (CTG.ID=CTGMTH.IDCATALOG)'
      +' LEFT JOIN "PL_ITEMS" PL ON (PL.VENDORCODE=CTG.VENDORCODE AND PL.IDOWNER=CTG.IDOWNER)'
      +' WHERE WTOI.FPASSED=0  /*and_group_string*/ /*and_where_string*/';
      //+' GROUP BY 2,3,4,5,6,7,8,9';

  fGridOrderNoFinded:= TwDBGrid.Create(fBase,aGridOrderNoFinded,_SQLText);
  fGridOrderNoFinded.MultiSelect:= true;
  fGridOrderNoFinded.SortON:= false;
  fGridOrderNoFinded.GroupField:='WTOI.ORDOWNER';
  fGridOrderNoFinded.GroupArray:=nil;

  _SQLText:='SELECT '
      +' WTOI.ID,'
      +' WTOI.ORDOWNER,'
      +' WTOI.ORDVENDORCODE,'
      +' WTOI.ORDLABEL,'
      +' WTOI.ORDSCOD,'
      +' WTOI.ORDNAME,'
      +' WTOI.ORDUNIT, '
      +' WTOI.ORDQUANTITY, '
      +' WTOI.MTHID, '
      +' WTOI.ORDPRICE, '
      +' WTOI.ORDSUM, '
      +' IIF((SELECT RESULT FROM WTOI_CHECK_MATCHING(WTOI.ID)),1,0) MTHRESULT '
      +' FROM W_TMP_ORDERS_IMPORT WTOI '
      +' WHERE WTOI.FPASSED=0  /*and_group_string*/ /*and_where_string*/';

  fGridInvoceNoFinded:= TwDBGrid.Create(fBase,aGridInvoceNoFinded,_SQLText);
  fGridInvoceNoFinded.MultiSelect:= true;
  fGridInvoceNoFinded.SortON:= false;
  //fGridOrderNoFinded.GroupField:='WTOI.ORDOWNER';
  //fGridOrderNoFinded.GroupArray:=nil;

  fTreeOwners:= TwDBTree.Create(fBase,aTreeOwners,'OWNER','IDPARENT,NAME',[]);
  //fTreeOwners.MultiSelect:= true;
  fTreeOwners.Expanded:= true;
  fTreeOwners.Tree.OnSelectionChanged:=@fTreeOwners_onSelectionChanged;
  fTreeOwners.ShowChildrenItems:= false;

  FFormatComboBox.OnChange:=@FFormatComboBox_onChange;

  fInvoce:= TInvoce.Create(fBase, nil);
  fInvoce.onSumCountChanged:= @onInvoceSumCountChanged;
  fInvoce.onAddNewItem:= @onAddNewItem;
  fGridInvoces:= TwDBGrid.Create(fBase, aGridInvoces, fInvoce.List);
  fGridInvoces.GroupField:=fInvoce.GroupField;
  fGridInvoces.SearchEntryArray:= fInvoce.SearchEntryArray;
  fGridInvoces.SearchParticleArray:= fInvoce.SearchParticleArray;
  fGridInvoces.SearchEdit:= TFmOrders(fOwnerForm).edInvoceSearch;
  fGridInvoces.MultiSelect:= true;
  fGridInvoces.SortTitleImagesIndex:=[2,3];
  fInvoce.Grid:= fGridInvoces;
end;

destructor TOrders.Destroy();
begin
  FreeAndNil(fInvoce);
  fGridInvoces.Destroy();
  fGridInvoceNoFinded.Destroy();
  fGridOrderFinded.Destroy();
  fGridOrderNoFinded.Destroy();
  fTreeOwners.Destroy();
  fImport.Destroy();
  cmbxClearData(TFmOrders(fOwnerForm).cbx_Format);
end;

procedure TOrders.FFormatComboBox_onChange(Sender:TObject);
var
  _arr: ArrayOfArrayVariant;
begin
  _arr:=nil;
  _arr:= fBase.SQLReadArr('FORMATS',[
                              'ID',
                              'NAME',
                              'REMARK',
                              'FCONVERTLIBRE',
                              'IDFILEFORMAT',
                              'SPREADSHEET',
                              'FIRSTLINE',
                              'VENDORCODE',
                              'FNAME',
                              'UNIT',
                              'QUANTITY',
                              'FSUM',
                              'CUSTOMSDECLARATION',
                              'COUNTRY',
                              'LABEL',
                              'SCOD',
                              'FREMARK',
                              'IDCODEPAGETEXT',
                              'IDCURRENCY',
                              'CURRENCYPERCENT',
                              'IDOWNER',
                              'IDCSVDELIMITER',
                              'IDVENDORCODEVARIANT',
                              'OUTCELLTEXT',
                              'ADDRCELLTEXT',
                              'IDSTOCKVARIANT',
                              'IDPRICEVARIANT'
                              ],
                              'ID='+IntToStr(cmbxSelectID(TComboBox(Sender))),
                              ''
                              );

  if not Assigned(_arr) then exit;

  fFormat.id:= _arr[0,0];
  fFormat.Name:= VarToStr(_arr[0,1]);
  fFormat.Remark:= VarToStr(_arr[0,2]);
  if Integer(_arr[0,3])=1 then
    fFormat.ConvertWithLibre:= true
  else
    fFormat.ConvertWithLibre:= false;

  fFormat.FileFormat:=_arr[0,4];

  fFormat.fFirstLine:= _arr[0,6];
  fFormat.fSpreadsheets:= fBase.MakeArrayArrayIntegerFromString(VarToStr(_arr[0,5]),fFormat.fFirstLine);

  fFormat.fVendorcode:= _arr[0,7];
  fFormat.fName:= _arr[0,8];
  fFormat.fUnit:= _arr[0,9];
  fFormat.fQuantity:= _arr[0,10];
  fFormat.fSum:= _arr[0,11];
  fFormat.fCustomDeclaration:= _arr[0,12];
  fFormat.fCountry:= _arr[0,13];
  fFormat.fLabel:= _arr[0,14];
  fFormat.fScod:= _arr[0,15];
  fFormat.fRemark:= _arr[0,16];
  fFormat.fCodePage:= _arr[0,17];
  fFormat.fCURRENCY:= _arr[0,18];
  fFormat.fCURRENCYPERCENT:= _arr[0,19];
  fFormat.IdOwner:= _arr[0,20];
  fFormat.fIDCSVDELIMITER:= _arr[0,21];
  fFormat.fIDVENDORCODEVARIANT:= _arr[0,22];
  fFormat.fOUTCELLTEXT:= VarToStr(_arr[0,23]);
  fFormat.fADDRCELLTEXT:= VarToStr(_arr[0,24]);
  fFormat.fIDSTOCKVARIANT:= _arr[0,25];
  fFormat.fIDPRICEVARIANT:= _arr[0,26];

  ChangeFileFilter();

end;


procedure TOrders.PassedMatching();
begin
  fBase.SQLUpdate('UPDATE W_TMP_ORDERS_IMPORT WTOI SET WTOI.FPASSED=1 WHERE WTOI.MTHID IS NOT NULL AND WTOI.MTHID<>0 AND WTOI.ORDOWNER='+IntTOStr(FormatOrder.IdOwner),true);

  GridOrderFindedFill();
  GridOrderNoFindedFill();
end;

procedure TOrders.OrderExportResult(aFileName: string);
var
  _SQLText: String;
begin
  if Assigned(fUtils) then exit;

 try
  _SQLText:='SELECT '
      +' CTG.VENDORCODE AS FOURVENDORCODE,'
      +' (select VSCOD from CTG_GET_SCOD(CTG.ID,true)) AS FOURSCOD, '
      +' CTG.NAME as FNAME,'
      +' WTOI.ORDUNIT AS FUNIT, '
      // расчитыаем количество и цену, в зависимости от фасовки
      +' CASE WHEN CTGMTH.QUANTITYINPACKING IS NOT NULL THEN '
      +'   CASE WHEN CTGMTH.QUANTITYINPACKING<1 THEN TRUNC(WTOI.ORDQUANTITY*CTGMTH.QUANTITYINPACKING,0) ELSE  WTOI.ORDQUANTITY*CTGMTH.QUANTITYINPACKING END'
      +' ELSE WTOI.ORDQUANTITY END AS FQUANTITY,'
      +' CASE WHEN CTGMTH.QUANTITYINPACKING IS NOT NULL THEN '
      +'   WTOI.ORDSUM/(CASE WHEN CTGMTH.QUANTITYINPACKING<1 THEN TRUNC(WTOI.ORDQUANTITY*CTGMTH.QUANTITYINPACKING,0) ELSE  WTOI.ORDQUANTITY*CTGMTH.QUANTITYINPACKING END) '
      +' ELSE WTOI.ORDPRICE END AS FPRICE, '
      //End расчитыаем количество и цену, в зависимости от фасовки
      +' WTOI.ORDSUM AS FSUM,'
      +' WTOI.ORDVENDORCODE AS FVENDORCODE, '
      +' WTOI.ORDSCOD AS FSCOD, '
      +' WTOI.ORDLABEL AS FLABEL, '
      +' WTOI.ORDCUSTOMSDECLARATION AS FCUSTOMSDECLARATION, '
      +' WTOI.ORDCOUNTRY AS FCOUNTRY, '
      +' WTOI.ORDREMARK AS FREMARK '
      +' FROM W_TMP_ORDERS_IMPORT WTOI '
      +' join CATALOG_MATCHING CTGMTH ON (CTGMTH.ID=WTOI.MTHID) '
      +' join CATALOG CTG ON (CTG.ID=CTGMTH.IDCATALOG) '
      +' WHERE (WTOI.FPASSED=1 AND WTOI.ORDOWNER='+IntToStr(TreeOwners.SelectedItems[0])+') ';

  fUtils:= TUtils.Create(fOwnerForm,fBase);
  fUtils.SQLCustomObject:=_SQLText;
  fUtils.onEndOperation:= @onEndOperation;
  fUtils.DefaultFilterIndex:= 2;
  fUtils.DefaultTemplateFileName:= aFileName;
  fUtils.OutStringArr:= fOutStringArr;
  fUtils.ExportData([eoCustomObject]);

   except
     on E: Exception do
       begin
          wLog('Formats','Ошибка [Export]: "' + E.Message + '"');
          ShowMessage('Ошибка [Export]: "' + E.Message + '"');
       end;
   end;
end;

procedure TOrders.OrderExportNewPosition(aFileName: string);
var
  _SQLText: String;
begin
  if Assigned(fUtils) then exit;
   try
    _SQLText:='SELECT '
      +' WTOI.ORDVENDORCODE AS FVENDORCODE,'
      +' WTOI.ORDNAME as FNAME,'
      +' WTOI.ORDUNIT AS FUNIT, '
      +' WTOI.ORDQUANTITY AS FQUANTITY,'
      +' WTOI.ORDPRICE AS FPRICE,'
      +' WTOI.ORDSUM AS FSUM,'
      +' WTOI.ORDSCOD AS FSCOD, '
      +' WTOI.ORDLABEL AS FLABEL, '
      +' WTOI.ORDCUSTOMSDECLARATION AS FCUSTOMSDECLARATION, '
      +' WTOI.ORDCOUNTRY AS FCOUNTRY, '
      +' WTOI.ORDREMARK AS FREMARK '
      +' FROM W_TMP_ORDERS_IMPORT WTOI '
      +' WHERE (WTOI.FPASSED=0 AND WTOI.ORDOWNER='+IntToStr(TreeOwners.SelectedItems[0])+') ';

    fUtils:= TUtils.Create(fOwnerForm,fBase);
    fUtils.SQLCustomObject:=_SQLText;
    fUtils.onEndOperation:= @onEndOperation;
    fUtils.DefaultFilterIndex:= 2;
    fUtils.DefaultTemplateFileName:= aFileName;
    fUtils.OutStringArr:= fOutStringArr;
    fUtils.ExportData([eoCustomObject]);

   except
     on E: Exception do
       begin
          wLog('Formats','Ошибка [Export]: "' + E.Message + '"');
          ShowMessage('Ошибка [Export]: "' + E.Message + '"');
       end;
   end;
end;

procedure TOrders.InvoceEdit(aId: integer);
var
  aVisible: Boolean;
begin
  aVisible:= true;

  fInvoce.InvoceEdit(aId, nil, aVisible);
end;

procedure TOrders.InvoceDel(aInvoces: ArrayOfInteger);
var
  aVisible: Boolean;
begin
  aVisible:= true;

  fInvoce.InvoceDel(nil, aVisible);
end;

procedure TOrders.InvoceClearRemark(aInvoces: ArrayOfInteger);
begin
   fInvoce.InvoceClearRemark(aInvoces);
   GridInvocesFill();
end;

procedure TOrders.mAnalogsFill(Sender: TObject; amAnalogs: TMenuItem);
begin
  fInvoce.mAnalogsFill(Sender, amAnalogs);
end;

procedure TOrders.ResetMatching(awGrid: TwDBGrid);
var
  _IDMatching, i: Integer;
  _arr: ArrayOfArrayVariant;
  _SelectedItems: ArrayOfInteger;
  _BookMarkFinded, _BookMarkNoFinded: TBookMark;
begin
  _SelectedItems:= awGrid.SelectedRows();

  if not fBase.LongTransaction then
    fBase.LongTransaction:= true;

  try
    for i:=0 to High(_SelectedItems) do
    begin
        _arr:= fBase.SQLReadArr('W_TMP_ORDERS_IMPORT',['MTHID'],'ID='+IntToStr(_SelectedItems[i]),'');

        if Assigned(_arr) and (_arr[0,0]<>null) then
              _IDMatching:=_arr[0,0] else
              _IDMatching:= 0;

        if _IDMatching>0 then
        begin
           fBase.SQLDelete('CATALOG_MATCHING','ID='+IntToStr(_IDMatching),false);
           fBase.SQLUpdate('W_TMP_ORDERS_IMPORT',['MTHID','FPASSED'],[integer(0),integer(0)],'MTHID='+IntToStr(_IDMatching)+' AND ORDOWNER='+IntToStr(fFormat.IdOwner),false);
        end;
    end;

    awGrid.SelectAll:= false;

    fBase.SQLTransactionEnd(true);

    if awGrid = GridOrderFinded then
    begin
      _BookMarkFinded:= GridOrderFinded.Grid.DataSource.DataSet.Bookmark;
      _BookMarkNoFinded:= GridOrderNoFinded.Grid.DataSource.DataSet.Bookmark;

      GridOrderFindedFill();
      GridOrderNoFindedFill();

      if GridOrderFinded.Grid.DataSource.DataSet.RecordCount>0 then
            GridOrderFinded.Grid.DataSource.DataSet.Bookmark    := _BookMarkFinded;
      if GridOrderNoFinded.Grid.DataSource.DataSet.RecordCount>0 then
            GridOrderNoFinded.Grid.DataSource.DataSet.Bookmark := _BookMarkNoFinded;
    end;

    if awGrid = GridOrderNoFinded then GridOrderNoFindedFill();
  except
    fBase.SQLTransactionEnd(false);
    raise;
  end;
end;

procedure TOrders.ChangeFileFilter;
begin

case fFormat.FileFormat of
  1: //xls
    TFmOrders(fOwnerForm).FileNameEdit1.Filter:='Excel files (*.xls)|*.xls';
  2: //xlsx
    TFmOrders(fOwnerForm).FileNameEdit1.Filter:='Excel files (*.xlsx, *.xlsm)|*.xlsx;*.xlsm';
  3: //ods
    TFmOrders(fOwnerForm).FileNameEdit1.Filter:='LibreOffice/OpenOffice spreadsheets (*.ods)|*.ods';
  4: //csv
    TFmOrders(fOwnerForm).FileNameEdit1.Filter:='Comma-separated text files (*.csv)|*.csv;*.txt'
  else
  begin
    ShowMessage('Указан неподдерживаемый формат файла!');
    TFmOrders(fOwnerForm).btnLoad.Enabled:= false;

  end;
end;

end;

procedure TOrders.fTreeOwners_onSelectionChanged(Sender:Tobject);
var
  _DataSource: TDataSource;
begin
  fGridOrderFinded.GroupArray:=fTreeOwners.SelectedItems;
  fGridOrderNoFinded.GroupArray:=fTreeOwners.SelectedItems;

  _DataSource:= fBase.SQLReadDS('FORMATS',['ID','NAME'],'IDOWNER='+IntToStr(fTreeOwners.SelectedItems[0])+' AND IDFMTS_CATEGORY=2','NAME');
  _DataSource.DataSet.Last;
  _DataSource.DataSet.First;
  cmbxClearData(TFmOrders(fOwnerForm).cbx_Format);
  cmbxFill(TFmOrders(fOwnerForm).cbx_Format,_DataSource,['NAME','ID']);
  _DataSource.DataSet.Close;

  if TFmOrders(fOwnerForm).cbx_Format.Items.Count>0 then
  begin
    FFormatComboBox.ItemIndex:= 0;
    FFormatComboBox.Enabled:= true;

    with  TFmOrders(fOwnerForm) do begin
      FileNameEdit1.Enabled:= true;
      FileNameEdit1.Text:='';

      btnLoad.Enabled:= true;

      btnRemark.Enabled:= true;
      btnOpenWithOuterProgram.Enabled:= true;

      FFormatComboBox_onChange(FFormatComboBox);
      fImport.FormatOrder:= fFormat;
    end;


    //fBase.SQLDelete('W_TMP_ORDERS_IMPORT','',true);
  end else
  begin
    FFormatComboBox.Enabled:= false;
    TFmOrders(fOwnerForm).FileNameEdit1.Text:='';
    TFmOrders(fOwnerForm).btnLoad.Caption:= 'Анализ';
    TFmOrders(fOwnerForm).btnLoad.Enabled:= false;
    TFmOrders(fOwnerForm).btnRemark.Enabled:= false;
    TFmOrders(fOwnerForm).btnOpenWithOuterProgram.Enabled:= false;
    fFormat.Remark:='';
  end;

  with TFmOrders(fOwnerForm) do begin
    if TreeOwners.SelectedItems[0] = IdMainOwner then
      btnLoad.Tag:= 0
    else
      btnLoad.Tag:= 1;


  case btnLoad.Tag of
      0:
        begin
            btnLoad.Caption:= 'Заказ';
            btnLoad.Hint:= 'Сформировать заказ на основании накладной';
            Images16.GetBitmap(28, btnLoad.Glyph);
            pcOrders.ActivePage:= tsInvoce;
            GridInvocesFill();
            GridInvoceNoFindedFill;
        end;
      1:
        begin
            btnLoad.Caption:= 'Анализ';
            btnLoad.Hint:= 'Анализ накладной';
            Images16.GetBitmap(26, btnLoad.Glyph);
            pcOrders.ActivePage:= tsNakl;
            GridOrderFindedFill();
            GridOrderNoFindedFill();
        end;
    end;
  end;
end;

procedure TOrders.UpdateMatchingsLinks(aBtnLoadTag: integer);
begin
case aBtnLoadTag of
    1:
      begin
        fBase.SQLUpdate('update W_TMP_ORDERS_IMPORT WTOI SET MTHID =('
                +' SELECT CTGMTH.ID FROM W_TMP_ORDERS_IMPORT WTOI2'
                +' LEFT JOIN PL_ITEMS PLI ON (PLI.VENDORCODE=WTOI2.ORDVENDORCODE AND PLI.IDOWNER=WTOI2.ORDOWNER)'
                +' LEFT JOIN CATALOG_MATCHING CTGMTH ON (CTGMTH.IDPL_ITEMS=PLI.ID) WHERE WTOI.ID=WTOI2.ID)');
        fBase.SQLUpdate('UPDATE W_TMP_ORDERS_IMPORT WTOI SET WTOI.FPASSED=1 WHERE WTOI.MTHID IS NOT NULL');
      end;
  end;

end;

procedure TOrders.GridOrderFindedFill();
var
  _arr: ArrayOfArrayVariant;
  _Sum: Double;
  _SQL: String;
  _Count: Integer;
begin
  fGridOrderFinded.Fill();
  //stOrderSum
  _Sum:=0;
  _arr:=nil;
  _SQL:='SELECT SUM (ORDSUM) FROM W_TMP_ORDERS_IMPORT WHERE FPASSED=1 AND ORDOWNER='+IntToStr(fFormat.IdOwner);

  _arr:= fBase.SQLReadArr(_SQL);
  if Assigned(_arr) then
    TryStrToFloat(VarToStr(_arr[0,0]),_Sum);
   _Count:= fBase.GetRowsCount(_SQL);
    TFmOrders(fOwnerForm).stOrderSum.Caption:= CurrToStrF(_Sum, ffCurrency, 2)+' | Строк: '+IntToStr(_Count)+' ';

end;

procedure TOrders.GridOrderNoFindedFill();
begin
  fGridOrderNoFinded.Fill();
end;

procedure TOrders.GridInvoceNoFindedFill();
begin
  fGridInvoceNoFinded.Fill();
end;

procedure TOrders.GridInvocesFill();
begin
  fInvoce.GridFill(nil, true);
end;

procedure TOrders.TreeOwnersFill();
begin
   fTreeOwners.Fill();
end;

procedure TOrders.LoadFromFile(aFileName: string; aBtnLoadTag: integer; const aSelected: ArrayOfInteger);
begin
  try
    fImport.Base:= fBase;

    screen.Cursor:= crSQLWait;
    fOutStringArr:= nil;

      fImport.FormatOrder:= FormatOrder;
      fImport.Import(ftNAKL,aFileName);

      while not fImport.EndThread do
      begin
        sleep(5);
        Application.ProcessMessages;
      end;

    try

      fOutStringArr:= fImport.OutStringArr;

      UpdateMatchingsLinks(aBtnLoadTag);

      case aBtnLoadTag of
        0:
          begin
            SetStatus('Формирование заказа...');

            if fBase.SQLUpdate(Format(fInvoce.AutoAdd,['false'])) then
            begin
              ShowMessage('Заказ успешно сформирован!');
              SetStatus('Заказ успешно сформирован!');
            end else
            begin
              ShowMessage('При формировании заказа произошли ошибки!');
              SetStatus('При формировании заказа произошли ошибки!');
            end;

            GridInvocesFill();
            GridInvoceNoFindedFill;
           end;
        1:
          begin
            GridOrderFindedFill();
            GridOrderNoFindedFill();

            if (fGridOrderFinded.Grid.DataSource.DataSet.RecordCount = 0) and
               (fGridOrderNoFinded.Grid.DataSource.DataSet.RecordCount = 0) then
               begin
                 SetStatus('Накладная обработана, но данные не найдены! Возможно формат не верен.');
                 raise Exception.Create('В накладной отсутствуют данные!');
               end;
            SetStatus('Накладная успешно загружена.');
          end;
        2:
          begin
            //fImport.Base.LongTransaction:= true;

            SetStatus('Формирование заказа...');
            if fBase.SQLUpdate(Format(fInvoce.AutoAdd,['true']), true) then
            begin
              ShowMessage('Заказ успешно сформирован!');
              SetStatus('Заказ успешно сформирован!');
            end else
            begin
              ShowMessage('При формировании заказа произошли ошибки!');
              SetStatus('При формировании заказа произошли ошибки!');
            end;

            //fBase.SQLTransactionEnd(true);

            GridInvocesFill();
            GridInvoceNoFindedFill;

          end;
        3:
          begin
            SetStatus('Выгрузка ...');
            ExportInvoceWithSelectOwnerCode(aSelected);
            fBase.SQLUpdate('DELETE FROM W_TMP_ORDERS_IMPORT WHERE ORDOWNER='+IntToStr(IdMainOwner));
            GridInvoceNoFindedFill;
          end;
      end;

    finally
     screen.Cursor:= crDefault;
    end;

  except
     fBase.SQLTransactionEnd(false);
    raise;
  end;
end;

procedure TOrders.FindMatching(aMode: integer);
var
  _SelectedItems: ArrayOfInteger;
  _ArrSearchPosition, _arrResult: ArrayOfArrayVariant;
  i: Integer;
  _Long: Boolean;
begin
  _SelectedItems:= GridOrderNoFinded.SelectedRows();
  //_ArrSearchPosition
  _ArrSearchPosition:=nil;

  if not fBase.LongTransaction then
      fBase.LongTransaction:= true;

  try
    case aMode of
      1: _ArrSearchPosition:= fBase.SQLReadArr('W_TMP_ORDERS_IMPORT',['ID','ORDSCOD'],fBase.PrepareWhereString('ID',_SelectedItems),'ID');
      2: _ArrSearchPosition:= fBase.SQLReadArr('W_TMP_ORDERS_IMPORT',['ID','ORDLABEL'],fBase.PrepareWhereString('ID',_SelectedItems),'ID');
    end;



    fBase.SQLDelete('W_TMP_TBL_NEUTRALSEARCH','',false);

    SetStatus('Поиск соответствий... Это может занять немного времени...');

    MatchingAutoFind(aMode,fBase,_ArrSearchPosition,0,IdMainOwner);

    SetStatus('Поиск завершен');

   _arrResult:= fBase.SQLReadArr('SELECT WTN.ID, WTN.IDMATCHPOSITION,WTI.ORDOWNER,WTI.ORDVENDORCODE FROM W_TMP_TBL_NEUTRALSEARCH WTN '
                 +' LEFT JOIN W_TMP_ORDERS_IMPORT WTI ON (WTN.IDMATCHPOSITION=WTI.ID) ORDER BY WTN.ID');

   TFmOrders(fOwnerForm).btnPassed.Enabled:= true;
   TFmOrders(fOwnerForm).btnResetMatching.Enabled:= true;

    for i:=0 to High(_arrResult) do
    begin
      InsertMaching(fBase, _arrResult[i,0],1,
          _arrResult[i,2],
          _arrResult[i,1],
          _arrResult[i,3]);
    end;

    fBase.SQLTransactionEnd(true);
  except
    fBase.SQLTransactionEnd(false);
    raise;
  end;

  GridOrderNoFindedFill();
end;

procedure TOrders.onAddNewItem(Sender: TObject; aValue: integer);
begin
if (Sender = GridInvoceNoFinded) then
     fInvoce.GridFill(nil, true);
end;

procedure TOrders.onEndOperation(Sender: TObject);
begin
  fUtils.Free;
  fUtils:=nil;;
end;

procedure TOrders.onInvoceSumCountChanged(Sender: TObject; aSum: double; aCount: integer);
begin
  TFmOrders(fOwnerForm).stInvoceSum.Caption:='Сумма: '+CurrToStrF(aSum, ffCurrency, 2)+' | Строк: '+IntToStr(aCount)+'  ';
end;

procedure TOrders.EditMatching(awGrid: TwDBGrid);
var
  _Form: TFmListSelect;
  _DataSet: TDataSet;
  _BoookMark: TBookMark;
begin
  _DataSet:= awGrid.Grid.DataSource.DataSet;

  if not fBase.LongTransaction then
      fBase.LongTransaction:= true;

  if _DataSet.RecordCount = 0 then exit;

  try
    _BoookMark:= _DataSet.Bookmark;

    _Form:= TFmListSelect.Create(awGrid.Grid);
    _Form.Base:= fBase;
    _Form.wFormMode:=1;
    _Form.MultiSelectGrid:= false;
    _Form.wDataSetLocateField:= 'ID';
    _Form.wDataSetLocateValue:= _DataSet.FieldByName('CTGID').AsVariant;
    Screen.Cursor:=crSQLWait;
    try
      _Form.ListFormInit(
               VarToStr(_DataSet.FieldByName('CTGID').AsVariant),
               VarToStr(_DataSet.FieldByName('ORDVENDORCODE').AsVariant),
               VarToStr(_DataSet.FieldByName('ORDNAME').AsVariant),
               VarToStr(_DataSet.FieldByName('ORDLABEL').AsVariant),
               VarToStr(_DataSet.FieldByName('ORDSCOD').AsVariant));

      _Form.ShowModal;

      if _Form.ModalResult = mrOK then
      begin
        TFmOrders(fOwnerForm).btnPassed.Enabled:= true;
        TFmOrders(fOwnerForm).btnResetMatching.Enabled:= true;

         InsertMaching(fBase, _Form.wSelectedRows[0],
            _Form.spQuantInPackLeft.Value/_Form.spQuantInPackRight.Value,
            _DataSet.FieldByName('ORDOWNER').AsInteger,
            _DataSet.FieldByName('ID').AsInteger,
            _DataSet.FieldByName('ORDVENDORCODE').AsString);
      end;

      fBase.SQLTransactionEnd(true);

    finally
     Screen.Cursor:=crDefault;
     _Form.Free;

      if awGrid = GridOrderFinded then GridOrderFindedFill();
      if awGrid = GridOrderNoFinded then GridOrderNoFindedFill();

     if _DataSet.RecordCount>0 then
        _DataSet.Bookmark:= _BoookMark;

    end;
  except
     fBase.SQLTransactionEnd(false);
  end;
end;

procedure TOrders.CopyData(Sender:TObject);
begin
case TMenuItem(Sender).Name of
    'mFindedVendorCode'   : Clipboard.AsText:= GridOrderFinded.Grid.DataSource.DataSet.FieldByName('ORDVENDORCODE').AsString;
    'mNoFindedVendorCode' : Clipboard.AsText:= GridOrderNoFinded.Grid.DataSource.DataSet.FieldByName('ORDVENDORCODE').AsString;
    'mFindedName'         : Clipboard.AsText:= GridOrderFinded.Grid.DataSource.DataSet.FieldByName('ORDNAME').AsString;
    'mNoFindedName'       : Clipboard.AsText:= GridOrderNoFinded.Grid.DataSource.DataSet.FieldByName('ORDNAME').AsString;
    'mFindedLabel'        : Clipboard.AsText:= GridOrderFinded.Grid.DataSource.DataSet.FieldByName('ORDLABEL').AsString;
    'mNoFindedLabel'      : Clipboard.AsText:= GridOrderNoFinded.Grid.DataSource.DataSet.FieldByName('ORDLABEL').AsString;
    'mFindedScod'         : Clipboard.AsText:= GridOrderFinded.Grid.DataSource.DataSet.FieldByName('ORDSCOD').AsString;
    'mNoFindedScod'       : Clipboard.AsText:= GridOrderNoFinded.Grid.DataSource.DataSet.FieldByName('ORDSCOD').AsString;
  end;
end;

procedure TOrders.DeleteMaching(aIDTMP: integer);
begin
  if not fBase.LongTransaction then
      fBase.LongTransaction:= true;

   fBase.SQLUpdate('W_TMP_ORDERS_IMPORT',['MTHID'],[integer(0)],'ID='+IntToStr(aIDTMP),false);

   fBase.SQLTransactionEnd(true);
end;

procedure TOrders.SummaryInvoce;
begin
  fInvoce.ExportSummaryInvoce;
end;

procedure TOrders.ExportInvoceWithSelectOwnerCode(aSelected: ArrayOfInteger);
begin
  fInvoce.ExportInvoceWithSelectOwnerCode(aSelected);
end;

procedure TOrders.ExportInvoceInOwnerPrice(aSelected: ArrayOfInteger);
begin
  fInvoce.ExportInvoceInOwnerPrice(aSelected);
end;

procedure TOrders.ExportInvoceToOwnerFiles;
begin
  Invoce.ExportInvoceToOwnerFiles;
end;

end.

