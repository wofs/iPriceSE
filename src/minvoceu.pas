unit mInvoceU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, Controls, db, Dialogs, FmArcViewU, FmInvoceU, FmListSelectU, Forms,
  LazFileUtils, LazUTF8, Menus, mUtilsU, SysUtils, fgl, UtilsU, wBaseU,
  wDBGridU, wDBImportU, wFuncU, wReportU, wTProgressU, wTViewerSpreadsheetU,
  wTypesU, fpspreadsheetctrls, fpspreadsheet, fpsTypes, wZipperU;
type

  { TInvoce }

  TInvoce = class (TObject)
   protected
     const
       uCreateTableSQL = 'CREATE TABLE INVOCES ('+wfLineEnding
          +' ID BIGINT NOT NULL, '+wfLineEnding
          +' IDOWNER BIGINT, '+wfLineEnding
          +' IDPL_ITEMS BIGINT, '+wfLineEnding
          +' QUANTITY NUMERIC(15,2),'+wfLineEnding
          +' REMARK VARCHAR(500))';

       uCreatePK ='ALTER TABLE INVOCES '+wfLineEnding
          +' ADD CONSTRAINT PK_INVOCES '+wfLineEnding
          +' PRIMARY KEY (ID)';

       uCreateGenerator ='CREATE SEQUENCE GEN_INVOCES_ID';

       uCreateTrigger ='CREATE TRIGGER INVOCES_BI FOR INVOCES '+wfLineEnding
          +' ACTIVE BEFORE INSERT POSITION 0 '+wfLineEnding
          +' AS '+wfLineEnding
          +' BEGIN '+wfLineEnding
          +'   IF (NEW.ID IS NULL) THEN '+wfLineEnding
          +'     NEW.ID = GEN_ID(GEN_INVOCES_ID,1); '+wfLineEnding
          +' END';

       uCreateFKOwner ='ALTER TABLE INVOCES '
          +' ADD CONSTRAINT FK_INVOCES_OWNER '
          +' FOREIGN KEY (IDOWNER) '
          +' REFERENCES OWNER(ID) '
          +' ON DELETE CASCADE';

       uCreateFKPL ='ALTER TABLE INVOCES '
         +' ADD CONSTRAINT FK_INVOCES_PL '
         +' FOREIGN KEY (IDPL_ITEMS) '
         +' REFERENCES PL_ITEMS(ID) '
         +' ON DELETE CASCADE';

       uCreateProcInvoceSearch ='create or alter procedure INVOCE_SEARCH_OUR_POS_IN_PRICES ( '+wfLineEnding
         +'     OUR_IDOWNER bigint, '+wfLineEnding
         +'     OUR_VENDORCODE varchar(300), '+wfLineEnding
         +'     IGNORESTOCKPRICE boolean = false) '+wfLineEnding
         +' returns ( '+wfLineEnding
         +'     PL_IDOWNER bigint, '+wfLineEnding
         +'     PL_ID bigint) '+wfLineEnding
         +' as '+wfLineEnding
         +' declare variable ID bigint; '+wfLineEnding
         +' BEGIN '+wfLineEnding
         +'   FOR '+wfLineEnding
         +'     SELECT PL.ID FROM PL_ITEMS PL WHERE PL.IDOWNER=:OUR_IDOWNER AND PL.VENDORCODE=:OUR_VENDORCODE '+wfLineEnding
         +'     INTO :ID '+wfLineEnding
         +'   DO '+wfLineEnding
         +'   BEGIN '+wfLineEnding
         +'     if (:ID IS NOT NULL) then '+wfLineEnding
         +'     begin '+wfLineEnding
         +'         if (IGNORESTOCKPRICE) then '+wfLineEnding
         +'         begin '+wfLineEnding
         +'             SELECT AP.IDOWNER, AP.ID  FROM ANALIS_SEL_ALL_ANALOG(:ID, true) AP ORDER BY AP.PRICE ASC ROWS 1 '+wfLineEnding
         +'             INTO :PL_IDOWNER, '+wfLineEnding
         +'                  :PL_ID; '+wfLineEnding
         +'         end else '+wfLineEnding
         +'         begin '+wfLineEnding
         +'             SELECT AP.IDOWNER, AP.ID  FROM ANALIS_SEL_ALL_ANALOG(:ID, true) AP WHERE AP.PRICE>0 AND AP.STOCK>0 ORDER BY AP.PRICE ASC ROWS 1 '+wfLineEnding
         +'             INTO :PL_IDOWNER, '+wfLineEnding
         +'                  :PL_ID; '+wfLineEnding
         +'         end '+wfLineEnding
         +'     end '+wfLineEnding
         +'     SUSPEND; '+wfLineEnding
         +'   END '+wfLineEnding
         +' END';

       uCreateProcInvoceAutoAdd = 'create or alter procedure INVOCE_ADD_ALANOG_FROM_PRICE ( '+wfLineEnding
         +'     IGNORESTOCKPRICE boolean = false) '+wfLineEnding
         +' as '+wfLineEnding
         +' declare variable IDOWNER bigint; '+wfLineEnding
         +' declare variable IDINVOCE bigint; '+wfLineEnding
         +' declare variable ID bigint; '+wfLineEnding
         +' declare variable VENDORCODE varchar(300); '+wfLineEnding
         +' declare variable QUANTITY numeric(15,2); '+wfLineEnding
         +' declare variable PL_IDOWNER bigint; '+wfLineEnding
         +' declare variable PL_ID bigint; '+wfLineEnding
         +' BEGIN '+wfLineEnding
         +' FOR '+wfLineEnding
         +' SELECT ORD.ID, ORD.ORDOWNER, ORD.ORDVENDORCODE, ORD.ORDQUANTITY FROM W_TMP_ORDERS_IMPORT ORD '+wfLineEnding
         +' INTO :ID, '+wfLineEnding
         +'      :IDOWNER, '+wfLineEnding
         +'      :VENDORCODE, '+wfLineEnding
         +'      :QUANTITY '+wfLineEnding
         +' DO '+wfLineEnding
         +' BEGIN '+wfLineEnding
         +'     SELECT INVS.PL_IDOWNER, INVS.PL_ID FROM INVOCE_SEARCH_OUR_POS_IN_PRICES(:IDOWNER, :VENDORCODE, :IGNORESTOCKPRICE) INVS '+wfLineEnding
         +'     INTO :PL_IDOWNER, '+wfLineEnding
         +'          :PL_ID; '+wfLineEnding
         +'          '+wfLineEnding
         +'     if (PL_ID IS NOT NULL) then '+wfLineEnding
         +'     begin '+wfLineEnding
         +'        INSERT INTO INVOCES (IDOWNER, IDPL_ITEMS, QUANTITY, REMARK) VALUES (:PL_IDOWNER, :PL_ID, :QUANTITY, ''auto'') RETURNING ID INTO :IDINVOCE; '+wfLineEnding
         +'        '+wfLineEnding
         +'        UPDATE W_TMP_ORDERS_IMPORT ORD SET ORD.FPASSED=1, ORD.MTHID=:IDINVOCE WHERE ORD.ID=:ID; '+wfLineEnding
         +'     end '+wfLineEnding
         +' END '+wfLineEnding
         +'      '+wfLineEnding
         +' END';

       uGetSummaryInvoce = 'SELECT  '
         +' CTG.VENDORCODE, '
         +' PL.VENDORCODE, '
         +' PL.LABEL, '
         +' PL.NAME, '
         +' PL.UNIT, '
         +' CAST(INV.QUANTITY AS INTEGER) QUANTITY, '
         +' PL.PRICECALC PRICEPL, '
         +' (PL.PRICECALC*INV.QUANTITY) FSUM, '
         +' INV.REMARK, '
         +' OWN.NAME OWNERNAME '
         +' FROM INVOCES INV '
         +' INNER JOIN PL_ITEMS PL ON (PL.ID=INV.IDPL_ITEMS) '
         +' INNER JOIN OWNER OWN ON (OWN.ID=INV.IDOWNER) '
         +' LEFT OUTER JOIN CATALOG_MATCHING MTH ON (MTH.IDPL_ITEMS=PL.ID) '
         +' LEFT OUTER JOIN CATALOG CTG ON (CTG.ID=MTH.IDCATALOG) '
         +' ORDER BY INV.IDOWNER, PL.NAME';

       uGetInvoceWithSelectOwnerCode = 'SELECT '
          +' ORDVENDORCODE, '
          +' ORDLABEL, '
          +' ORDSCOD, '
          +' ORDNAME, '
          +' ORDUNIT, '
          +' CAST(ORDQUANTITY AS INTEGER) ORDQUANTITY, '
          +' VENDORCODE, '
          +' ORDREMARK, '
          +' VENDORPRICE, '
          +' OWNERSEARCH '
          +' FROM W_TMP_ORDERS_GET_MTH';

       uGetListSQL = 'SELECT INV.ID, INV.QUANTITY, '
         +' INV.IDOWNER, '
         +' INV.IDPL_ITEMS, '
         +' PL.NAME, '
         +' PL.UNIT, '
         +' PL.PRICECALC PRICEPL, '
         +' PL.VENDORCODE, '
         +' PL.LABEL, '
         +' (PL.STOCK+PL.STOCK2+PL.STOCK3+PL.STOCK4+PL.STOCK5) QUANTITYPL, '
         +' (PL.PRICECALC*INV.QUANTITY) FSUM, '
         +' (SELECT VSCOD FROM PL_GET_SCOD(PL.ID, true)) AS SCOD,'
         +' INV.REMARK, '
         +' OWN.NAME OWNERNAME, '
         +' IIF((SELECT ID FROM ANALIS_SEL_ALL_ANALOG(INV.IDPL_ITEMS,true) WHERE IDOWNER<>%d ROWS 1) IS NOT NULL,1,0) MTHRESULT '
         +' FROM INVOCES INV '
         +' INNER JOIN PL_ITEMS PL ON (PL.ID=INV.IDPL_ITEMS) '
         +' INNER JOIN OWNER OWN ON (OWN.ID=INV.IDOWNER) '
         +' WHERE (1=1) /*and_group_string*/ /*and_search_string*/ '
         +' ORDER BY INV.IDOWNER, PL.NAME';

       uGetSumSQL ='SELECT '
         +' SUM((PL.PRICECALC*INV.QUANTITY)) '
         +' FROM INVOCES INV '
         +' INNER JOIN PL_ITEMS PL ON (PL.ID=INV.IDPL_ITEMS) '
         +' INNER JOIN OWNER OWN ON (OWN.ID=INV.IDOWNER) WHERE (1=1)';

       uGroupField = 'INV.IDOWNER';
       uSearchEntryArray = 'PL.NAME,INV.REMARK, PL.LABEL';
       uSearchParticleArray ='OWN.NAME, PL.VENDORCODE';

       uGetItemSQL = 'SELECT INV.ID, INV.QUANTITY, '
         +' INV.IDOWNER, '
         +' INV.IDPL_ITEMS, '
         +' PL.NAME, '
         +' PL.UNIT, '
         +' PL.LABEL, '
         +' PL.VENDORCODE, '
         +' PL.PRICECALC PRICEPL, '
         +' (PL.PRICECALC*INV.QUANTITY) FSUM, '
         +' INV.REMARK '
         +' FROM INVOCES INV '
         +' INNER JOIN PL_ITEMS PL ON (PL.ID=INV.IDPL_ITEMS) '
         +' WHERE INV.ID=:ID';

       uNewSQL = 'INSERT INTO INVOCES (IDOWNER, IDPL_ITEMS, QUANTITY, REMARK)'
         +' VALUES (:IDOWNER, :IDPL_ITEMS, :QUANTITY, :REMARK) RETURNING ID';

       uEditSQL = 'UPDATE INVOCES SET IDOWNER=:IDOWNER, IDPL_ITEMS=:IDPL_ITEMS, QUANTITY=:QUANTITY, REMARK=:REMARK WHERE ID=:ID RETURNING ID';
       uDelSQL = 'DELETE FROM INVOCES WHERE ID=:ID';

       uAutoAdd = 'EXECUTE PROCEDURE INVOCE_ADD_ALANOG_FROM_PRICE(%s)';

       uGetOwnersFromInvoce = 'SELECT DISTINCT(INV.IDOWNER), OWN.NAME FROM INVOCES INV '
         +' INNER JOIN OWNER OWN ON (OWN.ID=INV.IDOWNER)';

       uGetInvoceToExportIntoOwnerFiles = 'SELECT  '
         +' PL.VENDORCODE, '
         +' PL.LABEL, '
         +' PL.NAME, '
         +' PL.UNIT, '
         +' CAST(INV.QUANTITY AS INTEGER) QUANTITY, '
         +' (SELECT VSCOD FROM PL_GET_SCOD(PL.ID,true)) SCOD ,'
         +' INV.REMARK '
         +' FROM INVOCES INV '
         +' INNER JOIN PL_ITEMS PL ON (PL.ID=INV.IDPL_ITEMS) '
         +' ORDER BY INV.IDOWNER, PL.NAME';

       //uFindPositions = 'SELECT';
   var
     fBase:TwBase;
   private
     fEventBlock: Boolean;
     IdMainOwner: Int64;
     fGrid: TwDBGrid;
     fonAddNewItem: TIntValueNotify;
     fonSumCountChanged: TSumCountNotify;
     fProgress: TProgress;
     fReport: TwReport;

     function ConvertIdINVToIdPRICE(aSelectedItems: ArrayOfInteger): ArrayOfInteger;
     function ConvertIdWTOIToIdPRICE(aSelectedItems: ArrayOfInteger): ArrayOfInteger;
     procedure EventRegister;
     function GetAutoAdd: string;
     function GetCreateTable: ArrayOfString;
     function GetDelSQL: string;
     function GetEditSQL: string;
     function GetGetItemSQL: string;
     function GetGetListSQL: string;
     function GetGroupField: string;
     function GetInvoceToExportIntoOwnerFiles: string;
     function GetInvoceWithSelectOwnerCode: string;
     function GetItemAdjSQL(aWhere: string): string;
     function GetNewSQL: string;
     function GetOwnersFromInvoce: string;
     function GetSearchEntryArray: ArrayOfString;
     function GetSearchParticleArray: ArrayOfString;
     function ArrayFromString(aText: string): ArrayOfString;
     function GetSummaryInvoce: string;
     procedure InvoceChangeItem(aIdPL, aIdInvoce: integer);
     procedure mAnalogsClick(Sender: TObject);
     procedure onCbInvoiceChange(Sender: TObject; aValue: integer);
     procedure onEndThread(Sender: TObject);
     procedure OnEventAlert(Sender: TObject; EventName: string;
       EventCount: longint; var CancelAlerts: Boolean);
     procedure onGridFill(Sender: TObject);
     procedure onProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
     procedure onProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
     procedure onStopForce(Sender: TObject);

     procedure SetfGrid(aValue: TwDBGrid);
     function WriteWhere(aSQL: string; aWhere: string): string;

   public
     constructor Create(aBase: TwBase; aGrid: TwDBGrid);
     destructor Destroy; override;

     property Grid: TwDBGrid read fGrid write SetfGrid;

     property List: string read GetGetListSQL;

     property SummaryInvoce: string read GetSummaryInvoce;
     property InvoceWithSelectOwnerCode: string read GetInvoceWithSelectOwnerCode;

     property Item: string read GetGetItemSQL;
     property ItemAdj[aWhere:string]: string read GetItemAdjSQL;

     property New: string read GetNewSQL;
     property Edit: string read GetEditSQL;
     property Del: string read GetDelSQL;

     property AutoAdd: string read GetAutoAdd;

     property OwnersFromInvoce: string read GetOwnersFromInvoce;
     property InvoceToExportIntoOwnerFiles: string read GetInvoceToExportIntoOwnerFiles;

     property GroupField: string read GetGroupField;
     property SearchEntryArray: ArrayOfString read GetSearchEntryArray;
     property SearchParticleArray: ArrayOfString read GetSearchParticleArray;

     property CreateTable: ArrayOfString read GetCreateTable;
     property onSumCountChanged: TSumCountNotify read fonSumCountChanged write fonSumCountChanged;
     property onAddNewItem: TIntValueNotify read fonAddNewItem write fonAddNewItem;
     property EventBlock: Boolean read fEventBlock write fEventBlock;

     procedure InvoceEdit(aId: integer; const aTreeSelectedItems: ArrayOfInteger=nil; const aFill: boolean=false);
     procedure GridFill(aTreeSelectedItems: ArrayOfInteger; aFill: boolean);
     procedure InvoceDel(aTreeSelectedItems: ArrayOfInteger; aFill: boolean);
     procedure InvoceClearRemark(aInvoces: ArrayOfInteger);
     procedure InvoceAdd(aIdPricePosition: Int64; const aTreeSelectedItems: ArrayOfInteger=nil; const aFill: boolean=false);
     procedure InvoceAddNewItem(aGridInvoce: TwDBGrid; aFill: boolean);
     procedure InvoceAddNewItem(aGridNoFinded: TwDBGrid; aTreeSelectedItems: ArrayOfInteger; aFill: boolean);

     procedure mAnalogsFill(Sender: TObject; amAnalogs: TMenuItem);
     procedure ExportSummaryInvoce;
     procedure ExportInvoceWithSelectOwnerCode(aSelected: ArrayOfInteger);
     procedure ExportInvoceInOwnerPrice(aSelected: ArrayOfInteger);
     procedure ExportInvoceToOwnerFiles;

     procedure GetPositionAnalog(aSelectedItems: ArrayOfInteger; aFinded: boolean);

     procedure GetPriceFile(aOwnerID: int64; var aRecPriceFormat: TRecPriceFormat);
  end;

implementation

{ TInvoce }

function TInvoce.WriteWhere(aSQL: string; aWhere: string): string;
var
  _PosWhere, _PosWhereRes: PtrInt;
  _SQL: String;
begin

   _SQL:= UTF8UpperCase(aSQL);

   _PosWhere:=1;
   while _PosWhere>0 do begin
   _PosWhere:= UTF8Pos('WHERE',_SQL,_PosWhere+1);
   if (_PosWhere<>0) then
      _PosWhereRes:= _PosWhere+Length('WHERE');
  end;

   UTF8Delete(_SQL,_PosWhereRes,Length(_SQL)-_PosWhereRes+1);

   UTF8Insert(' '+aWhere+' ',_SQL,_PosWhereRes);

   Result:= _SQL;

end;

function TInvoce.ArrayFromString(aText: string): ArrayOfString;
var
  i:     integer;
  _List: TStringList;
begin
  Result := nil;

  if Length(aText)>0 then
  begin
    for i:= Length(aText) downto 1 do
    if not (aText[i] in ['<','>','=','-','+','*','/','(',')','.',',', '0'..'9', 'a'..'z', 'A'..'Z']) then Delete(aText, i, 1);

    _List := TStringList.Create;

    try
      ExtractStrings([','],[' '],PChar(aText),_List);

      SetLength(Result, _List.Count);

      for i := 0 to _List.Count-1 do
        Result[i] := _List[i];
    finally
      FreeAndNil(_List);
    end;
  end;
end;

function TInvoce.GetSummaryInvoce: string;
begin
  Result:= uGetSummaryInvoce;
end;

procedure TInvoce.mAnalogsClick(Sender: TObject);
var
  _ID: integer;
  _IDInvoce: LongInt;
begin

  _ID:= StrToInt(ReplaceStr(TMenuItem(Sender).Name,'m',''));
  _IDInvoce:= fGrid.FieldValue['ID'].AsInteger;

  InvoceChangeItem(_ID, _IDInvoce);
end;

procedure TInvoce.onCbInvoiceChange(Sender: TObject; aValue: integer);
var
  _Form: TFmInvoce;
  _DS: TDataSource;
begin
  _Form:= nil;
  _Form:= TFmInvoce(Sender);

  _DS:= nil;
  _DS:= fBase.SQLItemGetDS(Item,[aValue]);

  if Assigned(_DS) and (_DS.DataSet.RecordCount >0) then
  begin
    _Form.edQuantity.Value:= _DS.DataSet.FieldByName('QUANTITY').AsInteger;
    _Form.edRemark.Text:= _DS.DataSet.FieldByName('REMARK').AsString;
  end
  else
  begin
    _Form.edQuantity.Value:= 1;
    _Form.edRemark.Text:= '';
  end;
end;

procedure TInvoce.onEndThread(Sender: TObject);
begin
 try
   fProgress.ForceClose;

     if not fReport.Result then
        raise Exception.Create('Во время создания отчета произошла ошибка!');

 except
   on E: Exception do begin
      MessageDlg(E.Message,mtError, [mbOK], 0);
   end;
 end
end;

procedure TInvoce.OnEventAlert(Sender: TObject; EventName: string;
  EventCount: longint; var CancelAlerts: Boolean);
var
  fBookmark: TBookMark;
begin
  if EventBlock or not Assigned(Grid) then
  begin
    CancelAlerts := True;
    Exit;
  end;
  case EventName of
    'INVOCES_Change':
    begin
      fBookmark := Grid.Bookmark;
      Grid.Fill();
      Grid.Bookmark := fBookmark;
    end;
  end;
end;

procedure TInvoce.onGridFill(Sender: TObject);
var
  aSQL, aWhere: String;
  aArr: ArrayOfArrayVariant;
  aSum: Double;
  aCount: Integer;
begin
  aSQL:= uGetSumSQL;


  aWhere:= fBase.GetWhere(fGrid.SQL);

  aSQL:= fBase.WriteWhere(aSQL,aWhere, true);

  aArr:= fBase.SQLReadArr(aSQL);
  if Assigned(aArr) then
    TryStrToFloat(VarToStr(aArr[0,0]),aSum);

   aCount:= fBase.GetRowsCount(aSQL);

   if Assigned(onSumCountChanged) then onSumCountChanged(Self, aSum, aCount);
end;

procedure TInvoce.onProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
begin
  fProgress.InitBar(aProgressBarName, aValue);
end;

procedure TInvoce.onProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
begin
  fProgress.SetBar(aProgressBarName, aValue);
end;

procedure TInvoce.onStopForce(Sender: TObject);
begin
  fReport.Stop();
end;

procedure TInvoce.SetfGrid(aValue: TwDBGrid);
begin
  if fGrid=aValue then Exit;
  fGrid:=aValue;

  if Assigned(fGrid) then
    fGrid.onFill:= @onGridFill;
end;

function TInvoce.GetAutoAdd: string;
begin
  Result:= uAutoAdd;
end;

function TInvoce.GetCreateTable: ArrayOfString;
begin
  SetLength(Result,7);

  Result[0]:= uCreateTableSQL;
  Result[1]:= uCreateGenerator;
  Result[2]:= uCreateTrigger;
  Result[3]:= uCreateFKOwner;
  Result[4]:= uCreateFKPL;
  Result[5]:= uCreateProcInvoceSearch;
  Result[6]:= uCreateProcInvoceAutoAdd;
end;

function TInvoce.GetDelSQL: string;
begin
  Result:= uDelSQL;
end;

function TInvoce.GetEditSQL: string;
begin
  Result:= uEditSQL;
end;

function TInvoce.GetGetItemSQL: string;
begin
  Result:= uGetItemSQL;
end;

function TInvoce.GetGetListSQL: string;
begin
  Result:= Format(uGetListSQL,[IdMainOwner]);
end;

function TInvoce.GetGroupField: string;
begin
  Result:= uGroupField;
end;

function TInvoce.GetInvoceToExportIntoOwnerFiles: string;
begin
  Result:= uGetInvoceToExportIntoOwnerFiles;
end;

function TInvoce.GetInvoceWithSelectOwnerCode: string;
begin
  Result:= uGetInvoceWithSelectOwnerCode;
end;

function TInvoce.GetItemAdjSQL(aWhere: string): string;
begin
  Result:= WriteWhere(Item,aWhere);
end;

function TInvoce.GetNewSQL: string;
begin
  Result:= uNewSQL;
end;

function TInvoce.GetOwnersFromInvoce: string;
begin
  Result:= uGetOwnersFromInvoce;
end;

function TInvoce.GetSearchEntryArray: ArrayOfString;
begin
  Result:= ArrayFromString(uSearchEntryArray);
end;

function TInvoce.GetSearchParticleArray: ArrayOfString;
begin
  Result:= ArrayFromString(uSearchParticleArray);
end;

constructor TInvoce.Create(aBase: TwBase; aGrid: TwDBGrid);
begin
  fBase:= aBase;
  fGrid:= aGrid;
  IdMainOwner:= fBase.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист

  try
    inherited Create;
  finally
      EventRegister();
  end;
end;

destructor TInvoce.Destroy;
begin
  inherited Destroy;
end;

procedure TInvoce.EventRegister();
begin
  if fBase.RegisterEvents(['INVOCES_Change']) then
    fBase.EventDB.OnEventAlert := @OnEventAlert;
end;

procedure TInvoce.InvoceAdd(aIdPricePosition: Int64; const aTreeSelectedItems: ArrayOfInteger; const aFill: boolean);
var
  _Form: TFmInvoce;
  _arr: ArrayOfArrayVariant;
  _ID: String;
  _Result, _InvoceId: Integer;
  _DS: TDataSource;
begin

  _Form:= TFmInvoce.Create(nil);
  try
    _Form.Caption:='Добавление позиции в заказ';
    _ID:= IntToStr(aIdPricePosition);

    _arr:= fBase.SQLReadArr('SELECT PL.ID, '
        +' PL.NAME, '
        +' PL.UNIT, '
        +' PL.VENDORCODE, '
        +' (SELECT VSCOD FROM PL_GET_SCOD('+_ID+',true)) SCOD, '
        +' PL.LABEL, '
        +' PL.PRICECALC, '
        +' PL.IDOWNER '
        +' FROM PL_ITEMS PL '
        +' WHERE PL.ID='+_ID+'');

    _Form.edName.Text            := VarToStr(_arr[0,1]);
    _Form.edUnit.Text            := VarToStr(_arr[0,2]);
    _Form.edVendorCode.Text      := VarToStr(_arr[0,3]);
    _Form.edScod.Text            := VarToStr(_arr[0,4]);
    _Form.edLabel.Text           := VarToStr(_arr[0,5]);
    _Form.edPrice.Text           := VarToStr(_arr[0,6]);
    _Form.edQuantity.Value       := 1;
    _Form.onCbInvoiceChange:= @onCbInvoiceChange;

    _DS:= fBase.SQLItemGetDS(ItemAdj['INV.IDPL_ITEMS = :IDPL_ITEMS'],[integer(_arr[0,0])]);

    if Assigned(_DS) then
      begin
        cmbxFill(_Form.cbInvoced,_DS,['NAME','REMARK','ID']);
        _DS.DataSet.Close;
      end
      else
        _Form.cbInvoced.Enabled:= false;

      _Form.cbInvoced.Items.AddObject('Создать новую позицию...', TwData.Create(-1));
      _Form.cbInvoced.ItemIndex:= 0;
      onCbInvoiceChange(_Form,cmbxSelectID(_Form.cbInvoced));
      _Form.edQuantity.OnChange(_Form.edQuantity);

    _Form.ShowModal;

    if _Form.ModalResult = mrOK then
      begin
        EventBlock:= true;
        _InvoceId:= cmbxSelectID(_Form.cbInvoced);
        _Result:=0;

        if _Form.edQuantity.Value=0 then
          fBase.SQLItemUpdate(Del,[_InvoceId])
        else
          if (_InvoceId=-1) then
            _Result:= fBase.SQLItemUpdate(New,[_arr[0,7], _arr[0,0], _Form.edQuantity.Value, _Form.edRemark.Text],true)
          else
            _Result:= fBase.SQLItemUpdate(Edit,[_arr[0,7], _arr[0,0], _Form.edQuantity.Value, _Form.edRemark.Text, _InvoceId],true);
      end;

    if Assigned(fGrid) then
      begin
        GridFill(aTreeSelectedItems, aFill);
        if Assigned(fGrid.Grid.DataSource) and (fGrid.Grid.DataSource.DataSet.RecordCount>0) then
          fGrid.Grid.DataSource.DataSet.Locate('ID',_Result,[]);
      end;
  finally
     EventBlock:= false;
     cmbxClearData(_Form.cbInvoced);
    _Form.Free;
  end;
end;

procedure TInvoce.InvoceEdit(aId: integer; const aTreeSelectedItems: ArrayOfInteger; const aFill: boolean);
var
  _Form: TFmInvoce;
  _DS: TDataSource;
  _InvoceId, _Result: Integer;
  _Name, _Unit, _VendorCode, _Label, _Remark: String;
  _Quantity, _IdPL, _IDOwner: LongInt;
  _Price: Currency;
  _arr: ArrayOfArrayVariant;
begin
  if not Assigned(fGrid) then exit;

  _Form:= TFmInvoce.Create(nil);
  try
    _Form.Caption:='Редактирование позиции';

    _Result:= aId;

    _DS:= fBase.SQLItemGetDS(Item,[aId]);

    _Name:= _DS.DataSet.FieldByName('NAME').AsString;
    _Unit:= _DS.DataSet.FieldByName('UNIT').AsString;
    _VendorCode:= _DS.DataSet.FieldByName('VENDORCODE').AsString;
    _Label:= _DS.DataSet.FieldByName('LABEL').AsString;
    _Price:= _DS.DataSet.FieldByName('PRICEPL').AsCurrency;
    _Quantity:= _DS.DataSet.FieldByName('QUANTITY').AsInteger;
    _Remark:= _DS.DataSet.FieldByName('REMARK').AsString;
    _IDOwner:= _DS.DataSet.FieldByName('IDOWNER').AsInteger;

    _IdPL:= fGrid.Grid.DataSource.DataSet.FieldByName('IDPL_ITEMS').AsInteger;



    _Form.edName.Text            := _Name;
    _Form.edUnit.Text            := _Unit;
    _Form.edVendorCode.Text      := _VendorCode;
    _Form.edLabel.Text           := _Label;
    _Form.edPrice.Text           := FormatCurrValue(_Price);
    _Form.edQuantity.Value       := _Quantity;

    cmbxFill(_Form.cbInvoced,_DS,['NAME','REMARK','ID']);

    _DS.DataSet.Close;

    _arr:= nil;
    _arr:= fBase.SQLReadArr('SELECT VSCOD FROM PL_GET_SCOD('+IntToStr(_IdPL)+',true)');

    if Assigned(_arr) then
      _Form.edScod.Text:= VarToStr(_arr[0,0])
    else
      _Form.edScod.Text:= '';

    _Form.onCbInvoiceChange:= @onCbInvoiceChange;

    _Form.cbInvoced.Items.AddObject('Создать новую позицию...', TwData.Create(-1));
    _Form.cbInvoced.ItemIndex:= cmbxItemIndexByID(_Form.cbInvoced,aId);
    onCbInvoiceChange(_Form,cmbxSelectID(_Form.cbInvoced));
    _Form.edQuantity.OnChange(_Form.edQuantity);

    _Form.ShowModal;

    if _Form.ModalResult = mrOK then
      begin
        EventBlock:= true;
        _InvoceId:= cmbxSelectID(_Form.cbInvoced);
        _Result:=0;

        if _Form.edQuantity.Value=0 then
          fBase.SQLItemUpdate(Del,[_InvoceId])
        else
          if (_InvoceId=-1) then
            _Result:= fBase.SQLItemUpdate(New,[_IDOwner, _IdPL, _Form.edQuantity.Value, _Form.edRemark.Text],true)
          else
            _Result:= fBase.SQLItemUpdate(Edit,[_IDOwner, _IdPL, _Form.edQuantity.Value, _Form.edRemark.Text, _InvoceId],true);
      end;

    GridFill( aTreeSelectedItems, aFill);
    if Assigned(fGrid.Grid.DataSource) and (fGrid.Grid.DataSource.DataSet.RecordCount>0) then
      fGrid.Grid.DataSource.DataSet.Locate('ID',_Result,[]);
  finally
    EventBlock:= false;
     cmbxClearData(_Form.cbInvoced);
    _Form.Free;
  end;
end;

procedure TInvoce.GridFill(aTreeSelectedItems: ArrayOfInteger; aFill: boolean);
begin
  if not Assigned(fGrid) or  not Assigned(fBase)
     or not aFill then exit;

  fGrid.GroupArray:= aTreeSelectedItems;

  fGrid.Fill();
end;

procedure TInvoce.InvoceClearRemark(aInvoces: ArrayOfInteger);
var
  aSQLUpdate,aRemark: String;
  i: Integer;
begin
  aRemark:='';

  if not InputQuery('Установка примечаний','Установить новое примечание у выделенных позиций ('+IntToStr(Length(aInvoces))+')?',aRemark) then exit;

  aSQLUpdate:= 'UPDATE INVOCES SET REMARK=%s WHERE ID=%d';
  for i:=0 to High(aInvoces) do
    fBase.SQLUpdate(Format(aSQLUpdate,[QuotedStr(aRemark),aInvoces[i]]));
end;

procedure TInvoce.InvoceDel(aTreeSelectedItems: ArrayOfInteger; aFill: boolean);
var
  i: Integer;
  _BookMark: TBookMark;
  aInvoces: ArrayOfInteger;
begin
  if not Assigned(fGrid) then exit;

  if MessageDlg('Удалить выделенные позиции из заказа?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;
  _BookMark:= fGrid.Grid.DataSource.DataSet.Bookmark;

  aInvoces:= fGrid.SelectedRows;

  EventBlock:= true;

  try
    for i:=0 to High(aInvoces) do
      fBase.SQLItemUpdate(Del,[aInvoces[i]]);

    GridFill(aTreeSelectedItems, aFill);

    if fGrid.Grid.DataSource.DataSet.RecordCount>0 then
        fGrid.Grid.DataSource.DataSet.Bookmark := _BookMark;
  finally
    EventBlock:= false;
    fGrid.SelectAll:= false;
  end;

end;

procedure TInvoce.mAnalogsFill(Sender: TObject; amAnalogs: TMenuItem);
var
  _DS: TDataSet;
  _NewItem: TMenuItem;
  i: Integer;
begin
  if not Assigned(fGrid) or (fGrid.Grid.DataSource.DataSet.RecordCount=0) then exit;


  _DS:= fBase.SQLReadDS('SELECT * FROM ANALIS_SEL_ALL_ANALOG('+fGrid.FieldValue['IDPL_ITEMS'].AsString+',true) WHERE PRICE>0 AND IDOWNER<>'+IntToStr(IdMainOwner)+' ORDER BY PRICE').DataSet;

  if _DS.RecordCount>0 then
  begin
     amAnalogs.Visible:= true;
     amAnalogs.Clear;

     _NewItem:=nil;
     _NewItem:= NewItem(
                fGrid.FieldValue['OWNERNAME'].AsString+' | '
                +fGrid.FieldValue['NAME'].AsString
                +' | '+CurrToStrF(fGrid.FieldValue['PRICEPL'].AsCurrency, ffCurrency, 2)
                , 0, False, false, nil, 0, 'm'+fGrid.FieldValue['ID'].AsString);
     _NewItem.ImageIndex:=15;
       amAnalogs.Add(
           _NewItem
       );

       amAnalogs.Add(
            NewItem('-', 0, False, True, nil, 0, 'mSplitAnalogs')
        );

       for i:=0 to _DS.RecordCount-1 do
       begin
         _NewItem:= NewItem(_DS.FieldByName('OWNERNAME').AsString+' | '
                    +_DS.FieldByName('QUANTITYINPACKINGTEXT').AsString
                    +' | '+_DS.FieldByName('PLNAME').AsString+' | '
                    +_DS.FieldByName('STOCK').AsString+' | '
                    +CurrToStrF(_DS.FieldByName('PRICE').AsCurrency, ffCurrency, 2)
                    , 0, False, True, @mAnalogsClick, 0, 'm'+_DS.FieldByName('ID').AsString);
         _NewItem.ImageIndex:=24;

           amAnalogs.Add(
               _NewItem
           );
           _DS.Next;
       end;

  end else
    amAnalogs.Visible:= false;

  _DS.Close;
  _DS:= nil;
end;

procedure TInvoce.InvoceChangeItem(aIdPL, aIdInvoce: integer);
var
  _Quantity: LongInt;
  _Remark: String;
  _DS: TDataSet;
  _Result: Integer;
begin
  _Result:= -1;
  _Quantity:= fGrid.FieldValue['QUANTITY'].AsInteger;
  _Remark:= fGrid.FieldValue['REMARK'].AsString;

  fBase.LongTransaction:= true;

  try
    _DS:= fBase.SQLReadDS('PL_ITEMS',['ID','IDOWNER'],'ID='+IntToStr(aIdPL),'').DataSet;

    fBase.SQLItemUpdate(Del,[aIdInvoce],false,false);
    //:IDOWNER, :IDPL_ITEMS, :QUANTITY, :REMARK
    _Result:= fBase.SQLItemUpdate(New,[_DS.FieldByName('IDOWNER').AsInteger,
                      _DS.FieldByName('ID').AsInteger,
                      _Quantity,
                      _Remark],true,false);

  fBase.SQLTransactionEnd(true);

  with fGrid.Grid.DataSource.DataSet do
  begin
    Close;
    Open;
    DisableControls;
    Locate('ID',_Result,[]);
    EnableControls;
  end;

  except
    fBase.SQLTransactionEnd(false);
  end;

end;

procedure TInvoce.InvoceAddNewItem(aGridInvoce:TwDBGrid; aFill: boolean);
var
  _Form: TFmListSelect;
  _DataSet: TDataSet;
  aPriceId, aInvoceId, aPriceIdOwner: Integer;
  aOrderId: LongInt;
  _arr, aArr: ArrayOfArrayVariant;
  aQuantity: Double;
begin
  if not Assigned(aGridInvoce) then exit;

  aPriceId:= -1;
  aOrderId:= -1;
  aQuantity:= 1;

  _DataSet:= aGridInvoce.Grid.DataSource.DataSet;

  if _DataSet.RecordCount = 0 then exit;

  try

    _Form:= TFmListSelect.Create(aGridInvoce.Grid);
    _Form.Base:= fBase;
    _Form.wFormMode:=0;
    _Form.MultiSelectGrid:= false;

    aInvoceId:= _DataSet.FieldByName('ID').AsInteger;
    aPriceId:= _DataSet.FieldByName('IDPL_ITEMS').AsInteger;
    //_Form
    _Form.wDataSetLocateField:= 'ID';
    //_Form.wDataSetLocateValue:= aPriceId;

    Screen.Cursor:=crSQLWait;
    try
      _Form.ListFormInit(
               '',
               VarToStr(_DataSet.FieldByName('VENDORCODE').AsVariant),
               VarToStr(_DataSet.FieldByName('NAME').AsVariant),
               VarToStr(_DataSet.FieldByName('LABEL').AsVariant),
               VarToStr(_DataSet.FieldByName('SCOD').AsVariant));

      _Form.ShowModal;

      if _Form.ModalResult = mrOK then
      begin
        EventBlock:= true;

        if not fBase.LongTransaction then
            fBase.LongTransaction:= true;

        aPriceId:= _Form.wSelectedRows[0];
        aArr:= fBase.SQLReadArr('PL_ITEMS',['IDOWNER'],'ID='+IntToStr(aPriceId),'');
        if Assigned(aArr) then
          aPriceIdOwner:= integer(aArr[0,0]);

        aOrderId:= _DataSet.FieldByName('ID').AsInteger;
        aQuantity:= _DataSet.FieldByName('QUANTITY').AsFloat;

        //if MessageDlg('Добавить соответствие к выбранной позиции?',mtConfirmation, mbOKCancel, 0) = mrOK then
        //   InsertMaching(fBase,
        //      _DataSet.FieldByName('ORDOWNER').AsInteger,
        //      _DataSet.FieldByName('ORDVENDORCODE').AsString,
        //      aPriceIdOwner,
        //      aPriceId,
        //      _Form.wQuantityInPacked,
        //      aOrderId);

          _arr:= fBase.SQLReadArr('SELECT PL.IDOWNER FROM PL_ITEMS PL WHERE PL.ID='+IntToStr(aPriceId));

          //'UPDATE INVOCES SET IDOWNER=:IDOWNER, IDPL_ITEMS=:IDPL_ITEMS, QUANTITY=:QUANTITY, REMARK=:REMARK WHERE ID=:ID RETURNING ID'
          if Assigned(_arr) then
            aInvoceId:= fBase.SQLItemUpdate(Edit,[_arr[0,0], aPriceId, aQuantity/_Form.wQuantityInPacked, 'auto_add_manual',aInvoceId],true, false);

          fBase.SQLUpdate('W_TMP_ORDERS_IMPORT',['FPASSED'],[integer(1)],'ID='+IntToStr(aOrderId),false);

          fBase.SQLTransactionEnd(true);
          _DataSet.DisableControls;
      end;

    finally
      EventBlock:= false;
     Screen.Cursor:=crDefault;
     _Form.Free;

     GridFill(nil, aFill);

     aGridInvoce.Fill();

     if aInvoceId>0 then
       aGridInvoce.Grid.DataSource.DataSet.Locate('ID', aInvoceId,[]);

     _DataSet.EnableControls;
    end;

    if Assigned(onAddNewItem) then onAddNewItem(aGridInvoce, 0);

  except
     fBase.SQLTransactionEnd(false);
  end;

end;

procedure TInvoce.InvoceAddNewItem(aGridNoFinded:TwDBGrid; aTreeSelectedItems: ArrayOfInteger; aFill: boolean);
var
  _Form: TFmListSelect;
  _DataSet: TDataSet;
  _BoookMark: TBookMark;
  aPriceId, aInvoceId, aPriceIdOwner: Integer;
  aOrderId: LongInt;
  _arr, aArr: ArrayOfArrayVariant;
  aQuantity: Double;
begin
  if not Assigned(aGridNoFinded) then exit;

  aPriceId:= -1;
  aOrderId:= -1;
  aQuantity:= 1;

  _DataSet:= aGridNoFinded.Grid.DataSource.DataSet;

  if _DataSet.RecordCount = 0 then exit;

  try
    _BoookMark:= _DataSet.Bookmark;

    _Form:= TFmListSelect.Create(aGridNoFinded.Grid);
    _Form.Base:= fBase;
    _Form.wFormMode:=0;
    _Form.MultiSelectGrid:= false;
    //_Form
    _Form.wDataSetLocateField:= 'ID';
    //_Form.wDataSetLocateValue:= '';
    Screen.Cursor:=crSQLWait;
    try
      _Form.ListFormInit(
               '',
               VarToStr(_DataSet.FieldByName('ORDVENDORCODE').AsVariant),
               VarToStr(_DataSet.FieldByName('ORDNAME').AsVariant),
               VarToStr(_DataSet.FieldByName('ORDLABEL').AsVariant),
               VarToStr(_DataSet.FieldByName('ORDSCOD').AsVariant));

      _Form.ShowModal;

      if _Form.ModalResult = mrOK then
      begin
        EventBlock:= true;
        aPriceId:= _Form.wSelectedRows[0];
        aArr:= fBase.SQLReadArr('PL_ITEMS',['IDOWNER'],'ID='+IntToStr(aPriceId),'');
        if Assigned(aArr) then
          aPriceIdOwner:= integer(aArr[0,0]);

        aOrderId:= _DataSet.FieldByName('ID').AsInteger;
        aQuantity:= _DataSet.FieldByName('ORDQUANTITY').AsFloat;

        if not fBase.LongTransaction then
            fBase.LongTransaction:= true;

        if MessageDlg('Добавить соответствие к выбранной позиции?',mtConfirmation, mbOKCancel, 0) = mrOK then
           InsertMaching(fBase,
              _DataSet.FieldByName('ORDOWNER').AsInteger,
              _DataSet.FieldByName('ORDVENDORCODE').AsString,
              aPriceIdOwner,
              aPriceId,
              _Form.wQuantityInPacked,
              aOrderId);

          _arr:= fBase.SQLReadArr('SELECT PL.IDOWNER FROM PL_ITEMS PL WHERE PL.ID='+IntToStr(aPriceId));

          if Assigned(_arr) then
            aInvoceId:= fBase.SQLItemUpdate(New,[_arr[0,0], aPriceId, aQuantity/_Form.wQuantityInPacked, 'auto_add_new'],true, false);

          fBase.SQLUpdate('W_TMP_ORDERS_IMPORT',['FPASSED'],[integer(1)],'ID='+IntToStr(aOrderId),false);

          fBase.SQLTransactionEnd(true);
      end;

    finally
     EventBlock:= false;
     Screen.Cursor:=crDefault;
     _Form.Free;

     GridFill(nil, aFill);

     aGridNoFinded.Fill();

     if _DataSet.RecordCount>0 then
        _DataSet.Bookmark:= _BoookMark;

     if aInvoceId>0 then
       fGrid.Grid.DataSource.DataSet.Locate('ID', aInvoceId,[]);
    end;

    if Assigned(onAddNewItem) then onAddNewItem(self, 0);

  except
     fBase.SQLTransactionEnd(false);
  end;

end;

procedure TInvoce.ExportSummaryInvoce;
var
  fViewer: TwViewer;
  fWorkbookSource: TsWorkbookSource;
begin
 fViewer:= TwViewer.Create(nil);
 fViewer.Caption:= 'Сводный заказ';

 fWorkbookSource:= TsWorkbookSource.Create(Application);
 fWorkbookSource.Options:= fWorkbookSource.Options+[boFileStream];

 fProgress:= TProgress.Create(nil);
 fProgress.Caption:= 'Формирование отчета...';
 fProgress.ShowLog:= false;
 fProgress.onStopForce:= @onStopForce;
 fProgress.NoClose:= true;

 fReport:= TwReport.Create(true);
 //fReport.SelectedPriceItems:= aSelectedItems;
 fReport.Base:= fBase;
 fReport.ReportModes:= rmSummaryInvoce;
 fReport.WorkbookSource:= fWorkbookSource;
 fReport.onProgressInit:= @onProgressInit;
 fReport.onProgressUpdate:= @onProgressUpdate;
 fReport.onEndThread:= @onEndThread;

 screen.Cursor:= crSQLWait;

 fReport.start;

 try

   fProgress.ShowModal;

   fViewer.WorkbookSource:= fWorkbookSource;

   //TFmOrders(fOwnerForm).Repaint;
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

procedure TInvoce.ExportInvoceWithSelectOwnerCode(aSelected: ArrayOfInteger);
var
  fViewer: TwViewer;
  fWorkbookSource: TsWorkbookSource;
begin
 fViewer:= TwViewer.Create(nil);
 fViewer.Caption:= 'Заказ с кодами выбранных поставщиков';

 fWorkbookSource:= TsWorkbookSource.Create(Application);
 fWorkbookSource.Options:= fWorkbookSource.Options+[boFileStream];

 fProgress:= TProgress.Create(nil);
 fProgress.Caption:= 'Формирование отчета...';
 fProgress.ShowLog:= false;
 fProgress.onStopForce:= @onStopForce;
 fProgress.NoClose:= true;

 fReport:= TwReport.Create(true);
 fReport.SelectedPriceItems:= aSelected;
 fReport.Base:= fBase;
 fReport.ReportModes:= rmInvoceWithSelectOwnerCode;
 fReport.WorkbookSource:= fWorkbookSource;
 fReport.onProgressInit:= @onProgressInit;
 fReport.onProgressUpdate:= @onProgressUpdate;
 fReport.onEndThread:= @onEndThread;

 screen.Cursor:= crSQLWait;

 fReport.start;

 try

   fProgress.ShowModal;

   fViewer.WorkbookSource:= fWorkbookSource;

   //TFmOrders(fOwnerForm).Repaint;
   if fReport.Result then
     begin
        fViewer.ShowModal;
     end;

//   fReport.Terminate;
 finally
   screen.Cursor:=crDefault;
  if Assigned(fProgress) then
    fProgress.Free;
   fViewer.free;
 end;
end;

procedure TInvoce.GetPriceFile(aOwnerID: int64; var aRecPriceFormat: TRecPriceFormat);
var
  i, aFormatID: integer;
  aForm: TFmArcView;
  aFilesArr: ArrayOfArrayVariant;
  aArr: ArrayOfString;
  aUnPackPath, aWhere, aFileName: string;
  aZipper: TwZipper;
  aObject: TwData;
  aFieldsArray: wTypesU.ArrayOfString;
  aDataSet: TDataSet;
begin
  aFilesArr:= fBase.SQLReadArr('FORMATS', ['FILE', 'ID'], 'IDOWNER='+IntToStr(aOwnerID)+' AND IDFMTS_CATEGORY=1 AND FCLOSE=0', '');

  aFileName:= aFilesArr[0, 0];
  aFormatID:= integer(aFilesArr[0,1]);

  if Length(aFileName)=0 then Exit;

  aFieldsArray:=fBase.MakeArrayFromString(FormatImportFields);

  aDataSet:= fBase.SQLReadDS('FORMATS',aFieldsArray,' ID='+IntToStr(aFormatID),'').DataSet;


  if aDataSet.RecordCount=0 then exit;

  with aRecPriceFormat do
  begin
    fIDOWNER:= aDataSet.FieldByName('IDOWNER').AsVariant;
    fID:= aDataSet.FieldByName('ID').AsVariant;
    fNAME:= aDataSet.FieldByName('NAME').AsVariant;
    fFILE:= aDataSet.FieldByName('FILE').AsVariant;
    fFILEZIPNAMEDECODE:= aDataSet.FieldByName('FILEZIPNAMEDECODE').AsVariant;
    fFILEHASH:= aDataSet.FieldByName('FILEHASH').AsVariant;
    fURL:= aDataSet.FieldByName('URL').AsVariant;
    fIDFILEFORMAT:= aDataSet.FieldByName('IDFILEFORMAT').AsVariant;
    fFCONVERTLIBRE:= aDataSet.FieldByName('FCONVERTLIBRE').AsVariant;
    fIDCODEPAGETEXT:= aDataSet.FieldByName('IDCODEPAGETEXT').AsVariant;
    fIDCURRENCY:= aDataSet.FieldByName('IDCURRENCY').AsVariant;
    fCURRENCYPERCENT:= aDataSet.FieldByName('CURRENCYPERCENT').AsVariant;
    fSTORAGEDAYS:= aDataSet.FieldByName('STORAGEDAYS').AsVariant;
    fSTOCKONLY:= aDataSet.FieldByName('STOCKONLY').AsVariant;
    fSTOCKSYMBOLS:= fBase.MakeArrayArrayVariantFromString(aDataSet.FieldByName('STOCKSYMBOLS').AsString);
    fSTOCKONLYINFO:= aDataSet.FieldByName('STOCKONLYINFO').AsVariant;
    fYMLID:= aDataSet.FieldByName('YMLID').AsVariant;
    fYMLPRICE:= aDataSet.FieldByName('YMLPRICE').AsVariant;
    fYMLQUANTITY:= aDataSet.FieldByName('YMLQUANTITY').AsVariant;
    fFCLOSE:= aDataSet.FieldByName('FCLOSE').AsVariant;
    fGROUPSINROWS:= aDataSet.FieldByName('GROUPSINROWS').AsVariant;
    fGROUPALGORITHM:= aDataSet.FieldByName('GROUPALGORITHM').AsVariant;
    fGROUPS:= aDataSet.FieldByName('GROUPS').AsVariant;
    fSUBGROUPS1:= aDataSet.FieldByName('SUBGROUPS1').AsVariant;
    fSUBGROUPS2:= aDataSet.FieldByName('SUBGROUPS2').AsVariant;
    fSUBGROUPS3:= aDataSet.FieldByName('SUBGROUPS3').AsVariant;
    fFIRSTLINE:= aDataSet.FieldByName('FIRSTLINE').AsVariant;
    fVENDORCODE:= aDataSet.FieldByName('VENDORCODE').AsVariant;
    fFNAME:= aDataSet.FieldByName('FNAME').AsVariant;
    fUNIT:= aDataSet.FieldByName('UNIT').AsVariant;
    fQUANTITY:= aDataSet.FieldByName('QUANTITY').AsVariant;
    fSTOCK2:= aDataSet.FieldByName('STOCK2').AsVariant;
    fSTOCK3:= aDataSet.FieldByName('STOCK3').AsVariant;
    fSTOCK4:= aDataSet.FieldByName('STOCK4').AsVariant;
    fSTOCK5:= aDataSet.FieldByName('STOCK5').AsVariant;
    fTRANSIT:= aDataSet.FieldByName('TRANSIT').AsVariant;
    fPRICE:= aDataSet.FieldByName('PRICE').AsVariant;
    fPRICE2:= aDataSet.FieldByName('PRICE2').AsVariant;
    fPRICE3:= aDataSet.FieldByName('PRICE3').AsVariant;
    fPRICE4:= aDataSet.FieldByName('PRICE4').AsVariant;
    fPRICE5:= aDataSet.FieldByName('PRICE5').AsVariant;
    fPRICE6:= aDataSet.FieldByName('PRICE6').AsVariant;
    fPRICE7:= aDataSet.FieldByName('PRICE7').AsVariant;
    fPRICE8:= aDataSet.FieldByName('PRICE8').AsVariant;
    fPRICE9:= aDataSet.FieldByName('PRICE9').AsVariant;
    fPRICE10:= aDataSet.FieldByName('PRICE10').AsVariant;
    fLABEL:= aDataSet.FieldByName('LABEL').AsVariant;
    fSCOD:= aDataSet.FieldByName('SCOD').AsVariant;
    fFURL:= aDataSet.FieldByName('FURL').AsVariant;
    fFURLPICTURE:= aDataSet.FieldByName('FURLPICTURE').AsVariant;
    fFREMARK:= aDataSet.FieldByName('FREMARK').AsVariant;
    fFCOLOR:= aDataSet.FieldByName('FCOLOR').AsVariant;
    fIDFMTS_CATEGORY:= aDataSet.FieldByName('IDFMTS_CATEGORY').AsVariant;
    fSPREADSHEET:= fBase.MakeArrayArrayIntegerFromString(aDataSet.FieldByName('SPREADSHEET').AsString,aDataSet.FieldByName('FIRSTLINE').AsInteger);
    fIDVENDORCODEVARIANT:= aDataSet.FieldByName('IDVENDORCODEVARIANT').AsInteger;
    fIDCSVDELIMITER:= aDataSet.FieldByName('IDCSVDELIMITER').AsVariant;
    fIDSTOCKVARIANT:= aDataSet.FieldByName('IDSTOCKVARIANT').AsVariant;
    fIDPRICEVARIANT:= aDataSet.FieldByName('IDPRICEVARIANT').AsVariant;
    fADDRCELLFORINVOCE:= aDataSet.FieldByName('ADDRCELLFORINVOCE').AsVariant;
    fINVOCEDAYS:= aDataSet.FieldByName('INVOCEDAYS').AsVariant;
  end;

  aDataSet.Close;

   aZipper:= TwZipper.Create();

   try
     aArr:= aZipper.ParseComboFileName(aFileName);
     if Assigned(aArr) then
       begin
         if Length(aArr[1])>0 then
           begin
             aUnPackPath:= includeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
             aUnPackPath:= aUnPackPath+'tmp';

             if not DirectoryExistsUTF8(aUnPackPath) then ForceDirectoriesUTF8(aUnPackPath);
             aUnPackPath:=aUnPackPath+DirectorySeparator+IntTOStr(aOwnerID);
             if not DirectoryExistsUTF8(aUnPackPath) then ForceDirectoriesUTF8(aUnPackPath);

             //sgFormat.Cells[1,12]:= aFileName+'|'+_FileExtract;
             aZipper.ExtractOneFile(aArr[0], aArr[1], aUnPackPath);
             aFileName:=includeTrailingPathDelimiter(aUnPackPath)+aArr[1];
           end else
           aFileName:= aArr[0];
         //Length(
       end;
       if FileExists(aFileName) then
         begin
           if (aRecPriceFormat.fFCONVERTLIBRE = 1) then
             aFileName:= ConvertFileWithLibreOffice(aFileName);
         end
       else
       begin
         aFileName:= '';
         ShowMessage('Сохраненный локально файл не найден!');
       end;

       aRecPriceFormat.fFILE:= aFileName;
   finally
      aArr:=nil;
      aZipper.Destroy();
   end;
end;

procedure TInvoce.ExportInvoceInOwnerPrice(aSelected: ArrayOfInteger);
var
  fWorkbookSource: TsWorkbookSource;
begin
 //fViewer:= TwViewer.Create(nil);
 //fViewer.Caption:= 'Заказ с кодами выбранных поставщиков';

 fWorkbookSource:= TsWorkbookSource.Create(Application);
 fWorkbookSource.Options:= fWorkbookSource.Options+[boFileStream];

 fProgress:= TProgress.Create(nil);
 fProgress.Caption:= 'Экспорт в прайс-лист...';
 fProgress.ShowLog:= false;
 fProgress.onStopForce:= @onStopForce;
 fProgress.NoClose:= true;

 fReport:= TwReport.Create(true);
 fReport.SelectedPriceItems:= aSelected;
 fReport.Base:= fBase;
 fReport.ReportModes:= rmInvoceToPrice;
 fReport.WorkbookSource:= fWorkbookSource;
 fReport.onProgressInit:= @onProgressInit;
 fReport.onProgressUpdate:= @onProgressUpdate;
 fReport.onEndThread:= @onEndThread;

 screen.Cursor:= crSQLWait;

 fReport.start;

 try

   fProgress.ShowModal;

   //fViewer.WorkbookSource:= fWorkbookSource;

   //TFmOrders(fOwnerForm).Repaint;
   //if fReport.Result then
   //  begin
   //     fViewer.ShowModal;
   //  end;

  // fReport.Terminate;
 finally
   screen.Cursor:=crDefault;
  if Assigned(fProgress) then
    fProgress.Free;
   //fViewer.free;
 end;

end;

procedure TInvoce.ExportInvoceToOwnerFiles;
var
  fWorkbookSource: TsWorkbookSource;
  SelectDirectoryDialog: TSelectDirectoryDialog;
  aPathFiles: String;
begin

 SelectDirectoryDialog:= TSelectDirectoryDialog.Create(nil);

 try
   if not SelectDirectoryDialog.Execute then exit;
   aPathFiles:= SelectDirectoryDialog.FileName;
 finally
   SelectDirectoryDialog.Free;
 end;


 fWorkbookSource:= TsWorkbookSource.Create(Application);
 fWorkbookSource.Options:= fWorkbookSource.Options+[boFileStream];

 fProgress:= TProgress.Create(nil);
 fProgress.Caption:= 'Экспорт...';
 fProgress.ShowLog:= false;
 fProgress.onStopForce:= @onStopForce;
 fProgress.NoClose:= true;

 fReport:= TwReport.Create(true);
 //fReport.SelectedPriceItems:= aSelectedItems;
 fReport.Base:= fBase;
 fReport.ReportModes:= rmToOwnerFiles;
 fReport.WorkbookSource:= fWorkbookSource;
 fReport.onProgressInit:= @onProgressInit;
 fReport.onProgressUpdate:= @onProgressUpdate;
 fReport.onEndThread:= @onEndThread;
 fReport.PathToFiles:= IncludeTrailingBackslash(aPathFiles);
 screen.Cursor:= crSQLWait;

 fReport.start;

 try

   fProgress.ShowModal;

   //fViewer.WorkbookSource:= fWorkbookSource;

   //TFmOrders(fOwnerForm).Repaint;
   if fReport.Result then
     begin
        ShowMessage('Экспорт завершен!');
     end;

  // fReport.Terminate;
 finally
   screen.Cursor:=crDefault;
  if Assigned(fProgress) then
    fProgress.Free;
 end;

end;

function TInvoce.ConvertIdWTOIToIdPRICE(aSelectedItems:ArrayOfInteger):ArrayOfInteger;
var
  aSelect, aSQL: String;
  aArr: ArrayOfArrayVariant;
  i: Integer;
begin
 Result:= nil;
 aSelect:= fBase.MakeStringFromArray(aSelectedItems);
 aSQL:='SELECT PL.ID FROM W_TMP_ORDERS_IMPORT WTOI '
        +' INNER JOIN PL_ITEMS PL ON (PL.IDOWNER=WTOI.ORDOWNER AND PL.VENDORCODE=WTOI.ORDVENDORCODE) '
        +' WHERE WTOI.ID IN (%s)';

 aArr:= fBase.SQLReadArr(Format(aSQL,[aSelect]));
 SetLength(Result,Length(aArr));

 for i:=0 to High(aArr) do
   Result[i]:= aArr[i,0];

 aArr:= nil;
end;

function TInvoce.ConvertIdINVToIdPRICE(aSelectedItems:ArrayOfInteger):ArrayOfInteger;
var
  aSelect, aSQL: String;
  aArr: ArrayOfArrayVariant;
  i: Integer;
begin
 Result:= nil;
 aSelect:= fBase.MakeStringFromArray(aSelectedItems);
 aSQL:='SELECT INV.IDPL_ITEMS FROM INVOCES INV '
        +' WHERE INV.ID IN (%s)';

 aArr:= fBase.SQLReadArr(Format(aSQL,[aSelect]));
 SetLength(Result,Length(aArr));

 for i:=0 to High(aArr) do
   Result[i]:= aArr[i,0];

 aArr:= nil;
end;

procedure TInvoce.GetPositionAnalog(aSelectedItems: ArrayOfInteger; aFinded: boolean);
var
  fViewer: TwViewer;
  fWorkbookSource: TsWorkbookSource;
begin
 ///W_TMP_ORDERS_IMPORT
 fViewer:= TwViewer.Create(nil);
 fViewer.Caption:= 'Аналоги позиций';

 fWorkbookSource:= TsWorkbookSource.Create(Application);
 fWorkbookSource.Options:= fWorkbookSource.Options+[boFileStream];

 fProgress:= TProgress.Create(nil);
 fProgress.Caption:= 'Формирование отчета...';
 fProgress.ShowLog:= false;
 fProgress.onStopForce:= @onStopForce;
 fProgress.NoClose:= true;

 fReport:= TwReport.Create(true);
 if aFinded then
   fReport.SelectedPriceItems:= ConvertIdINVToIdPRICE(aSelectedItems)
 else
   fReport.SelectedPriceItems:= ConvertIdWTOIToIdPRICE(aSelectedItems);

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

   if fReport.Result then
     begin
        fViewer.ShowModal;
     end;

//   fReport.Terminate;
 finally
   screen.Cursor:=crDefault;
   if Assigned(fProgress) then
     fProgress.Free;
   fViewer.free;
 end;

end;

end.

