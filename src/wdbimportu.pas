unit wDBImportU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

// Назначение:
// Импорт данных в БД из различных источников
// СВЯЗИ: wDBaseU, wLogU


{$mode objfpc}{$H+}

interface

uses
  Classes, Controls, StdCtrls, SysUtils, fgl, DB, Graphics, LazUTF8, LConvEncoding, LCLProc,
  LazFileUtils,
  fpspreadsheetctrls, fpstypes, fpscsv, fpsallformats, fpspreadsheet, fpsNumFormat, gvector, Dialogs,
  {Для импорта электронных таблиц xls,xlsx,ods}

  //zexmlss, zeodfs, zexmlssutils, zeformula, zsspxml,zexlsx,
  //{Для альтернативного импорта электронных таблиц xlsx}
  DOM, xmlread,
  wBaseU, UtilsU,
  wLogU, wFuncU, wZipperU, wYMLparser, wGetU, wTypesU, wCustomClassThreadU,
  Forms;

  const
    FormatImportFields = 'IDOWNER, ID, NAME, FILE, FILEZIPNAMEDECODE, FILEHASH, URL, IDFILEFORMAT, FCONVERTLIBRE, '
      +' IDCODEPAGETEXT, IDCURRENCY, CURRENCYPERCENT, STORAGEDAYS, STOCKONLY, STOCKSYMBOLS, STOCKONLYINFO, YMLID, YMLPRICE, YMLQUANTITY, '
      +' SPREADSHEET, FCLOSE, GROUPSINROWS, GROUPALGORITHM, GROUPS, SUBGROUPS1, SUBGROUPS2, SUBGROUPS3, FIRSTLINE, VENDORCODE, FNAME, UNIT, '
      +' QUANTITY, STOCK2, STOCK3, STOCK4, STOCK5, TRANSIT, PRICE, PRICE2, PRICE3, PRICE4, PRICE5, PRICE6, PRICE7, PRICE8, PRICE9, PRICE10, LABEL, SCOD, FURL, FURLPICTURE, FREMARK, '
      +' FCOLOR, IDFMTS_CATEGORY, IDVENDORCODEVARIANT, IDCSVDELIMITER, IDSTOCKVARIANT, IDPRICEVARIANT, ADDRCELLFORINVOCE, INVOCEDAYS ';
type

  TGroupAlgoritm = (gaBgPrice, gaBgIdent);

  TwImportFormat = (wSpreadSheet);
  TwColorList = specialize TFPGList<TsColor>;  // типизированный список

  TRecOrderFormat = Record
    id: integer;
    IdOwner: integer;
    Name: string;
    Remark: string;
    ConvertWithLibre: boolean;
    FileFormat: integer;
    fSpreadsheets: ArrayOfArrayInteger;
    fFirstLine: integer;
    fVendorcode: integer;
    fName: integer;
    fUnit: integer;
    fQuantity: integer;
    fSum: integer;
    fCustomDeclaration: integer;
    fCountry: integer;
    fLabel: integer;
    fScod: integer;
    fRemark: integer;
    fCodePage: integer;
    fIDCSVDELIMITER: integer;
    fCURRENCY: integer;
    fCURRENCYPERCENT: double;
    fIDVENDORCODEVARIANT : integer;
    fIDSTOCKVARIANT : integer;
    fIDPRICEVARIANT : integer;
    fOUTCELLTEXT: string;
    fADDRCELLTEXT: string;
  end;

  TwResult = record
    fVENDORCODE: string;
    fFNAME: string;
    fUNIT: string;
    fQUANTITY: string;
    fSTOCK2: string;
    fSTOCK3: string;
    fSTOCK4: string;
    fSTOCK5: string;
    fPRICE: string;
    fPRICE2: string;
    fPRICE3: string;
    fPRICE4: string;
    fPRICE5: string;
    fPRICE6: string;
    fPRICE7: string;
    fPRICE8: string;
    fPRICE9: string;
    fPRICE10: string;
    fLABEL: string;
    fSCOD: string;
    fGROUPS: string;
    fSUBGROUPS1: string;
    fSUBGROUPS2: string;
    fSUBGROUPS3: string;
    fTRANSIT: string;
    fFURL: string;
    fFURLPICTURE: string;
    fFREMARK: string;
    fFCOLOR: string;
  end;

    //
    TRecPriceFormat = record
      fIDOWNER: integer;
      fID: integer;
      fNAME: string;
      fFILE: string;
      fFILEZIPNAMEDECODE: integer;
      fFILEHASH: string;
      fURL: string;
      fIDFILEFORMAT: integer;
      fFCONVERTLIBRE: integer;
      fIDCODEPAGETEXT: integer;
      fIDCURRENCY: integer;
      fCURRENCYPERCENT: double;
      fSTORAGEDAYS: integer;
      fSTOCKONLY: integer;
      fSTOCKSYMBOLS: ArrayOfArrayVariant;
      fSTOCKONLYINFO: integer;
      fYMLID: integer;
      fYMLPRICE: integer;
      fYMLQUANTITY: integer;
      fFCLOSE: integer;
      fGROUPSINROWS: integer;
      fGROUPALGORITHM: integer;
      fGROUPS: integer;
      fSUBGROUPS1: integer;
      fSUBGROUPS2: integer;
      fSUBGROUPS3: integer;
      fFIRSTLINE: integer;
      fVENDORCODE: integer;
      fFNAME: integer;
      fUNIT: integer;
      fQUANTITY: integer; //fSTOCK
      fSTOCK2: integer;
      fSTOCK3: integer;
      fSTOCK4: integer;
      fSTOCK5: integer;
      fTRANSIT: integer;
      fPRICE: integer;
      fPRICE2: integer;
      fPRICE3: integer;
      fPRICE4: integer;
      fPRICE5: integer;
      fPRICE6: integer;
      fPRICE7: integer;
      fPRICE8: integer;
      fPRICE9: integer;
      fPRICE10: integer;
      fLABEL: integer;
      fSCOD: integer;
      fFURL: integer;
      fFURLPICTURE: integer;
      fFREMARK: integer;
      fFCOLOR: integer;
      fIDFMTS_CATEGORY: integer;
      fSPREADSHEET: ArrayOfArrayInteger;
      fIDVENDORCODEVARIANT : integer;
      fIDSTOCKVARIANT : integer;
      fIDPRICEVARIANT : integer;
      fIDCSVDELIMITER: integer;
      fADDRCELLFORINVOCE: integer; //нет в CreateFormat
      fINVOCEDAYS: integer; //нет в CreateFormat
    end;

    TwPriceFormats = specialize TVector<TRecPriceFormat>;

  { TwDBImport }

  TwDBImport = class
  private
    fFormatsPrice: TwPriceFormats;

    fFormName: string;
    fOutBase: Boolean;
    wGroupRootIndex: integer;
    wOwnerID: integer;
    //FormatID: string;

    fErrorMessage: string;
    fOKMessage: string;

    fEndThread: boolean;

    wParent: TComponent;
    wMemo: TMemo;
    wSettingsCellArray: ArrayOfArrayVariant;
    wSQLStrings: TStringList;
    //wSQLStringsNotClearPrice: TStringList;
    fBase: TwBase;
    fBaseOuter: Boolean;

//    XMLSS: TZEXMLSS;       // импорт xlsx
    XMLSS_ON: boolean;

   // FImportedFieldNames: TStringList;
    wTimeStamp: string;
    procedure _onEndThread(Sender: TObject);

   private
     FFormatOrder: TRecOrderFormat;
     fOutStringArr: ArrayOfString;
    wIgnoreVersion: boolean;
    //FImportedRowCells: Array of TCell;
    //FDateTemplateCell: PCell;
    //FImportedFieldNames: TStringList;

    procedure Log(_Text: string);
    procedure SetStatus(_Text: string; const _Status: boolean = false; const aLogSection: boolean = true); // вывод статуса
    // открыть источник данных для импорта

    property Parent: TComponent read wParent write wParent;

    //property SQLStrings: TStringList read wSQLStrings write wSQLStrings;
    //
    //property GroupRootIndex: integer read wGroupRootIndex write wGroupRootIndex;
    //
    //
    //property OwnerID: integer read wOwnerID write wOwnerID;
    //property TimeStamp: string read wTimeStamp write wTimeStamp;
  public

    constructor Create(Sender: TObject; const _Memo: TMemo = nil);

    destructor Destroy();

    procedure Import(const aFormatType: TFormatType = ftPRICE; const aFileName: string = '');
    procedure ImportKursValut(aSilent: boolean);

    function CreateFormatPrice(aIDOWNER, aID: integer; aNAME, aFILE: string; aFILEZIPNAMEDECODE: integer; aFILEHASH, aURL: string; aIDFILEFORMAT, aFCONVERTLIBRE,
      aIDCODEPAGETEXT, aIDCURRENCY: integer; aCURRENCYPERCENT: double; aSTORAGEDAYS, aSTOCKONLY: integer; aSTOCKSYMBOLS: ArrayOfArrayVariant; aSTOCKONLYINFO,
  aYMLID, aYMLPRICE, aYMLQUANTITY, aFCLOSE, aGROUPSINROWS, aGROUPALGORITHM, aGROUPS, aSUBGROUPS1, aSUBGROUPS2, aSUBGROUPS3, aFIRSTLINE, aVENDORCODE, aFNAME,
  aUNIT, aQUANTITY, aSTOCK2, aSTOCK3, aSTOCK4, aSTOCK5, aTRANSIT, aPRICE, aPRICE2, aPRICE3, aPRICE4, aPRICE5, aPRICE6, aPRICE7, aPRICE8, aPRICE9, aPRICE10,
  aLABEL, aSCOD, aFURL, aFURLPICTURE, aFREMARK,
  aCOLOR, aIDFMTS_CATEGORY: integer; aSPREADSHEET: ArrayOfArrayInteger; aIDVENDORCODEVARIANT,aIDCSVDELIMITER,aIDSTOCKVARIANT,aIDPRICEVARIANT:integer): TRecPriceFormat;

    property IgnoreVersion: boolean read wIgnoreVersion write wIgnoreVersion;

    property Memo: TMemo read wMemo write wMemo;
    property Base: TwBase read fBase write fBase;
    property SettingsCellArray: ArrayOfArrayVariant
      read wSettingsCellArray write wSettingsCellArray;
    property FormatsPrice: TwPriceFormats read fFormatsPrice write fFormatsPrice;

    property FormatOrder: TRecOrderFormat read FFormatOrder write FFormatOrder;

    property ErrorMessage: string read  fErrorMessage write fErrorMessage;
    property OKMessage: string read fOKMessage write fOKMessage;
    property EndThread: boolean read FEndThread write fEndThread;
    property OutStringArr: ArrayOfString read fOutStringArr;

  end;

  { TReadDataOrdersThread }

  TReadDataOrdersThread = class(TwCustomThread)
   private
     FonEndThread: TNotifyEvent;
     FormatOrder: TRecOrderFormat;
     Base: TwBase;
     DataFileName: string;
     DBImport: TwDBImport;

     FWorkBook: TsWorkbook;

     procedure FOpenWorkBook(Sender: TObject);
     procedure fOpenWorkBookInsertInBase(const _Kurs: Double; const _fRemark: String; const _fCountry: String; const _fCustomDeclaration: String;
       const _fScod: String; const _fLabel: String; const _fSum: String; const _fQuantity: String; const _fUnit: String; const _fName: String;
       const _fVendorcode: String);
     procedure ShowStatus;
   protected
     procedure Execute; override;
     procedure SetStatus(aText: string);

   public
    Constructor Create(CreateSuspended : boolean);
  end;

{ TReadDataPriceThread }

TReadDataPriceThread = class(TwCustomThread)
private
  //CatalogArr: ArrayOfArrayVariant;
  FonEndThread: TNotifyEvent;
  fStatus: string;
  fStatusLog: boolean;
  //fRowCount: Integer;
  fRowCurrent: Integer;
  fParent: TComponent;
  Base: TwBase;
  DBImport: TwDBImport;
  FormatRecords: TwPriceFormats;
  FormatRecordsCuttentIndex: integer;
  CountImportedRecords: integer;
  FWorksheet: TsWorksheet;
  FWorkBook: TsWorkbook;
  CollArray: ArrayOfInteger;
  IgnoreVersion: boolean;
  GroupAlgorithm: TGroupAlgoritm;
  GroupInRows: boolean;
  GroupRootIndex: integer;
  GroupCunnrentLevel: integer;

  KURS_AND_PERCENT: string;

  FGroup: TStringList;
  OwnerID: integer;
  TimeStamp: string;
  FormatID: string;
  StockOnly: boolean;
  StockSymbols: boolean;
  IdMainOwner: integer;
  __ARRAYFPREADSHEET: ArrayOfArrayInteger;
  __ARRAYSTOCKSYMBOLS : ArrayOfArrayVariant;
  __IDVENDORCODEVARIANT: Byte;
  __IDSTOCKVARIANT : Byte;
  __IDPRICEVARIANT: Byte;

  function ColorInUsed(_ADataCell: PCell; _GroupColorList: TStringList
    ): integer;
  function FGetStock(aStock: string; aArray: ArrayOfArrayVariant): string;
  procedure FOpenWorkBook(Sender: TObject);
  procedure FReadPriceDataBufStreamMode;
  procedure fReadPriceDataBufStreamModeSQLInsert(const _FCOLOR: Longint; const _PriceNomenclature5: string; const _PriceNomenclature4: string;
    const _StockNomenclature5: string; const _StockNomenclature4: string; const _StockNomenclature3: string; const _StockNomenclature2: string;
    const _PriceNomenclature3: string; const _PriceNomenclature2: string; _IdPL: integer; var ResultRecord: TwResult; const _PriceNomenclature: string;
    const _StockNomenclature: string; const _GroupNomenclature: string; const _PriceNomenclature6: string; const _PriceNomenclature7: string; const _PriceNomenclature8: string
    ; const _PriceNomenclature9: string; const _PriceNomenclature10: string);
  procedure fReadPriceDataBufStreamModeSQLInsertRow(const GroupCunnrentIndex3: integer; const GroupCunnrentIndex2: integer; const GroupCunnrentIndex1: integer;
    const GroupCunnrentIndex: integer; var ResultRecord: TwResult);
  function GetBackground(_ACell: PCell): TsColor;
  function GetFormat(_Row, _Cell: cardinal): TsUsedFormattingFields;
  function GetKursValute(aGet: TwGet): boolean;
  procedure ImportFormat(aIndexFormat: integer; aDataFileName, aCurrentDataFileHash: string; aFormatRecords: TwPriceFormats);
  procedure ImportYML(aYML: TYML);
  function LOadPriceFromInternet(aGet: TwGet; aXMLText: string): boolean;
  procedure ShowStatus;
  procedure StringListClear(_List: TStringList);
  procedure TransferYMLInBase(aYML: TYML; aCategoriesArr: ArrayOfCategories;
    idInBase: Integer; const aCategory: integer);
protected

  procedure Execute; override;
  procedure SetStatus(aText: string; const aStatus: boolean = false);
public
  Constructor Create(CreateSuspended : boolean);
  property onEndThread: TNotifyEvent read FonEndThread write FonEndThread;
end;

{ TUpdateKursValutThread }

TUpdateKursValutThread = class(TThread)
private
  FonEndThread: TNotifyEvent;
  fStatus: string;
  fStatusLogSetcion: boolean;
  fParent: TComponent;
  Base: TwBase;
  DBImport: TwDBImport;
  FormatsArr: ArrayOfArrayVariant;
  procedure ShowStatus;

protected

  procedure Execute; override;
  procedure SetStatus(aText: string; const aStatusLogSetcion: boolean = true);
public
  Constructor Create(CreateSuspended : boolean);
  property onEndThread: TNotifyEvent read FonEndThread write FonEndThread;
end;


const

  // File formats corresponding to the items of the RgFileFormat radiogroup
  // Items in RadioGroup in Export tab match this order
  FILE_FORMATS: array[0..4] of TsSpreadsheetFormat = (
    sfExcel2, sfExcel5, sfExcel8, sfOOXML, sfOpenDocument
    );
  // Spreadsheet files will get the TABLENAME and have one of these extensions.
  FILE_EXT: array[0..4] of string = (
    '_excel2.xls', '_excel5.xls', '.xls', '.xlsx', '.ods');

implementation

{ TReadDataOrdersThread }

procedure TReadDataOrdersThread.FOpenWorkBook(Sender: TObject);
var
  iSpreadSheet, i: Integer;
  FWorksheet: TsWorksheet;
  ADataCell: PCell;
  ARow, ACol,CurrentRow: Cardinal;
  _fVendorcode, _fName, _fUnit, _fQuantity, _fSum, _fLabel, _fScod, _fCustomDeclaration, _fCountry, _fRemark, _OUTCELLTEXT, _ADDRCELLTEXT: String;
  _arr: ArrayOfArrayVariant;
  _Kurs: Double;

  procedure ClearVariables();
  begin
    _fVendorcode:='';
    _fName:='';
    _fUnit:='';
    _fQuantity:='';
    _fSum:='';
    _fLabel:='';
    _fScod:='';
    _fCustomDeclaration:='';
    _fCountry:='';
    _fRemark:='';
  end;
begin

CurrentRow:=0;
_Kurs:=0;
_OUTCELLTEXT:='';
_ADDRCELLTEXT:= '';

ClearVariables();

_arr:= Base.SQLReadArr('CURRENCY',['KURS'],'ID='+IntToStr(FormatOrder.fCURRENCY),'');
if Assigned(_arr) then
   _Kurs:= _arr[0,0];

   _Kurs:= (_Kurs*(1+FormatOrder.fCURRENCYPERCENT/100));


//_KURS_AND_PERCENT:=
try
  for iSpreadSheet:=0 to High(FormatOrder.fSpreadsheets) do begin

    FWorksheet := FWorkBook.GetWorksheetByIndex(FormatOrder.fSpreadsheets[iSpreadSheet,0]-1); // получаем лист по номеру
    _OUTCELLTEXT:= FormatOrder.fOUTCELLTEXT;

    if Length(FormatOrder.fADDRCELLTEXT)>0 then
      _ADDRCELLTEXT:= GetDataString(FWorksheet.GetCell(FormatOrder.fADDRCELLTEXT));

    OutStringArr:=[
                   _OUTCELLTEXT,
                   _ADDRCELLTEXT
                  ];

  if Assigned(FWorksheet) then
    begin
       SetStatus('Обрабатываю лист № '+IntToStr(FormatOrder.fSpreadsheets[iSpreadSheet,0])+' ...');

       for ADataCell in FWorksheet.Cells do
       begin
         ARow := ADataCell^.Row;
         ACol := ADataCell^.Col;

       if ARow >= FormatOrder.fFirstLine-1 then
          begin
            if CurrentRow <> ARow then
            begin
              if ARow mod 100 = 0 then
                   SetStatus('Обработано строк: '+IntToStr(ARow));

              fOpenWorkBookInsertInBase(_Kurs, _fRemark, _fCountry, _fCustomDeclaration, _fScod, _fLabel, _fSum, _fQuantity, _fUnit, _fName, _fVendorcode);

              CurrentRow := ARow;

              ClearVariables();
            end;

           if ACol = FormatOrder.fVendorcode-1 then
             begin
                case FormatOrder.fIDVENDORCODEVARIANT of
                  0: _fVendorcode:= GetDataString(ADataCell);
                  1: _fVendorcode:= GetDataString(ADataCell, vtNumber);
                  2: _fVendorcode:= GetDataString(ADataCell, vtString);
                end;
             end;
           if ACol = FormatOrder.fName-1 then _fName:= GetDataString(ADataCell);
           if ACol = FormatOrder.fUnit-1 then _fUnit:= GetDataString(ADataCell);
           if ACol = FormatOrder.fQuantity-1 then
           begin
             case FormatOrder.fIDSTOCKVARIANT of
               0: _fQuantity:= GetDataString(ADataCell);
               1: _fQuantity:= GetDataString(ADataCell, vtNumber);
             end;
           end;
           if ACol = FormatOrder.fSum-1 then
           begin
             case FormatOrder.fIDPRICEVARIANT of
               0: _fSum:= GetDataString(ADataCell);
               1: _fSum:= GetDataString(ADataCell, vtNumber);
             end;
           end;
           if ACol = FormatOrder.fCustomDeclaration-1 then _fCustomDeclaration:= GetDataString(ADataCell);
           if ACol = FormatOrder.fCountry-1 then _fCountry:= GetDataString(ADataCell);
           if ACol = FormatOrder.fLabel-1 then _fLabel:= GetDataString(ADataCell);
           if ACol = FormatOrder.fScod-1 then _fScod:= GetDataString(ADataCell);
           if ACol = FormatOrder.fRemark-1 then _fRemark:= GetDataString(ADataCell);

          end;
       end;
       fOpenWorkBookInsertInBase(_Kurs, _fRemark, _fCountry, _fCustomDeclaration, _fScod, _fLabel, _fSum, _fQuantity, _fUnit, _fName, _fVendorcode);
     end;

  end;
except
  on E: Exception do begin
    Base.SQLTransactionEnd(false);
    SetStatus('Ошибка: '+E.Message);
  end;
end;

end;

procedure TReadDataOrdersThread.fOpenWorkBookInsertInBase(const _Kurs: Double; const _fRemark: String; const _fCountry: String;
  const _fCustomDeclaration: String; const _fScod: String; const _fLabel: String; const _fSum: String; const _fQuantity: String; const _fUnit: String;
  const _fName: String; const _fVendorcode: String);
var
  _Sum: Double;
  _Quantity: Double;
begin
  TryStrToFloat(_fQuantity, _Quantity);
  TryStrToFloat(_fSum, _Sum);

  _Sum:=_Sum*_Kurs;

  if (Length(_fVendorcode)>0) and (_Quantity>0) then
      Base.SQLInsert('W_TMP_ORDERS_IMPORT', [
        'ORDOWNER',
        'ORDVENDORCODE',
        'ORDNAME',
        'ORDUNIT',
        'ORDQUANTITY',
        'ORDSUM',
        'ORDSCOD',
        'ORDLABEL',
        'ORDCUSTOMSDECLARATION',
        'ORDCOUNTRY',
        'ORDREMARK',
        'MTHID'],
        [
        FormatOrder.IdOwner,
        UTF8Copy(_fVendorcode, 1, 300),
        UTF8Copy(_fName, 1, 1000),
        UTF8Copy(_fUnit, 1, 15),
        _Quantity,
        _Sum,
        UTF8Copy(_fScod, 1, 30),
        UTF8Copy(_fLabel, 1, 255),
        UTF8Copy(_fCustomDeclaration, 1, 300),
        UTF8Copy(_fCountry, 1, 300),
        UTF8Copy(_fRemark, 1, 1000),
        integer(0)
        ]);
end;

procedure TReadDataOrdersThread.ShowStatus;
begin
   DBImport.SetStatus(Status,true);
end;

procedure TReadDataOrdersThread.Execute;
var
  fmterror: Boolean;
  //fUserFormat: TImportFileFormat;
  fmt: TsSpreadsheetFormat;
begin
  try
    fmterror := False;

    //fUserFormat:= mfNONE;
    case FormatOrder.FileFormat of
      1: //xls
      begin
        if pos(FILE_EXT[0], DataFileName) > 0 then
          fmt := sfExcel2
        else
        if pos(FILE_EXT[1], DataFileName) > 0 then
          fmt := sfExcel5
        else
          fmt := sfExcel8;
      end;
      2: //xlsx
        fmt := sfOOXML;
      3: //ods
        fmt := sfOpenDocument;
      4: //csv
        begin
          fmt := sfCSV;
          CSVParams.Encoding:= Base.SQLReadArr('SELECT CODE FROM "CODEPAGETEXT" WHERE ID='+IntTOStr(FormatOrder.fCodePage))[0,0];
          case FormatOrder.fIDCSVDELIMITER of
            0: CSVParams.Delimiter:= ';';
            1: CSVParams.Delimiter:= ';';
            2: CSVParams.Delimiter:= ',';
            3: CSVParams.Delimiter:= '$';
          end;
        end;
      else
        fmterror := True;
    end;

    if fmterror then  raise Exception.Create('Неизвестный формат файла');

    FWorkBook := TsWorkbook.Create();   // импорт с fspspreadsheet

    if FormatOrder.ConvertWithLibre then // если конвертируем, то
    begin
       SetStatus('Конвертация файла с помощью LibreOffice...');
       try
         DataFileName:= ConvertFileWithLibreOffice(DataFileName);
         SetStatus('Конвертация файла с помощью LibreOffice... [OK]');
         fmt:= sfExcel8;
       except
         raise;
       end;
    end;

    FWorkBook.Options := FWorkBook.Options + [boBufStream];
    FWorkbook.OnOpenWorkbook := @FOpenWorkBook;
    SetStatus('Обработка накладной...');



    Base.SQLDelete('W_TMP_ORDERS_IMPORT','ORDOWNER='+IntToStr(FormatOrder.IdOwner),true);

    try
      //if not Base.LongTransaction then Base.LongTransaction:= true;
      FWorkBook.ReadFromFile(UTF8ToSys(DataFileName), fmt);  // все, что не xlsx - fpcspreadsheets
      //Base.SQLTransactionEnd(true);
    finally
      FWorkBook.Free;
      onEndThread(self);
    end;

    SetStatus('Вывод результата...');
  except
    on E: Exception do
    begin
      Base.SQLTransactionEnd(false);
      SetStatus('Ошибка: '+E.Message);
    end;
  end;

end;

procedure TReadDataOrdersThread.SetStatus(aText: string);
begin
  Status := aText;
  Synchronize(@Showstatus);
end;

constructor TReadDataOrdersThread.Create(CreateSuspended: boolean);
begin
  FreeOnTerminate := true;
  inherited Create(CreateSuspended);
end;

{ TUpdateKursValutThread }

procedure TUpdateKursValutThread.ShowStatus;
begin
  DBImport.SetStatus(fStatus,true,fStatusLogSetcion);
end;

procedure TUpdateKursValutThread.Execute;
var
  wGet: TwGet;
  _XMLText, FormatID, _KURS_AND_PERCENT: String;
  xdoc: TXMLDocument;
  _XMLString: TStringStream;
  Node: TDOMNode;
  _TimeStampKurs: DOMString;
  i: Integer;
begin
  try
    //fUserFormat:= mfNONE;
    try

      wGet:= TwGet.Create(fParent);


      SetStatus('Обновляю курсы валют с сайта ЦБ РФ...');
      //TimeStamp:= FormatDateTime('DD.MM.YYYY', now);

      _XMLText:='';

      try
        while Length(_XMLText)=0 do
            _XMLText:= wGet.GetKursValute;
      except
        on E: Exception do
         begin
           _XMLText:='ERROR: '+E.Message;
         end;
      end;

      if UTF8Pos('Error',_XMLText)>0 then SetStatus('Обновляю курсы валют с сайта ЦБ РФ... [ОШИБКА!][Парсинг данных] ['+_XMLText+']')
      else
      begin
            xdoc:=nil;
            _XMLString:= TStringStream.Create(_XMLText);
            _XMLString.SaveToFile(PathLogFiles_Unsafe+'log-kursvalute.txt');

          try

              ReadXMLFile(xdoc,_XMLString);
              Node:= xdoc.FindNode('ValCurs');

              _TimeStampKurs:= TDOMElement(Node).GetAttribute('Date');

              Node:= Node.FirstChild;

              while Assigned(Node) do
              begin
              case TDOMElement(Node).GetAttribute('ID') of
                'R01235':   //USD
                        begin
                          with Node.ChildNodes do
                             begin
                                try
                                  for i:=0 to Count-1 do begin
                                     if Item[i].NodeName = 'Value' then Base.SQLUpdate('CURRENCY',['KURS','FTIMESTAMP'],[Item[i].TextContent,_TimeStampKurs],'IDVALUTE='+QuotedStr(TDOMElement(Node).GetAttribute('ID')));
                                  end;
                                finally
                                  Free;
                                end;
                             end;
                        end;
                'R01239':   //EUR
                        begin
                          with Node.ChildNodes do
                             begin
                                try
                                  for i:=0 to Count-1 do begin
                                     if Item[i].NodeName = 'Value' then Base.SQLUpdate('CURRENCY',['KURS','FTIMESTAMP'],[Item[i].TextContent,_TimeStampKurs],'IDVALUTE='+QuotedStr(TDOMElement(Node).GetAttribute('ID')));
                                  end;
                                finally
                                  Free;
                                end;
                             end;
                        end;
                'R01335':   //KZT
                        begin
                          with Node.ChildNodes do
                             begin
                                try
                                  for i:=0 to Count-1 do begin
                                     if Item[i].NodeName = 'Value' then Base.SQLUpdate('CURRENCY',['KURS','FTIMESTAMP'],[Item[i].TextContent,_TimeStampKurs],'IDVALUTE='+QuotedStr(TDOMElement(Node).GetAttribute('ID')));
                                  end;
                                finally
                                  Free;
                                end;
                             end;
                        end;
                'R01720':   //UAH
                        begin
                          with Node.ChildNodes do
                             begin
                                try
                                  for i:=0 to Count-1 do begin
                                     if Item[i].NodeName = 'Value' then Base.SQLUpdate('CURRENCY',['KURS','FTIMESTAMP'],[Item[i].TextContent,_TimeStampKurs],'IDVALUTE='+QuotedStr(TDOMElement(Node).GetAttribute('ID')));
                                  end;
                                finally
                                  Free;
                                end;
                             end;
                        end;
              end;

                Node:= Node.NextSibling;
              end;
            finally
              xdoc.Free;
              _XMLString.Free;
            end;
            SetStatus('Обновляю курсы валют с сайта ЦБ РФ...[OK]');
      end;

      if Length(FormatsArr) > 0 then
      begin

        SetStatus('Обновление цен прайс-листов согласно текущему курсу валют... Это может занять немного времени.');

        for i := 0 to High(FormatsArr) do
        begin
          Base.LongTransaction:= true;

          //OwnerID:= FormatRecords.Items[i].Idowner;
          FormatID := IntToStr(FormatsArr[i,0]);
          _KURS_AND_PERCENT:= FloatToStr(FormatsArr[i,1]);

          ImportUpdatePRICECALC(Base,_KURS_AND_PERCENT,FormatID);

          SetStatus('Обновлено '+IntToStr(i+1)+' из '+IntToStr(High(FormatsArr)+1));
        end; //for i...

          Base.SQLTransactionEnd(true);

          SetStatus('Обновление курсов валют [ОК]');

          DBImport.fOKMessage:= 'Обновление курсов валют успешно завершено!';

      end;
    finally
      SetStatus('Освобождаю объект загрузчика данных...');
      wGet.Destroy();
      SetStatus('Освобождаю объект загрузчика данных... ОК');

      SetStatus('Завершаю поток обновления...');
      onEndThread(self);
    end;
  except
    on E: Exception do
    begin
       SetStatus('Обновляю курсы валют с сайта ЦБ РФ... [ОШИБКА!] [Error: '+E.Message+']');
       DBImport.fErrorMessage:= E.Message;
       Base.SQLTransactionEnd(false);
       onEndThread(self);
    end;
  end;

end;

procedure TUpdateKursValutThread.SetStatus(aText: string; const aStatusLogSetcion: boolean);
begin
  fStatus := aText;
  fStatusLogSetcion:= aStatusLogSetcion;
  Synchronize(@Showstatus);
end;

constructor TUpdateKursValutThread.Create(CreateSuspended: boolean);
begin
  FreeOnTerminate := true;
  inherited Create(CreateSuspended);
end;

{ TReadDataPriceThread }

procedure TReadDataPriceThread.TransferYMLInBase(aYML:TYML; aCategoriesArr:ArrayOfCategories; idInBase:Integer; const aCategory:integer);
var
   i, iParent, j: integer;
   OffersArr: ArrayOfOffers;
   _Barcode, _Picture, _Name, _Vendorcode: String;
   _Price, _Price2, _Price3: double;
   _Stock, _Stock2, _Stock3,_Stock4,_Stock5, _IDPL: integer;
begin
    iParent:= idInBase;
    i:=0;
    OffersArr:= nil;

    _Price:= 0;
    _Price2:= 0;
    _Price3:= 0;
    _Stock:= 0;
    _Stock2:= 0;
    _Stock3:= 0;
    _Stock4:= 0;
    _Stock5:= 0;
    _Vendorcode:='';

    while Length(aCategoriesArr)>i do
    begin
       if aCategoriesArr[i].parentId = aCategory then
       begin
         if iParent = 0 then iParent:= GroupRootIndex;

         idInBase:= Base.SQLInsert('PL_GROUP', ['IDPARENT', 'NAME', 'IDOWNER','IDFORMATS','FTIMESTAMP'],
                    [iParent, Trim(aYML.DecodeText(aCategoriesArr[i].name,true)), OwnerID ,FormatID ,TimeStamp],'IDOWNER, NAME, IDPARENT', False);

         OffersArr:= aYML.GetOffersByCategory(aCategoriesArr[i].id);
         for j:=0 to High(OffersArr) do
          begin
             Inc(fRowCurrent);
             if fRowCurrent mod 300 = 0 then
                begin
                  SetStatus('Импортировано строк: '+IntToStr(fRowCurrent)+' из '+IntToStr(CountImportedRecords),true);
       //           Application.ProcessMessages;
                end;

            _Barcode:='';
            _Barcode:= Base.MakeStringFromArray(OffersArr[j].barcode);

            if Assigned(OffersArr[j].picture) then _Picture:= OffersArr[j].picture[0] else _Picture:= '';


            case FormatRecords.Items[FormatRecordsCuttentIndex].fYMLQUANTITY of
              0:
                begin
                  _Stock:= 0;
                  _Stock2:= 0;
                  _Stock3:= 0;
                  _Stock4:= 0;
                  _Stock5:= 0;

                  if Assigned(OffersArr[j].outlets) then
                  begin
                    _Stock:= OffersArr[j].outlets[0].instock;

                    if High(OffersArr[j].outlets)>0 then
                      _Stock2:= OffersArr[j].outlets[1].instock;

                    if High(OffersArr[j].outlets)>1 then
                      _Stock3:= OffersArr[j].outlets[2].instock;

                    if High(OffersArr[j].outlets)>2 then
                      _Stock4:= OffersArr[j].outlets[3].instock;

                    if High(OffersArr[j].outlets)>3 then
                      _Stock5:= OffersArr[j].outlets[4].instock;
                  end;
                end;
              1:
                begin
                  _Stock:= OffersArr[j].quantity;
                  _Stock2:= 0;
                  _Stock3:= 0;
                  _Stock4:= 0;
                  _Stock5:= 0;
                end;
            end;
            case FormatRecords.Items[FormatRecordsCuttentIndex].fYMLID of
              0: _Vendorcode:= UTF8Copy(aYML.DecodeText(OffersArr[j].id),1,300);
              1: _Vendorcode:= UTF8Copy(aYML.DecodeText(IntToStr(OffersArr[j].product_id_1c)),1,300);
              2: _Vendorcode:= UTF8Copy(aYML.DecodeText(OffersArr[j].vendorCode),1,300);
            end;

              _Vendorcode:=Trim(_Vendorcode);

            case FormatRecords.Items[FormatRecordsCuttentIndex].fYMLPRICE of
              0:
                begin
                  _Price:= OffersArr[j].price;
                  _Price2:= OffersArr[j].oldprice;
                  _Price3:= 0;
                end;
              1:
                begin
                  _Price:= OffersArr[j].key_partner;
                  _Price2:= OffersArr[j].price;
                  _Price3:= OffersArr[j].oldprice;
                end;
              2:
                begin
                  _Price:= OffersArr[j].oldprice;
                  _Price2:= OffersArr[j].price;
                  _Price3:= 0;
                end;
            end;



            _Name:= '';
            if Length(OffersArr[j].typePrefix)>0 then _Name:=OffersArr[j].typePrefix;
            if Length(_Name)>0 then _Name:= _Name+' '+OffersArr[j].name else _Name:=OffersArr[j].name;
            if Length(_Name)>0 then _Name:= _Name+' '+OffersArr[j].model else _Name:=OffersArr[j].model;

            _Name:= Trim(_Name);
            //_Price:= ReplaceStr(FloatToStr(_Price),',','.');

            if StockSymbols then
            begin
              TryStrToInt(FGetStock('',__ARRAYSTOCKSYMBOLS),_Stock);
              _Stock2:= 0;
              _Stock3:= 0;
              _Stock4:= 0;
              _Stock5:= 0;
            end;

            if StockOnly then // если только в наличии
              begin
                if _Stock>0 then
                begin
                    _IDPL:= Base.SQLInsert('PL_ITEMS',
                      ['IDPL_GROUP',
                      'IDOWNER',
                      'NAME',
                      'UNIT',
                      'LABEL',
                      'VENDORCODE',
                      'FTIMESTAMP',
                      'IDFORMATS',
                      'FURL',
                      'FURLPICTURE',
                      'REMARK',
                      'PRICE',
                      'PRICE2',
                      'PRICE3',
                      'STOCK',
                      'STOCK2',
                      'STOCK3',
                      'STOCK4',
                      'STOCK5'
                      ],
                      [idInBase,
                        OwnerID,
                        UTF8Copy(Trim(aYML.DecodeText(_Name)),1,500),
                        '',
                        UTF8Copy(Trim(aYML.DecodeText(OffersArr[j].vendorCode)),1,300),
                        _Vendorcode,
                        TimeStamp,
                        FormatID,
                        UTF8Copy(OffersArr[j].url,1,1000),
                        UTF8Copy(_Picture,1,1000),
                        UTF8Copy(aYML.DecodeText(OffersArr[j].description),1,3000),
                        _Price,
                        _Price2,
                        _Price3,
                        _Stock,
                        _Stock2,
                        _Stock3,
                        _Stock4,
                        _Stock5
                        ],
                        'IDOWNER, VENDORCODE',
                      false
                      );
                    Base.SQLInsert('PL_VERSIONS',
                      ['IDPL_ITEMS',
                      'IDOWNER',
                      'FTIMESTAMP',
                      'IDFORMATS',
                      'PRICE',
                      'PRICE2',
                      'PRICE3',
                      'STOCK',
                      'STOCK2',
                      'STOCK3',
                      'STOCK4',
                      'STOCK5'
                      ],
                      [_IDPL,
                        OwnerID,
                        TimeStamp,
                        FormatID,
                        _Price,
                        _Price2,
                        _Price3,
                        _Stock,
                        _Stock2,
                        _Stock3,
                        _Stock4,
                        _Stock5
                        ],
                      false
                      );
                end;
              end else
              begin
                _IDPL:= Base.SQLInsert('PL_ITEMS',
                  ['IDPL_GROUP',
                  'IDOWNER',
                  'NAME',
                  'UNIT',
                  'LABEL',
                  //'SCOD',
                  'VENDORCODE',
                  'FTIMESTAMP',
                  'IDFORMATS',
                  'FURL',
                  'FURLPICTURE',
                  'REMARK',
                  'PRICE',
                  'PRICE2',
                  'PRICE3',
                  'STOCK',
                  'STOCK2',
                  'STOCK3',
                  'STOCK4',
                  'STOCK5'
                  ],
                  [idInBase,
                    OwnerID,
                    UTF8Copy(Trim(aYML.DecodeText(_Name)),1,500),
                    '',
                    UTF8Copy(Trim(aYML.DecodeText(OffersArr[j].vendorCode)),1,300),
                    //UTF8Copy(_Barcode,1,13),
                    _Vendorcode,
                    TimeStamp,
                    FormatID,
                    UTF8Copy(OffersArr[j].url,1,1000),
                    UTF8Copy(_Picture,1,1000),
                    UTF8Copy(aYML.DecodeText(OffersArr[j].description),1,3000),
                    _Price,
                    _Price2,
                    _Price3,
                    _Stock,
                    _Stock2,
                    _Stock3,
                    _Stock4,
                    _Stock5
                    ],
                    'IDOWNER, VENDORCODE',
                  false
                  );
                Base.SQLInsert('PL_VERSIONS',
                  ['IDPL_ITEMS',
                  'IDOWNER',
                  'FTIMESTAMP',
                  'IDFORMATS',
                  'PRICE',
                  'PRICE2',
                  'PRICE3',
                  'STOCK',
                  'STOCK2',
                  'STOCK3',
                  'STOCK4',
                  'STOCK5'
                  ],
                  [_IDPL,
                    OwnerID,
                    TimeStamp,
                    FormatID,
                    _Price,
                    _Price2,
                    _Price3,
                    _Stock,
                    _Stock2,
                    _Stock3,
                    _Stock4,
                    _Stock5
                    ],
                  false
                  );
              end;

            if Length(_Barcode)>0 then
                   Base.SQLUpdate('EXECUTE PROCEDURE PL_SET_SCOD('+IntToStr(OwnerID)+','+IntToStr(_IDPL)+','+QuotedStr(_Barcode)+','','')',false);
          end;


         TransferYMLInBase(aYML,aCategoriesArr, idInBase, aCategoriesArr[i].id);
       end;
       Inc(i);
    end;
end;


procedure TReadDataPriceThread.ImportYML(aYML:TYML);
var
  aCategoriesArr: ArrayOfCategories;
  i: Integer;
begin
  fRowCurrent:= 0;
  CountImportedRecords:= High(aYML.Offers)+1;
  aCategoriesArr:= aYML.SortedCategoriesByParentId(aYML.Categories);
  TransferYMLInBase(aYML,aCategoriesArr,0,0);
end;

procedure TReadDataPriceThread.FOpenWorkBook(Sender: TObject);
begin
  try
    FReadPriceDataBufStreamMode();
  except
    raise;
  end;
end;

function TReadDataPriceThread.GetFormat(_Row, _Cell: cardinal): TsUsedFormattingFields;
begin
  Result := FWorksheet.ReadUsedFormatting(FWorksheet.Cells.FindCell(_Row, _Cell));
end;

function TReadDataPriceThread.GetBackground(_ACell: PCell): TsColor;
begin
     Result := FWorksheet.ReadBackground(_ACell).FgColor;
end;

procedure TReadDataPriceThread.StringListClear(_List: TStringList);
var
  i: integer;
begin
  for  i := _List.Count - 1 downto 0 do
    TwData(_List.Objects[i]).Free;

  _List.Clear;
end;

function TReadDataPriceThread.ColorInUsed(_ADataCell: PCell;
  _GroupColorList: TStringList): integer;
var
  i: integer;
  _Color: TsColor;
begin
  _Color := GetBackground(_ADataCell);
  Result := -1;

  for i := _GroupColorList.Count - 1 downto 1 do
  begin
    if _Color = TwData(_GroupColorList.Objects[i]).Color then
    begin
      Result := i;
      exit;
    end;
  end;

end;

function TReadDataPriceThread.FGetStock(aStock: string; aArray: ArrayOfArrayVariant):string;
var
  i: Integer;
  _ValueStock, _ValueArr: String;
begin
  Result:='';
  Result:= UTF8LowerCase(Trim(aStock));

  if Length(aArray)=0 then
  begin
    if not Base.IsNum(Result) then Result:='-1';
    if (Length(aStock)=0) then Result:='0';
    exit;
  end;

  //aStock:= (aStock);

  for i:=0 to High(aArray) do
    if (Length(Result)=0) then
    begin
      if (aArray[i,0]='null') then
            Result:= VarToStr(aArray[i,1]) else
            Result:='0';
    end else
    begin
      _ValueStock:= UTF8LowerCase(VarToStr(aArray[i,0]));
      _ValueArr:= VarToStr(aArray[i,1]);

      if Length(_ValueStock)=1 then
         Result:= StringReplace(Result,_ValueStock,_ValueArr,[rfReplaceAll]) else
      if Result = _ValueStock then
      Result:= StringReplace(Result,_ValueStock,_ValueArr,[]);
    end;

  if not Base.IsNum(Result) then Result:='-1';

end;

procedure TReadDataPriceThread.FReadPriceDataBufStreamMode;
var
  //FWorksheet: TsWorksheet;
  ADataCell: PCell;
  i: integer;
  _bgColorGroup, _bgColorCell: TsColor;
  _FormatsSaved: boolean;

  _ColFromSetings, _LevelFromColor, igroup, DetectGroupColumn,
    CurrentRow: integer;
  ACol, ARow, _CountImportedRecords: cardinal;
  ResultRecord: TwResult;
  GroupCunnrentIndex, GroupCunnrentIndex1,GroupCunnrentIndex2,GroupCunnrentIndex3, iSpreadSheet, iSCods: integer;
  GroupCunnrentName,GroupCunnrentName1,GroupCunnrentName2,GroupCunnrentName3: string;
  aTransit: Longint;

begin

  try
    _CountImportedRecords:= 0;
    CountImportedRecords:= 0;

    for iSpreadSheet:=0 to High(__ARRAYFPREADSHEET) do begin

      FWorksheet := FWorkBook.GetWorksheetByIndex(__ARRAYFPREADSHEET[iSpreadSheet,0]-1); // получаем лист по номеру

    if Assigned(FWorksheet) then
      begin
        SetStatus('Импортирую лист № '+IntToStr(__ARRAYFPREADSHEET[iSpreadSheet,0])+' ...');

        _CountImportedRecords:= FWorksheet.GetCellCountInCol(CollArray[2] - 1);// Rows.Count;

        CountImportedRecords:= CountImportedRecords+_CountImportedRecords;

        _FormatsSaved := False;

        case GroupAlgorithm of
          gaBgPrice: DetectGroupColumn:= 10;
          gaBgIdent: DetectGroupColumn:= 2;
        end;

        //GroupCunnrentLevel:=0;
        GroupCunnrentIndex:=0;
        GroupCunnrentIndex1:=0;
        GroupCunnrentIndex2:=0;
        GroupCunnrentIndex3:=0;
        GroupCunnrentName:='';
        GroupCunnrentName1:='';
        GroupCunnrentName2:='';
        GroupCunnrentName3:='';

        for ADataCell in FWorksheet.Cells do
        begin
          ARow := ADataCell^.Row;
          ACol := ADataCell^.Col;

            if ARow = 0 then
            begin
              CurrentRow := __ARRAYFPREADSHEET[iSpreadSheet,1];//CollArray[1] - 1;
              with ResultRecord do
                 begin
                   fVENDORCODE:='';
                   fFNAME:='';
                   fUNIT:='';
                   fQUANTITY:='';
                   fSTOCK2:='0';
                   fSTOCK3:='0';
                   fSTOCK4:='0';
                   fSTOCK5:='0';
                   fPRICE:='';
                   fPRICE2:='';
                   fPRICE3:='';
                   fPRICE4:='';
                   fPRICE5:='';
                   fPRICE6:='';
                   fPRICE7:='';
                   fPRICE8:='';
                   fPRICE9:='';
                   fPRICE10:='';
                   fLABEL:='';
                   fSCOD:='';
                   fGROUPS:='';
                   fSUBGROUPS1:='';
                   fSUBGROUPS2:='';
                   fSUBGROUPS3:='';
                   fTRANSIT:='';
                   fFURL:='';
                   fFURLPICTURE:='';
                   fFREMARK:='';
                   fFCOLOR:= '';
                 end;
            end
            else
            begin
              if ARow >= CollArray[1] - 1 then
              begin


                if CurrentRow <> ARow then
                begin

                  if ARow mod 300 = 0 then
                       SetStatus('Импортировано строк: '+IntToStr(ARow)+' из '+IntToStr(_CountImportedRecords),true);


    /////////////
                 if ((GroupAlgorithm = gaBgPrice) and (Length(ResultRecord.fPRICE) > 0))
                    or ((GroupAlgorithm = gaBgIdent) and (Length(ResultRecord.fVENDORCODE) > 0))
                    and (UTF8LowerCase(ResultRecord.fFNAME)<>'наименование')
                 then // если цена или идентификатор (зависит от настройки GroupAlgorithmCol) не пусто
                      fReadPriceDataBufStreamModeSQLInsertRow(GroupCunnrentIndex3, GroupCunnrentIndex2, GroupCunnrentIndex1, GroupCunnrentIndex, ResultRecord);
    /////////////

                with ResultRecord do
                   begin
                     fVENDORCODE:='';
                     fFNAME:='';
                     fUNIT:='';
                     fQUANTITY:='';
                     fSTOCK2:='0';
                     fSTOCK3:='0';
                     fSTOCK4:='0';
                     fSTOCK5:='0';
                     fPRICE:='';
                     fPRICE2:='';
                     fPRICE3:='';
                     fPRICE4:='';
                     fPRICE5:='';
                     fPRICE6:='';
                     fPRICE7:='';
                     fPRICE8:='';
                     fPRICE9:='';
                     fPRICE10:='';
                     fLABEL:='';
                     fSCOD:='';
                     fGROUPS:='';
                     fSUBGROUPS1:='';
                     fSUBGROUPS2:='';
                     fSUBGROUPS3:='';
                     fTRANSIT:='';
                     fFURL:='';
                     fFURLPICTURE:='';
                     fFREMARK:='';
                     fFCOLOR:='';
                   end;
                  CurrentRow := ARow;
                end;

                for i := 0 to Length(CollArray) - 1 do
                begin
                  _ColFromSetings := CollArray[i] - 1;

                  //if ARow mod 300 = 0 then
                  //  Application.ProcessMessages;

                  //TsWorksheet(CurCell^.Worksheet).ReadBackground(CurCell).BgColor
                  //wCunnrentGroupName
                  //wCunnrentGroupLevel

                  if ACol = _ColFromSetings then
                  begin
                    case i of
                      2 :  // VENDORCODE
                        begin
                          case __IDVENDORCODEVARIANT of
                            0: ResultRecord.fVENDORCODE:= Trim(GetDataString(ADataCell));
                            1: ResultRecord.fVENDORCODE:= Trim(GetDataString(ADataCell,vtNumber));
                            2: ResultRecord.fVENDORCODE:= Trim(GetDataString(ADataCell,vtString));
                          end;
                        end;
                      3 : ResultRecord.fFNAME:= Trim(GetDataString(ADataCell));// FNAME
                      4 : ResultRecord.fUNIT:= GetDataString(ADataCell);// UNIT
                      5 : // QUANTITY
                      begin
                        case __IDSTOCKVARIANT of
                          0: ResultRecord.fQUANTITY:= GetDataString(ADataCell);
                          1: ResultRecord.fQUANTITY:= GetDataString(ADataCell,vtNumber);
                        end;
                      end;
                      6 : // QUANTITY
                      begin
                        case __IDSTOCKVARIANT of
                          0: ResultRecord.fSTOCK2:= GetDataString(ADataCell);
                          1: ResultRecord.fSTOCK2:= GetDataString(ADataCell,vtNumber);
                        end;
                      end;
                      7 : // QUANTITY
                      begin
                        case __IDSTOCKVARIANT of
                          0: ResultRecord.fSTOCK3:= GetDataString(ADataCell);
                          1: ResultRecord.fSTOCK3:= GetDataString(ADataCell,vtNumber);
                        end;
                      end;
                      8 : // QUANTITY
                      begin
                        case __IDSTOCKVARIANT of
                          0: ResultRecord.fSTOCK4:= GetDataString(ADataCell);
                          1: ResultRecord.fSTOCK4:= GetDataString(ADataCell,vtNumber);
                        end;
                      end;
                      9 : // QUANTITY
                      begin
                        case __IDSTOCKVARIANT of
                          0: ResultRecord.fSTOCK5:= GetDataString(ADataCell);
                          1: ResultRecord.fSTOCK5:= GetDataString(ADataCell,vtNumber);
                        end;
                      end;
                     10 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;
                     11 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE2:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE2:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;
                     12 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE3:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE3:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;
                     13 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE4:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE4:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;
                     14 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE5:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE5:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;
                     15 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE6:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE6:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;
                     16 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE7:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE7:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;
                     17 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE8:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE8:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;
                     18 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE9:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE9:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;
                     19 : // PRICE
                     begin
                       case __IDPRICEVARIANT of
                         0: ResultRecord.fPRICE10:= GetDataString(ADataCell);
                         1: ResultRecord.fPRICE10:= GetDataString(ADataCell,vtNumber);
                       end;
                     end;


                     20 : ResultRecord.fLABEL:= Trim(GetDataString(ADataCell));// LABEL
                     21 : ResultRecord.fSCOD:= GetDataString(ADataCell);// SCOD
                     22 :  // GROUP
                      begin
                        if not GroupInRows then
                        begin
                          if (GroupCunnrentName <> ADataCell^.UTF8StringValue) and
                            (Length(ADataCell^.UTF8StringValue) > 0) then
                          begin
                            GroupCunnrentName := UTF8Copy(Trim(ADataCell^.UTF8StringValue),1,255);
                            GroupCunnrentIndex :=
                              Base.SQLInsert('PL_GROUP', ['IDPARENT', 'NAME', 'IDOWNER','IDFORMATS','FTIMESTAMP'],
                              [GroupRootIndex, GroupCunnrentName, OwnerID ,FormatID ,TimeStamp],'IDOWNER, NAME, IDPARENT', False);
                            GroupCunnrentLevel := 0;
                          end;
                        end else
                        begin
                          if not _FormatsSaved then
                           begin
                            _bgColorGroup := GetBackground(ADataCell);
                            _FormatsSaved := True;
                           end;

                           _bgColorCell:= GetBackground(ADataCell);

                           if (_bgColorGroup = _bgColorCell) and (Length(GetDataString(FWorksheet.Cells.FindCell(ADataCell^.Row, CollArray[DetectGroupColumn]-1))) = 0) then
                           begin
                            StringListClear(FGroup);

                            GroupCunnrentName := UTF8Copy(Trim(ADataCell^.UTF8StringValue),1,255);
                            GroupCunnrentIndex:= Base.SQLInsert('PL_GROUP',['IDPARENT','NAME','IDOWNER','IDFORMATS','FTIMESTAMP'],[GroupRootIndex,GroupCunnrentName,OwnerID ,FormatID ,TimeStamp],'IDOWNER, NAME, IDPARENT',false);

                            GroupCunnrentLevel := FGroup.AddObject(IntToStr(GroupRootIndex), TwData.Create(GroupCunnrentIndex,_bgColorCell));

                           end
                           else   // иначе
                           begin
                               // GroupAlgorithmCol / проверяем не пусто ли 0 - цена, 1 - идентификатор
                              if (Length(GetDataString(FWorksheet.Cells.FindCell(ADataCell^.Row, CollArray[DetectGroupColumn]-1))) = 0) and (Length(ADataCell^.UTF8StringValue)>0) then
                              begin
                                _LevelFromColor := ColorInUsed(ADataCell, FGroup);

                                if _LevelFromColor > 0 then
                                begin

                                  GroupCunnrentName := UTF8Copy(Trim(ADataCell^.UTF8StringValue),1,255);
                                  GroupCunnrentLevel:= _LevelFromColor;

                                  for igroup:= FGroup.Count-1 downto GroupCunnrentLevel+1 do
                                   begin
                                     TwData(FGroup.Objects[igroup]).Free;
                                     FGroup.Delete(igroup);
                                   end;
                                  GroupCunnrentIndex:= Base.SQLInsert('PL_GROUP',['IDPARENT','NAME','IDOWNER','IDFORMATS','FTIMESTAMP'],[FGroup.Strings[GroupCunnrentLevel],GroupCunnrentName,OwnerID ,FormatID ,TimeStamp],'IDOWNER, NAME, IDPARENT',false);

                                  TwData(FGroup.Objects[GroupCunnrentLevel]).ID:=GroupCunnrentIndex;
                                end
                                else
                                begin
                                  GroupCunnrentName := UTF8Copy(Trim(ADataCell^.UTF8StringValue),1,255);
                                  GroupCunnrentIndex:= Base.SQLInsert('PL_GROUP',['IDPARENT','NAME','IDOWNER','IDFORMATS','FTIMESTAMP'],[TwData(FGroup.Objects[GroupCunnrentLevel]).ID,GroupCunnrentName,OwnerID ,FormatID ,TimeStamp],'IDOWNER, NAME, IDPARENT',false);


                                  GroupCunnrentLevel := FGroup.AddObject(IntToStr(TwData(FGroup.Objects[GroupCunnrentLevel]).ID), TwData.Create(GroupCunnrentIndex,_bgColorCell));
                                end;
                              end;
                           end;
                        end;
                      end;
                      23:  // SUBGROUPS1
                      begin
                       if not GroupInRows then
                       begin
                          if (GroupCunnrentName1 <> ADataCell^.UTF8StringValue) and
                            (Length(ADataCell^.UTF8StringValue) > 0) then
                          begin
                            GroupCunnrentName1 := UTF8Copy(Trim(ADataCell^.UTF8StringValue),1,255);
                            GroupCunnrentIndex1 :=
                              Base.SQLInsert('PL_GROUP', ['IDPARENT', 'NAME', 'IDOWNER','IDFORMATS','FTIMESTAMP'],
                              [GroupCunnrentIndex, GroupCunnrentName1, OwnerID ,FormatID ,TimeStamp],'IDOWNER, NAME, IDPARENT', False);
                            GroupCunnrentLevel := 1;
                          end;
                       end;
                      end;
                      24:  // SUBGROUPS2
                      begin
                          if not GroupInRows then
                          begin
                            if (GroupCunnrentName2 <> ADataCell^.UTF8StringValue) and
                              (Length(ADataCell^.UTF8StringValue) > 0) then
                            begin
                              GroupCunnrentName2 := UTF8Copy(Trim(ADataCell^.UTF8StringValue),1,255);
                              GroupCunnrentIndex2 :=
                                Base.SQLInsert('PL_GROUP', ['IDPARENT', 'NAME', 'IDOWNER','IDFORMATS','FTIMESTAMP'],
                                [GroupCunnrentIndex1, GroupCunnrentName2, OwnerID ,FormatID ,TimeStamp],'IDOWNER, NAME, IDPARENT', False);
                              GroupCunnrentLevel := 2;
                            end;
                          end;
                      end;
                      25:  // SUBGROUPS3
                      begin
                          if not GroupInRows then
                          begin
                            if (GroupCunnrentName3 <> ADataCell^.UTF8StringValue) and
                              (Length(ADataCell^.UTF8StringValue) > 0) then
                            begin
                              GroupCunnrentName3 := UTF8Copy(Trim(ADataCell^.UTF8StringValue),1,255);
                              GroupCunnrentIndex3 :=
                                Base.SQLInsert('PL_GROUP', ['IDPARENT', 'NAME', 'IDOWNER','IDFORMATS','FTIMESTAMP'],
                                [GroupCunnrentIndex2, GroupCunnrentName3, OwnerID ,FormatID ,TimeStamp],'IDOWNER, NAME, IDPARENT', False);
                              GroupCunnrentLevel := 3;
                            end;
                          end;
                      end;
                      26:
                      begin
                        ResultRecord.fTRANSIT:= GetDataString(ADataCell);// TRANSIT
                        if TryStrToInt(ResultRecord.fTRANSIT,aTransit) then
                           if aTransit>0 then
                             ResultRecord.fTRANSIT:=IntToStr(aTransit)
                           else
                             ResultRecord.fTRANSIT:='';
                      end;
                      27: ResultRecord.fFURL:= GetDataString(ADataCell);// fFURL
                      28: ResultRecord.fFURLPICTURE:= GetDataString(ADataCell);// fFURLPICTURE
                      29: ResultRecord.fFREMARK:= GetDataString(ADataCell);// FREMARK
                      30: ResultRecord.fFCOLOR:= GetDataString(ADataCell);// FCOLOR
                    end;
                  end;
                end; // for i:=0...
              end;
            end;
        end;

        /////////////
         if ((GroupAlgorithm = gaBgPrice) and (Length(ResultRecord.fPRICE) > 0))
            or ((GroupAlgorithm = gaBgIdent) and (Length(ResultRecord.fVENDORCODE) > 0))
            and (UTF8LowerCase(ResultRecord.fFNAME)<>'наименование')
         then // если цена или идентификатор (зависит от настройки GroupAlgorithmCol) не пусто
              fReadPriceDataBufStreamModeSQLInsertRow(GroupCunnrentIndex3, GroupCunnrentIndex2, GroupCunnrentIndex1, GroupCunnrentIndex, ResultRecord);
        /////////////
        StringListClear(FGroup);
      end else
         SetStatus('Импортирую лист № '+IntToStr(__ARRAYFPREADSHEET[iSpreadSheet,0])+' [ОШИБКА - лист не найден в документе!]');
    end; // End перебор листов
  except
    on E: Exception do
    begin
      SetStatus('Ошибка [FReadPriceDataBufStreamMode]: "' + E.Message + '"');
      __Log.SaveLogError(E);
      SetStatus('Ошибка [FReadPriceDataBufStreamMode]: "' + E.Message + '"');
      raise;
    end;
  end;
end;

procedure TReadDataPriceThread.fReadPriceDataBufStreamModeSQLInsert(const _FCOLOR: Longint; const _PriceNomenclature5: string;
  const _PriceNomenclature4: string; const _StockNomenclature5: string; const _StockNomenclature4: string; const _StockNomenclature3: string;
  const _StockNomenclature2: string; const _PriceNomenclature3: string; const _PriceNomenclature2: string; _IdPL: integer; var ResultRecord: TwResult;
  const _PriceNomenclature: string; const _StockNomenclature: string; const _GroupNomenclature: string; const _PriceNomenclature6: string;
  const _PriceNomenclature7: string; const _PriceNomenclature8: string; const _PriceNomenclature9: string; const _PriceNomenclature10: string);
begin
  _IdPL:= Base.SQLInsert('PL_ITEMS',
    ['IDPL_GROUP',
    'IDOWNER',
    'NAME',
    'UNIT',
    'LABEL',
    'VENDORCODE',
    'FTIMESTAMP',
    'IDFORMATS',
    'FURL',
    'FURLPICTURE',
    'REMARK',
    'FCOLOR',
    'PRICE',
    'PRICE2',
    'PRICE3',
    'PRICE4',
    'PRICE5',
    'PRICE6',
    'PRICE7',
    'PRICE8',
    'PRICE9',
    'PRICE10',
    'STOCK',
    'STOCK2',
    'STOCK3',
    'STOCK4',
    'STOCK5',
    'TRANSIT'],
    [_GroupNomenclature,
      OwnerID,
      UTF8Copy(ResultRecord.fFNAME, 1, 500),
      UTF8Copy(ResultRecord.fUNIT, 1, 15),
      UTF8Copy(ResultRecord.fLABEL, 1, 255),
      //UTF8Copy(ResultRecord.fSCOD,1,13),
      UTF8Copy(ResultRecord.fVENDORCODE, 1, 300),
      TimeStamp,
      FormatID,
      UTF8Copy(ResultRecord.fFURL, 1, 1000),
      UTF8Copy(ResultRecord.fFURLPICTURE, 1, 1000),
      UTF8Copy(ResultRecord.fFREMARK, 1, 3000),
      _FCOLOR,
      _PriceNomenclature,
      _PriceNomenclature2,
      _PriceNomenclature3,
      _PriceNomenclature4,
      _PriceNomenclature5,
      _PriceNomenclature6,
      _PriceNomenclature7,
      _PriceNomenclature8,
      _PriceNomenclature9,
      _PriceNomenclature10,
      _StockNomenclature,
      _StockNomenclature2,
      _StockNomenclature3,
      _StockNomenclature4,
      _StockNomenclature5,
       UTF8Copy(ResultRecord.fTRANSIT, 1, 120)],
      'IDOWNER,VENDORCODE',
    false
    );

   Base.SQLInsert('PL_VERSIONS',
     ['IDPL_ITEMS',
     'IDOWNER',
     'PRICE',
     'PRICE2',
     'PRICE3',
     'PRICE4',
     'PRICE5',
     'PRICE6',
     'PRICE7',
     'PRICE8',
     'PRICE9',
     'PRICE10',
     'STOCK',
     'STOCK2',
     'STOCK3',
     'STOCK4',
     'STOCK5',
     'FTIMESTAMP',
     'IDFORMATS',
     'TRANSIT'],
     [_IdPL,
       OwnerID,
       _PriceNomenclature,
       _PriceNomenclature2,
       _PriceNomenclature3,
       _PriceNomenclature4,
       _PriceNomenclature5,
       _PriceNomenclature6,
       _PriceNomenclature7,
       _PriceNomenclature8,
       _PriceNomenclature9,
       _PriceNomenclature10,
       _StockNomenclature,
       _StockNomenclature2,
       _StockNomenclature3,
       _StockNomenclature4,
       _StockNomenclature5,
       TimeStamp,
       FormatID,
       UTF8Copy(ResultRecord.fTRANSIT, 1, 120)],
     false
     );

  if Length(ResultRecord.fSCOD)>0 then
         Base.SQLUpdate('EXECUTE PROCEDURE PL_SET_SCOD('+IntToStr(OwnerID)+','+IntToStr(_IdPL)+','+QuotedStr(ResultRecord.fSCOD)+','','')', false);
end;

procedure TReadDataPriceThread.fReadPriceDataBufStreamModeSQLInsertRow(const GroupCunnrentIndex3: integer; const GroupCunnrentIndex2: integer;
  const GroupCunnrentIndex1: integer; const GroupCunnrentIndex: integer; var ResultRecord: TwResult);
var
  _FCOLOR: Longint;
  _PriceNomenclature5: string;
  _PriceNomenclature4: string;
  _StockNomenclature5: string;
  _StockNomenclature4: string;
  _StockNomenclature3: string;
  _StockNomenclature2: string;
  _PriceNomenclature3: string;
  _PriceNomenclature2: string;
  _IdPL: integer;
  _PriceNomenclature: string;
  _StockNomenclature: string;
  _GroupNomenclature: string;
  _PriceNomenclature6: string;
  _PriceNomenclature7: string;
  _PriceNomenclature8: string;
  _PriceNomenclature9: string;
  _PriceNomenclature10: string;
begin
  _GroupNomenclature:='';

  try
    if not GroupInRows then
    begin
      case GroupCunnrentLevel of
        0: _GroupNomenclature := IntTOStr(GroupCunnrentIndex);
        1: _GroupNomenclature := IntTOStr(GroupCunnrentIndex1);
        2: _GroupNomenclature := IntTOStr(GroupCunnrentIndex2);
        3: _GroupNomenclature := IntTOStr(GroupCunnrentIndex3);
        else
          _GroupNomenclature := IntTOStr(GroupRootIndex);
      end;
    end else
    begin
      if GroupCunnrentLevel < FGroup.Count then
         _GroupNomenclature := IntToStr(TwData(FGroup.Objects[GroupCunnrentLevel]).ID);
    end;


    _PriceNomenclature := ReplaceStr(ResultRecord.fPRICE, '.', ',');
    if not Base.IsNum(_PriceNomenclature) then
       _PriceNomenclature:='0';
    _PriceNomenclature := ReplaceStr(_PriceNomenclature, ',', '.');

    _PriceNomenclature2 := ReplaceStr(ResultRecord.fPRICE2, '.', ',');
    if not Base.IsNum(_PriceNomenclature2) then
       _PriceNomenclature2:='0';
    _PriceNomenclature2 := ReplaceStr(_PriceNomenclature2, ',', '.');

    _PriceNomenclature3 := ReplaceStr(ResultRecord.fPRICE3, '.', ',');
    if not Base.IsNum(_PriceNomenclature3) then
       _PriceNomenclature3:='0';
    _PriceNomenclature3 := ReplaceStr(_PriceNomenclature3, ',', '.');

    _PriceNomenclature4 := ReplaceStr(ResultRecord.fPRICE4, '.', ',');
    if not Base.IsNum(_PriceNomenclature4) then
       _PriceNomenclature4:='0';
    _PriceNomenclature4 := ReplaceStr(_PriceNomenclature4, ',', '.');

    _PriceNomenclature5 := ReplaceStr(ResultRecord.fPRICE5, '.', ',');
    if not Base.IsNum(_PriceNomenclature5) then
       _PriceNomenclature5:='0';
    _PriceNomenclature5 := ReplaceStr(_PriceNomenclature5, ',', '.');

    _PriceNomenclature6 := ReplaceStr(ResultRecord.fPRICE6, '.', ',');
    if not Base.IsNum(_PriceNomenclature6) then
       _PriceNomenclature6:='0';
    _PriceNomenclature6 := ReplaceStr(_PriceNomenclature6, ',', '.');

    _PriceNomenclature7 := ReplaceStr(ResultRecord.fPRICE7, '.', ',');
    if not Base.IsNum(_PriceNomenclature7) then
       _PriceNomenclature7:='0';
    _PriceNomenclature7 := ReplaceStr(_PriceNomenclature7, ',', '.');

    _PriceNomenclature8 := ReplaceStr(ResultRecord.fPRICE8, '.', ',');
    if not Base.IsNum(_PriceNomenclature8) then
       _PriceNomenclature8:='0';
    _PriceNomenclature8 := ReplaceStr(_PriceNomenclature8, ',', '.');

    _PriceNomenclature9 := ReplaceStr(ResultRecord.fPRICE9, '.', ',');
    if not Base.IsNum(_PriceNomenclature9) then
       _PriceNomenclature9:='0';
    _PriceNomenclature9 := ReplaceStr(_PriceNomenclature9, ',', '.');

    _PriceNomenclature10 := ReplaceStr(ResultRecord.fPRICE10, '.', ',');
    if not Base.IsNum(_PriceNomenclature10) then
       _PriceNomenclature10:='0';
    _PriceNomenclature10 := ReplaceStr(_PriceNomenclature10, ',', '.');

    _StockNomenclature:= FGetStock(ReplaceStr(ResultRecord.fQUANTITY, ',', '.'), __ARRAYSTOCKSYMBOLS);
    _StockNomenclature2:= FGetStock(ReplaceStr(ResultRecord.fSTOCK2, ',', '.'), __ARRAYSTOCKSYMBOLS);
    _StockNomenclature3:= FGetStock(ReplaceStr(ResultRecord.fSTOCK3, ',', '.'), __ARRAYSTOCKSYMBOLS);
    _StockNomenclature4:= FGetStock(ReplaceStr(ResultRecord.fSTOCK4, ',', '.'), __ARRAYSTOCKSYMBOLS);
    _StockNomenclature5:= FGetStock(ReplaceStr(ResultRecord.fSTOCK5, ',', '.'), __ARRAYSTOCKSYMBOLS);

    _IdPL:=0;
    TryStrToInt(ResultRecord.fFCOLOR, _FCOLOR);

    if Length(_GroupNomenclature)=0 then
    begin
    __Log.Add('DBImport', 'GroupNomenclature ='+_GroupNomenclature);
    __Log.Add('DBImport', 'OwnerID ='+IntToStr(OwnerID));
    __Log.Add('DBImport', 'fFNAME ='+ResultRecord.fFNAME);
    __Log.Add('DBImport', 'fUNIT ='+ResultRecord.fUNIT);
    __Log.Add('DBImport', 'fLABEL ='+ResultRecord.fLABEL);
    __Log.Add('DBImport', 'fVENDORCODE ='+ResultRecord.fVENDORCODE);
    __Log.Add('DBImport', 'TimeStamp ='+TimeStamp);
    __Log.Add('DBImport', 'FormatID ='+FormatID);
    __Log.Add('DBImport', 'fFURL ='+ResultRecord.fFURL);
    __Log.Add('DBImport', 'fFURLPICTURE ='+ResultRecord.fFURLPICTURE);
    __Log.Add('DBImport', 'fFREMARK ='+ResultRecord.fFREMARK);
    __Log.Add('DBImport', 'FCOLOR ='+IntToStr(_FCOLOR));
    __Log.Add('DBImport', 'PriceNomenclature ='+_PriceNomenclature);
    __Log.Add('DBImport', 'PriceNomenclature2 ='+_PriceNomenclature2);
    __Log.Add('DBImport', 'PriceNomenclature3 ='+_PriceNomenclature3);
    __Log.Add('DBImport', 'PriceNomenclature4 ='+_PriceNomenclature4);
    __Log.Add('DBImport', 'PriceNomenclature5 ='+_PriceNomenclature5);
    __Log.Add('DBImport', 'PriceNomenclature6 ='+_PriceNomenclature6);
    __Log.Add('DBImport', 'PriceNomenclature7 ='+_PriceNomenclature7);
    __Log.Add('DBImport', 'PriceNomenclature8 ='+_PriceNomenclature8);
    __Log.Add('DBImport', 'PriceNomenclature9 ='+_PriceNomenclature9);
    __Log.Add('DBImport', 'PriceNomenclature10 ='+_PriceNomenclature10);
    __Log.Add('DBImport', 'StockNomenclature ='+_StockNomenclature);
    __Log.Add('DBImport', 'StockNomenclature2 ='+_StockNomenclature2);
    __Log.Add('DBImport', 'StockNomenclature3 ='+_StockNomenclature3);
    __Log.Add('DBImport', 'StockNomenclature4 ='+_StockNomenclature4);
    __Log.Add('DBImport', 'StockNomenclature5 ='+_StockNomenclature5);
    __Log.Add('DBImport', 'fTRANSIT ='+ResultRecord.fTRANSIT);
      raise Exception.Create('Ошибка определения группы позиции! Проверьте формат прайс-листа.');
    end;

    if StockOnly then // если только в наличии
      begin
        if (StrToFLoat(_StockNomenclature)>0)
           or (StrToFLoat(_StockNomenclature2)>0)
           or (StrToFLoat(_StockNomenclature3)>0)
           or (StrToFLoat(_StockNomenclature4)>0)
           or (StrToFLoat(_StockNomenclature5)>0) then
              fReadPriceDataBufStreamModeSQLInsert(_FCOLOR, _PriceNomenclature5, _PriceNomenclature4, _StockNomenclature5, _StockNomenclature4,
                _StockNomenclature3, _StockNomenclature2, _PriceNomenclature3, _PriceNomenclature2, _IdPL, ResultRecord, _PriceNomenclature,
                _StockNomenclature, _GroupNomenclature, _PriceNomenclature6, _PriceNomenclature7, _PriceNomenclature8, _PriceNomenclature9, _PriceNomenclature10);

      end else
          fReadPriceDataBufStreamModeSQLInsert(_FCOLOR, _PriceNomenclature5, _PriceNomenclature4, _StockNomenclature5, _StockNomenclature4,
            _StockNomenclature3, _StockNomenclature2, _PriceNomenclature3, _PriceNomenclature2, _IdPL, ResultRecord, _PriceNomenclature,
            _StockNomenclature, _GroupNomenclature, _PriceNomenclature6, _PriceNomenclature7, _PriceNomenclature8, _PriceNomenclature9, _PriceNomenclature10);
  except
    on E: Exception do
    begin
      __Log.SaveLogError(E);
      raise;
    end;
  end;
end;

procedure TReadDataPriceThread.ShowStatus;
begin
  DBImport.SetStatus(fStatus,fStatusLog);
end;

function TReadDataPriceThread.GetKursValute(aGet:TwGet):boolean;
var
  _XMLText: String;
  xdoc: TXMLDocument;
  _XMLString: TStringStream;
  _TimeStampKurs: DOMString;
  Node: TDOMNode;
  i: Integer;
begin
  Result:= false;
  _XMLText:='';

  try
    while Length(_XMLText)=0 do
        _XMLText:= aGet.GetKursValute;
  except
    on E: Exception do
     begin
       _XMLText:='ERROR: '+E.Message;
     end;
  end;


  if UTF8Pos('Error',_XMLText)>0 then SetStatus('Обновляю курсы валют с сайта ЦБ РФ... [ОШИБКА!] ['+_XMLText+']')
  else
  begin
      try
        xdoc:=nil;
        _XMLString:= TStringStream.Create(_XMLText);
        _XMLString.SaveToFile(PathLogFiles_Unsafe+'log-kursvalute.txt');
      try

          ReadXMLFile(xdoc,_XMLString);
          Node:= xdoc.FindNode('ValCurs');

          _TimeStampKurs:= TDOMElement(Node).GetAttribute('Date');

          Node:= Node.FirstChild;

          while Assigned(Node) do
          begin
          case TDOMElement(Node).GetAttribute('ID') of
            'R01235':   //USD
                    begin
                      with Node.ChildNodes do
                         begin
                            try
                              for i:=0 to Count-1 do begin
                                 if Item[i].NodeName = 'Value' then Base.SQLUpdate('CURRENCY',['KURS','FTIMESTAMP'],[Item[i].TextContent,_TimeStampKurs],'IDVALUTE='+QuotedStr(TDOMElement(Node).GetAttribute('ID')));
                              end;
                            finally
                              Free;
                            end;
                         end;
                    end;
            'R01239':   //EUR
                    begin
                      with Node.ChildNodes do
                         begin
                            try
                              for i:=0 to Count-1 do begin
                                 if Item[i].NodeName = 'Value' then Base.SQLUpdate('CURRENCY',['KURS','FTIMESTAMP'],[Item[i].TextContent,_TimeStampKurs],'IDVALUTE='+QuotedStr(TDOMElement(Node).GetAttribute('ID')));
                              end;
                            finally
                              Free;
                            end;
                         end;
                    end;
            'R01335':   //KZT
                    begin
                      with Node.ChildNodes do
                         begin
                            try
                              for i:=0 to Count-1 do begin
                                 if Item[i].NodeName = 'Value' then Base.SQLUpdate('CURRENCY',['KURS','FTIMESTAMP'],[Item[i].TextContent,_TimeStampKurs],'IDVALUTE='+QuotedStr(TDOMElement(Node).GetAttribute('ID')));
                              end;
                            finally
                              Free;
                            end;
                         end;
                    end;
            'R01720':   //UAH
                    begin
                      with Node.ChildNodes do
                         begin
                            try
                              for i:=0 to Count-1 do begin
                                 if Item[i].NodeName = 'Value' then Base.SQLUpdate('CURRENCY',['KURS','FTIMESTAMP'],[Item[i].TextContent,_TimeStampKurs],'IDVALUTE='+QuotedStr(TDOMElement(Node).GetAttribute('ID')));
                              end;
                            finally
                              Free;
                            end;
                         end;
                    end;
          end;

            Node:= Node.NextSibling;
          end;
        finally
          xdoc.Free;
          _XMLString.Free;
        end;
        Result:= true;
      except
        on E: Exception do
        begin
           SetStatus('Обновляю курсы валют с сайта ЦБ РФ... [ОШИБКА!] [Error: '+E.Message+']')
        end;
      end;
  end;
end;

function TReadDataPriceThread.LOadPriceFromInternet(aGet: TwGet; aXMLText: string): boolean;
var
  _arr: ArrayArrayOfString;
  _Pos: PtrInt;
  i: Integer;
begin
  Result:= false;

  try
    try
      _arr:= aGet.ExecuteXML(aXMLText);
    except
      on E: Exception do
      begin
        __Log.SaveLogError(E);
        SetStatus('Ошибка [ExecuteXML]: "' + E.Message + '"');
        exit;
      end;
    end;

    for i:=0 to High(_arr) do
    begin
      _Pos:= UTF8Pos('File loaded in ',_arr[i,1]);
      if _Pos>0 then
         begin
           Result:= true;
           Break;
         end;
    end;

  finally
    aXMLText:=_arr[0,0];
    _arr:= nil;
    if not Result then
        SetStatus('Загружаю файл из сети интернет... [ОШИБКА!] ['+aXMLText+']');
  end;
end;

procedure TReadDataPriceThread.ImportFormat(aIndexFormat:integer; aDataFileName,aCurrentDataFileHash: string; aFormatRecords:TwPriceFormats);
var
  fmterror: Boolean;
  //ext: String;
  fmt: TsSpreadsheetFormat;
  fUserFormat: TFileFormat;
  wYML: TYML;
  _FileStream: TStringList;
  _PosXMLStart: PtrInt;
  i, k: Integer;
  _S: String;
begin

fmterror := False;
//ext := lowercase(ExtractFileExt(aDataFileName));
fUserFormat:= ffNONE;
case FormatRecords.Items[aIndexFormat].fIDFILEFORMAT of
  1: //xls
  begin
    if pos(FILE_EXT[0], aDataFileName) > 0 then
      fmt := sfExcel2
    else
    if pos(FILE_EXT[1], aDataFileName) > 0 then
      fmt := sfExcel5
    else
      fmt := sfExcel8;
  end;
  2: //xlsx
    fmt := sfOOXML;
  3: //ods
    fmt := sfOpenDocument;
  4: //csv
    begin
      fmt := sfCSV;
      CSVParams.Encoding:= Base.SQLReadArr('SELECT CODE FROM "CODEPAGETEXT" WHERE ID='+IntTOStr(FormatRecords.Items[aIndexFormat].fIDCODEPAGETEXT))[0,0];
    end;
  6: //YML
    begin
      fmt := sfUser;
      fUserFormat:=ffYML;
      CSVParams.Encoding:= Base.SQLReadArr('SELECT CODE FROM "CODEPAGETEXT" WHERE ID='+IntTOStr(FormatRecords.Items[aIndexFormat].fIDCODEPAGETEXT))[0,0];
    end
  else
  begin
    fmterror := True;
  end;
end;

 try

    if fmterror then  Exception.Create('Неизвестный формат файла');

    FWorkBook := TsWorkbook.Create();   // импорт с fspspreadsheet

    if (fmt <> sfUser) and (fmt <> sfCSV) then
      begin
        if FormatRecords.Items[aIndexFormat].fGROUPSINROWS = 1 then
        begin
          GroupInRows := True;
          SetStatus('Группы в строках: ВКЛ');
        end
        else
        begin
           GroupInRows := False;  // если группы в строках, то устанавливаем флаг
           SetStatus('Группы в строках: ВЫКЛ');
        end;
      end else
        GroupInRows := False;  // если группы в строках, то устанавливаем флаг

    if FormatRecords.Items[aIndexFormat].fSTOCKONLY = 1 then
    begin
        StockOnly:= true;
        SetStatus('Только в наличии: ВКЛ');
    end
    else
    begin
        StockOnly:= false;
        SetStatus('Только в наличии: ВЫКЛ');
    end;

    if Length(FormatRecords.Items[aIndexFormat].fSTOCKSYMBOLS)>0 then
    begin
      SetStatus('Автозамена остатка в прайс-листе [ВКЛ]');
      StockSymbols:= true;
    end else
      StockSymbols:= false;

    if FormatRecords.Items[aIndexFormat].fFCONVERTLIBRE = 1 then // если конвертируем, то
    begin
       SetStatus('Конвертация файла с помощью LibreOffice...');
       try
         aDataFilename:= ConvertFileWithLibreOffice(aDataFilename);
         SetStatus('Конвертация файла с помощью LibreOffice... [OK]');
         fmt:= sfExcel8;
       except
         raise;
       end;
    end;

    if (fmt <> sfUser) and (fmt <> sfCSV) then
      case FormatRecords.Items[aIndexFormat].fGROUPALGORITHM of // алгоритм поиска группы (только для группы в строках)
        0:
        begin
          GroupAlgorithm:= gaBgPrice;  // если цена не пусто
          if GroupInRows then SetStatus('Алгоритм поиска группы: Фон + Цена');
        end;
        1:
        begin
          GroupAlgorithm:= gaBgIdent; // если идентификатор не пусто
          if GroupInRows then SetStatus('Алгоритм поиска группы: Фон + Идент.');
        end;
      end;

     //FWorkBook.Options.CS;
     FWorkBook.Options := FWorkBook.Options + [boBufStream];

     case FormatRecords.Items[aIndexFormat].fIDCSVDELIMITER of
       0: CSVParams.Delimiter:= ';';
       1: CSVParams.Delimiter:= ';';
       2: CSVParams.Delimiter:= ',';
       3: CSVParams.Delimiter:= '$';
     end;

     FWorkbook.OnOpenWorkbook := @FOpenWorkBook;


      if FormatRecords.Items[aIndexFormat].fSTORAGEDAYS = 0 then  //STORAGEDAYS
      begin
          SetStatus('Очистка прайс-листа контрагента: ВКЛ');
          try
            Base.SQLDelete('PRICELISTS_TIMESTAMPS', 'IDOWNER=' + IntToStr(OwnerID)+' AND IDFORMATS='+FormatID, false);
            Base.SQLDelete('PL_VERSIONS', 'IDOWNER=' +  IntToStr(OwnerID)+' AND IDFORMATS='+FormatID, false);
            Base.SQLDelete('PRICELISTS_TIMESTAMPS', 'IDOWNER=' +  IntToStr(OwnerID)+' AND IDFORMATS='+FormatID, false);

          except
              SetStatus('Очистка прайс-листа... [ERROR]');
              raise;
          end;
      end else
      begin
        SetStatus('Включена очистка прайс-листа контрагента старше: '+IntToStr(FormatRecords.Items[aIndexFormat].fSTORAGEDAYS)+' дн.');
        try
          Base.SQLUpdate('DELETE FROM "PRICELISTS_TIMESTAMPS" WHERE IDOWNER='+ IntToStr(OwnerID)+' AND IDFORMATS='+FormatID+' AND FTIMESTAMP<DATEADD(DAY, -'+IntToStr(FormatRecords.Items[aIndexFormat].fSTORAGEDAYS)+', CURRENT_DATE)',false);
          Base.SQLUpdate('DELETE FROM "PL_VERSIONS" WHERE IDOWNER='+ IntToStr(OwnerID)+' AND IDFORMATS='+FormatID+' AND FTIMESTAMP<DATEADD(DAY, -'+IntToStr(FormatRecords.Items[aIndexFormat].fSTORAGEDAYS)+', CURRENT_DATE)',false);
        except
            SetStatus('Очистка прайс-листа... [ERROR]');
            raise;
        end;

      end;

      Base.SQLUpdate('PL_ITEMS',['STOCK','STOCK2','STOCK3','STOCK4','STOCK5'],[0,0,0,0,0],'IDOWNER=' + IntToStr(OwnerID)+' AND IDFORMATS='+FormatID,false);
      Base.SQLUpdate('FORMATS',['FILEHASH'],[aCurrentDataFileHash],'ID='+FormatID, false);
      Base.SQLUpdate('FORMATS',['FTIMESTAMPLASTIMPORT'],[TimeStamp],'ID='+FormatID, false);
      Base.SQLInsert('PRICELISTS_TIMESTAMPS',['IDFORMATS','IDOWNER','IDUSER','FTIMESTAMP'],[StrToInt(FormatID),OwnerID,integer(1),TimeStamp], false);

      GroupRootIndex := Base.SQLReadArr('PL_GROUP',['ID'],'IDOWNER='+IntTOStr(OwnerID)+' AND IDPARENT=0','')[0,0];

      GroupCunnrentLevel := GroupRootIndex;

      FGroup := TStringList.Create;

      SetStatus('Чтение данных. Ждите...');
      SetStatus('...',true);

////////////////////////
        try
          if fmt<>sfUser then
          begin
             FWorkBook.ReadFromFile(UTF8ToSys(aDataFilename), fmt);  // все, что не xlsx - fpcspreadsheets
          end else //fUserFormat
          begin
            case fUserFormat of
              ffYML:
                    begin
                    _FileStream:= TStringList.Create;
                    try
                      _FileStream.LoadFromFile(UTF8ToSys(aDataFilename));

                      for i:=0 to _FileStream.Count-1 do
                      begin
                       _S:= '';
                       _S:= _FileStream.Strings[i];
                       _PosXMLStart:= UTF8Pos('<?xml',UTF8LowerCase(_S),1);
                       if _PosXMLStart>0 then
                       begin
                         _S:= _FileStream.Strings[i];
                         UTF8Delete(_S,1,_PosXMLStart-1);
                         _FileStream.Strings[i]:= _S;

                         if i>0 then
                           for k:=i-1 downto 0 do
                               _FileStream.Delete(k);
                         Break;
                       end;

                      end;

                      _FileStream.SaveToFile(UTF8ToSys(aDataFilename));
                    finally
                      _FileStream.free;
                    end;

                    wYML:= TYML.Create(UTF8ToSys(aDataFilename));
                    try
                      wYML.Open();
                      //TimeStamp := DateTimeToStr(Now());
                      ImportYML(wYML); // процедура импорт YML
                    finally
                      wYML.Destroy;
                    end;

                    end
               else
                   begin
                     SetStatus('Открываю источник данных ... [ERROR]');
                     SetStatus('Неизвестный формат файла.');
                   end;
            end;
          end;
//////////////////////
        finally
          CollArray := nil;
          if FWorkBook <> nil then
            FWorkBook.Free;
          if FGroup <> nil then
            FGroup.Free;
        end;


      //SetStatus('Чтение данных... [ОК]',true);

      SetStatus('Обновление цен прайс-листа согласно текущему курсу валют...');

      SetStatus('Обновление данных...',true);

      KURS_AND_PERCENT:= Base.SQLReadArr('select KURS FROM GET_KURS_AND_PERCENT('+FormatID+')')[0,0];

      ImportUpdatePRICECALC(Base,KURS_AND_PERCENT,FormatID);

      try
        Base.SQLTransactionEnd(True);
      finally
        SetStatus('Импорт формата завершен. Импортировано позиций: '+IntToStr(CountImportedRecords));
      end;


    except
      on E: Exception do
      begin
        Base.SQLTransactionEnd(False);
        SetStatus('Открываю источник данных ... [ERROR]');
        SetStatus('Импорт завершен c ошибкой: "' + E.Message + '"');
        if FWorkBook <> nil then
          FWorkBook.Free;
{ TODO : отчего выходит совсем?
}
      end;
    end;
end;

procedure TReadDataPriceThread.Execute;
var
  i: integer;
  DataFileName: string;
  _ErrorDownloadFiles, _UnzipResult: boolean;
   DataFileHash, _CurrentDataFileHash, _UnPackPath, _OwnerName: string;
  _countSources: integer;
  wZipper: TwZipper;
  _arrZipper: ArrayOfString;

  wGet: TwGet;
  _XMLText: string;
  _arrDouble: ArrayOfArrayVariant;

begin
  try

    wZipper:= TwZipper.Create();
    try

      wGet:= TwGet.Create(fParent);

      SetStatus('Обновляю курсы валют с сайта ЦБ РФ...');
      StockOnly:= false;
      StockSymbols:= false;

      if GetKursValute(wGet) then
         SetStatus('Обновляю курсы валют с сайта ЦБ РФ...[OK]');


      if FormatRecords.Size > 0 then
      begin
        _countSources:= FormatRecords.Size;
        if _countSources > 1 then
          SetStatus('Выбрано несколько источников данных ('+IntTOStr(_countSources)+')');

        for i := 0 to FormatRecords.Size - 1 do
        begin

          FormatRecordsCuttentIndex:= i;
          Base.LongTransaction:= true;
          if FormatRecords.Items[i].fFILEZIPNAMEDECODE = 1 then wZipper.DecodeFileName:= true else wZipper.DecodeFileName:= false;
          OwnerID:= FormatRecords.Items[i].fIDOWNER;

          _OwnerName:= Base.SQLReadArr('OWNER', ['NAME'], 'ID=' + IntToStr(OwnerID), '')[0, 0];
          _XMLText:='';
          _XMLText:= Base.SQLReadArr('FORMATS',['URL'],'ID='+IntToStr(FormatRecords.Items[i].fID),'')[0,0];

          SetStatus('');
          if Length(_XMLText)>0 then
                SetStatus('-= ' + '[ ' + _OwnerName + ' ] Формат: ' + FormatRecords.Items[i].fNAME + ' | Файл: ' + FormatRecords.Items[i].fFILE + ' [Файл будет загружен из сети] =-') else
                SetStatus('-= ' + '[ ' + _OwnerName + ' ] Формат: ' + FormatRecords.Items[i].fNAME + ' | Файл: ' + FormatRecords.Items[i].fFILE + ' =-');

          _ErrorDownloadFiles:= false;
          if Length(_XMLText)>0 then
          begin
            SetStatus('Загружаю файл из сети интернет...');
            if LoadPriceFromInternet(wGet,_XMLText) then
               SetStatus('Загружаю файл из сети интернет... [OK]') else
               _ErrorDownloadFiles:= true;
          end;

          if not _ErrorDownloadFiles then
          begin
          //выборка прайс-листа

              _arrZipper:= wZipper.ParseComboFileName(FormatRecords.Items[i].fFILE);



              _UnPackPath:=includeTrailingPathDelimiter(wZipper.GetUnPackPath())+IntTOStr(OwnerID);
              if not DirectoryExistsUTF8(_UnPackPath) then ForceDirectoriesUTF8(_UnPackPath);

              _UnzipResult:= true;

              if Assigned(_arrZipper) then
              begin
                try
                wZipper.ExtractOneFile(SafePath(_arrZipper[0]),_arrZipper[1],_UnPackPath);
                except
                  on E: Exception do
                  begin
                    SetStatus('Error! '+E.Message);
                    _UnzipResult:= false;
                  end;
                end;

                DataFileName:= SafePath(includeTrailingPathDelimiter(_UnPackPath)+_arrZipper[1]);
                _arrZipper:= nil;
              end else
                    DataFileName := SafePath(FormatRecords.Items[i].fFILE);

              DataFileHash := FormatRecords.Items[i].fFILEHASH;
              FormatID := IntToStr(FormatRecords.Items[i].fID);

             //if _UnzipResult then

              //CalcMD5File
             _CurrentDataFileHash:= CalcMD5File(UTF8ToSys(DataFileName)); // вычисляем хэш прайс-листа;

              TimeStamp := DateTimeToStr(Now());
              CountImportedRecords:=0;

              if not IgnoreVersion and (_CurrentDataFileHash = DataFileHash) and _UnzipResult then // если уже загружали, то на второй круг
              begin
                    SetStatus('Данная версия файла уже была загружена в БД и будет пропущена!');
                    SetStatus('Обновление цен прайс-листа согласно текущему курсу валют...');
                    KURS_AND_PERCENT:= Base.SQLReadArr('select KURS FROM GET_KURS_AND_PERCENT('+FormatID+')')[0,0];

                    ImportUpdatePRICECALC(Base,KURS_AND_PERCENT,FormatID);

                    Base.SQLTransactionEnd(true);
              end else
              begin


               if IgnoreVersion then SetStatus('Игнорирование версии прайс-листа: ВКЛ');

               CollArray := nil;
               SetLength(CollArray, 31);
               __ARRAYFPREADSHEET:= FormatRecords.Items[i].fSPREADSHEET;
               __ARRAYSTOCKSYMBOLS:= FormatRecords.Items[i].fSTOCKSYMBOLS;
               __IDVENDORCODEVARIANT:= FormatRecords.Items[i].fIDVENDORCODEVARIANT;
               __IDSTOCKVARIANT:= FormatRecords.Items[i].fIDSTOCKVARIANT;
               __IDPRICEVARIANT:= FormatRecords.Items[i].fIDPRICEVARIANT;

               CollArray[0]:= 0;
               CollArray[1]:= FormatRecords.Items[i].fFIRSTLINE;
               CollArray[2]:= FormatRecords.Items[i].fVENDORCODE;
               CollArray[3]:= FormatRecords.Items[i].fFNAME;
               CollArray[4]:= FormatRecords.Items[i].fUNIT;
               CollArray[5]:= FormatRecords.Items[i].fQUANTITY;
               CollArray[6]:= FormatRecords.Items[i].fSTOCK2;
               CollArray[7]:= FormatRecords.Items[i].fSTOCK3;
               CollArray[8]:= FormatRecords.Items[i].fSTOCK4;
               CollArray[9]:= FormatRecords.Items[i].fSTOCK5;
               CollArray[10]:= FormatRecords.Items[i].fPRICE;
               CollArray[11]:= FormatRecords.Items[i].fPRICE2;
               CollArray[12]:= FormatRecords.Items[i].fPRICE3;
               CollArray[13]:= FormatRecords.Items[i].fPRICE4;
               CollArray[14]:= FormatRecords.Items[i].fPRICE5;
               CollArray[15]:= FormatRecords.Items[i].fPRICE6;
               CollArray[16]:= FormatRecords.Items[i].fPRICE7;
               CollArray[17]:= FormatRecords.Items[i].fPRICE8;
               CollArray[18]:= FormatRecords.Items[i].fPRICE9;
               CollArray[19]:= FormatRecords.Items[i].fPRICE10;

               CollArray[20]:= FormatRecords.Items[i].fLABEL;
               CollArray[21]:= FormatRecords.Items[i].fSCOD;
               CollArray[22]:= FormatRecords.Items[i].fGROUPS;
               CollArray[23]:= FormatRecords.Items[i].fSUBGROUPS1;
               CollArray[24]:= FormatRecords.Items[i].fSUBGROUPS2;
               CollArray[25]:= FormatRecords.Items[i].fSUBGROUPS3;
               CollArray[26]:= FormatRecords.Items[i].fTRANSIT;
               CollArray[27]:= FormatRecords.Items[i].fFURL;
               CollArray[28]:= FormatRecords.Items[i].fFURLPICTURE;
               CollArray[29]:= FormatRecords.Items[i].fFREMARK;
               CollArray[30]:= FormatRecords.Items[i].fFCOLOR;


                SetStatus('Открываю источник данных ...');

                if not FileExists(DataFileName) then
                begin
                  SetStatus('Открываю источник данных ... [ERROR]');
                  SetStatus('Файл не существует.');
                end
                else
                 if _UnzipResult then
// ИМПОРТ прайс-листа
                   ImportFormat(i,DataFileName,_CurrentDataFileHash,FormatRecords) else
                   begin
                      SetStatus('Обновление цен прайс-листа согласно текущему курсу валют...');
                      SetStatus('Обновление данных...',true);
                      KURS_AND_PERCENT:= Base.SQLReadArr('select KURS FROM GET_KURS_AND_PERCENT('+FormatID+')')[0,0];
                      ImportUpdatePRICECALC(Base,KURS_AND_PERCENT,FormatID);
                   end;
              end; // конец проверки хэша файла

          end;

          dec(_countSources);
          SetStatus('Осталось импортировать прайс-листов : ' + IntToStr(_countSources));
        end; //for i...

              // активация индексов
          SetStatus('Пересчет статистики индексов... Это может занять некоторое время...');

              Base.SQLUpdate('SET STATISTICS INDEX PL_VERSIONS_FTIMESTAMP;');

              Base.SQLUpdate('SET STATISTICS INDEX PL_ITEMS_VENDORCODE;');

          SetStatus('Пересчет статистики индексов... [ОК]');


  // обновление каталога товаров на основании импортированного основного прайс-листа

            IdMainOwner:= Base.ReadSettingByName('setDefaultOwner'); // считываем настройки - текущий основной прайс-лист

            _arrDouble:=nil;
            _arrDouble:= Base.SQLReadArr('select first 1 NAME from CATALOG WHERE IDOWNER='+IntTOStr(IdMainOwner)+' group by NAME having count(*)>1');

            if (Base.SQLReadDS('SELECT first 1 ID FROM PL_ITEMS WHERE IDOWNER='+IntTOStr(IdMainOwner)).DataSet.RecordCount>0)  then
            begin
              SetStatus('Синхронизация каталога товаров с основным прайс-листом...');
              if Length(_arrDouble)>0 then
              begin
                 SetStatus('Синхронизация каталога товаров с основным прайс-листом... [ОШИБКА!!!]');
                 SetStatus('В каталоге найдены ДУБЛИРУЮЩИЕСЯ ПОЗИЦИИ!!!');
                 SetStatus('Синхронизация каталога с прайс-листом отменено!!!');
                 __Log.SaveLogError(nil);
              end
              else
              begin

                 if (Base.SQLReadDS('SELECT first 1 ID FROM CATALOG_GROUP WHERE IDOWNER='+IntTOStr(IdMainOwner)).DataSet.RecordCount=1)
                    and (Base.SQLReadDS('SELECT first 1 ID FROM CATALOG WHERE IDOWNER='+IntTOStr(IdMainOwner)).DataSet.RecordCount=0)
                    then
                    Base.SQLUpdate('DELETE FROM CATALOG_GROUP WHERE IDOWNER='+IntTOStr(IdMainOwner));

                 Base.SQLUpdate('EXECUTE PROCEDURE CTG_UPDT_PL('+IntTOStr(IdMainOwner)+',true)'); // вызываем ХП, обновляющую каталог товаров
                 Base.SQLUpdate('SET STATISTICS INDEX CTG_VENDORCODE;');
                 Base.SQLUpdate('SET STATISTICS INDEX CTG_NAME;');
                 SetStatus('Синхронизация каталога товаров с основным прайс-листом... [OK]');
              end;
            end;

            _arrDouble:=nil;
  ///End обновление каталога

          SetStatus('-=== Все операции импорта завершены ===-');
          SetStatus('Все операции импорта завершены.',true);

        // конец всего импорта
        if FormatRecords.Size > 1 then
        begin
          SetStatus('Обработано прайс-листов: '+IntToStr(FormatRecords.Size));
        end;

      end
      else
      begin
        SetStatus('Открываю источник данных ... [ERROR].');
        SetStatus('Источник данных не найден.');
      end;
    finally
      wZipper.Destroy();
      wGet.Destroy();
      if Assigned(onEndThread) then onEndThread(self);
    end;
  except
    on E: Exception do
    begin
      SetStatus('Импорт завершен c ошибкой: "' + E.Message + '"');
      raise;
    end;
  end;
end;

procedure TReadDataPriceThread.SetStatus(aText: string; const aStatus: boolean);
begin
  fStatus := aText;
  fStatusLog:= aStatus;
  Synchronize(@Showstatus);
end;

constructor TReadDataPriceThread.Create(CreateSuspended: boolean);
begin
  FreeOnTerminate := true;
  inherited Create(CreateSuspended);
end;

{ TwDBImport }

procedure TwDBImport.Log(_Text: string);
begin
  // здесь напишите процедуру ведения лог-файла
  // если вы не ведете лог-файл, то оставьте тело функции пустым
  wLog('[' + fFormName + '] ' + '[DBImport] ', _Text);
end;

procedure TwDBImport.SetStatus(_Text: string; const _Status: boolean; const aLogSection: boolean);
var
  _memoLine: integer;
begin
  if Memo = nil then
    wStatus(fFormName, _Text, aLogSection)
  else
  begin
    if _Status then
    begin
      wStatus(fFormName, _Text, aLogSection)
    end
    else
      Memo.Lines.Add(DateTimeToStr(now()) + ' | ' + _Text);
      Log('[console.log] ' + _Text);
  end;
end;

constructor TwDBImport.Create(Sender: TObject; const _Memo: TMemo);
begin
  try
    Parent := (Sender as TComponent);
    fFormName := TForm(Sender).Name;
    Memo := _Memo;
    Base:=nil;
    fBaseOuter:= true;
    IgnoreVersion:= false;
    XMLSS_ON:= false;

    fFormatsPrice:= TwPriceFormats.Create;

    FEndThread:= true;

    fErrorMessage:= '';
    fOKMessage:= '';

    if Memo <> nil then
      Memo.Clear;

    wLog('DBImport','Создаем экземпляр импорта в БД... Инициатор: ' + fFormName);


    Log('Экземпляр импорта в БД успешно создан.');
    SetStatus('Экземпляр импорта в БД успешно создан.');
    SetStatus('Готов к работе.');
  except
    on E: Exception do
    begin
      __Log.SaveLogError(E);
      wLog('DBImport','Ошибка [Create]: "' + E.Message + '"');
      raise;
    end;
  end;
end;

destructor TwDBImport.Destroy();
begin
  fFormatsPrice.Free;
end;

procedure TwDBImport._onEndThread(Sender: TObject);
begin
  if not fOutBase then
  begin
    fBase.Destroy();
    fBase:=nil;
  end;
  EndThread:= true;
  if Assigned(TwCustomThread(Sender).OutStringArr) then
     fOutStringArr:= TwCustomThread(Sender).OutStringArr;
end;

procedure TwDBImport.Import(const aFormatType: TFormatType; const aFileName: string);
var
  ReadDataThread: TReadDataPriceThread;
  ReadOrder: TReadDataOrdersThread;
begin
 // if Assigned(ReadDataThread) then Exit;
  if not FEndThread then
  begin
    ShowMessage('Импорт уже запущен! Дождитесь окончания операции.');
    exit;
  end;

  SetStatus('===============');
  fOutBase:= true;
case aFormatType of
  ftPRICE:
          begin
            ReadDataThread := TReadDataPriceThread.Create(True);
            ReadDataThread.DBImport:= TwDBImport(Self);
            ReadDataThread.FormatRecords:= TwDBImport(Self).FormatsPrice;
            ReadDataThread.IgnoreVersion:= IgnoreVersion;
            ReadDataThread.fParent:= Parent;

            if not Assigned(Base) then
            begin
              fBase:= TwBase.Create(Parent);
              fOutBase:= false;
            end;
            ReadDataThread.Base:= fBase;
            ReadDataThread.onEndThread:=@_onEndThread;
            FEndThread:= false;

            ReadDataThread.Start;
          end;

  ftNAKL:
          begin
            ReadOrder:= TReadDataOrdersThread.Create(true);
            ReadOrder.DBImport:= TwDBImport(Self);
            ReadOrder.FormatOrder:= FFormatOrder;
            ReadOrder.DataFileName:= aFileName;
            if not Assigned(Base) then
            begin
              fBase:= TwBase.Create(Parent);
              fOutBase:= false;
            end;
            ReadOrder.Base:= fBase;
            ReadOrder.onEndThread:=@_onEndThread;
            FEndThread:= false;

            ReadOrder.Start;
          end;
end;
end;

procedure TwDBImport.ImportKursValut(aSilent: boolean);
var
  UpdateKursValutThread: TUpdateKursValutThread;
  i: Integer;
  _SQLText: String;
  _arr: ArrayOfArrayVariant;
begin
    if not FEndThread then
     begin
       ShowMessage('Обновление курса валют уже запущено! Дождитесь окончания операции.');
       exit;
     end;

     SetStatus('===============');

     UpdateKursValutThread := TUpdateKursValutThread.Create(True);
     UpdateKursValutThread.DBImport:= TwDBImport(Self);
//создание списка форматов для обновления

    if not Assigned(fBase) then
    begin
      fBaseOuter:= false;
      fBase:= TwBase.Create(Parent);
    end;

   _SQLText:='SELECT "FORMATS".ID, ("CURRENCY".kurs*(1+"FORMATS".currencypercent/100)) as KURS '
    +' from  "FORMATS" '
    +' left outer join "CURRENCY" ON "CURRENCY".ID="FORMATS".idcurrency '
    +' where KURS>1 ';

   _arr:= fBase.SQLReadArr(_SQLText);

   UpdateKursValutThread.FormatsArr:= _arr;


     UpdateKursValutThread.fParent:= Parent;

     UpdateKursValutThread.Base:= fBase;
     UpdateKursValutThread.Resume;

     //aSilent
     UpdateKursValutThread.onEndThread:=@_onEndThread;

     FEndThread:= false;
end;

function TwDBImport.CreateFormatPrice(aIDOWNER, aID: integer; aNAME, aFILE: string; aFILEZIPNAMEDECODE: integer; aFILEHASH, aURL: string; aIDFILEFORMAT,
  aFCONVERTLIBRE, aIDCODEPAGETEXT, aIDCURRENCY: integer; aCURRENCYPERCENT: double; aSTORAGEDAYS, aSTOCKONLY: integer; aSTOCKSYMBOLS: ArrayOfArrayVariant;
  aSTOCKONLYINFO, aYMLID, aYMLPRICE, aYMLQUANTITY, aFCLOSE, aGROUPSINROWS, aGROUPALGORITHM, aGROUPS, aSUBGROUPS1, aSUBGROUPS2, aSUBGROUPS3, aFIRSTLINE,
  aVENDORCODE, aFNAME, aUNIT, aQUANTITY, aSTOCK2, aSTOCK3, aSTOCK4, aSTOCK5, aTRANSIT, aPRICE, aPRICE2, aPRICE3, aPRICE4, aPRICE5, aPRICE6, aPRICE7, aPRICE8,
  aPRICE9, aPRICE10, aLABEL, aSCOD, aFURL, aFURLPICTURE, aFREMARK, aCOLOR, aIDFMTS_CATEGORY: integer; aSPREADSHEET: ArrayOfArrayInteger; aIDVENDORCODEVARIANT,
  aIDCSVDELIMITER, aIDSTOCKVARIANT, aIDPRICEVARIANT: integer): TRecPriceFormat;
begin
  Result.fIDOWNER:= aIDOWNER;
  Result.fID:= aID;
  Result.fNAME:= aNAME;
  Result.fFILE:= aFILE;
  Result.fFILEZIPNAMEDECODE:= aFILEZIPNAMEDECODE;
  Result.fFILEHASH:= aFILEHASH;
  Result.fURL:= aURL;
  Result.fIDFILEFORMAT:= aIDFILEFORMAT;
  Result.fFCONVERTLIBRE:= aFCONVERTLIBRE;
  Result.fIDCODEPAGETEXT:= aIDCODEPAGETEXT;
  Result.fIDCURRENCY:= aIDCURRENCY;
  Result.fCURRENCYPERCENT:= aCURRENCYPERCENT;
  Result.fSTORAGEDAYS:= aSTORAGEDAYS;
  Result.fSTOCKONLY:= aSTOCKONLY;
  Result.fSTOCKSYMBOLS:= aSTOCKSYMBOLS;
  REsult.fSTOCKONLYINFO:= aSTOCKONLYINFO;
  Result.fYMLID:= aYMLID;
  Result.fYMLPRICE:= aYMLPRICE;
  Result.fYMLQUANTITY:= aYMLQUANTITY;
  Result.fFCLOSE:= aFCLOSE;
  Result.fGROUPSINROWS:= aGROUPSINROWS;
  Result.fGROUPALGORITHM:= aGROUPALGORITHM;
  Result.fGROUPS:= aGROUPS;
  Result.fSUBGROUPS1:= aSUBGROUPS1;
  Result.fSUBGROUPS2:= aSUBGROUPS2;
  Result.fSUBGROUPS3:= aSUBGROUPS3;
  Result.fFIRSTLINE:= aFIRSTLINE;
  Result.fVENDORCODE:= aVENDORCODE;
  Result.fFNAME:= aFNAME;
  Result.fUNIT:= aUNIT;
  Result.fQUANTITY:= aQUANTITY;
  Result.fSTOCK2:= aSTOCK2;
  Result.fSTOCK3:= aSTOCK3;
  Result.fSTOCK4:= aSTOCK4;
  Result.fSTOCK5:= aSTOCK5;
  Result.fTRANSIT:= aTRANSIT;
  Result.fPRICE:= aPRICE;
  Result.fPRICE2:= aPRICE2;
  Result.fPRICE3:= aPRICE3;
  Result.fPRICE4:= aPRICE4;
  Result.fPRICE5:= aPRICE5;
  Result.fPRICE6:= aPRICE6;
  Result.fPRICE7:= aPRICE7;
  Result.fPRICE8:= aPRICE8;
  Result.fPRICE9:= aPRICE9;
  Result.fPRICE10:= aPRICE10;
  Result.fLABEL:= aLABEL;
  Result.fSCOD:= aSCOD;
  Result.fFURL:= aFURL;
  Result.fFURLPICTURE:= aFURLPICTURE;
  Result.fFREMARK:= aFREMARK;
  Result.fFCOLOR:= aCOLOR;
  Result.fIDFMTS_CATEGORY:= aIDFMTS_CATEGORY;
  Result.fSPREADSHEET:= aSPREADSHEET;
  Result.fIDVENDORCODEVARIANT:= aIDVENDORCODEVARIANT;
  Result.fIDCSVDELIMITER:= aIDCSVDELIMITER;
  Result.fIDSTOCKVARIANT:= aIDSTOCKVARIANT;
  Result.fIDPRICEVARIANT:= aIDPRICEVARIANT;
end;

end.
