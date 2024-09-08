unit wReportU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, Controls, dateutils, db, Dialogs, Forms, fpsSearch, fpsutils, Grids, fpspreadsheet, fpspreadsheetctrls, fpsTypes, Graphics, gvector, LazUTF8,
  LCLIntf, messages, mUtilsU, SysUtils, wCustomClassThreadU, wBaseU, wDBImportU, wFormulaU, wFuncU, wLogU, wTProgressU, wTypesU, wTViewerSpreadsheetU, wZipperU;

  type
     // Category
     TwGroup = packed record
        id: integer;
        parentId: integer;
        name: string;
     end;

     TwGroups = specialize TVector<TwGroup>;

    TwReportStruct = packed record
      SQLText: string;
      ColumnWidth: array of integer;
      HeadsToPrint: array of string;
      Fields: TStringList;
      Prices: string;
      Stocks: string;
      HeadsPrice: array of string;
      HeadsStock: array of string;
      Groups: TwGroups;
      WorkSheetGroup:TsWorksheet;
      WorkSheetList:TsWorksheet;
      WorkBookSizeUnits: TsSizeUnits;
    end;

    { TwReport }

    TwReport = class (TwCustomThreadWithProgressBar)
      protected
        fBase: TwBase;
        fReportModes: TReportModes;
        fSelectedPriceItems: ArrayOfInteger;
        fSelectedStockItems: ArrayOfInteger;
        fSelectedOwners: ArrayOfInteger;
        fPriceBase, fPriceCompare: TPriceType;
        fReport: TwViewer;
        fReportStruct: TwReportStruct;
        fOwnerForm: TForm;

        procedure Execute; override;
      private
        fFlagBool: Boolean;
        fPathToFiles: string;
        fTemplate: string;
        fWorkbookSource: TsWorkbookSource;
        procedure GetAnalogs;
        procedure GetCatalogExportCSV;
        procedure GetCatalogExportSpreadSheetStrubalin;
        procedure GetCatalogExportSpreadSheet;
        procedure GetCompareHorisontal;
        function GetOwnerColumn(aOwnerId: integer; aOwnerItems: ArrayOfArrayVariant): integer;
        procedure GetOwnerFiles;
        procedure GetSummaryInvoce;
        procedure GetInvoceWithSelectOwnerCode;
        procedure GetInvoceInPrice;
        procedure GetPriceDate;
        procedure GetCatalogExportSpreadSheet_WriteGroups(var aRowGroup: integer; var aRow: Integer; const uCol: Integer; Delimiter: string; var aGroups: TwGroups;
          const uGroup: integer; const uHeadsCount: integer; R, G, B: Byte);
        procedure GetCatalogExportSpreadSheet_WritePositions(var aRow: Integer; const uCol: Integer; aGroupId: integer);
        procedure SyncEnd;

      public
        constructor Create(CreateSuspended: boolean);
        destructor Destroy(); override;


        property ReportModes: TReportModes read fReportModes write fReportModes;
        property Base: TwBase read fBase write fBase;
        property SelectedPriceItems: ArrayOfInteger read fSelectedPriceItems write fSelectedPriceItems;
        property SelectedStockItems: ArrayOfInteger read fSelectedStockItems write fSelectedStockItems;
        property SelectedOwners: ArrayOfInteger read fSelectedOwners write fSelectedOwners;
        property WorkbookSource: TsWorkbookSource read fWorkbookSource write fWorkbookSource;
        property PriceBase: TPriceType read fPriceBase write fPriceBase;
        property PriceCompare: TPriceType read fPriceCompare write fPriceCompare;
        property PathToFiles: string read fPathToFiles write fPathToFiles;
        property FlagBool: Boolean read fFlagBool write fFlagBool;
        property Template: string read fTemplate write fTemplate;
        property OwnerForm: TForm read fOwnerForm write fOwnerForm;
    end;

implementation

uses
  mInvoceU;

{ TwReport }

procedure TwReport.SyncEnd;
begin
  if Assigned(fOwnerForm) then fOwnerForm.Repaint;
end;

procedure TwReport.Execute;
begin
  try
    if not Assigned(fBase) then
       raise Exception.Create('Error initialisation Thread!');

    case fReportModes of
      rmAnalogs                    : GetAnalogs;
      rmCompareHorisontal          : GetCompareHorisontal;
      rmSummaryInvoce              : GetSummaryInvoce;
      rmToOwnerFiles               : GetOwnerFiles;
      //rmCatalogExportSpreadSheet   : GetCatalogExportSpreadSheetStrubalin;
      rmCatalogExportSpreadSheet   : GetCatalogExportSpreadSheet;
      rmCatalogExportCSV           : GetCatalogExportCSV;
      rmInvoceWithSelectOwnerCode  : GetInvoceWithSelectOwnerCode;
      rmInvoceToPrice              : GetInvoceInPrice;
      rmPriceDate                  : GetPriceDate;
    end;

    Result:= true;
    if Assigned(onEndThread) then onEndThread(self);
  except
    on E: Exception do begin
      Result:= false;
      SetStatus('Error: '+E.Message);
      if Assigned(onEndThread) then onEndThread(self);
      Synchronize(@SyncEnd);
    end;
  end;
end;

function TwReport.GetOwnerColumn(aOwnerId: integer; aOwnerItems: ArrayOfArrayVariant): integer;
var
  i: Integer;
begin
  Result:=-1;

  for i:= 0 to High(aOwnerItems) do
    if aOwnerId = integer(aOwnerItems[i,0]) then
      begin
        Result:= integer(aOwnerItems[i,2]);
        Break;
      end;
end;

procedure TwReport.GetCompareHorisontal;
type
   TPriceMinValue = Record
     iCol: integer;
     Value: Double;
     iColEnd: integer;
   end;

var
  _FontColor: TColor;
  _DS: TDataSet;
  _arrItems, _arrOwners: ArrayOfArrayVariant;
  _ColHeaders: ArrayOfString;
  _ColWidth,_ColWidhResult : ArrayOfInteger;
  iCol: Integer;
  iSel: Integer;
  i: Integer;
  iRow, iOwner, iColOwner, iColOwnerLast: Integer;
  _SQLITems: String;
  _SQLText, _PriceBase, _PriceCompare, _SelectedOwners: String;
  _PriceBaseValue: Double;

  _PriceMinValue: TPriceMinValue;
const
  PriceCol = 5;
begin
  if Assigned(fSelectedPriceItems) then
   ProgressInit(pbBottom, High(fSelectedPriceItems)+1);

  _PriceBase:='';
  _PriceCompare:='';
  i:= 0;
case PriceBase of
    ptBase     : _PriceBase:='PL.PRICECALC';
    ptPrice2   : _PriceBase:='PL.PRICECALC2';
    ptPrice3   : _PriceBase:='PL.PRICECALC3';
    ptPrice4   : _PriceBase:='PL.PRICECALC4';
    ptPrice5   : _PriceBase:='PL.PRICECALC5';
  end;

case PriceCompare of
    ptBase     : _PriceCompare:='PRICE';
    ptPrice2   : _PriceCompare:='PRICE2';
    ptPrice3   : _PriceCompare:='PRICE3';
    ptPrice4   : _PriceCompare:='PRICE4';
    ptPrice5   : _PriceCompare:='PRICE5';
  end;

   _SQLText:='SELECT '
       +' IDOWNER, '
       +' VENDORCODE, '
       +_PriceCompare+','
       +' QUANTITYINPACKINGTEXT '
       +' FROM ANALIS_SEL_ALL_ANALOG(%d, true) '
       +' WHERE '+_PriceCompare+'>0'
       +' AND (%s)'
       +' ORDER BY IDOWNER, '+_PriceCompare;

    iRow:=2;

    _ColWidth:= [10,10, 15, 12, 60, 6, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10,
                10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10,
                10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10];
    _ColHeaders:=['Контр.','Код', 'Штрих-код', 'Артикул', 'Наименование', 'Ед.', 'Цена'];

    SetColWidth(fWorkbookSource.Worksheet, _ColWidth);

     fWorkbookSource.Worksheet.Options:= fWorkbookSource.Worksheet.Options+ [soHasFrozenPanes];
     fWorkbookSource.Worksheet.TopPaneHeight:= 3;

     for i:=0 to High(_ColHeaders) do
          WriteValue(fWorkbookSource.Worksheet, iRow, i, _ColHeaders[i], [fssBold], ReportHeaderColor);

     _arrOwners:= fBase.SQLReadArr('SELECT DISTINCT(OWN.ID), OWN.NAME, 0 SCOLUMN FROM OWNER OWN'
                  +' INNER JOIN FORMATS FMTS ON (FMTS.IDOWNER=OWN.ID)'
                  +' WHERE '+fBase.PrepareWhereString('OWN.ID',fSelectedOwners)
                  +' ORDER BY OWN.NAME');

     iColOwner:= i+1;
     // заполняем таблицу "вширь" контрагентами

     for iOwner:=0 to High(_arrOwners) do
       begin

         WriteValue(fWorkbookSource.Worksheet, iRow-1, iColOwner, _arrOwners[iOwner,1], [fssBold], ReportHeaderColor);
         fWorkbookSource.Worksheet.MergeCells(iRow-1,iColOwner,iRow-1,iColOwner+2);
         fWorkbookSource.Worksheet.WriteHorAlignment(iRow-1,iColOwner,haCenter);

         WriteValue(fWorkbookSource.Worksheet, iRow, iColOwner, 'Код', [fssBold], ReportHeaderColor);
         _arrOwners[iOwner,2]:= iColOwner;
         inc(iColOwner);
         WriteValue(fWorkbookSource.Worksheet, iRow, iColOwner, 'Цена', [fssBold], ReportHeaderColor);
         inc(iColOwner);
         WriteValue(fWorkbookSource.Worksheet, iRow, iColOwner, 'Фас.', [fssBold], ReportHeaderColor);
         fWorkbookSource.Worksheet.WriteColWidth(iColOwner,5,suChars,cwtCustom);

         inc(iColOwner);
       end;

     _PriceMinValue.iColEnd:= iColOwner;
     WriteValue(fWorkbookSource.Worksheet, iRow, _PriceMinValue.iColEnd, 'Лучшая', [fssBold], ReportHeaderColor);

     inc(iRow);

    _SelectedOwners:= fBase.PrepareWhereString('IDOWNER',fSelectedOwners);

    for iSel:=0 to High(fSelectedPriceItems) do begin

     if StopForce then raise Exception.Create('Прервано пользователем!');

     ProgressUpdate(pbBottom);

    _SQLITems:='SELECT '
        +' OWN.NAME OWNERNAME, '
        +' PL.VENDORCODE, '
        +' (select VSCOD FROM PL_GET_SCOD(PL.ID,true)) SCOD,'
        +' PL.LABEL, '
        +' PL.NAME, '
        +' PL.UNIT, '
        +_PriceBase+' '
        +' FROM PL_ITEMS PL '
        +' INNER JOIN OWNER OWN ON (OWN.ID=PL.IDOWNER) '
        +' WHERE PL.ID='+IntToStr(fSelectedPriceItems[iSel]);

    // выбираем выбранные позиции (левая часть)
    _arrItems:= nil;
    _arrItems:= fBase.SQLReadArr(_SQLITems);

    for iCol:=0 to High(_arrItems[0]) do
        WriteValue(fWorkbookSource.Worksheet, iRow, iCol, _arrItems[0, iCol]);

    // сохраняем базовую цену
    _PriceBaseValue:= _arrItems[0, 6];

    with _PriceMinValue do begin
      iCol:= -1;
      Value:= _PriceBaseValue;
    end;

    // выбираем аналоги (правая часть)
      _DS:= nil;
      _DS:= fBase.SQLReadDS(Format(_SQLText, [fSelectedPriceItems[iSel],_SelectedOwners]), true).DataSet;
      _DS.First;

      ProgressInit(pbTop, _DS.RecordCount);

     if _DS.RecordCount = 0 then
      begin
        //dec(iRow);
        fWorkbookSource.Worksheet.DeleteRow(iRow);
      end else
      begin
        iColOwnerLast:= -1;
        for i:=0 to _DS.RecordCount-1 do
          begin
           iColOwner:= GetOwnerColumn(_DS.Fields[0].AsInteger,_arrOwners);
           if (iColOwner>-1) and (iColOwnerLast<>iColOwner) then
            begin
               iColOwnerLast:= iColOwner;

               if _DS.Fields[2].AsFloat < _PriceMinValue.Value then
                begin
                   _PriceMinValue.Value:= _DS.Fields[2].AsFloat;
                   _PriceMinValue.iCol:= iColOwner;
                end;

               WriteValue(fWorkbookSource.Worksheet, iRow, iColOwner, _DS.Fields[1], [], clDefault);
               WriteValue(fWorkbookSource.Worksheet, iRow, iColOwner+1, _DS.Fields[2], [], clDefault);
               WriteValue(fWorkbookSource.Worksheet, iRow, iColOwner+2, _DS.Fields[3], [], clDefault);
            end;

           _DS.Next;

           ProgressUpdate(pbTop);
          end;

        if  _PriceMinValue.Value <> _PriceBaseValue then
         begin
           fWorkbookSource.Worksheet.WriteFontColor(iRow,_PriceMinValue.iCol+1,clRed);
           WriteValue(fWorkbookSource.Worksheet, iRow,_PriceMinValue.iColEnd, _PriceMinValue.Value, [], clDefault, clRed);
         end;

        inc(iRow);

      end;

      _DS.Close;

    end;
    _arrItems:= nil;
end;

procedure TwReport.GetOwnerFiles;
var
  _FontColor: TColor;
  aDS: TDataSet;
  _ColHeaders: ArrayOfString;
  _ColWidth: ArrayOfInteger;
  iCol: Integer;
  i: Integer;
  iRow, iOwner: Integer;
  aSQLText, aOwner: String;
  aInvoce: TInvoce;
  _arr: ArrayOfArrayVariant;

begin
  aInvoce:= TInvoce.Create(fBase, nil);

  try
    aSQLText:= aInvoce.InvoceToExportIntoOwnerFiles;

    _arr:= nil;
    _arr:= fBase.SQLReadArr(aInvoce.OwnersFromInvoce);

  finally
    FreeAndNil(aInvoce);
  end;


  for iOwner:=0 to High(_arr) do
    begin
      iRow:=1;

      _ColWidth:= [10, 15, 15, 60, 10, 10, 15, 20];
      _ColHeaders:=['№', 'Код', 'Артикул', 'Наименование', 'Ед.', 'Кол-во', 'Штрих-код','Примечание'];

      SetColWidth(fWorkbookSource.Worksheet, _ColWidth);



       for i:=0 to High(_ColHeaders) do
         begin
            WriteValue(fWorkbookSource.Worksheet, iRow, i, _ColHeaders[i], [fssBold], ReportHeaderColor);
         end;

       inc(iRow);


        aDS:= nil;

        aDS:= fBase.SQLReadDS(fBase.WriteWhereEx(aSQLText,'INV.IDOWNER='+VarToStr(_arr[iOwner,0])),true).DataSet;
        aDS.First;

        ProgressInit(pbTop, aDS.RecordCount);

        aOwner:= VarToStr(_arr[iOwner,1]);
        i:=0;
        while not aDS.EOF do
          begin
          if StopForce then raise Exception.Create('Прервано пользователем!');

          WriteValue(fWorkbookSource.Worksheet, iRow, 0, i+1, [], clDefault, clBlack);
          for iCol:=0 to aDS.Fields.Count-1 do
            begin
            _FontColor:= clBlack;
            WriteValue(fWorkbookSource.Worksheet, iRow, iCol+1, aDS.Fields[iCol], [], clDefault, _FontColor);
            end;

           inc(iRow);
           inc(i);
           aDS.Next;

           ProgressUpdate(pbTop);
          end;

        aDS.Close;

        fWorkbookSource.SaveToSpreadsheetFile(PathToFiles+'Заказ_'+aOwner+'.xls');
        fWorkbookSource.Worksheet.Clear;
    end;
end;

procedure TwReport.GetCatalogExportSpreadSheet_WritePositions(var aRow: Integer; const uCol: Integer; aGroupId: integer);
var
  aDataSet: TDataSet;
  i: Integer;
  aStock: LongInt;
  aStockWhere: String;
begin
  aStockWhere:='';
  if FlagBool then
   aStockWhere:='((PLOUR.STOCK+PLOUR.STOCK2+PLOUR.STOCK3+PLOUR.STOCK4+PLOUR.STOCK5)>0 '
   +' OR PLFP.PRICEPL >0 '
   +' OR CHAR_LENGTH(PLOUR.TRANSIT)>0) AND ';

  aDataSet:= Base.SQLReadDS(Base.WriteWhereEx(fReportStruct.SQLText,aStockWhere+' CATALOG.IDCTG_GROUP='+IntToStr(aGroupId),true)).DataSet;
  Inc(aRow);
  try
    while not aDataSet.EOF do
      begin
        for i:= 0 to fReportStruct.Fields.Count-1 do
          begin
            case fReportStruct.Fields.Strings[i] of
              'STOCK1':
                begin
                  aStock:= VarToInt(aDataSet.FieldByName(fReportStruct.Fields.Strings[i]).AsVariant);

                  if aStock>0 then
                    WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i, aStock)
                  else
                    WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i, UTF8LowerCase(aDataSet.FieldByName('TOINVOCE').AsString));
                end;
              'STOCK2': WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i, aDataSet.FieldByName(fReportStruct.Fields.Strings[i]).AsVariant);
              'STOCK3': WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i, aDataSet.FieldByName(fReportStruct.Fields.Strings[i]).AsVariant);
              'STOCK4': WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i, aDataSet.FieldByName(fReportStruct.Fields.Strings[i]).AsVariant);
              'STOCK5': WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i, aDataSet.FieldByName(fReportStruct.Fields.Strings[i]).AsVariant);
              'NAME'  : WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i, aDataSet.FieldByName(fReportStruct.Fields.Strings[i]).AsVariant);

              //'HYPERLINK':
              //  begin
              //    WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i,'описание',[],clDefault,clBlue);
              //    fReportStruct.WorkSheetList.WriteHyperlink(aRow, uCol+i,'http://ink-service.ru/?page=desc&id='+aDataSet.FieldByName('SCOD').AsString,'описание');
              //  end;
              //'COLORNEW':
              //    WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i, aDataSet.FieldByName('COLORNEW').AsString,[fssBold],clDefault,clRed);

                else
                  WriteValue(fReportStruct.WorkSheetList, aRow, uCol+i, aDataSet.FieldByName(fReportStruct.Fields.Strings[i]).AsVariant);
            end;

          end;
        Inc(aRow);
        aDataSet.Next;
      end;
  finally
    dec(aRow);
    aDataSet.Close;
  end;

end;

procedure TwReport.GetCatalogExportSpreadSheet_WriteGroups(var aRowGroup: integer; var aRow: Integer; const uCol: Integer; Delimiter: string; var aGroups: TwGroups; const uGroup: integer; const uHeadsCount: integer; R, G, B: Byte);
var
   i: integer;
begin
    i:=0;
    //BkColor:= BkColor;
    while aGroups.Size>i do
    begin

       if aGroups[i].parentId = uGroup then
       begin
         {write group}
         WriteValue(fReportStruct.WorkSheetGroup, aRowGroup, uCol, Delimiter+aGroups[i].name, [fssBold], RGBToColor(R,G,B));
         fReportStruct.WorkSheetGroup.WriteHyperlink(aRowGroup,uCol,'#'+fReportStruct.WorkSheetList.Name+'!'+GetColString(uCol)+IntToStr(aRow+1));

         Inc(aRowGroup);

         {write list}
         WriteValue(fReportStruct.WorkSheetList, aRow, uCol, Delimiter+aGroups[i].name, [fssBold], RGBToColor(R,G,B));
         fReportStruct.WorkSheetList.MergeCells(aRow,uCol,aRow,uHeadsCount);

         ProgressUpdate(pbBottom);
         GetCatalogExportSpreadSheet_WritePositions(aRow, uCol, aGroups[i].id);

         inc(aRow);
         GetCatalogExportSpreadSheet_WriteGroups(aRowGroup, aRow, uCol, Delimiter+'  ', aGroups, aGroups[i].id, uHeadsCount, R,G,B-20);
       end;
       Inc(i);
    end;
end;

procedure TwReport.GetCatalogExportCSV;
var
  fFormula: TFormula;
  aDataSet: TDataSet;
  i: Integer;
  aPriceFormula, aFieldsToPrint: String;
  aPackFiles: string;
  aZipper: TwZipper;
begin
    PathToFiles:= IncludeTrailingBackslash(PathToFiles);

    fReportStruct.SQLText:='SELECT * FROM CATALOG_GROUP ORDER BY IDPARENT, NAME';
    Base.ExportTableToCSVFile(fReportStruct.SQLText,PathToFiles+'CATALOG_GROUP.CSV',nil,'',10000);

    aFieldsToPrint:='ID,IDCTG_GROUP,NAME,UNIT,LABEL,VENDORCODE,SCOD,FCOLOR,STOCK1,STOCK2,STOCK3,STOCK4,STOCK5,TRANSIT,TOINVOCE,FURLPICTURE,REMARK';
    fReportStruct.Fields:= TStringList.Create;
    fReportStruct.Fields.Delimiter:=',';
    fReportStruct.Fields.DelimitedText:= aFieldsToPrint;
    fFormula:= TFormula.Create(nil);

    try
      aDataSet:= Base.SQLReadDS('SELECT NAME,FORMULA FROM PRICEFIELD WHERE FCLOSE=0 ORDER BY PRIORITY',true).DataSet;
      aDataSet.Last;


      SetLength(fReportStruct.HeadsPrice, aDataSet.RecordCount);

      fReportStruct.Prices:='';

      for i:=0 to aDataSet.RecordCount-1 do
       begin
         if i>0 then fReportStruct.Prices:= fReportStruct.Prices+',';
         aPriceFormula:= fFormula.Prepare(aDataSet.FieldByName('FORMULA').AsString);

         fReportStruct.Prices:= fReportStruct.Prices+ wfLineEnding+'IIF('+aPriceFormula+'IS NULL,0,'+aPriceFormula+')'+' EXPORTPRICE'+IntTOStr(i);

         fReportStruct.Fields.Add('EXPORTPRICE'+IntTOStr(i));
         fReportStruct.HeadsPrice[i]:= aDataSet.FieldByName('NAME').AsString;
         aDataSet.Prior;
       end;

      aDataSet.Close;

        fReportStruct.SQLText:=' SELECT CATALOG.ID, '+wfLineEnding
        +' CATALOG.IDCTG_GROUP, '+wfLineEnding
        +' CATALOG.NAME AS NAME, '+wfLineEnding
        +' CATALOG.UNIT, '+wfLineEnding
        +' (select VSCOD from CTG_GET_SCOD(CATALOG.ID,true)) SCOD, '+wfLineEnding
        +' CATALOG.LABEL, '+wfLineEnding
        +' CATALOG.PRICE AS PRICE, '+wfLineEnding
        +' CATALOG.VENDORCODE, '+wfLineEnding
        +' CATALOG.FCOLOR, '+wfLineEnding
        +' CATALOG.FURLPICTURE, '+wfLineEnding
        +' CATALOG.REMARK, '+wfLineEnding
        +' CATALOG.FTIMESTAMP, '+wfLineEnding
        +' CAST((PLOUR.STOCK+PLOUR.STOCK2+PLOUR.STOCK3+PLOUR.STOCK4+PLOUR.STOCK5) AS INTEGER) AS STOCK, '+wfLineEnding
        +' IIF(PLFP.PRICEPL >0 AND FMTS.INVOCEDAYS>0,FMTS.INVOCEDAYS,'''') AS TOINVOCE, '+wfLineEnding
        +' PLOUR.STOCK AS STOCK1, '+wfLineEnding
        +' PLOUR.STOCK2 AS STOCK2, '+wfLineEnding
        +' PLOUR.STOCK3 AS STOCK3, '+wfLineEnding
        +' PLOUR.STOCK4 AS STOCK4, '+wfLineEnding
        +' PLOUR.STOCK5 AS STOCK5, '+wfLineEnding
        +' CATALOG.PN,CATALOG.PM,CATALOG.PD,CATALOG.PC,CATALOG.PK, '+wfLineEnding
        +'  PLFP.PRICEPL AS PRICEPL, '+wfLineEnding
        +'  PLFP.PRICEPL2 AS PRICEPL2, '+wfLineEnding
        +'  PLFP.PRICEPL3 AS PRICEPL3, '+wfLineEnding
        +'  PLFP.PRICEPL4 AS PRICEPL4, '+wfLineEnding
        +'  PLFP.PRICEPL5 AS PRICEPL5, '+wfLineEnding
        +'  PLFP.PRICEPL6 AS PRICEPL6, '+wfLineEnding
        +'  PLFP.PRICEPL7 AS PRICEPL7, '+wfLineEnding
        +'  PLFP.PRICEPL8 AS PRICEPL8, '+wfLineEnding
        +'  PLFP.PRICEPL9 AS PRICEPL9, '+wfLineEnding
        +'  PLFP.PRICEPL10 AS PRICEPL10, '+wfLineEnding
        +' PLOUR.PRICECALC AS PRICEOUR, '+wfLineEnding
        +' PLOUR.PRICECALC2 AS PRICEOUR2, '+wfLineEnding
        +' PLOUR.PRICECALC3 AS PRICEOUR3, '+wfLineEnding
        +' PLOUR.PRICECALC4 AS PRICEOUR4, '+wfLineEnding
        +' PLOUR.PRICECALC5 AS PRICEOUR5, '+wfLineEnding
        +' PLOUR.PRICECALC6 AS PRICEOUR6, '+wfLineEnding
        +' PLOUR.PRICECALC7 AS PRICEOUR7, '+wfLineEnding
        +' PLOUR.PRICECALC8 AS PRICEOUR8, '+wfLineEnding
        +' PLOUR.PRICECALC9 AS PRICEOUR9, '+wfLineEnding
        +' PLOUR.PRICECALC10 AS PRICEOUR10, '+wfLineEnding
        +' IIF(PLOUR.TRANSIT IS NULL,'''',PLOUR.TRANSIT) TRANSIT,'+wfLineEnding
        +''+fReportStruct.Prices+' '+wfLineEnding
        +'  FROM CATALOG '+wfLineEnding
        +' LEFT JOIN CATALOG_PL_MIN_PRICE(CATALOG.ID) PLFP ON (1=1)'+wfLineEnding
        +'  LEFT OUTER JOIN "PL_ITEMS" PLOUR ON ( '+wfLineEnding
        +'  CATALOG.VENDORCODE = PLOUR.VENDORCODE AND CATALOG.IDOWNER = PLOUR.IDOWNER) '+wfLineEnding
        +' LEFT JOIN FORMATS FMTS ON (PLFP.IDFORMATS = FMTS.ID) '+wfLineEnding
        +' ';

        aFieldsToPrint:='';
        for i:=0 to fReportStruct.Fields.Count-1 do
         begin
           if i>0 then
             aFieldsToPrint:= aFieldsToPrint+',';
           aFieldsToPrint:= aFieldsToPrint+fReportStruct.Fields[i];
         end;

        Base.ExportTableToCSVFile(fReportStruct.SQLText,PathToFiles+'CATALOG_ITEMS.CSV',nil,aFieldsToPrint,10000);

        aZipper:= TwZipper.Create();

        try
          aZipper.PackFiles('CATALOG.zip',PathToFiles, 'CATALOG_*.CSV');
        finally
          FreeAndNil(aZipper);
        end;

  finally
    FreeAndNil(fReportStruct.Fields);
    FreeAndNil(fFormula);
  end;
end;

procedure TwReport.GetCatalogExportSpreadSheetStrubalin;
  function CreateGroup(aId, aParentId: integer; aName: string):TwGroup;
  begin
    Result.id:= aId;
    Result.parentId:= aParentId;
    Result.name:= aName;
  end;

var
  aDataSet: TDataSet;
  iCol: Integer;
  i: Integer;
  iRow, aCurrentColCount, aNewCol, iRowGroup: Integer;
  fFormula: TFormula;
  aFieldsToPrint, aPriceFormula: String;
  IsTemplate: Boolean;
  aCell: PCell;
  aHeadColor: TsColor;
begin
    Result:= false;
    WorkbookSource:= TsWorkbookSource.Create(nil);

    IsTemplate:= Length(Template)>0;

    if IsTemplate then
    begin
      WorkbookSource.LoadFromSpreadsheetFile(Template, GetSpreadSheetFormat(Template));
      fReportStruct.WorkBookSizeUnits:= WorkbookSource.Workbook.Units;
      fReportStruct.WorkSheetGroup:= WorkbookSource.Workbook.GetWorksheetByIndex(0);
      fReportStruct.WorkSheetGroup.Name:= 'Groups';

      fReportStruct.WorkSheetList:= WorkbookSource.Workbook.GetWorksheetByIndex(1);
      fReportStruct.WorkSheetList.Name:= 'Price';
    end else

    begin
      fReportStruct.WorkSheetGroup:= WorkbookSource.Workbook.ActiveWorksheet;
      fReportStruct.WorkSheetGroup.Name:= 'Groups';

      fReportStruct.WorkSheetList:= WorkbookSource.Workbook.AddWorksheet('Price');
    end;

    WriteValue(fReportStruct.WorkSheetList, 0, 9, now(), [fssBold], RGBToColor(240,251,255));

    fReportStruct.Groups:= TwGroups.Create;

    aFieldsToPrint:='HYPERLINK,SCOD,NAME,UNIT,COLORNEW,STOCK1,STOCK2,STOCK3,TRANSIT';

    fReportStruct.HeadsToPrint:= ['Ссылка','Ш-код','Наименование','Ед.','New','Влж','СЦ','ММ','Транзит'];

    fReportStruct.Fields:= TStringList.Create;
    fReportStruct.Fields.Delimiter:=',';
    fReportStruct.Fields.DelimitedText:= aFieldsToPrint;

    //aSQLText:= 'SELECT CTG.VENDORCODE, CTG.NAME, CTG.UNIT, (CTG.STOCK+CTG.STOCK2+CTG.STOCK3+CTG.STOCK4+CTG.STOCK5) STOCK, FROM CATALOG CTG';
    try
      fFormula:= TFormula.Create(nil);

      aDataSet:= Base.SQLReadDS('SELECT NAME,FORMULA FROM PRICEFIELD WHERE '+Base.PrepareWhereString('ID', SelectedPriceItems),true).DataSet;
      aDataSet.Last;

      fReportStruct.Prices:='';

      SetLength(fReportStruct.HeadsPrice, aDataSet.RecordCount);

      //ShowMessage(Format('Heads %d | Count %d',[Length(fReportStruct.HeadsPrice), aDataSet.RecordCount]));
      for i:=0 to aDataSet.RecordCount-1 do
       begin
         if i>0 then fReportStruct.Prices:= fReportStruct.Prices+',';
         aPriceFormula:= fFormula.Prepare(aDataSet.FieldByName('FORMULA').AsString);

         fReportStruct.Prices:= fReportStruct.Prices+ wfLineEnding+'IIF('+aPriceFormula+'IS NULL,0,'+aPriceFormula+')'+' EXPORTPRICE'+IntTOStr(i);

         fReportStruct.Fields.Add('EXPORTPRICE'+IntTOStr(i));
         fReportStruct.HeadsPrice[i]:= aDataSet.FieldByName('NAME').AsString;
         aDataSet.Prior;
       end;

      aDataSet.Close;

      fReportStruct.SQLText:=' SELECT CATALOG.ID, '+wfLineEnding
        +''+fReportStruct.Prices+', '+wfLineEnding
        +' CATALOG.IDCTG_GROUP, '+wfLineEnding
        +' 1 as HYPERLINK,'
        +' (select VSCOD from CTG_GET_SCOD(CATALOG.ID,true)) SCOD, '+wfLineEnding
        +' CATALOG.NAME AS NAME, '+wfLineEnding
        +' CATALOG.UNIT, '+wfLineEnding
        +' IIF(CATALOG.FCOLOR=15,''new'','''') COLORNEW, '+wfLineEnding
        +' CAST(PLOUR.STOCK as INTEGER) AS STOCK1, '+wfLineEnding
        +' IIF(PLOUR.STOCK2>0,CAST(PLOUR.STOCK2 as INTEGER),'''') AS STOCK2, '+wfLineEnding
        +' IIF(PLOUR.STOCK3>0,CAST(PLOUR.STOCK3 as INTEGER),'''') AS STOCK3, '+wfLineEnding
        +' IIF(PLOUR.STOCK4>0,CAST(PLOUR.STOCK4 as INTEGER),'''') AS STOCK4,'+wfLineEnding
        +' IIF(PLOUR.STOCK5>0,CAST(PLOUR.STOCK5 as INTEGER),'''') AS STOCK5, '+wfLineEnding
        +' IIF(PLOUR.TRANSIT IS NULL,'''',PLOUR.TRANSIT) TRANSIT,'+wfLineEnding
        +' CATALOG.LABEL, '+wfLineEnding
        +' CATALOG.PRICE AS PRICE, '+wfLineEnding
        +' CATALOG.VENDORCODE, '+wfLineEnding
        +' CATALOG.FTIMESTAMP, '+wfLineEnding
        +' CAST((PLOUR.STOCK+PLOUR.STOCK2+PLOUR.STOCK3+PLOUR.STOCK4+PLOUR.STOCK5) AS INTEGER) AS STOCK, '+wfLineEnding
        +' IIF(PLFP.PRICEPL >0,IIF(FMTS.INVOCEDAYS>0,''заказ ''||FMTS.INVOCEDAYS||'' дн.'',''заказ '') ,'''') AS TOINVOCE, '+wfLineEnding
        +' CATALOG.PN,CATALOG.PM,CATALOG.PD,CATALOG.PC,CATALOG.PK, '+wfLineEnding
        //+' CASE WHEN (SELECT * FROM CATALOG_SELECT_MTHRESULT(CATALOG.ID))>0 THEN 1 ELSE 0 END MTHRESULT, '
        +'  PLFP.PRICEPL AS PRICEPL, '+wfLineEnding
        +'  PLFP.PRICEPL2 AS PRICEPL2, '+wfLineEnding
        +'  PLFP.PRICEPL3 AS PRICEPL3, '+wfLineEnding
        +'  PLFP.PRICEPL4 AS PRICEPL4, '+wfLineEnding
        +'  PLFP.PRICEPL5 AS PRICEPL5, '+wfLineEnding
        +'  PLFP.PRICEPL6 AS PRICEPL6, '+wfLineEnding
        +'  PLFP.PRICEPL7 AS PRICEPL7, '+wfLineEnding
        +'  PLFP.PRICEPL8 AS PRICEPL8, '+wfLineEnding
        +'  PLFP.PRICEPL9 AS PRICEPL9, '+wfLineEnding
        +'  PLFP.PRICEPL10 AS PRICEPL10, '+wfLineEnding
        +' PLOUR.PRICECALC AS PRICEOUR, '+wfLineEnding
        +' PLOUR.PRICECALC2 AS PRICEOUR2, '+wfLineEnding
        +' PLOUR.PRICECALC3 AS PRICEOUR3, '+wfLineEnding
        +' PLOUR.PRICECALC4 AS PRICEOUR4, '+wfLineEnding
        +' PLOUR.PRICECALC5 AS PRICEOUR5, '+wfLineEnding
        +' PLOUR.PRICECALC6 AS PRICEOUR6, '+wfLineEnding
        +' PLOUR.PRICECALC7 AS PRICEOUR7, '+wfLineEnding
        +' PLOUR.PRICECALC8 AS PRICEOUR8, '+wfLineEnding
        +' PLOUR.PRICECALC9 AS PRICEOUR9, '+wfLineEnding
        +' PLOUR.PRICECALC10 AS PRICEOUR10 '+wfLineEnding
        +'  FROM CATALOG '+wfLineEnding
        +' LEFT JOIN CATALOG_PL_MIN_PRICE(CATALOG.ID) PLFP ON (1=1)'+wfLineEnding
        +'  LEFT OUTER JOIN "PL_ITEMS" PLOUR ON ( '+wfLineEnding
        +'  CATALOG.VENDORCODE = PLOUR.VENDORCODE AND CATALOG.IDOWNER = PLOUR.IDOWNER) '+wfLineEnding
        +' LEFT JOIN FORMATS FMTS ON (PLFP.IDFORMATS = FMTS.ID) '+wfLineEnding
        +' ORDER  BY CATALOG.NAME ';

      aCurrentColCount:= Length(fReportStruct.HeadsToPrint);
      aNewCol:= aCurrentColCount+Length(fReportStruct.HeadsPrice);
      SetLength(fReportStruct.HeadsToPrint, aNewCol);

      if not IsTemplate then
      begin
        for i:=0 to High(fReportStruct.HeadsPrice) do
          fReportStruct.HeadsToPrint[i+aCurrentColCount]:= fReportStruct.HeadsPrice[i];

          fReportStruct.ColumnWidth:= [5, 15, 60, 10, 15];
      end;
      aDataSet:= fBase.SQLReadDS('SELECT ID, IDPARENT, NAME FROM CATALOG_GROUP ORDER BY IDPARENT, NAME').DataSet;

      while not aDataSet.EOF do
        begin
          fReportStruct.Groups.PushBack(CreateGroup(aDataSet.FieldByName('ID').AsInteger, aDataSet.FieldByName('IDPARENT').AsInteger, aDataSet.FieldByName('NAME').AsString));
          aDataSet.Next;
        end;

      if not IsTemplate then
      begin
        SetColWidth(fReportStruct.WorkSheetList, fReportStruct.ColumnWidth);
        SetColWidth(fReportStruct.WorkSheetGroup, [5,100]);
      end;

       iRow:= 6;
       iCol:= 1;

       ProgressInit(pbBottom, fReportStruct.Groups.Size);

      if not IsTemplate then
      begin
         for i:=0 to High(fReportStruct.HeadsToPrint) do
           begin
              WriteValue(fReportStruct.WorkSheetList, iRow, i+iCol, fReportStruct.HeadsToPrint[i], [fssBold], RGBToColor(240,251,255));
           end;
      end
      else
      begin
        aCell:= fReportStruct.WorkSheetList.FindCell(iRow,iCol);
        aHeadColor:= fReportStruct.WorkSheetList.ReadCellFormat(aCell).Background.BgColor;
        for i:=0 to High(fReportStruct.HeadsPrice) do
          begin
             WriteValue(fReportStruct.WorkSheetList, iRow, High(fReportStruct.HeadsToPrint)+i+iCol-1, fReportStruct.HeadsPrice[i], [fssBold], aHeadColor);
          end;
      end;

         inc(iRow);

         iRowGroup:= 2;

         GetCatalogExportSpreadSheet_WriteGroups(iRowGroup, iRow, 1, '', fReportStruct.Groups,fReportStruct.Groups[0].id, Length(fReportStruct.HeadsToPrint),240,251,235);

         //fReportStruct.WorkSheetList.TopPaneHeight:= iRow+2;
         //WorkbookSource.on
         WorkbookSource.SaveToSpreadsheetFile(PathToFiles, GetSpreadSheetFormat(PathToFiles));
         Result:= true;
    finally
      fReportStruct.Fields.Free;
      fReportStruct.Groups.Free;
      fFormula.Free;
      WorkbookSource.Free;
    end;

end;

procedure TwReport.GetCatalogExportSpreadSheet;
function CreateGroup(aId, aParentId: integer; aName: string):TwGroup;
begin
  Result.id:= aId;
  Result.parentId:= aParentId;
  Result.name:= aName;
end;

var
aDataSet: TDataSet;
iCol: Integer;
i,k: Integer;
iRow, aCurrentColCount, aNewCol, iRowGroup, aSelectStock: Integer;
fFormula: TFormula;
aFieldsToPrint, aPriceFormula, aStockColumn: String;
IsTemplate: Boolean;
aCell: PCell;
aHeadColor: TsColor;
begin
  Result:= false;
  WorkbookSource:= TsWorkbookSource.Create(nil);

  IsTemplate:= Length(Template)>0;

  if IsTemplate then
  begin
    WorkbookSource.LoadFromSpreadsheetFile(Template, GetSpreadSheetFormat(Template));
    fReportStruct.WorkBookSizeUnits:= WorkbookSource.Workbook.Units;
    fReportStruct.WorkSheetGroup:= WorkbookSource.Workbook.GetWorksheetByIndex(0);
    fReportStruct.WorkSheetGroup.Name:= 'Groups';

    fReportStruct.WorkSheetList:= WorkbookSource.Workbook.GetWorksheetByIndex(1);
    fReportStruct.WorkSheetList.Name:= 'Price';
  end else

  begin
    fReportStruct.WorkSheetGroup:= WorkbookSource.Workbook.ActiveWorksheet;
    fReportStruct.WorkSheetGroup.Name:= 'Groups';

    fReportStruct.WorkSheetList:= WorkbookSource.Workbook.AddWorksheet('Price');
  end;

  WriteValue(fReportStruct.WorkSheetList, 0,5, now(), [fssBold], RGBToColor(240,251,255)); // пишем дату

  fReportStruct.Groups:= TwGroups.Create;

  aFieldsToPrint:='CTGID, NAME, UNIT';

  fReportStruct.HeadsToPrint:= ['Код','Наименование','Ед.'];

  fReportStruct.Fields:= TStringList.Create;
  fReportStruct.Fields.Delimiter:=',';
  fReportStruct.Fields.DelimitedText:= aFieldsToPrint;

  try
    //// Формирование колонок остатков
    //
    aDataSet:= Base.SQLReadDS('SELECT NAME,FORMULA FROM PRICEFIELD WHERE '+Base.PrepareWhereString('ID', SelectedStockItems),true).DataSet;
    aDataSet.Last;

    fReportStruct.Stocks:='';

    SetLength(fReportStruct.HeadsStock, Length(SelectedStockItems));

    for i:=0 to High(SelectedStockItems) do
     begin
       if i>0 then fReportStruct.Stocks:= fReportStruct.Stocks+',';

       aSelectStock := SelectedStockItems[i];

       if (aSelectStock = 1) then
          fReportStruct.Stocks:= fReportStruct.Stocks+ wfLineEnding+
          Format('IIF(PLOUR.STOCK2>0,CAST(PLOUR.STOCK as INTEGER),'''') AS STOCK%d',[aSelectStock])
        else
          fReportStruct.Stocks:= fReportStruct.Stocks+ wfLineEnding+
          Format('IIF(PLOUR.STOCK2>0,CAST(PLOUR.STOCK%d as INTEGER),'''') AS STOCK%d',[aSelectStock, aSelectStock]);

       fReportStruct.Fields.Add('STOCK'+IntTOStr(aSelectStock));
       fReportStruct.HeadsStock[i]:= Format('Отдел %d',[aSelectStock]);
     end;

    // Формирование колонок цен
    fFormula:= TFormula.Create(nil);

    aDataSet:= Base.SQLReadDS('SELECT NAME,FORMULA FROM PRICEFIELD WHERE '+Base.PrepareWhereString('ID', SelectedPriceItems),true).DataSet;
    aDataSet.Last;

    fReportStruct.Prices:='';

    SetLength(fReportStruct.HeadsPrice, aDataSet.RecordCount);

    for i:=0 to aDataSet.RecordCount-1 do
     begin
       if i>0 then fReportStruct.Prices:= fReportStruct.Prices+',';
       aPriceFormula:= fFormula.Prepare(aDataSet.FieldByName('FORMULA').AsString);

       fReportStruct.Prices:= fReportStruct.Prices+ wfLineEnding+'IIF('+aPriceFormula+'IS NULL,0,'+aPriceFormula+')'+' EXPORTPRICE'+IntTOStr(i);

       fReportStruct.Fields.Add('EXPORTPRICE'+IntTOStr(i));
       fReportStruct.HeadsPrice[i]:= aDataSet.FieldByName('NAME').AsString;
       aDataSet.Prior;
     end;

    aDataSet.Close;

    fReportStruct.SQLText:=' SELECT CATALOG.ID AS CTGID, '+wfLineEnding
      +''+fReportStruct.Prices+', '+wfLineEnding
      +''+fReportStruct.Stocks+', '+wfLineEnding
      +' CATALOG.IDCTG_GROUP, '+wfLineEnding
      //+' (select VSCOD from CTG_GET_SCOD(CATALOG.ID,true)) SCOD, '+wfLineEnding
      +' CATALOG.NAME AS NAME, '+wfLineEnding
      +' CATALOG.UNIT, '+wfLineEnding
      +' CATALOG.LABEL, '+wfLineEnding
      +' CATALOG.PRICE AS PRICE, '+wfLineEnding
      +' CATALOG.VENDORCODE, '+wfLineEnding
      +' CATALOG.FTIMESTAMP, '+wfLineEnding
      +' CAST((PLOUR.STOCK+PLOUR.STOCK2+PLOUR.STOCK3+PLOUR.STOCK4+PLOUR.STOCK5) AS INTEGER) AS STOCK, '+wfLineEnding
      +' IIF(PLFP.PRICEPL >0,IIF(FMTS.INVOCEDAYS>0,''заказ ''||FMTS.INVOCEDAYS||'' дн.'',''заказ '') ,'''') AS TOINVOCE, '+wfLineEnding
      +' CATALOG.PN,CATALOG.PM,CATALOG.PD,CATALOG.PC,CATALOG.PK, '+wfLineEnding
      //+' CASE WHEN (SELECT * FROM CATALOG_SELECT_MTHRESULT(CATALOG.ID))>0 THEN 1 ELSE 0 END MTHRESULT, '
      +'  PLFP.PRICEPL AS PRICEPL, '+wfLineEnding
      +'  PLFP.PRICEPL2 AS PRICEPL2, '+wfLineEnding
      +'  PLFP.PRICEPL3 AS PRICEPL3, '+wfLineEnding
      +'  PLFP.PRICEPL4 AS PRICEPL4, '+wfLineEnding
      +'  PLFP.PRICEPL5 AS PRICEPL5, '+wfLineEnding
      +'  PLFP.PRICEPL6 AS PRICEPL6, '+wfLineEnding
      +'  PLFP.PRICEPL7 AS PRICEPL7, '+wfLineEnding
      +'  PLFP.PRICEPL8 AS PRICEPL8, '+wfLineEnding
      +'  PLFP.PRICEPL9 AS PRICEPL9, '+wfLineEnding
      +'  PLFP.PRICEPL10 AS PRICEPL10, '+wfLineEnding
      +' PLOUR.PRICECALC AS PRICEOUR, '+wfLineEnding
      +' PLOUR.PRICECALC2 AS PRICEOUR2, '+wfLineEnding
      +' PLOUR.PRICECALC3 AS PRICEOUR3, '+wfLineEnding
      +' PLOUR.PRICECALC4 AS PRICEOUR4, '+wfLineEnding
      +' PLOUR.PRICECALC5 AS PRICEOUR5, '+wfLineEnding
      +' PLOUR.PRICECALC6 AS PRICEOUR6, '+wfLineEnding
      +' PLOUR.PRICECALC7 AS PRICEOUR7, '+wfLineEnding
      +' PLOUR.PRICECALC8 AS PRICEOUR8, '+wfLineEnding
      +' PLOUR.PRICECALC9 AS PRICEOUR9, '+wfLineEnding
      +' PLOUR.PRICECALC10 AS PRICEOUR10 '+wfLineEnding
      +'  FROM CATALOG '+wfLineEnding
      +' LEFT JOIN CATALOG_PL_MIN_PRICE(CATALOG.ID) PLFP ON (1=1)'+wfLineEnding
      +'  LEFT OUTER JOIN "PL_ITEMS" PLOUR ON ( '+wfLineEnding
      +'  CATALOG.VENDORCODE = PLOUR.VENDORCODE AND CATALOG.IDOWNER = PLOUR.IDOWNER) '+wfLineEnding
      +' LEFT JOIN FORMATS FMTS ON (PLFP.IDFORMATS = FMTS.ID) '+wfLineEnding
      +' ORDER  BY CATALOG.NAME ';

    aCurrentColCount:= Length(fReportStruct.HeadsToPrint);

    if not IsTemplate then
    begin
      aNewCol:= aCurrentColCount+Length(fReportStruct.HeadsStock)+Length(fReportStruct.HeadsPrice);
      SetLength(fReportStruct.HeadsToPrint, aNewCol);

      for i:=0 to High(fReportStruct.HeadsStock) do
        fReportStruct.HeadsToPrint[i+aCurrentColCount]:= fReportStruct.HeadsStock[i];

      aCurrentColCount := aCurrentColCount + High(fReportStruct.HeadsStock);

      for i:=0 to High(fReportStruct.HeadsPrice) do
        fReportStruct.HeadsToPrint[i+aCurrentColCount]:= fReportStruct.HeadsPrice[i];

        fReportStruct.ColumnWidth:= [15, 60];
    end;

    aDataSet:= fBase.SQLReadDS('SELECT ID, IDPARENT, NAME FROM CATALOG_GROUP ORDER BY IDPARENT, NAME').DataSet;

    while not aDataSet.EOF do
      begin
        fReportStruct.Groups.PushBack(CreateGroup(aDataSet.FieldByName('ID').AsInteger, aDataSet.FieldByName('IDPARENT').AsInteger, aDataSet.FieldByName('NAME').AsString));
        aDataSet.Next;
      end;

    if not IsTemplate then
    begin
      SetColWidth(fReportStruct.WorkSheetList, fReportStruct.ColumnWidth);
      SetColWidth(fReportStruct.WorkSheetGroup, [15,100]);
    end;

     iRow:= 6;
     iCol:= 2;

     ProgressInit(pbBottom, fReportStruct.Groups.Size);

    if not IsTemplate then
    begin
       for i:=0 to High(fReportStruct.HeadsToPrint) do
         begin
            WriteValue(fReportStruct.WorkSheetList, iRow, i+iCol, fReportStruct.HeadsToPrint[i], [fssBold], RGBToColor(240,251,255));
         end;
    end
    else
    begin
      aCell:= fReportStruct.WorkSheetList.FindCell(iRow,iCol);
      aHeadColor:= fReportStruct.WorkSheetList.ReadCellFormat(aCell).Background.BgColor;

      for i:=0 to High(fReportStruct.HeadsStock) do
        begin
           inc(aCurrentColCount);
           WriteValue(fReportStruct.WorkSheetList, iRow, aCurrentColCount, fReportStruct.HeadsStock[i], [fssBold], aHeadColor);
        end;

      for k:=0 to High(fReportStruct.HeadsPrice) do
        begin
           inc(aCurrentColCount);
           WriteValue(fReportStruct.WorkSheetList, iRow, aCurrentColCount, fReportStruct.HeadsPrice[k], [fssBold], aHeadColor);
        end;
    end;

       inc(iRow);

       iRowGroup:= 2;

       GetCatalogExportSpreadSheet_WriteGroups(iRowGroup, iRow, 1, '', fReportStruct.Groups,fReportStruct.Groups[0].id, aCurrentColCount,240,251,235);

       //fReportStruct.WorkSheetList.TopPaneHeight:= iRow+2;
       //WorkbookSource.on
       WorkbookSource.SaveToSpreadsheetFile(PathToFiles, GetSpreadSheetFormat(PathToFiles));
       Result:= true;
  finally
    fReportStruct.Fields.Free;
    fReportStruct.Groups.Free;
    fFormula.Free;
    WorkbookSource.Free;
  end;

end;

procedure TwReport.GetSummaryInvoce;
var
  _FontColor: TColor;
  aDS: TDataSet;
  _ColHeaders: ArrayOfString;
  _ColWidth: ArrayOfInteger;
  iCol: Integer;
  i: Integer;
  iRow: Integer;
  aSum: Currency;
  aSQLText, aOwner: String;
  aInvoce: TInvoce;

  procedure WriteSum;
  begin
    WriteValue(fWorkbookSource.Worksheet, iRow, 0, aSum, [fssBold], ReportHeaderColor);

    fWorkbookSource.Worksheet.MergeCells(iRow, 0, iRow, High(_ColWidth));
    fWorkbookSource.Worksheet.WriteHorAlignment(iRow, 0, haRight);

    aOwner:= aDS.FieldByName('OWNERNAME').AsString;
    Inc(iRow);
    aSum:= 0;
  end;

begin
  aInvoce:= TInvoce.Create(fBase, nil);

  try
    aSQLText:= aInvoce.SummaryInvoce;
  finally
    FreeAndNil(aInvoce);
  end;

    iRow:=1;

    _ColWidth:= [10,10, 15, 80, 10, 10, 10, 10, 15, 15];
    _ColHeaders:=['Код Свой','Код', 'Артикул', 'Наименование', 'Ед.', 'Выписано', 'Цена', 'Сумма', 'Примечание', 'Контрагент'];

    SetColWidth(fWorkbookSource.Worksheet, _ColWidth);



     for i:=0 to High(_ColHeaders) do
       begin
          WriteValue(fWorkbookSource.Worksheet, iRow, i, _ColHeaders[i], [fssBold], ReportHeaderColor);
       end;

     inc(iRow);


      aDS:= nil;
      aDS:= fBase.SQLReadDS(aSQLText, true).DataSet;
      aDS.First;

      ProgressInit(pbTop, aDS.RecordCount);

      aOwner:= aDS.FieldByName('OWNERNAME').AsString;
      aSum:= 0;

      for i:=0 to aDS.RecordCount-1 do
        begin
        if StopForce then raise Exception.Create('Прервано пользователем!');

        if aOwner <> aDS.FieldByName('OWNERNAME').AsString then
         begin
           WriteSum;
         end;

        for iCol:=0 to aDS.Fields.Count-1 do
          begin
          _FontColor:= clBlack;
          WriteValue(fWorkbookSource.Worksheet, iRow, iCol, aDS.Fields[iCol], [], clDefault, _FontColor);
          end;

         aSum:= aSum+aDS.FieldByName('FSUM').AsCurrency;
         inc(iRow);
         aDS.Next;

         ProgressUpdate(pbTop);
        end;

      WriteSum;

      aDS.Close;
end;

procedure TwReport.GetInvoceWithSelectOwnerCode;
var
  _FontColor, _CellColor: TColor;
  aDS: TDataSet;
  _ColHeaders: ArrayOfString;
  _ColWidth: ArrayOfInteger;
  iCol: Integer;
  i: Integer;
  iRow: Integer;
  aOrderVendorCode: string;
  aSum: Currency;
  aSQLText: String;
  aOwnerID: Int64;
  aInvoce: TInvoce;
  aArrOwner: ArrayOfArrayVariant;

  procedure WriteOwner;
  var
    aOwnerName: String;
    iArr: Integer;
  begin
    aOwnerName:='';

    for iArr:=0 to High(aArrOwner) do
      begin
        if aArrOwner[iArr,0] = aOwnerID then
          begin
            aOwnerName:= VarToStr(aArrOwner[iArr,1]);
            break;
          end;
      end;

    WriteValue(fWorkbookSource.Worksheet, iRow, 0, aOwnerName, [fssBold], ReportHeaderColor);

    fWorkbookSource.Worksheet.MergeCells(iRow, 0, iRow, High(_ColWidth));
    fWorkbookSource.Worksheet.WriteHorAlignment(iRow, 0, haLeft);

    Inc(iRow);
  end;

begin
  aInvoce:= TInvoce.Create(fBase, nil);

  fBase.SQLUpdate('DELETE FROM W_TMP_ORDERS_GET_MTH');

  aSQLText:= 'EXECUTE PROCEDURE WTOI_GET_ANALOGS(%d)';

  aArrOwner:= fBase.SQLReadArr('SELECT ID, NAME FROM OWNER WHERE (ID IN ('+fBase.MakeStringFromArray(SelectedPriceItems)+')) AND (SELECT COUNT(*) FROM OWNER T WHERE T.IDPARENT=OWNER.ID)=0');

  for i:=0 to High(aArrOwner) do
    begin
      fBase.SQLUpdate(Format(aSQLText,[int64(aArrOwner[i,0])]));
    end;


  try
    aSQLText:= aInvoce.InvoceWithSelectOwnerCode;
  finally
    FreeAndNil(aInvoce);
  end;

    iRow:=1;

    _ColWidth:= [10, 10, 10, 80, 5, 10, 10, 10, 10];
    _ColHeaders:=['Код', 'Артикул','Штрих-код', 'Наименование', 'Ед.', 'Выписано','Код контрагента', 'Примечание', 'Цена'];

    SetColWidth(fWorkbookSource.Worksheet, _ColWidth);



     for i:=0 to High(_ColHeaders) do
       begin
          WriteValue(fWorkbookSource.Worksheet, iRow, i, _ColHeaders[i], [fssBold], ReportHeaderColor);
       end;

     inc(iRow);

      aDS:= nil;
      aDS:= fBase.SQLReadDS(aSQLText).DataSet;
      aDS.First;

      ProgressInit(pbTop, aDS.RecordCount);

      aOwnerID:= aDS.FieldByName('OWNERSEARCH').AsLargeInt;
      WriteOwner;
      //aSum:= 0;

      aOrderVendorCode:= '';

      while not aDS.EOF do
        begin
        if StopForce then raise Exception.Create('Прервано пользователем!');

        if aOwnerID <> aDS.FieldByName('OWNERSEARCH').AsLargeInt then
          begin
            aOwnerID:= aDS.FieldByName('OWNERSEARCH').AsLargeInt;
            WriteOwner;
          end;
        _CellColor:= clDefault;

        for iCol:=0 to aDS.Fields.Count-1 do
          begin
            _FontColor:= clBlack;
            if aDS.Fields[iCol].FieldName <> 'OWNERSEARCH' then
              begin
                if aDS.Fields[iCol].FieldName = 'ORDVENDORCODE' then
                 begin
                   if aOrderVendorCode = aDS.FieldByName('ORDVENDORCODE').AsString then
                    begin
                      _CellColor:= clYellow;
                      for i:=0 to aDS.Fields.Count-2 do
                        begin
                          fWorkbookSource.Worksheet.WriteBackgroundColor(iRow-1,i,_CellColor);
                        end;
                    end;

                    aOrderVendorCode:= aDS.FieldByName('ORDVENDORCODE').AsString;
                  end;

                WriteValue(fWorkbookSource.Worksheet, iRow, iCol, aDS.Fields[iCol], [], _CellColor, _FontColor);
              end;
          end;

         inc(iRow);
         aDS.Next;

         ProgressUpdate(pbTop);
        end;

      //WriteSum;

      aDS.Close;

end;

function GetListIndex(aText:string; aStringList: TStringList):int64;
var
  i: Integer;
begin
  Result:= -1;

  for i:=0 to aStringList.Count-1 do
    if Trim(aText) = Trim(aStringList.Strings[i]) then
     begin
       Result:= i;
       break;
     end;

end;

procedure TwReport.GetInvoceInPrice;
var
  aDS: TDataSet;
  aSQLText, aFileName, aValue, aSQLUpdate: String;
  aInvoce: TInvoce;

  aRecPriceFormat: TRecPriceFormat;
  iSpreadSheet, iRow, aVendorCol, aFirstLine, i, aInvoceCol, aVendorIndex, aId, aQuantity: Integer;
  aWorkSheet: TsWorksheet;
  aCount, ARow, ACol, aCurrentRow: Int64;
  aDataCell: PCell;
  aVendorCodes: TStringList;
begin
  aFileName:= '';

  aInvoce:= TInvoce.Create(fBase, nil);

  try
    aSQLText:= aInvoce.List;
    aSQLUpdate:= 'UPDATE INVOCES SET REMARK=%s WHERE ID=%d';

    aSQLText:= fBase.WriteWhereEx(aSQLText,'INV.IDOWNER='+IntToStr(SelectedPriceItems[0]));

    aInvoce.GetPriceFile(SelectedPriceItems[0], aRecPriceFormat);

    aFileName:= aRecPriceFormat.fFILE;
  finally
    FreeAndNil(aInvoce);
  end;

  if Length(aFileName) =0 then exit;
  try

    fWorkbookSource.LoadFromSpreadsheetFile(aFileName, GetSpreadSheetFormat(aFileName));
  //aWorkSheet:= nil;

    aDS:= fBase.SQLReadDS(aSQLText).DataSet;
    aDS.First;

    aVendorCodes:= TStringList.Create;
    aVendorCodes.Clear;

    while not aDS.EOF do
      begin
        aVendorCodes.AddObject(aDS.FieldByName('VENDORCODE').AsString, TwData.Create(aDS.FieldByName('ID').AsLargeInt,aDS.FieldByName('QUANTITY').AsInteger));
        aDS.Next;
      end;

    aDS.Close;

    for iSpreadSheet:=0 to High(aRecPriceFormat.fSPREADSHEET) do
    begin
       aWorkSheet:= fWorkbookSource.Workbook.GetWorksheetByIndex(aRecPriceFormat.fSPREADSHEET[iSpreadSheet,0]-1);
       aFirstLine:= aRecPriceFormat.fSPREADSHEET[iSpreadSheet,1];
       aVendorCol:= aRecPriceFormat.fVENDORCODE-1;
       aCount:= aWorkSheet.GetCellCountInCol(aRecPriceFormat.fVENDORCODE-1);
       aInvoceCol:= aRecPriceFormat.fADDRCELLFORINVOCE;

       aCurrentRow:= 0;

       if aInvoceCol=0 then aInvoceCol:= aWorkSheet.Cols.Count+1;

       ProgressInit(pbTop, aCount);

       for ADataCell in aWorkSheet.Cells do
       begin
         ARow := ADataCell^.Row;
         ACol := ADataCell^.Col;

         if StopForce then raise Exception.Create('Прервано пользователем!');

         if ARow >= aFirstLine-1 then
          begin
            if aCurrentRow<>aRow then
             begin
               ProgressUpdate(pbTop);
               aCurrentRow:= aRow;
             end;

            if aCol = aVendorCol then
             begin
               case aRecPriceFormat.fIDVENDORCODEVARIANT of
                 0: aValue:= GetDataString(ADataCell);
                 1: aValue:= GetDataString(ADataCell, vtNumber);
                 2: aValue:= GetDataString(ADataCell, vtString);
               end;
               aVendorIndex:= GetListIndex(aValue,aVendorCodes);
               if aVendorIndex>-1 then
                begin
                  aId:= TwData(aVendorCodes.Objects[aVendorIndex]).ID;
                  aQuantity:= TwData(aVendorCodes.Objects[aVendorIndex]).Value;

                  WriteValue(aWorkSheet,aCurrentRow,aInvoceCol-1, aQuantity,[fssBold], clYellow, clRed);
                  fBase.SQLUpdate(Format(aSQLUpdate,[QuotedStr('отмечено в прайс-листе'),aId]));
                end;
             end;
          end;

       end;
    end;

      fWorkbookSource.SaveToSpreadsheetFile(aFileName,GetSpreadSheetFormat(aFileName));

      if MessageDlg('Открыть прайс-лист контрагента во внешней программе просмотра?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

      OpenDocument(aFileName);
    finally
      for i:=0 to aVendorCodes.Count-1 do
        TwData(aVendorCodes.Objects[i]).Free;

      FreeAndNil(aVendorCodes);
    end;

end;

procedure TwReport.GetPriceDate;
var
  aFontColor: TColor;
  aDS: TDataSet;
  aColHeaders: ArrayOfString;
  aColWidth: ArrayOfInteger;
  iCol: Integer;
  i: Integer;
  iRow: Integer;
  aSQLText: String;
begin

   aSQLText:='select OWN.NAME OWNERNAMR, FMTS.NAME FORMATNAME, FMTS.FTIMESTAMPLASTIMPORT '
      +' from FORMATS FMTS '
      +' INNER JOIN OWNER OWN ON (OWN.ID=FMTS.IDOWNER) '
      +' WHERE FMTS.IDFMTS_CATEGORY=1 '
      +' ORDER BY FMTS.FTIMESTAMPLASTIMPORT DESC, OWN.NAME';

    iRow:=1;

    WriteValue(fWorkbookSource.Worksheet, iRow, 0, '< 1 дня', [], clDefault, clGreen);
    Inc(iRow);
    WriteValue(fWorkbookSource.Worksheet, iRow, 0, '> 1 дня', [], clDefault, clBlue);
    Inc(iRow);
    WriteValue(fWorkbookSource.Worksheet, iRow, 0, '> 7 дней', [], clDefault, clGray);
    Inc(iRow);
    WriteValue(fWorkbookSource.Worksheet, iRow, 0, '> 14 дней', [], clDefault, clMaroon);
    Inc(iRow);
    WriteValue(fWorkbookSource.Worksheet, iRow, 0, '> 30 дней', [], clDefault, clRed);
    Inc(iRow);
    Inc(iRow);

    aColWidth:= [15, 15, 15];
    aColHeaders:=['Контрагент', 'Формат', 'Дата'];

    SetColWidth(fWorkbookSource.Worksheet, aColWidth);



     for i:=0 to High(aColHeaders) do
       begin
          WriteValue(fWorkbookSource.Worksheet, iRow, i, aColHeaders[i], [fssBold], ReportHeaderColor);
       end;

      inc(iRow);
      aDS:= nil;
      aDS:= fBase.SQLReadDS(aSQLText, true).DataSet;
      aDS.First;

      ProgressInit(pbBottom, aDS.RecordCount);

      aFontColor:= clBlack;

      for i:=0 to aDS.RecordCount-1 do
        begin
          if StopForce then raise Exception.Create('Прервано пользователем!');
          if aDS.FieldByName('FTIMESTAMPLASTIMPORT').AsDateTime >= IncDay(Now,-1) then
            aFontColor:= clGreen;

          if aDS.FieldByName('FTIMESTAMPLASTIMPORT').AsDateTime < IncDay(Now,-1) then
            aFontColor:= clBlue;

          if aDS.FieldByName('FTIMESTAMPLASTIMPORT').AsDateTime < IncDay(Now,-7) then
            aFontColor:= clGray;

          if aDS.FieldByName('FTIMESTAMPLASTIMPORT').AsDateTime < IncDay(Now,-14) then
            aFontColor:= clMaroon;

          if aDS.FieldByName('FTIMESTAMPLASTIMPORT').AsDateTime < IncDay(Now,-30) then
            aFontColor:= clRed;


          for iCol:=0 to aDS.Fields.Count-1 do
            WriteValue(fWorkbookSource.Worksheet, iRow, iCol, aDS.Fields[iCol], [], clDefault, aFontColor);

         inc(iRow);
         aDS.Next;

         ProgressUpdate(pbTop);
        end;

      aDS.Close;

end;

procedure TwReport.GetAnalogs;
var
  _FontColor: TColor;
  _DS: TDataSet;
  _arrItems: ArrayOfArrayVariant;
  _ColHeaders: ArrayOfString;
  _ColWidth: ArrayOfInteger;
  iCol: Integer;
  iSel: Integer;
  i: Integer;
  iRow: Integer;
  _SQLITems: String;
  _SQLText: String;
const
  PriceCol = 5;
begin
  if Assigned(fSelectedPriceItems) then
   ProgressInit(pbBottom, High(fSelectedPriceItems)+1);

   _SQLText:='SELECT '
       +' VENDORCODE, '
       +' LABEL, '
       +' PLNAME, '
       +' UNIT, '
       +' STOCK, '
       +' PRICE, '
       +' QUANTITYINPACKINGTEXT, '
       +' FTIMESTAMP, '
       +' OWNERNAME '
       +' FROM ANALIS_SEL_ALL_ANALOG(%d, true) '
       +' ORDER BY PRICE';

    iRow:=1;

    _ColWidth:= [10, 10, 80, 10, 10, 10, 10, 15, 10];
    _ColHeaders:=['Код', 'Артикул', 'Наименование', 'Ед.', 'Остаток', 'Цена', 'Фасовка', 'Дата Импорта', 'К'
      +'онтрагент'];

    SetColWidth(fWorkbookSource.Worksheet, _ColWidth);



     for i:=0 to High(_ColHeaders) do
       begin
          WriteValue(fWorkbookSource.Worksheet, iRow, i, _ColHeaders[i], [fssBold], ReportHeaderColor);
       end;
     inc(iRow);

    for iSel:=0 to High(fSelectedPriceItems) do begin

     if StopForce then raise Exception.Create('Прервано пользователем!');

     ProgressUpdate(pbBottom);

    _SQLITems:='SELECT '
        +' PL.VENDORCODE, '
        +' PL.LABEL, '
        +' PL.NAME, '
        +' PL.UNIT, '
        +' CAST((PL.STOCK+PL.STOCK2+PL.STOCK3+PL.STOCK4+PL.STOCK5) AS INTEGER) STOCK, '
        +' PL.PRICECALC, '
        +' '''' QUANTITYINPACKINGTEXT, '
        +' PL.FTIMESTAMP, '
        +' OWN.NAME OWNERNAME '
        +' FROM PL_ITEMS PL '
        +' INNER JOIN OWNER OWN ON (OWN.ID=PL.IDOWNER) '
        +' WHERE PL.ID='+IntToStr(fSelectedPriceItems[iSel]);

    _arrItems:= nil;
    _arrItems:= fBase.SQLReadArr(_SQLITems);

    for iCol:=0 to High(_arrItems[0]) do
        WriteValue(fWorkbookSource.Worksheet, iRow, iCol, _arrItems[0, iCol], [], $cefbff);

    inc(iRow);
      _DS:= nil;
      _DS:= fBase.SQLReadDS(Format(_SQLText, [fSelectedPriceItems[iSel]]), true).DataSet;
      _DS.First;

      ProgressInit(pbTop, _DS.RecordCount);

      if _DS.RecordCount = 0 then
       begin
         dec(iRow);
         fWorkbookSource.Worksheet.DeleteRow(iRow);
       end;

      for i:=0 to _DS.RecordCount-1 do
        begin
        for iCol:=0 to _DS.Fields.Count-1 do
          begin
          _FontColor:= clBlack;
            if (i=0) and (iCol = PriceCol) and (_DS.Fields[PriceCol].AsFloat < _arrItems[0, PriceCol]) then
              _FontColor:= clRed;

            WriteValue(fWorkbookSource.Worksheet, iRow, iCol, _DS.Fields[iCol], [], clDefault, _FontColor);
          end;


         inc(iRow);
         _DS.Next;

         ProgressUpdate(pbTop);
        end;

      _DS.Close;

    end;
    _arrItems:= nil;
end;

constructor TwReport.Create(CreateSuspended: boolean);
begin
  fBase:= nil;
  fOwnerForm:= nil;
  inherited Create(CreateSuspended);
end;

destructor TwReport.Destroy();
begin
  inherited Destroy();
end;
end.

