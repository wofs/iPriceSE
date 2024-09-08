unit wTypesU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, db, fpspreadsheet, fpspreadsheetctrls, fpsTypes, Graphics, SysUtils;

type

  TThreadExceptionEvent = procedure(
    Thread: TThread; E: Exception
  ) of object;

  TIntValueNotify = procedure (Sender:TObject; aValue:integer) of object;
  TFloatValueNotify = procedure (Sender:TObject; aValue:double) of object;
  TStringValueNotify = procedure (Sender:TObject; aValue:string) of object;
  TSumCountNotify = procedure (Sender:TObject; aSum: double; aCount: integer) of object;

  TFileFormat = (ffXLS, ffXLSX, ffODS, ffCSV, ffYML, ffNONE);
  TwfExportFormat = (eftCSV, eftSpreadSheet);

  TFormatType = (ftPRICE, ftNAKL);
  TValueType = (vtDefault, vtNumber, vtString);
  TPriceType = (ptBase, ptPrice2, ptPrice3, ptPrice4, ptPrice5);

  TReportModes = (rmPriceDate, rmAnalogs, rmCompareHorisontal, rmSummaryInvoce, rmInvoceWithSelectOwnerCode, rmInvoceToPrice, rmToOwnerFiles, rmCatalogExportSpreadSheet, rmCatalogExportCSV);

  ArrayOfInteger  = array of integer;
  ArrayOfArrayInteger = array of array of integer;
  ArrayOfInt64    = array of int64;
  ArrayOfDouble   = array of double;
  ArrayOfCurrency   = array of Currency;
  ArrayOfString   = array of string;
  ArrayOfConst    = array of TVarRec;
  ArrayOfVariant  = array of variant;
  ArrayOfDateTime = array of TDateTime;
  ArrayOfArrayVariant = array of array of variant;

  const
    wfLineEnding = #10;
    wfLE = #10#13;

implementation


end.

