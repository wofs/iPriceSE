program project1;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads, cmem
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, lazcontrols, datetimectrls, tachartlazaruspkg, lazdbexport, dbflaz, ibexpress, FmMainU, fmmasteru, wLogU, wPlugin, FmNomenclatureEditU, wFuncU,
  FmTreeU, wTabU, FmPriceFieldsU, FmNomenclatureEditMassU, wFormatsGridU, wDBImportU, FmAboutU, FmListSelectU, FmQuantityInPackingU, pkgUtilsU,
  FmMatchingEditU, FmWaitU, FmMatchingAddU, FmArcViewU, wTViewerSpreadsheetU, wZipperU, win1251decoder, FmURLsU, mCatalogU, wBaseU, wDBGridU, mPriceLists,
  pkgAnalisisU, mAnalisisU, wGetU, FmFormatsImportU, pkgOrdersU, FmReportU, mInvoceU, FmInvoceU, wReportU, FmCatalogExportU;

{$R *.res}

begin
  Application.Title:='iPriceSE';
  RequireDerivedFormResource:=True;
  Application.Initialize;
  Application.CreateForm(TFmMain, FmMain);
  Application.Run;
end.

