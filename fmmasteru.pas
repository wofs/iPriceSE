unit fmmasteru;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, fpsutils, LCLType, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ExtCtrls,
  Buttons, StdCtrls, ComCtrls, Menus, ActnList, Grids, LazUTF8, FmArcViewU, LazFileUtils,
  LCLIntf,
  wLogU, wBaseU, wFuncU, wFormatsGridU, wZipperU, wGetU, wTypesU,
  fpspreadsheetctrls, fpstypes, fpsallformats, fpspreadsheetgrid, fpscsv, wymlparser;

type

  { TFmMaster }

  TFmMaster = class(TForm)
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    cb_VidFormat: TComboBox;
    e_Owner: TEdit;
    e_File: TEdit;
    ImageListGrid: TImageList;
    Images16: TImageList;
    Label1: TLabel;
    lbLinkCBR: TLabel;
    od1: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    m_xlGrid: TPopupMenu;
    btnOwner: TSpeedButton;
    btnFile: TSpeedButton;
    sgFormat: TStringGrid;
    btnOpenFileWithExternalProg: TSpeedButton;
    Splitter1: TSplitter;
    TabCategory: TTabControl;
    xlBS: TsWorkbookSource;
    xlTab: TsWorkbookTabControl;
    xlGrid: TsWorksheetGrid;
    procedure btnCancelClick(Sender: TObject);
    procedure btnOpenFileWithExternalProgClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure lbLinkCBRClick(Sender: TObject);
    procedure btnFileClick(Sender: TObject);
    procedure xlGridClick(Sender: TObject);
    procedure xlGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure xlGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
  private
    FormIDent: string;
    wFormatName: string;
    fBase: TwBase;
    fFormatsGrid: TwFormatsGrid;

    procedure AsyncStringGridRepaint(Data: PtrInt);
    procedure SetStatus(_Text: string);

    property wFormID: string read FormIDent write FormIDent;
    { private declarations }
  public
    { public declarations }
    OwnerID: integer;
    property FormatName: string read wFormatName write wFormatName;
    procedure OpenFile(aFileName: string; const aIDFormat: integer=0);

    procedure _onChangeFormat(Sender: TObject);
    procedure _onFillGrid(Sender: TObject);
    procedure _onSavedFormat(Sender: TObject);
  end;

var
  FmMaster: TFmMaster;

implementation

{$R *.lfm}

{ TFmMaster }

procedure TFmMaster.OpenFile(aFileName:string; const aIDFormat: integer = 0);

  procedure SetFileFormatCmbx(aExt: string);
  var
    _arr: ArrayOfArrayVariant;
     _Ext: string;
  begin
     _Ext:= ReplaceStr(aExt,'.','');
     _arr:= fBase.SQLReadArr('SELECT ID FROM "FILEFORMAT" WHERE LOWER(CODE)='''+_Ext+'''');
     if Assigned(_arr) then
          fFormatsGrid.ComboBoxFileFormat.ItemIndex := cmbxItemIndexByID(fFormatsGrid.ComboBoxFileFormat ,integer(_arr[0,0]));
     _arr:=nil;

  end;

var
  _FileExt, _FileExtract, _TempPath: string;
  _FCONVERTLIBRE: boolean;
  _Form: TFmArcView;
  wZipper: TwZipper;
  _UnPackPath: RawByteString;
  fYML: Boolean;
  YMLFile: TYML;
  _FILEZIPNAMEDECODE: Integer;
begin
 fYML:= false;
 xlGrid.Clear;
 xlGrid.BeginUpdate;
 screen.Cursor:= crHourGlass;

 btnSave.Enabled:= true;
 aFileName:= SafePath(aFileName);

 _FileExt:= UTF8LowerCase(ExtractFileExt(aFileName));

 try
 xlBS.AutodetectFormat := false;

   case _FileExt of
     '.xls'       :
                  begin
                    xlBS.AutoDetectFormat:= true;
                    SetFileFormatCmbx(_FileExt);
                  end;
     '.xlsx'      :
                  begin
                    xlBS.AutoDetectFormat:= true;
                    SetFileFormatCmbx(_FileExt);
                  end;
     '.xlsm'      :
                  begin
                    xlBS.AutoDetectFormat:= true;
                    SetFileFormatCmbx(_FileExt);
                  end;
     '.ods'       :
                  begin
                    xlBS.FileFormat:= sfOpenDocument;
                    SetFileFormatCmbx(_FileExt);
                  end;
     '.htm'       :
                  begin
                    xlBS.FileFormat:= sfHTML;
                    SetFileFormatCmbx(_FileExt);
                  end;
     '.html'      :
                  begin
                    xlBS.FileFormat:= sfHTML;
                    SetFileFormatCmbx(_FileExt);
                  end;
     '.csv'       :
                  begin
                    xlBS.FileFormat:= sfCSV;
                    SetFileFormatCmbx(_FileExt);
                    CSVParams.Encoding:= fBase.SQLReadArr('SELECT CODE FROM "CODEPAGETEXT" WHERE ID='+IntTOStr(cmbxSelectID(fFormatsGrid.ComboBoxCodePage)))[0,0];
                    case fFormatsGrid.ComboBoxCSVDelimiter.ItemIndex of
                      0: CSVParams.Delimiter:= ';';
                      1: CSVParams.Delimiter:= ';';
                      2: CSVParams.Delimiter:= ',';
                      3: CSVParams.Delimiter:= '$';
                    end;
                  end;
     '.xml'       :
                  begin
                    //xlBS.FileFormat:= sfCSV;
                    fYML:=true;
                    SetFileFormatCmbx('yml');
                    fFormatsGrid.ChangeView('yml');
                    //CSVParams.Encoding:= _DBase.SQLReadArrRaw('SELECT CODE FROM "CODEPAGETEXT" WHERE ID='+IntTOStr(cmbxSelectID(fFormatsGrid.p_ComboBoxCodepage)))[0,0];
                  end;
     '.yml'       :
                  begin
                    //xlBS.FileFormat:= sfCSV;
                    fYML:=true;
                    SetFileFormatCmbx('yml');
                    fFormatsGrid.ChangeView('yml');
                    //CSVParams.Encoding:= _DBase.SQLReadArrRaw('SELECT CODE FROM "CODEPAGETEXT" WHERE ID='+IntTOStr(cmbxSelectID(fFormatsGrid.p_ComboBoxCodepage)))[0,0];
                  end;
     '.zip'       :
         begin
           _Form:= TFmArcView.Create(self);

             wZipper:= TwZipper.Create();

             if aIDFormat>0 then
                  begin
                    _FILEZIPNAMEDECODE:= fBase.SQLReadArr('FORMATS',['FILEZIPNAMEDECODE'],'ID='+IntToStr(aIDFormat),'')[0,0];
                    if _FILEZIPNAMEDECODE = 1 then wZipper.DecodeFileName:= true else wZipper.DecodeFileName:= false;
                  end;
             _Form.Zipper:= wZipper;
             _Form.ArcFileName := aFileName;
             _Form.ListFiles.Items:= wZipper.ReadFileList(aFileName);
           try
             _Form.ShowModal;

           finally
              Repaint;
              _FileExtract:= _Form.SelectedFileName;
              fFormatsGrid.SetGridCellValue('FILEZIPNAMEDECODE',_Form.DecodeFileName);

              if Length(_FileExtract)>0 then
               begin
                 _UnPackPath:= includeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
                 _UnPackPath:= _UnPackPath+'tmp';

                 if not DirectoryExistsUTF8(_UnPackPath) then ForceDirectoriesUTF8(_UnPackPath);
                 _UnPackPath:=_UnPackPath+DirectorySeparator+IntTOStr(OwnerID);
                 if not DirectoryExistsUTF8(_UnPackPath) then ForceDirectoriesUTF8(_UnPackPath);

                 fFormatsGrid.SetGridCellValue('FILE',UnsafePath(PathApplication_Unsafe,aFileName)+'|'+_FileExtract);
                 wZipper.ExtractOneFile(aFileName,_FileExtract,_UnPackPath);
                 aFileName:=_UnPackPath+DirectorySeparator+_FileExtract;
               end;

             wZipper.Destroy();
             _Form.Free;
           end;
         end;
     else
       begin
         ShowMessage('Формат не поддерживается.');
         exit;
       end;
   end;

 Application.ProcessMessages;
             _FileExt:=  ExtractFileExt(aFileName);
             case _FileExt of
                '.xls'       :
                             begin
                               xlBS.AutoDetectFormat:= true;
                               SetFileFormatCmbx(_FileExt);
                             end;
                '.xlsx'      :
                             begin
                               xlBS.AutoDetectFormat:= true;
                               SetFileFormatCmbx(_FileExt);
                             end;
                '.xlsm'      :
                             begin
                               xlBS.AutoDetectFormat:= true;
                               SetFileFormatCmbx(_FileExt);
                             end;
                '.ods'       :
                             begin
                               xlBS.FileFormat:= sfOpenDocument;
                               SetFileFormatCmbx(_FileExt);
                             end;
                '.htm'       :
                             begin
                               xlBS.FileFormat:= sfHTML;
                               SetFileFormatCmbx(_FileExt);
                             end;
                '.html'      :
                             begin
                               xlBS.FileFormat:= sfHTML;
                               SetFileFormatCmbx(_FileExt);
                             end;
                '.csv'       :
                             begin
                               xlBS.FileFormat:= sfCSV;
                               SetFileFormatCmbx(_FileExt);
                               CSVParams.Encoding:= fBase.SQLReadArr('SELECT CODE FROM "CODEPAGETEXT" WHERE ID='+IntTOStr(cmbxSelectID(fFormatsGrid.ComboBoxCodePage)))[0,0];

                               case fFormatsGrid.ComboBoxCSVDelimiter.ItemIndex of
                                 0: CSVParams.Delimiter:= ';';
                                 1: CSVParams.Delimiter:= ';';
                                 2: CSVParams.Delimiter:= ',';
                                 3: CSVParams.Delimiter:= '$';
                               end;
                             end;
                '.xml'      :
                             begin
                               //xlBS.FileFormat:= sfHTML;
                               if not fYML then
                                begin
                                  fYML:= true;
                                  SetFileFormatCmbx('yml');
                                  fFormatsGrid.ChangeView('yml');
                                end;
                             end;
                '.yml'      :
                             begin
                               //xlBS.FileFormat:= sfHTML;
                               if not fYML then
                                begin
                                  fYML:= true;
                                  SetFileFormatCmbx('yml');
                                  fFormatsGrid.ChangeView('yml');
                                end;
                             end;
           end;

 if _FileExt <> '.zip' then
  begin
    if fYML then
     begin

     try
       YMLFile:= TYML.Create(aFileName);
       try
         YMLFile.Open();
         ShowMessage('Файл успешно открыт и корректен! Указания полей не требуется.');
       except
        on E: Exception do
        begin
           wLog('YML Open','Ошибка : "' + E.Message + '"');
           ShowMessage('Ошибка открытия файла: "' + E.Message + '"');
        end;
       end;
     finally
        YMLFile.Destroy;
     end;


     end else
     begin
       if (_FileExt = '.xlsx') and (FileSize(aFileName)>5000000) then
        begin
         ShowMessage('Файл прайс-листа XLSX слишком большой для его визуализации. Откройте файл в программе просмотра и пропишите настройки колонок вручную. Столбцы указываются цифрами!');
         if MessageDlg('Открыть файл в программе по умолчанию? ',mtConfirmation, mbOKCancel, 0) = mrOK then
                OpenDocument(aFileName);
        end else
        begin
          _FCONVERTLIBRE:= fFormatsGrid.GetCellCheckBox('FCONVERTLIBRE');
          if _FCONVERTLIBRE then // если конвертируем, то
          begin
             aFileName:= ConvertFileWithLibreOffice(aFileName);
          end;

          xlBS.FileName :=aFileName;
        end;
     end;

  end
   else
   begin
      ShowMessage('Выберите файл в архиве!');
      fFormatsGrid.SetGridCellValue('FILE','');
      //sgFormat.Cells[1,12]:='';
   end;


 finally
        Screen.Cursor := crDefault;
       // xlGrid.SelectSheetByIndex(0);
        xlGrid.EndUpdate(true);
        xlGrid.Row:=1;
        xlGrid.Col:=1;
 end;

end;

procedure TFmMaster._onChangeFormat(Sender: TObject);
begin
   if fFormatsGrid.FillGridOn then exit;

   if not fBase.LongTransaction then
      fBase.LongTransaction:= true;

   btnSave.Enabled:= true;
end;

procedure TFmMaster._onFillGrid(Sender: TObject);
begin
   btnSave.Enabled:= false;
end;

procedure TFmMaster._onSavedFormat(Sender: TObject);
begin
  fBase.SQLTransactionEnd(true);
  //fBase.LongTransaction:= true;
  ShowMessage('Формат успешно сохранен!');
  btnSave.Enabled:=false;
end;

procedure TFmMaster.btnFileClick(Sender: TObject);
var
  _File: String;
begin
  if od1.Execute then
   begin
     _File:= UnsafePath(PathApplication_Unsafe,od1.FileName);
     e_File.Text:=_File;

     fFormatsGrid.SetGridCellValue('FILE',_File);

   end else
   begin
     exit();
   end;
   OpenFile(UTF8ToSys(_File));

end;

procedure TFmMaster.xlGridClick(Sender: TObject);
begin
   fFormatsGrid.xGridCurrentColRow:=[xlGrid.Col,xlGrid.Row,xlGrid.Workbook.GetWorksheetIndex(xlGrid.Worksheet)];
   fFormatsGrid.xGridCurrentColRowAddr:=GetCellString(xlGrid.Row-1,xlGrid.Col-1);
end;

procedure TFmMaster.xlGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  _Zoom: Double;
begin

  if (ssCtrl in Shift) and (Key in [VK_LCL_EQUAL, VK_LCL_MINUS, VK_0..VK_9])then
  begin
    _Zoom:= TsWorksheetGrid(Sender).ZoomFactor;

    if Key = VK_LCL_EQUAL then
      _Zoom:= _Zoom+0.1;

    if Key = VK_LCL_MINUS then
      _Zoom:= _Zoom-0.1;

    if Key = VK_0 then
      _Zoom:= 1;

    if Key in [VK_1..VK_9] then
    case Key of
      VK_1: _Zoom:= 0.1;
      VK_2: _Zoom:= 0.2;
      VK_3: _Zoom:= 0.3;
      VK_4: _Zoom:= 0.4;
      VK_5: _Zoom:= 0.5;
      VK_6: _Zoom:= 0.6;
      VK_7: _Zoom:= 0.7;
      VK_8: _Zoom:= 0.8;
      VK_9: _Zoom:= 0.9;
    end;

    TsWorksheetGrid(Sender).ZoomFactor:= _Zoom;
    TsWorksheetGrid(Sender).Repaint;
  end;

end;

procedure TFmMaster.xlGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  _Col, _Row: Longint;
begin
 _Col:= 0;
 _Row:= 0;

TsWorksheetGrid(Sender).MouseToCell(X, Y, _Col, _Row);

 xlGrid.Col:= _Col;
 xlGrid.Row:= _Row;
end;

procedure TFmMaster.SetStatus(_Text: string);
begin
  wStatus(wFormID,_Text,true);
end;

procedure TFmMaster.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if btnSave.Enabled then
       if MessageDlg(' Закрыть мастер без сохранения ?',mtConfirmation, mbOKCancel, 0) = mrCancel then CanClose:=false;
end;

procedure TFmMaster.btnSaveClick(Sender: TObject);
begin
  try
    fFormatsGrid.Save(Sender);

  except
    on E: Exception do
    begin
      __Log.SaveLogError(E);
        SetStatus('Сбой сохранения формата!');
        wLog('FmMaster','Ошибка: "' + E.Message + '"');
        wLog('FmMaster','Сбой сохранения формата!');
        ShowMessage('Ошибка: "' + E.Message + '"');

     end;
  end;


end;

procedure TFmMaster.btnCancelClick(Sender: TObject);
begin
  fBase.SQLTransactionEnd(false);
  Close();
end;

procedure TFmMaster.btnOpenFileWithExternalProgClick(Sender: TObject);
begin
  if MessageDlg('Открыть накладную во внешней программе просмотра?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;
  if Length(e_File.Text)=0 then exit;

  OpenDocument(e_File.Text);
end;

procedure TFmMaster.FormCreate(Sender: TObject);
var
  _res: ArrayOfVariant;
begin
wFormID:=Self.Name;
FormatName:= '';

 wLog('FmMaster','Инициализация формы... ['+wFormID+']');

 try

 //DBase
 fbase:= TwBase.Create(Sender);


 cmbxFill(cb_VidFormat,fBase.SQLReadDS('FORMATS_CATEGORY',['NAME','ID'],'FCLOSE=0','ID'),['NAME','ID']);

 fFormatsGrid:= TwFormatsGrid.Create(Sender,sgFormat,cb_VidFormat,true,fBase);
 fFormatsGrid.TabCategory:= TabCategory;
 fFormatsGrid.onChangedFormat:= @_onChangeFormat;
 fFormatsGrid.onFillGrid:= @_onFillGrid;
 fFormatsGrid.onSavedFormat:=@_onSavedFormat;
 fFormatsGrid.xGridPopupMenu:= m_xlGrid;
 fFormatsGrid.MasterMode:= true;
 wLog('FmMaster','Инициализация формы успешно завершена.');

 except
   on E: Exception do
   begin
     __Log.SaveLogError(E);
       SetStatus('Сбой инициализации формы.');
       wLog('FmMaster','Ошибка [FmCreate]: "' + E.Message + '"');
       wLog('FmMaster','Сбой инициализации формы.');
       ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

    end;
 end;
end;

procedure TFmMaster.FormDestroy(Sender: TObject);
begin
try
  wLog('FmPriceFields','Выгрузка формы...');

   // выгружаем DBase
  fBase.Destroy();

   // выгружаем wFormatsGrid;
  fFormatsGrid.Destroy();

  cmbxClearData(cb_VidFormat);

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

procedure TFmMaster.FormResize(Sender: TObject);
begin
  sgFormat.Repaint;
end;

procedure TFmMaster.AsyncStringGridRepaint(Data: PtrInt);
begin
 sgFormat.Repaint;
end;

procedure TFmMaster.FormShow(Sender: TObject);
begin
  fFormatsGrid.FillGrid(sgFormat.Tag, self);
  Application.QueueAsyncCall(@AsyncStringGridRepaint,0);
end;

procedure TFmMaster.lbLinkCBRClick(Sender: TObject);
begin
     OpenURL('http://www.cbr.ru/currency_base/daily.aspx');
end;

end.

