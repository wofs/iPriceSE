unit wDBGridU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

// Назначение:
// Простая работа с DBGrid

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DBGrids, Grids, Dialogs, Controls, fgl, LCLType, Graphics, StdCtrls,Buttons, ComCtrls, Forms,
  LazUTF8, LCLIntf, wCustomClassThreadU,
  fpcsvexport, fpDBExport, fpsexport, Clipbrd,
  db,
  IBCustomDataSet, IBDatabase,
  wBaseU,wFormulaU, wDBTreeU,
  wLogU, wFuncU, wTProgressU, wTypesU, wYMLparser
  ;
type
  TIntegerList = specialize TFPGList<Int64>;
  { TwDBGrid }

  TSelectEvent =
      procedure(Sender: TObject) of object;

  { TDataExportThread }

  TDataExportThread = class (TwCustomThreadWithProgressBar)
    protected
      procedure Execute; override;
    private
      fCSV: TCSVExporter;
      fBase: TwBase;
      fSQL: string;
      fGrid:   TDBGrid;
      fFileName: string;
      fFileFormat: TFileFormat;
      fSpreadSheet: TFPSExport;
      procedure OnProgress(Sender: TObject; const ItemNo: Integer);
    public
      constructor Create(CreateSuspended: boolean);
      destructor Destroy(); override;

      property DBGrid: TDBGrid read fGrid write fGrid;
      property FileName: string read fFileName write fFileName;
      property FileFormat: TFileFormat read fFileFormat write fFileFormat;

      procedure Stop;

      function ExportGridToFile(aDBGrid: TDBGrid; aFileName: string; aExportFormat: TFileFormat): boolean;
      function ExportGridToCSVFile(aDBGrid: TDBGrid; aFileName: string): boolean;
      function ExportGridToSpreadSheetFile(aDBGrid: TDBGrid; aFileName: string; aExportFormat: TFileFormat): boolean;
  end;


  TwDBGrid = class

    private
      fDataExport: TDataExportThread;
      fonFill: TNotifyEvent;
      fProgress: TProgress;
      const
        cEdSearchDropCountHistory = 15;

      var
      fFormulaText: string;
      fResult: boolean;
      fSearchEditHistoryFile: String;

      fGrid: TDBGrid; // Грид, которым управляем
      fGroupArray: ArrayOfInteger; // Массив группировок для фильтрации
      fGroupField: string; // Плоле, по которому будет фильтрации
      fFormula: TFormula;
      fIDCurrentRecord: integer;
      FonGridCellClick: TNotifyEvent;
      fSearchSplitStringBtn: TSpeedButton;
      fSortON: boolean;
      FSortTitleImagesIndex: ArrayOfInteger;
      fSQL: string; // запрос на выборку
      fOrderBy: string;
      fRow: integer;
      fCol: integer;
      fFieldName: string;
      fStaticTextSelection: TStaticText;

      fFillGridNow: boolean; // флаг, означающий что грид занят загрузкой
      fMultiSelect: boolean; // разрешает мультивыделение

      fdbDataSet: TIBDataSet;
      fdbDataSource: TDataSource;
      fdbTransaction: TIBTransaction;
      fBase: TwBase;
      fTree: TwDBTree;

      fSearchEdit: TComboBox; // Поле ввода искомого текста
      fSearchStrings: ArrayOfString; //список строк для поиска
      fSearchEntryArray: ArrayOfString; // вхождение
      fSearchParticleArray: ArrayOfString; //точно

      fSearchComplete: Boolean; // флаг окончания поиска

      fSearchPreventiveBtn: TSpeedButton; //кнопка превентивного поиска

      fSelectedRowsList: TIntegerList; // список, содержащий в себе список выбеленных строк
      fSelectEvent: TSelectEvent;

      fShiftState: TShiftState; //
      fmLeft_kCtrl: Boolean; // флаг нажатия ЛКМ
      fWhere: string;

      fGridImageList: TImageList;

      procedure Edit_DropDown(Sender: TObject);
      procedure Edit_Enter(Sender: TObject);
      procedure Edit_KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure Edit_SaveInputValues(aEdit: TComboBox);
      procedure ExportData(aFileName: string; aFileFormat: TFileFormat);
      function getBookmark: TBookMark;
      function GetFieldValue(aFieldName: string ): TField;

      function GetSelectedRowsCount: integer;
      function GetSQL: string;
      function GetTableName(aGroupField: string): string;
      function GetFieldName(aGroupField: string): string;
      function IsExistsField(aFieldName: string): boolean;
      procedure Log(aText: string);
      procedure onEndThread(Sender: TObject);
      procedure onExceptionEvent(Thread: TThread; E: Exception);

      procedure onKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure onKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure onMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
      procedure onMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

      procedure onDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);

      procedure onDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
      procedure onEndDrag(Sender, Target: TObject; X, Y: Integer);
      procedure onCellClick(Column: TColumn);

      procedure DataSet_onCalcFields(DataSet: TDataSet);

      procedure Edit_onChange(Sender: TObject);
      procedure Edit_onExit(Sender: TObject);
      procedure Edit_onKeyPress(Sender: TObject; var Key: char);
      procedure onProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
      procedure onProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
      procedure onStopForce(Sender: TObject);
      procedure SearchSplitStringBtn_onClick(Sender: TObject);

      procedure SelectRow(Sender: TObject);
      procedure setBookmark(AValue: TBookMark);
      procedure SetfSearchSplitStringBtn(aValue: TSpeedButton);
      procedure SetSearchEdit(AValue: TComboBox);
      procedure SetSearchText(AValue: string);
      procedure SetSelectAll(AValue: boolean);
      procedure SelectionChanged();

      function SorTwDBGrid(aColumn: TColumn): string;
      function MakeSearchString(aText: string):string;
    protected
      function CopyToClipboardWriteString(const aDataSet: TDataSet; const aCurrencyFields: ArrayOfString; const aFieldsArr: ArrayOfString): string;
    public
      property Base: TwBase read fBase;
      property Tree: TwDBTree read fTree write fTree;
      property SQL: string read GetSQL write fSQL;
      property FillGridNow: boolean read fFillGridNow write fFillGridNow;
      property Grid: TDBGrid read fGrid write fGrid;
      property MultiSelect: boolean read fMultiSelect write fMultiSelect;
      property SortON: boolean read fSortON write fSortON;
      property SortTitleImagesIndex: ArrayOfInteger read FSortTitleImagesIndex write FSortTitleImagesIndex;
      property SearchEdit: TComboBox read fSearchEdit write SetSearchEdit;
      property SearchText: string write SetSearchText;
      property SearchPreventiveBtn: TSpeedButton read fSearchPreventiveBtn write fSearchPreventiveBtn;
      property SearchSplitStringBtn: TSpeedButton read fSearchSplitStringBtn write SetfSearchSplitStringBtn;
      property SearchEntryArray: ArrayOfString write fSearchEntryArray;
      property SearchComplete: Boolean write fSearchComplete;
      property SearchParticleArray: ArrayOfString write fSearchParticleArray;
      property GroupArray: ArrayOfInteger read fGroupArray write fGroupArray;
      property GroupField: string write fGroupField;
      property Where: string read fWhere write fWhere;

      property SelectAll:boolean write SetSelectAll;
      property SelectedRowsCount:integer read GetSelectedRowsCount;
      property SelectedRowsList: TIntegerList read fSelectedRowsList write fSelectedRowsList;
      property onSelect:TSelectEvent read fSelectEvent write fSelectEvent;

      property onGridCellClick: TNotifyEvent read FonGridCellClick write FonGridCellClick;

      property FieldName: string read fFieldName write fFieldName;
      property ShiftState: TShiftState read fShiftState write fShiftState;

      //FORMULA
      property Formula: TFormula read fFormula write fFormula;
      property FormulaText: string read fFormulaText write fFormulaText;

      property StaticTextSelection: TStaticText read fStaticTextSelection write fStaticTextSelection;

      property FieldValue[aFieldName:string ]:TField read GetFieldValue;
      property Bookmark: TBookMark read getBookmark write setBookmark;
      property onFill: TNotifyEvent read fonFill write fonFill;

      constructor Create(aBase:TwBase; aGrid: TDBGrid; aSQL: string);
      destructor Destroy(); override;

      procedure Fill(const aSQL: string ='');
      procedure onTitleClick(aColumn: TColumn);
      procedure SelectedRowsClear();
      procedure HighLightText(Sender: TObject; const aFieldName: string; Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
      procedure ExportData();
      procedure CopyToClipboard(aFieldsArr: ArrayOfString; aCurrencyFields: ArrayOfString; const aIDField:string = 'ID');
      procedure CopyToClipboard(aDataSet: TDataSet; aFieldsArr: ArrayOfString; aCurrencyFields: ArrayOfString);
      procedure SetColumnCaption(aFieldName, aCaption: string);

      function SelectedRowsListIndexOf(aID: integer): integer;
      function SelectedRows: ArrayOfInteger;
  end;

implementation

{ TDataExportThread }

procedure TDataExportThread.Execute;
begin
    try
      try
        onProgressInit(pbTop,fBase.GetRowsCount(fSQL));

         Result:= ExportGridToFile(fGrid, fFileName, fFileFormat);

         Result:= not StopForce;
       finally
        onEndThread(self);
      end;
    except
      on E: Exception do
        onExceptionEvent(self,E);
    end;
end;

procedure TDataExportThread.OnProgress(Sender: TObject; const ItemNo: Integer);
begin
  if Assigned(onProgressUpdate) then
     onProgressUpdate(pbTop,ItemNo);
end;

constructor TDataExportThread.Create(CreateSuspended: boolean);
begin
  inherited Create(CreateSuspended);
end;

destructor TDataExportThread.Destroy();
begin
  inherited Destroy();
end;

procedure TDataExportThread.Stop;
begin
  if Assigned(fCSV) then
     fCSV.Cancel;
  if Assigned(fSpreadSheet) then
     fSpreadSheet.Cancel;

  inherited Stop;
end;

function TDataExportThread.ExportGridToFile(aDBGrid: TDBGrid; aFileName: string; aExportFormat: TFileFormat): boolean;
var
  _BookMark: TBookMark;
begin
  StopForce:= false;

  _BookMark:= fGrid.DataSource.DataSet.Bookmark;
  fGrid.DataSource.DataSet.DisableControls;

  try
    case aExportFormat of
      ffCSV: Result:= ExportGridToCSVFile(aDBGrid, aFileName);
      ffODS,
      ffXLS,
      ffXLSX : Result:= ExportGridToSpreadSheetFile(aDBGrid, aFileName, aExportFormat)
      else
        Result:= false;
    end;
  finally
    fGrid.DataSource.DataSet.Bookmark:= _BookMark;
    fGrid.DataSource.DataSet.EnableControls;
  end;
end;

function TDataExportThread.ExportGridToCSVFile(aDBGrid: TDBGrid; aFileName: string): boolean;
var
  i: Integer;
  _Field: TExportFieldItem;
begin
  Result:= true;
  try
    fCSV:= TCSVExporter.Create(fGrid);
    fCSV.FileName:= aFileName;
    fCSV.Dataset:= fGrid.DataSource.DataSet;
    fCSV.FormatSettings.FieldDelimiter:= ';';
    fCSV.OnProgress:= @OnProgress;

    fGrid.DataSource.DataSet.First;

    for i:=0 to fGrid.Columns.Count-1 do
    begin
      if Length(fGrid.Columns[i].Title.Caption)>0 then
         begin
          _Field:= fCSV.ExportFields.AddField(fGrid.Columns[i].FieldName);
          _Field.ExportedName:= fGrid.Columns[i].Title.Caption;
         end;
    end;

    try

     fCSV.Execute;
    finally
     fCSV.Free;
     fCSV:= nil;
    end;

  except
    Result:= false;
  end;
end;

function TDataExportThread.ExportGridToSpreadSheetFile(aDBGrid: TDBGrid; aFileName: string; aExportFormat: TFileFormat): boolean;
  var
    i: Integer;
    _Field: TExportFieldItem;
    _ExportFormat: TExportFormat;
  begin
    Result:= true;
    try
      fSpreadSheet:= TFPSExport.Create(fGrid);
      fSpreadSheet.FileName:= aFileName;
      fSpreadSheet.Dataset:= fGrid.DataSource.DataSet;
      fSpreadSheet.FormatSettings.HeaderRow:= true;
      //fSpreadSheet.on

      fSpreadSheet.OnProgress:= @OnProgress;
      fGrid.DataSource.DataSet.First;

      for i:=0 to fGrid.Columns.Count-1 do
      begin
        if Length(fGrid.Columns[i].Title.Caption)>0 then
           begin
            _Field:= fSpreadSheet.ExportFields.AddField(fGrid.Columns[i].FieldName);
            _Field.ExportedName:= fGrid.Columns[i].Title.Caption;
           end;
      end;

      try

      case aExportFormat of
        ffXLS    : _ExportFormat:= efXLS;
        ffXLSX   : _ExportFormat:= efXLSX;
        ffODS    : _ExportFormat:= efODS;
      end;

       fSpreadSheet.FormatSettings.ExportFormat:= _ExportFormat;
       fSpreadSheet.Execute;
      finally
       fSpreadSheet.Free;
       fSpreadSheet:= nil;
      end;

    except
      Result:= false;
    end;
end;

{ TwDBGrid }

procedure TwDBGrid.Log(aText: string);
begin
  if __onLog and assigned (__Log) then
     wLog('[Grid] ', aText);
end;

procedure TwDBGrid.onEndThread(Sender: TObject);
var
  _Status: String;
begin
    if Assigned(fDataExport) then
    begin
      fResult := fDataExport.Result;
      fDataExport.Terminate;
      fDataExport := nil;
    end;

    if fResult then
      _Status:='Экспорт успешно завершен.'
    else
      _Status:= 'При экспорте произошла ошибка!';

    ShowMessage(_Status);

    fProgress.ForceClose;
end;

procedure TwDBGrid.onExceptionEvent(Thread: TThread; E: Exception);
begin
    MessageDlg(E.Message,
            mtError, [mbOK], 0);
end;

constructor TwDBGrid.Create(aBase: TwBase; aGrid: TDBGrid; aSQL: string);
var
  i: Integer;
begin
  fGrid:= aGrid;

  fGrid.OnTitleClick:= @onTitleClick;
  fGrid.onMouseDown:= @onMouseDown;
  fGrid.OnMouseUp:= @onMouseUp;
  fGrid.OnKeyDown:= @onKeyDown;
  fGrid.OnKeyUp:= @onKeyUp;
  if not Assigned(fGrid.OnDrawColumnCell) then fGrid.OnDrawColumnCell:= @onDrawColumnCell;
  fGrid.OnDragOver:= @onDragOver;
  fGrid.OnEndDrag:= @onEndDrag;
  fGrid.onCellClick:= @onCellClick;

  fGrid.SelectedColor:= fGrid.FixedHotColor;
  fGrid.Options:=fGrid.Options-[dgMultiselect]+[dgRowHighlight]+[dgAlwaysShowSelection]+[dgTruncCellHints]-[dgEditing];//-[dgIndicator];//+[dgRowHighlight];
  //fGrid.Options:=fGrid.Options-[dgMultiselect]+[dgRowHighlight]+[dgAlwaysShowSelection]+[dgTruncCellHints]-[dgEditing];//-[dgIndicator];//+[dgRowHighlight];

  MultiSelect:= false;
  fSortON:= true;

  fSearchEdit:= nil;
  fSearchEditHistoryFile:= PathApplication_Unsafe+'iPrice.search';
  fSearchPreventiveBtn:= nil;
  fSearchComplete:= true;
  fGroupArray:= nil;
  fGroupField:= '';
  fWhere:= '';
  fOrderBy:= '';
  fFormula:= nil;

   for i:=0 to fGrid.Columns.Count-1 do
       fGrid.Columns[i].Tag:=0;

  fSelectedRowsList:= TIntegerList.Create;

  fSQL:= aSQL;
  fBase:= aBase;
  fdbDataSet:= TIBDataSet.Create(fGrid);
  fdbDataSource:= TDataSource.Create(fGrid);
  fdbTransaction:= TIBTransaction.Create(fGrid);

  with fdbTransaction do
  begin
    DefaultDatabase := Base.DataBase;
    Params.Add('read');
    Params.Add('read_committed');
    Params.Add('rec_version');
    Params.Add('nowait');
  end;

  with fdbDataSet do
  begin
    Database := Base.DataBase;
    //Transaction := fdbTransaction;
    AutoCalcFields := False;
  end;
  fdbDataSource.DataSet:= fdbDataSet;


end;

destructor TwDBGrid.Destroy();
begin
  fGrid:=nil;
  fBase:= nil;

  if Assigned(fFormula) then fFormula.Destroy();
  FreeAndNil(fdbDataSet);
  FreeAndNil(fdbDataSource);
  FreeAndNil(fdbTransaction);
  fSelectedRowsList.Free;

  inherited Destroy();
end;

function TwDBGrid.GetTableName(aGroupField: string): string;
var
  _StartPos: PtrInt;
begin
  Result:='';
  _StartPos:= UTF8Pos(#34,aGroupField)+1;
  if _StartPos>0 then
     Result:= UTF8Copy(aGroupField,_StartPos, UTF8Pos(#34,aGroupField,_StartPos)-_StartPos);
end;

function TwDBGrid.GetSelectedRowsCount: integer;
begin
  Result:= fSelectedRowsList.Count;
end;

function TwDBGrid.GetSQL: string;
begin
  Result:= fdbDataSet.SelectSQL.Text;
end;

function TwDBGrid.GetFieldName(aGroupField: string): string;
var
  _StartPos: PtrInt;
begin
  Result:='';
  _StartPos:= UTF8Pos('.',aGroupField)+1;
  if _StartPos>0 then
     Result:= UTF8Copy(aGroupField,_StartPos, UTF8Length(aGroupField));
end;

procedure TwDBGrid.Fill(const aSQL: string);
var
  _SearchEntry: string; // вхождение
  _SearchParticle: string; //точно
  i, iEntry, iParticle: Integer;
  _SQL, _GroupFilter, _WhereFilter, _SearchEntry2: string;
  _Field: TFloatField;
  _PosOrderBy: PtrInt;
begin
  FillGridNow:=true;
  Log('Заполняем Grid... ['+Grid.Name+']');

  if Length(aSQL)>0 then fSQL:= aSQL;

  if Assigned(Grid.DataSource) then
      fGrid.DataSource.DataSet.DisableControls;

  try
    _SearchEntry:='';
    _SearchParticle:='';
    _GroupFilter:='';

    _SQL:= fSQL;

    if Length(_SQL)=0 then exit;

    if Assigned (fSearchStrings) then
     begin
       for i:=0 to High(fSearchStrings) do
       begin
          if Assigned(fSearchEntryArray) then
           begin
             if i>0 then _SearchEntry:= _SearchEntry+' AND ';
             _SearchEntry2:='';
             for iEntry:=0 to High(fSearchEntryArray) do
             begin
               if iEntry>0 then _SearchEntry2:= _SearchEntry2+' OR ';
               _SearchEntry2:= _SearchEntry2+' UPPER('+fSearchEntryArray[iEntry]+') LIKE UPPER('+QuotedStr('%'+fSearchStrings[i]+'%')+')';
             end;
               _SearchEntry:= _SearchEntry+'('+_SearchEntry2+')';
           end;

          if Assigned(fSearchParticleArray) and (High(fSearchStrings)=0) then
           begin
             if i>0 then _SearchParticle:= _SearchParticle+' OR ';
             for iParticle:=0 to High(fSearchParticleArray) do
             begin
                if iParticle>0 then _SearchParticle:= _SearchParticle+' OR ';
                if UTF8Pos('=',fSearchParticleArray[iParticle])=0 then
                   _SearchParticle:= _SearchParticle+fSearchParticleArray[iParticle]+'='+QuotedStr(fSearchStrings[i])
                else
                   _SearchParticle:= _SearchParticle+ Format(fSearchParticleArray[iParticle],[VarToStr(fSearchStrings[i])]);




             end;
             if Length(_SearchEntry)>0 then
             _SearchParticle:= ' OR ('+_SearchParticle+')' else
             _SearchParticle:= '('+_SearchParticle+')';
           end;
       end;
       _SQL:= StringReplace(_SQL,'/*and_search_string*/',' AND ('+_SearchEntry+' '+_SearchParticle+') ',[rfReplaceAll]);
       _SQL:= StringReplace(_SQL,'/*search_string*/',' ('+_SearchEntry+' '+_SearchParticle+') ',[rfReplaceAll]);
     end;

    if Assigned(fGroupArray) and (Length(fGroupField)>0) then
    begin
      _GroupFilter:= Base.PrepareWhereString(fGroupField, fGroupArray);
      _SQL:= StringReplace(_SQL,'/*and_group_string*/',' AND ('+_GroupFilter+') ',[rfReplaceAll]);
      _SQL:= StringReplace(_SQL,'/*group_string*/',' ('+_GroupFilter+') ',[rfReplaceAll]);
    end;

    if (Length(fWhere)>0) then
    begin
      _WhereFilter:= fWhere;//Base.PrepareWhereString(fGroupField, fGroupArray);
      _SQL:= StringReplace(_SQL,'/*and_where_string*/',' AND ('+_WhereFilter+') ',[rfReplaceAll]);
      _SQL:= StringReplace(_SQL,'/*where_string*/',' ('+_WhereFilter+') ',[rfReplaceAll]);
    end;

    if Assigned(Formula) and ((Length(fFormulaText)>0)) then
    begin
      _SQL:= StringReplace(_SQL,'/*formula*/',' ('+fFormulaText+') as '+fFormula.CalculateField+', ',[rfIgnoreCase]);
    end;

    if Length(fOrderBy)>0 then
    begin
       _PosOrderBy:= UTF8Pos('ORDER BY',_SQL);
       if _PosOrderBy>0 then
          _SQL:= UTF8Copy(_SQL,1,_PosOrderBy-1);

       _SQL:=_SQL+' ORDER BY '+fOrderBy;
    end;

    with fdbDataSet do begin
      if Assigned(Formula) and (Length(fFormulaText)=0) then
      begin
        Close;
        SelectSQL.Text:=_SQL;

        if not AutoCalcFields then
        begin
          /////

          FieldDefs.Clear;
          FieldDefs.Update;
          {добавляем поля из запроса}
          for i := 0 to FieldDefs.Count - 1 do
            FieldDefs[i].CreateField(fdbDataSet);

          {добавляем вычисляемое поле}
          _Field := TFloatField.Create(fdbDataSet);
          _Field.FieldName := Formula.CalculateField;
          _Field.FieldKind := fkCalculated;
          _Field.Calculated := True;

          with fdbDataSet.FieldDefs.AddFieldDef do
          begin
            Name := Formula.CalculateField;
            DataType := ftFloat;
          end;

          _Field.DataSet := fdbDataSet;
          onCalcFields := @DataSet_onCalcFields;    // назначаем обработчик
          AutoCalcFields := True;
          //////
        end;
        Base.SQLReadDS(fdbDataSet, nil, fdbDataSource, '');
      end
      else
      begin
        AutoCalcFields := False;
        Base.SQLReadDS(fdbDataSet, nil, fdbDataSource, _SQL);
      end;
    end;

    Grid.DataSource:= fdbDataSource;

  //  Base.BaseFormula:=nil;
    FillGridNow:=false;

    if Assigned(fonFill) then fonFill(self);
  finally
   if Assigned(Grid.DataSource) then
      fGrid.DataSource.DataSet.EnableControls;
  end;
end;

function TwDBGrid.SorTwDBGrid(aColumn: TColumn):string;
var
  i: integer;
begin
  Result:= '';

  for i:=1 to Grid.Columns.Count-1 do
  begin
   if aColumn.FieldName = Grid.Columns[i].FieldName then
    begin
       case Grid.Columns[i].Tag of
          0 : Grid.Columns[i].Tag:= 1;
          1 : Grid.Columns[i].Tag:= -1;
         -1 : Grid.Columns[i].Tag:= 0;
       end;
    end else
        Grid.Columns[i].Tag:= 0;
  end;

  for i:=1 to Grid.Columns.Count-1 do
  begin
    case Grid.Columns[i].Tag of
       0 : Grid.Columns[i].Title.ImageIndex:=-1;
       1 :
         begin
             Result:= Grid.Columns[i].FieldName+' ASC';
          if Assigned(FSortTitleImagesIndex) then
             Grid.Columns[i].Title.ImageIndex:=FSortTitleImagesIndex[0];
         end;
      -1 :
        begin
             Result:= Grid.Columns[i].FieldName+' DESC';
         if Assigned(FSortTitleImagesIndex) then
             Grid.Columns[i].Title.ImageIndex:=FSortTitleImagesIndex[1];
        end;
    end;
  end;

end;

function TwDBGrid.MakeSearchString(aText: string): string;
begin
  if Length(aText)=0 then
   begin
      Result:='';
      Exit;
   end;

end;

function TwDBGrid.CopyToClipboardWriteString(const aDataSet: TDataSet; const aCurrencyFields: ArrayOfString; const aFieldsArr: ArrayOfString): string;
var
  i: Integer;
begin
  with aDataSet do
  begin
     Result:= '';
     for i:=0 to High(aFieldsArr) do begin
       if Length(VarToStr(aFieldsArr[i]))>0 then
         if  Length(FieldByName(aFieldsArr[i]).AsString)>0 then
         begin
           if Length(Result)>0 then Result:= Result+' | ';
              if  fBase.TextInArray(aFieldsArr[i], aCurrencyFields) then
                Result:= Result+ CurrToStrF(FieldByName(aFieldsArr[i]).AsCurrency, ffCurrency, 2)
              else
                Result:= Result+ FieldByName(aFieldsArr[i]).AsString;
         end;
     end;
  end;
end;

procedure TwDBGrid.onTitleClick(aColumn: TColumn);
var
  _ID: integer;
begin
   if not fSortON then exit;

  if aColumn.FieldName<>'ID' then
   begin
    fOrderBy:= SorTwDBGrid(aColumn);
    Fill();
   end;

   Log('ORDER BY '+fdbDataSet.OrderFields);
end;

procedure TwDBGrid.SelectRow(Sender: TObject);
var
  _Dataset: TDataSet;
  _Index: integer;
begin
  try
  if fMultiSelect then
  begin

     if not Assigned(TDBGrid(Sender).DataSource) then exit;

     _DataSet:=TDBGrid(Sender).DataSource.DataSet;


     if _DataSet.RecordCount=0 then exit;

     _Index:= fSelectedRowsList.IndexOf(_DataSet.FieldByName('ID').AsInteger);

     if _Index = -1 then
        fSelectedRowsList.Add(_DataSet.FieldByName('ID').AsInteger) else
        fSelectedRowsList.Delete(_Index);

     grid.Repaint;

     SelectionChanged();
    end;

  except
     on E: Exception do
     begin
       __Log.SaveLogError(E);
       Log('Ошибка [SelectRow]: "' + E.Message + '"');
       ShowMessage('Ошибка [SelectRow]: "' + E.Message + '"');
       raise;
     end;
  end;
end;

procedure TwDBGrid.setBookmark(AValue: TBookMark);
begin
  if Assigned(Grid.DataSource.DataSet) and (Grid.DataSource.DataSet.RecordCount > 0) then
     Grid.DataSource.DataSet.Bookmark:= AValue;
end;

procedure TwDBGrid.SetfSearchSplitStringBtn(aValue: TSpeedButton);
begin
  fSearchSplitStringBtn:=aValue;
  if Assigned(fSearchSplitStringBtn) then
     fSearchSplitStringBtn.OnClick:= @SearchSplitStringBtn_onClick;
end;

procedure TwDBGrid.SetSearchEdit(AValue: TComboBox);
begin
  if fSearchEdit=AValue then Exit;
  fSearchEdit:=AValue;
  if Assigned(fSearchEdit) then
  begin
    with fSearchEdit do begin
       OnChange:= @Edit_onChange;
       OnExit:= @Edit_onExit;
       OnKeyPress:= @Edit_onKeyPress;
       OnKeyDown:= @Edit_KeyDown;
       OnDropDown:=@Edit_DropDown;
       OnEnter:=@Edit_Enter;
       //Style:= csSimple;
       Hint:='Чтобы показать список последних поисковых запросов - нажмите F3.';
       ParentShowHint:= false;
       ShowHint:= true;
       DropDownCount:= cEdSearchDropCountHistory;
       AutoSelect:= false;
       fSearchEdit.DroppingDown:= true;

       if not FileExists(fSearchEditHistoryFile) then
         fSearchEdit.Items.SaveToFile(fSearchEditHistoryFile)
       else
         fSearchEdit.Items.LoadFromFile(fSearchEditHistoryFile);

    end;

  end else
    fSearchStrings:=nil;
end;

procedure TwDBGrid.SetSearchText(AValue: string);
begin
  fSearchStrings:= Base.MakeArrayFromString(AValue);
end;

procedure TwDBGrid.SetSelectAll(AValue: boolean);
var
  ABookmark: TBookMark;
  _RowCount: Integer;
begin
  if AValue then
  begin
    Grid.Cursor:= crSQLWait;

   try
     _RowCount:= fBase.GetRowsCount(fdbDataSet.SelectSQL.Text);
   finally
     Grid.Cursor:= crDefault;
   end;

    if _RowCount > 100000 then
        if MessageDlg('Невозможно выделить ('+IntToStr(_RowCount)+') позиций! Примените один из доступных фильтров и попробуйте снова.',mtError, [mbOK], 0) = mrOK then exit;

    if _RowCount > 50000 then
    begin
        if MessageDlg('Вы действительно выделить ('+IntToStr(_RowCount)+') позиций?',mtWarning, mbOKCancel, 0) = mrCancel then exit;
    end else
    begin
     if _RowCount > 3000 then
          if MessageDlg('Вы действительно хотите выделить ('+IntToStr(_RowCount)+') позиций?',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;
    end;

    //fSelectedRowsList.Clear;
    Grid.Cursor:= crSQLWait;

    with Grid.DataSource.Dataset do begin
      if (BOF and EOF) then Exit;
      DisableControls;
      try
        ABookmark := GetBookmark;
        try
          First;
          while not EOF do begin
            if fSelectedRowsList.IndexOf(Grid.DataSource.DataSet.FieldByName('ID').AsInteger) < 0 then
                         fSelectedRowsList.Add(Grid.DataSource.DataSet.FieldByName('ID').AsInteger);
            Next;
          end;
        finally
          try
            GotoBookmark(ABookmark);
          except
          end;
          FreeBookmark(ABookmark);
        end;
      finally
        EnableControls;
        Grid.Cursor:= crDefault;
      end;
    end;
  end else
    SelectedRowsClear();

    SelectionChanged();
end;

procedure TwDBGrid.SelectionChanged();
begin
  if Assigned(fStaticTextSelection) then fStaticTextSelection.Caption:='  Выделено: '+IntToStr(SelectedRowsCount);

  if Assigned(onSelect) then onSelect(self);
end;

function TwDBGrid.SelectedRowsListIndexOf(aID: integer): integer;
begin
   if fSelectedRowsList.Count=0 then
     Result:= -1
   else
     result:= fSelectedRowsList.IndexOf(aID);
end;

procedure TwDBGrid.SelectedRowsClear();
begin
  fSelectedRowsList.Clear;
  Grid.Repaint;
end;

function TwDBGrid.SelectedRows: ArrayOfInteger;
var
  _GridDataSet: TDataSet;
  i: integer;
begin
  try
    result:= nil;
    if not Assigned(Grid.DataSource) then exit;

    _GridDataSet:= Grid.DataSource.DataSet;

    if fSelectedRowsList.Count <> 0 then
    begin
      SetLength(result,fSelectedRowsList.Count);
      for i := 0 to fSelectedRowsList.Count-1 do
        begin
            result[i]:= fSelectedRowsList.Items[i];
        end;
    end else
    begin
     SetLength(result,1);
     result[0]:= _GridDataSet.FieldByName('ID').AsInteger;
    end;
  except
    on E: Exception do
    begin
      __Log.SaveLogError(E);
      Log('Ошибка [SelectedRows]: "' + E.Message + '"');
      raise;
    end;
  end;
end;

procedure TwDBGrid.onKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
  if ((ssShift in Shift) ) and ((Key = VK_UP) or (Key = VK_DOWN)) then SelectRow(Sender);
end;

procedure TwDBGrid.onKeyUp(Sender: TObject; var Key: Word;  Shift: TShiftState);
begin
    if Key = VK_SPACE then
     begin
      SelectRow(Sender);
      TDBGrid(Sender).DataSource.DataSet.Next;
     end;
end;

procedure TwDBGrid.onMouseUp(Sender: TObject;  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
   if  fmLeft_kCtrl then SelectRow(Sender);
end;

procedure TwDBGrid.onMouseDown(Sender: TObject;  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
   if (ssLeft in Shift) and (ssCtrl in Shift) then fmLeft_kCtrl:= true else fmLeft_kCtrl:= false;
   fShiftState:= Shift;
end;


procedure TwDBGrid.HighLightText(Sender: TObject; const aFieldName: string; Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  i,tmpPos: integer;
  s, e, f2: string;
  r: TRect;
begin
   if Column.FieldName = aFieldName then
   if Assigned(fSearchStrings) then
    begin
         with TDBGrid(Sender).Canvas do
         begin
           r := Rect;
           FillRect(r);

             e := AnsiUpperCase(fSearchStrings[High(fSearchStrings)]);
             f2 := AnsiUpperCase(Column.Field.Text);
             i := pos(e, f2);
             if i <> 0 then
             begin
               s := copy(Column.Field.Text, 1, i - 1);
               Font.Color := clBlack;
               Font.Style := [];
               TextOut(r.Left, r.Top, s);
               r.Left := r.Left + TextWidth(s);
               Brush.Color := $00CCFFFF;
               tmpPos := Pos(e, f2);
               s := Copy(Column.Field.Text, tmpPos, Length(e));
               Font.Color := clMaroon;
               Font.Style := [fsBold];
               TextOut(r.Left, r.Top, s);
               r.Left := r.Left + TextWidth(s);
               Brush.Color := clWhite;
               s := copy(Column.Field.Text, i + length(e), length(f2));
               Font.Color := clBlack;
               Font.Style := [];
               TextOut(r.Left, r.Top, s);
             end else
             begin
               TextOut(r.Left, r.Top, Column.Field.Text);
             end;
         end;
    end;
end;

function TwDBGrid.IsExistsField(aFieldName:string):boolean;
var
  aDataSet: TDataSet;
  i: Integer;
begin
  aDataSet:= fGrid.DataSource.DataSet;
  Result:= false;
  for i:=0 to aDataSet.FieldCount-1 do
  begin
    Result:= aDataSet.Fields[i].FieldName = aFieldName;
    if Result then Break;
  end;

end;

procedure TwDBGrid.onDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  _ColumnText: string;
  _ImageIndex: integer;
  _FieldValue: Double;
begin
 if (gdFocused in State) then // если строка не выделена, то
   begin
     TDBGrid(Sender).Canvas.Brush.Color:= TDBGrid(Sender).FixedHotColor;
     TDBGrid(Sender).Canvas.Font.Color:= clBlack;

     TDBGrid(Sender).Canvas.FillRect(Rect);
     TDBGrid(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
   end;

 if (Column.FieldName = 'ID') then
  begin
    if Assigned(TDBGrid(Sender).TitleImageList) then
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);
        TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

        if SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         begin
                  _ImageIndex:= 0;
                  // А теперь пусть ImageList нарисует ее на канве DBGrid'а
                  TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
         end;
     end else
     begin
        TDBGrid(Sender).Canvas.FillRect(Rect);

        if SelectedRowsListIndexOf(Column.Field.AsInteger)>-1 then
         _ColumnText:='+'
            else
         _ColumnText:=' ';

         TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
     end;
  end;

 if Column.FieldName = 'QUANTITYINPACKING' then
  begin
    with TDBGrid(Sender).Canvas do
    begin
      _FieldValue:= TDBGrid(Sender).DataSource.DataSet.FieldByName('QUANTITYINPACKING').AsFloat;
      if _FieldValue <> 0 then
        begin
          FillRect(Rect);
          if _FieldValue<1 then
            _ColumnText:= '1 к '+FloatToStr(_wRND(1/_FieldValue))
          else
            _ColumnText:= FloatToStr(1*_FieldValue)+' к 1';

          TextOut(Rect.Right - 2 - TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
        end else
        begin
           FillRect(Rect);
           _ColumnText:='';
           TextOut(Rect.Right - 2 - TextWidth(_ColumnText), Rect.Top + 2, _ColumnText);
        end;
    end;
  end;

   if Column.FieldName = 'MTHRESULT' then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      TDBGrid(Sender).Canvas.TextOut(Rect.Right - 2 - TDBGrid(Sender).Canvas.TextWidth(' '), Rect.Top + 2, ' ');

       if TDBGrid(Sender).DataSource.DataSet.FieldByName('MTHRESULT').AsInteger>0 then
        begin
     	    _ImageIndex:= 1;
     	    // А теперь пусть ImageList нарисует ее на канве DBGrid'а
          TDBGrid(Sender).TitleImageList.Draw(TDBGrid(Sender).Canvas,Rect.Left,Rect.Top, _ImageIndex );
        end;
    end;

   HighLightText(Sender,'NAME', Rect,DataCol,Column,State); // подсветить часть текста
   HighLightText(Sender,'LABEL', Rect,DataCol,Column,State); // подсветить часть текста
end;

procedure TwDBGrid.onDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
   if fSelectedRowsList.Count>1 then Grid.DragCursor:=crMultiDrag else Grid.DragCursor:=crDrag;
end;

procedure TwDBGrid.onEndDrag(Sender, Target: TObject; X, Y: Integer);
var
  _ReceiverNodeData, i: Integer;
  _TableName: String;
  _GroupField: string;
begin
   if (Target is TTreeVIew) and (fTree <> nil) then
    begin
      //_DBTree:= __DBTree(FormName,TreeView);
      _TableName:= GetTableName(fGroupField);
      _GroupField:= GetFieldName(fGroupField);
      _ReceiverNodeData:= fTree.ReceiverNodeData;

      try
       if fSelectedRowsList.Count>0 then
        begin
         if MessageDlg('Выделено '+IntToStr(fSelectedRowsList.Count)+' позиций. Переместить их все?',mtConfirmation, mbOKCancel, 0) = mrOK then
          begin
           for i:=0 to fSelectedRowsList.Count-1 do
           begin
             Base.SQLUpdate(_TableName,[_GroupField],[_ReceiverNodeData],'ID'+'='+IntToStr(fSelectedRowsList.Items[i]))
           end;
             fSelectedRowsList.Clear;
          end;
        end else
        begin
           Base.SQLUpdate(_TableName,[_GroupField],[_ReceiverNodeData],'ID'+'='+IntToStr(fIDCurrentRecord))
        end;

      finally
        fTree.FindNodeWithDataInt(TTreeData(fTree.Tree.Selected.Data).Value);
       Fill();
      //  FillGrid([TTreeData(DBTree.Tree.Selected.Data).Value]);
        Grid.DataSource.DataSet.Locate('ID',fIDCurrentRecord,[]);
      end;

    end;
end;

procedure TwDBGrid.onCellClick(Column: TColumn);
begin
  fIDCurrentRecord:=fGrid.DataSource.DataSet.FieldByName('ID').AsInteger;
  fRow:= fGrid.DataSource.DataSet.RecNo;
  fCol:= Column.Field.Index;
  fFieldName:= Column.FieldName;

  if Assigned(onGridCellClick) then onGridCellClick(self);

end;

procedure TwDBGrid.DataSet_onCalcFields(DataSet: TDataSet);
begin
  if Formula = nil then  exit;

      DataSet.FieldByName(Formula.CalculateField).AsFloat := Formula.Calc(DataSet.FieldByName('FORMULA').AsString);
end;

procedure TwDBGrid.Edit_onChange(Sender: TObject);
function PrepareString(aText: string): string;
begin
  Result:= StringReplace(aText, #39, #39+#39,[rfReplaceAll]);
  Result:= UTF8Copy(Result,1,500);
end;

var
  _Edit:TComboBox;
begin
     _Edit:=(Sender as TComboBox);

  fSearchComplete:= false;

  if Length(ReplaceStr(_Edit.Text,' ',''))=0 then fSearchStrings:=nil else
    if Assigned(fSearchSplitStringBtn) and not fSearchSplitStringBtn.Down then
    begin
      SetLength(fSearchStrings,1);
      fSearchStrings[0]:= PrepareString(_Edit.Text);
    end else
      fSearchStrings:=Base.MakeArrayFromString(PrepareString(_Edit.Text)); // указываем строку для поиска


  if SearchPreventiveBtn <> nil then
  begin
     if SearchPreventiveBtn.Down then
     begin
        if  Length(_Edit.Text)>0 then _Edit.Color:=clSkyBlue;
        Fill(); // заполнение грида
     end else
        _Edit.Color:=clMoneyGreen;

  end else
  begin
    if  Length(_Edit.Text)>0 then _Edit.Color:=clSkyBlue;
    Fill(); // заполнение грида
  end;

  if  Length(_Edit.Text)=0 then
  begin
    _Edit.Color:=clDefault;
    Fill(); // заполнение грида
  end;

end;

procedure TwDBGrid.SearchSplitStringBtn_onClick(Sender: TObject);
var
  _Edit:TComboBox;
  _Stat: Boolean;
begin
  _Edit:=SearchEdit;

  _Stat:= SearchPreventiveBtn.Down;
  SearchPreventiveBtn.Down:= true;
    Edit_onChange(_Edit);
  SearchPreventiveBtn.Down:= _Stat;

end;

procedure TwDBGrid.Edit_onExit(Sender: TObject);
var
  _Edit:TComboBox;
begin
     _Edit:=SearchEdit;
  if fSearchComplete then exit;
  if (fSearchPreventiveBtn <> nil) and not fSearchPreventiveBtn.Down and (Length(_Edit.Text)>0) then
  begin
     _Edit.Color:=clSkyBlue;
     Fill(); // заполнение грида

     Edit_SaveInputValues(_Edit);
  end;
end;

procedure TwDBGrid.Edit_SaveInputValues(aEdit: TComboBox);
var
  i: Integer;
begin
   with aEdit do begin
     //Items.Append(Text);
     if Items.Count>cEdSearchDropCountHistory-1 then
        for i:= Items.Count downto cEdSearchDropCountHistory-1 do
          Items.Delete(i);

     Items.Insert(0,Text);
   end;
   aEdit.Items.SaveToFile(fSearchEditHistoryFile);
end;

procedure TwDBGrid.Edit_KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  _Edit: TComboBox;
begin
  _Edit:= TComboBox(Sender);
  if (Key = 114) then
    if _Edit.DroppedDown then
       _Edit.DroppedDown:= false else
       _Edit.DroppedDown:= true;
end;

procedure TwDBGrid.Edit_DropDown(Sender: TObject);
begin
  TComboBox(Sender).Items.LoadFromFile(fSearchEditHistoryFile);
end;

procedure TwDBGrid.Edit_Enter(Sender: TObject);
begin
  TComboBox(Sender).Items.LoadFromFile(fSearchEditHistoryFile);
end;

procedure TwDBGrid.Edit_onKeyPress(Sender: TObject; var Key: char);
var
  _Edit:TComboBox;
begin
     _Edit:=SearchEdit;

  if (Key = #13) and (fSearchPreventiveBtn <> nil) and not fSearchPreventiveBtn.Down and (Length(_Edit.Text)>0) then
  begin
    fSearchComplete:= true;

     _Edit.Color:=clSkyBlue;
     Fill() // заполнение грида
  end;

  if (Key = #13) then
      Edit_SaveInputValues(_Edit);
end;

procedure TwDBGrid.onProgressInit(const aProgressBarName: TProgressBarName; aValue: integer);
begin
  fProgress.InitBar(pbTop,aValue);
end;

procedure TwDBGrid.onProgressUpdate(const aProgressBarName: TProgressBarName; aValue: integer);
begin
  fProgress.SetBar(aProgressBarName,aValue);
end;

procedure TwDBGrid.onStopForce(Sender: TObject);
begin
  fDataExport.Stop;
end;

procedure TwDBGrid.ExportData();
var
  SaveDialog: TSaveDialog;
  _FileFormat: TFileFormat;
begin
  SaveDialog:= TSaveDialog.Create(fGrid);
  SaveDialog.Options:= [ofOverwritePrompt];

try
    SaveDialog.Filter:='OpenDocument (*.ods)|*.ods| Excel (*.xls)|*.xls| Excel (*.xlsx)|*.xlsx|Comma Text (*.csv)|*.csv';
    SaveDialog.FilterIndex:=1;

    SaveDialog.FileName:='Экспорт данных';


  if not SaveDialog.Execute then exit;

    case SaveDialog.FilterIndex of
      1: _FileFormat:= ffODS;
      2: _FileFormat:= ffXLS;
      3: _FileFormat:= ffXLSX;
      4: _FileFormat:= ffCSV;
    end;

  ExportData(SaveDialog.FileName, _FileFormat);
finally
  SaveDialog.Free;
end;

end;

procedure TwDBGrid.ExportData(aFileName:string; aFileFormat: TFileFormat);
begin
  if Assigned(fDataExport) then
  begin
    ShowMessage('Экспорт уже запущен!');
    exit;
  end;
  fResult:= false;

  fProgress:= TProgress.Create(fGrid);
  fProgress.Caption:= 'Экспорт...';
  fProgress.ShowLog:= false;
  fProgress.ShowBottom:= false;
  fProgress.ShowMaxMinButtons:= False;
  fProgress.onStopForce:= @onStopForce;

  screen.Cursor:=crSQLWait;

  fDataExport := TDataExportThread.Create(True);
  fDataExport.fFileName := aFileName;
  fDataExport.fGrid := fGrid;
  fDataExport.fBase:= fBase;
  fDataExport.fSQL:= SQL;

  //fDataExport.onStatusUpdate := @onStatusUpdate;
  fDataExport.onEndThread:= @onEndThread;
  fDataExport.fFileFormat:= aFileFormat;
  //fDataExport.onProgressInit:= @onProgressInit;
  fDataExport.onProgressUpdate:= @onProgressUpdate;
  fDataExport.onExceptionEvent:= @onExceptionEvent;
  fDataExport.onProgressInit:= @onProgressInit;

  fDataExport.start;

  try
    fProgress.ShowModal;
  finally
    screen.Cursor:=crDefault;
   if Assigned(fProgress) then
     fProgress.Free;
  end;
  if fResult then
    if MessageDlg('Открыть полученный файл в программе просмотра?',
        mtConfirmation, mbOKCancel, 0) = mrOK then
        OpenDocument(aFileName);
end;

function TwDBGrid.getBookmark: TBookMark;
begin
  if Assigned(Grid.DataSource) then
    result:= Grid.DataSource.DataSet.Bookmark;
end;

function TwDBGrid.GetFieldValue(aFieldName: string ): TField;
begin
  Result:= Grid.DataSource.DataSet.FieldByName(aFieldName);
end;

procedure TwDBGrid.CopyToClipboard(aFieldsArr: ArrayOfString; aCurrencyFields: ArrayOfString; const aIDField: string);
var
  _Text, _SQL: String;
  i: Integer;
  _SelectedRows: ArrayOfInteger;
  _DataSet: TDataSet;
begin
  if not Assigned(self.Grid.DataSource) then exit;

  _DataSet:= self.Grid.DataSource.DataSet;
  if not Assigned(aFieldsArr) then
  begin
    SetLength(aFieldsArr,self.Grid.Columns.Count);
    for i:=0 to self.Grid.Columns.Count-1 do
    if Length(self.Grid.Columns[i].Title.Caption)>0 then
        aFieldsArr[i]:= self.Grid.Columns[i].FieldName;
  end;

  _SelectedRows:= SelectedRows;

  if Length(_SelectedRows)>1 then
  begin
    _SQL:= Base.WriteWhere(SQL,aIDField+' in ('+Base.MakeStringFromArray(_SelectedRows)+')');
    _DataSet:= Base.SQLReadDS(_SQL,true).DataSet;
    _DataSet.First;
  end;

  _Text:= '';

  for i:=0 to High(_SelectedRows) do
  begin
    if i>0 then _Text:= _Text+LineEnding;
    _Text:= _Text + CopyToClipboardWriteString(_DataSet, aCurrencyFields, aFieldsArr);

    if Length(_SelectedRows)>1 then
       _DataSet.Next;
  end;

  Clipboard.AsText:= _Text;
end;

procedure TwDBGrid.CopyToClipboard(aDataSet:TDataSet; aFieldsArr: ArrayOfString; aCurrencyFields: ArrayOfString);
var
  _Text: String;
  i: Integer;
begin
  if not Assigned(aDataSet) then exit;

  if not Assigned(aFieldsArr) then
  begin
    SetLength(aFieldsArr,aDataSet.Fields.Count);
    for i:=0 to aDataSet.Fields.Count-1 do
      aFieldsArr[i]:= aDataSet.Fields[i].FieldName;
  end;

  _Text:= CopyToClipboardWriteString(aDataSet, aCurrencyFields, aFieldsArr);

  Clipboard.AsText:= _Text;
end;

procedure TwDBGrid.SetColumnCaption(aFieldName, aCaption:string);
var
  i: Integer;
begin
  for i:=0 to fGrid.Columns.Count-1 do
  if fGrid.Columns.Items[i].FieldName = aFieldName then
     begin
       fGrid.Columns.Items[i].Title.Caption:= aCaption;
       Break;
     end;
end;

end.

