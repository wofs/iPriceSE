unit wFuncU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, ExtCtrls, fpspreadsheet, Graphics, SysUtils, StdCtrls, ComCtrls,
  ExtDlgs, DB, Math, DateUtils, LazUTF8, fpsTypes, md5, Process, Dialogs, Menus,
  FileInfo, LazFileUtils, DBGrids, ExtGraphics, TypInfo, wBaseU, wLogU, wTypesU;

type

  { TwData }

  TwData = class
  private
    wColor: TsColor;
    wValue: integer;
    wID: integer;

  public
    constructor Create(_Value: integer);
    constructor Create(_Color: TsColor);
    constructor Create(_ID: integer; _Color: TsColor);
    constructor Create(aID, aValue: integer);

    property Value: integer read wValue write wValue;
    property ID: integer read wID write wID;
    property Color: TsColor read wColor write wColor;
  end;

const

  __SCODPREFIX = '29';
//  префикс для штрих-кода (потом вынести в настройки)

var
  _FileMD5: TMemoryStream;

procedure CatchUnhandledException(Obj: TObject; Addr: Pointer;
  FrameCount: longint; Frames: PPointer);

// работа с ComboBox
procedure cmbxFill(CMBX: TComboBox; DS: TDataSource; _arr: array of variant);
// заполнение Combobox по схеме текст+значение  _arr[0] - поле с текстом, _arr[1] - поле со значением
function cmbxSelectID(cmbx: TCombobox): integer;
// получение выбранного идентификатора
procedure cmbxClearData(cmbx: TCombobox);
// очистка идентификаторов
function cmbxItemIndexByID(cmbx: TCombobox; aID: integer): integer;
// выбор значения по ID


// работа с ListBox
procedure lbxFill(lbx: TListBox; DS: TDataSource; _arr: array of variant);
procedure lbxFill(lbx: TListBox; _arr: array of string);
// заполнение ListBox по схеме текст+значение  _arr[0] - поле с текстом, _arr[1] - поле со значением
function lbxSelectID(lbx: TListBox): integer;
// получение выбранного идентификатора
function lbxSelectIDs(lbx: TListBox): ArrayOfInteger;
// получение выбранного идентификатора
procedure lbxClearData(lbx: TListBox);
// очистка идентификаторов
function lbxItemIndexByID(lbx: TListBox; aID: integer): integer;
// выбор значения по ID

procedure Razdelitel(Sender: TObject; Dec: integer; ent: boolean);
//разделители
function EditValue(_Edit: TEdit): double;
function ReplaceStr(_Str, _SearchStr, _ReplaceStr: string): string;
// функция замены текста
function FilterSimvol(Sender: TObject; var Key: char; minus: boolean): char;
// фильтрация вводимых в поле символов
procedure CalcOpen(edt: TEdit; clc: TCalculatorDialog);
// вызов калькулятора для заполнения Edit
procedure CheckClear(Sender: TObject; Dec: integer; txt: string);
// проверка поля ввода на пустоту

function CalcSumFromPercent(_Value, _Percent: double): double;
// расчет суммы наценки по цене и проценту
function CalcPercentFromSum(_Value, _Sum: double): double;
// расчет процента наценки из суммы
function SeparStrToDouble(_Value: string): double;
function CalcMD5File(_File: string): string;

function DateToUnix(const AValue: TDateTime): int64;
function UnixToDate(const AValue: int64): TDateTime;

function HTMLEntrToUTF8(const S: WideString): WideString;

function VarToStr(Value: variant): string;
function VarToDouble(Value: variant): double;
function VarToInt(Value: variant): integer;

//создание штрих-кода
function CalculateEANCheckSum(Str: string): char;
function GenEAN(pref, Num, suf: string): string;
// ИСПОЛЬЗОВАНИЕ: GenEAN(__SCODPREFIX, '', IntToStr(ID));

function _wKurs(_val: string): double;
function _wSRNDTO(const AValue: double; const Digits: TRoundToRange = -2): double;
function _wR5(r: real): double; // до 0.5
function _wRN(const AValue: double; const Digits: double = 0): double;
function _wRNDTO(const AValue: double; const Digits: TRoundToRange = -2): double;
function _wRND(r: real): integer;

function CalcProbel(s: string; flag: boolean): integer;
// flag = true - считаем слова, иначе - пробелы

function GetMaxFTimeStampPricesArr(aBase: TwBase): ArrayOfDateTime;


function SafePath(const APath: string): string;
function UnsafePath(const ARootPath, ATarget: string): string;

procedure WriteUTF8String(aFileStream: TFileStream; aText: RawByteString);
function DecodeHTMLEntrities(aText: string; const QuotedShield: boolean = False;
  aClearCLCF: boolean = False): string; // decode htmlEnt
function GetLibreOfficeInstallation(): string;
function ConvertFileWithLibreOffice(aFileName: string): string;

function RPosUTF8(c: char; const S: ansistring): SizeInt; overload;

function GetVersion: string;
function CheckVersionFile(aFileName: string): boolean;
function CheckLessVersion(aStringVersion1: string;
  const aStringVersion2: string = ''): boolean;

procedure DBGridClearOrderBy(aDBGrid: TDbGrid);
function BitTest(var aVal; aNumOfbits:integer):integer;
function GetDataString(_ACell: PCell; const aValueType: TValueType = vtDefault): string;
procedure WriteDataToCell(aField: TField; const ARow, ACol: Cardinal; var aWorkSheet: TsWorksheet);

{FPSpreadSheet}
procedure SetColWidth(aWorksheet:TsWorksheet; aWidths: ArrayOfInteger);

procedure WriteValue(aWorksheet:TsWorksheet; aRow, aCol: integer; aField: TField;
  const aFontStyles: TsFontStyles = [];
  aCellColor: TsColor = clDefault;
  aFontColor: TsColor = clBlack;
  aBorders: boolean = true);

procedure WriteValue(aWorksheet:TsWorksheet; aRow, aCol: integer; aValue: variant;
  const aFontStyles: TsFontStyles = [];
  aCellColor: TsColor = clDefault;
  aFontColor: TsColor = clBlack;
  aBorders: boolean = true);

function CompareCellValue(aWorksheet: TsWorksheet; aRow, aCol: integer; aValue: variant; aValueType: TValueType = vtDefault): boolean;

{}
function MinValue(aArray: ArrayOfArrayVariant; aCol: integer): Double;
function MinValue(aDS: TDataSet; aCol: integer): Double;

function FormatCurrValue(aValue: Currency):string;
function GetValue(aEdit: TEdit):Currency;
function GetValue(aEdit: TLabeledEdit):Currency;
procedure FormatValue(aEdit: TEdit);
procedure FormatValue(aEdit: TLabeledEdit);

function getUTFSymbol(str:string;i:integer):string; //Функция достает нужный мне символ по номеру

function GetSpreadSheetFormat(aFileName:string):TsSpreadsheetFormat;

function RGBtoBGR(aRGB:string):TColor;
implementation

{ TwData }

constructor TwData.Create(_Value: integer);
begin
  Value := _Value;
end;

constructor TwData.Create(_Color: TsColor);
begin
  Color := _Color;
end;

constructor TwData.Create(_ID: integer; _Color: TsColor);
begin
  ID := _ID;
  Color := _Color;
end;

constructor TwData.Create(aID, aValue: integer);
begin
  wID:= aID;
  wValue:= aValue;
end;

function RGBtoBGR(aRGB:string):TColor;
begin
  //$F0FBFF
  Result:= StrToInt('$'+Copy(aRGB,6,2)+Copy(aRGB,4,2)+Copy(aRGB,2,2));
end;

procedure CatchUnhandledException(Obj: TObject; Addr: Pointer;
  FrameCount: longint; Frames: PPointer);
var
  Message: string;
  i: longint;
  hstdout: ^Text;
begin
  hstdout := @stdout;
  Writeln(hstdout^, 'An unhandled exception occurred at $',
    HexStr(PtrUInt(Addr), SizeOf(PtrUInt) * 2), ' :');
  if Obj is Exception then
  begin
    Message := Exception(Obj).ClassName + ' : ' + Exception(Obj).Message;
    Writeln(hstdout^, Message);
  end
  else
    Writeln(hstdout^, 'Exception object ', Obj.ClassName, ' is not of class Exception.');
  Writeln(hstdout^, BackTraceStrFunc(Addr));
  if (FrameCount > 0) then
  begin
    for i := 0 to FrameCount - 1 do
      Writeln(hstdout^, BackTraceStrFunc(Frames[i]));
  end;
  Writeln(hstdout^, '');
end;

function GetMaxFTimeStampPricesArr(aBase: TwBase): ArrayOfDateTime;
var
  _arr: ArrayOfArrayVariant;
  i: integer;
begin
  Result := nil;

  _arr := aBase.SQLReadArr('FORMATS', ['FTIMESTAMPLASTIMPORT'], '', '');

  if not Assigned(_arr) then
    exit;

  SeTLength(Result, High(_arr) + 1);

  for i := 0 to High(_arr) do
    Result[i] := _arr[i, 0];

end;


function SafePath(const APath: string): string;
begin
  Result := ExpandFileName(APath);
end;

function UnsafePath(const ARootPath, ATarget: string): string;
  //ARootPath - каталог с программой, документом и т.п., относительно которого нужно создать путь.
  //ATarget - имя файла, которое нужно сохранить в конфигурацию, документ и т.п.
  //Result - собственно результат, который нужно сохранить
begin
  Result := ExtractRelativePath(ARootPath, ATarget);
end;

procedure WriteUTF8String(aFileStream: TFileStream; aText: RawByteString);
begin
  aText := aText + LineEnding;
  SetCodePage(aText, CP_UTF8, True);
  SetCodePage(aText, CP_NONE, False);
  aFileStream.WriteBuffer(aText[1], Length(aText));
end;

function DecodeHTMLEntrities(aText: string; const QuotedShield: boolean;
  aClearCLCF: boolean): string; // decode htmlEnt
const
  TagArr: array[1..6] of string = ('&lt;', '&gt;', '&amp;', '&quot;', '&apos;', '&br;');
  CodeArr: array[1..6] of string = (#60, #62, #38, #34, #39, #10);//< > & " '
var
  i: integer;
  //<       &lt;
  //>       &gt;
  //&       &amp;
  //"       &quot;
  //'       &apos;
begin
  if QuotedShield then
    Result := StringReplace(aText, #39, #39 + #39, [rfReplaceAll, rfIgnoreCase])
  else
    Result := aText;
  if aClearCLCF then
  begin
    Result := StringReplace(aText, #13#10, '', [rfReplaceAll]);
    Result := StringReplace(aText, #10, '', [rfReplaceAll]);
  end;

  for i := 1 to High(TagArr) do
  begin
    Result := StringReplace(Result, TagArr[i], CodeArr[i], [rfReplaceAll, rfIgnoreCase]);
    //&amp;
  end;
end;

function GetLibreOfficeInstallation(): string;
var
  i: integer;
begin
{$IFDEF WINDOWS}

Result := GetEnvironmentVariableUTF8('SYSTEMDRIVE') +'\Program Files\LibreOffice\program\soffice.exe';

{$ENDIF}

end;

function ConvertFileWithLibreOffice(aFileName: string): string;
var
  AProcess: TProcess;
  _TempPath: string;
begin
  try
    AProcess := TProcess.Create(nil);
    //soffice -env:UserInstallation=file://$HOME/.libreoffice-headless --headless --convert-to txt:Text mydoc.odt
    Result := '';
    try
      _TempPath := IncludeTrailingPathDelimiter(PathApplication_Unsafe) + 'tmp';
      _TempPath := IncludeTrailingPathDelimiter(_TempPath) + 'convert';

      if not DirectoryExistsUTF8(_TempPath) then
        ForceDirectoriesUTF8(_TempPath);
      {$IFDEF WINDOWS}
      if not FileExists(PathLibreOffice) then
        Exception.Create(
          'LibreOffice не найден! Конвертация файла невозможна!');
      AProcess.Executable := '"' + PathLibreOffice + '" '
        //+'-env:UserInstallation='+PathApplication_Unsafe+'libreoffice-headless '
        + ' --headless --convert-to xls:"MS Excel 97" --outdir "' + _TempPath +
        '" "' + aFileName + '"';
      wLog('[LibreOfficeConvert]', AProcess.Executable);
      {$ENDIF}
      AProcess.Options := AProcess.Options + [poWaitOnExit, poUsePipes];


      //SetCurrentDir(_TempPath);
      try
        AProcess.Execute;
        _TempPath := IncludeTrailingPathDelimiter(_TempPath);
        Result := _TempPath + StringReplace(ExtractFileName(aFileName),
          ExtractFileExt(aFileName), '', []) + '.xls';
      except
        raise;
      end;

    finally
      //SetCurrentDir(PathApplication_Unsafe);
      AProcess.Free;
    end;
  except
    on E: Exception do
    begin
      __Log.SaveLogError(E);
      wLog('LibreOfficeConvert', 'Ошибка: "' + E.Message + '"');
      wLog('LibreOfficeConvert', 'Сбой сохранения формата!');
      //raise;
    end;
  end;
end;

function _wSRNDTO(const AValue: double; const Digits: TRoundToRange = -2): double;

var
  RV: double;

begin
  RV := IntPower(10, -Digits);
  if AValue < 0 then
    Result := Trunc((AValue * RV) - 0.5) / RV
  else
    Result := Trunc((AValue * RV) + 0.5) / RV;
end;


function _wRND(r: real): integer;
begin
  Result := Trunc(r) + Ord(abs(Frac(r)) > 0.5);
end;

function DateToUnix(const AValue: TDateTime): int64;
begin
  Result := DateTimeToUnix(AValue);
end;

function UnixToDate(const AValue: int64): TDateTime;
begin
  Result := UnixToDateTime(AValue);
end;

function VarToStr(Value: variant): string;
begin
  Result := '';
    case TVarData(Value).VType of
      varnull: Result:='';
      varSmallInt,
      varint64,
      varInteger: Result := IntToStr(Value);
      varSingle,
      varDouble,
      varCurrency: Result := FloatToStr(Value);
      varDate: Result := FormatDateTime('dd.mm.yyyy', Value);
      varBoolean: if Value then
          Result := 'T'
        else
          Result := 'F';
      varString: Result := Value
      else
        WriteStr(Result, TVarData(Value).VType)
    end;
end;

function VarToDouble(Value: variant): double;
begin
  Result := 0;
  case TVarData(Value).VType of
    varnull: Result:=0.0;
    varSmallInt,
    varint64,
    varInteger: Result := Value;
    varSingle,
    varDouble,
    varCurrency: Result := Value;
    varBoolean: if Value then
        Result := 1
      else
        Result := 0;
    varString: TryStrToFloat(Value, Result);
  end;
end;

function VarToInt(Value: variant): integer;
begin
  Result := 0;
  case TVarData(Value).VType of
    varnull: Result:=0;
    varSmallInt,
    varint64,
    varInteger: Result := Value;
    varSingle,
    varDouble,
    varCurrency: Result := Value;
    varBoolean: if Value then
        Result := 1
      else
        Result := 0;
    varString: TryStrToInt(Value, Result);
  end;
end;

//создание штрих-кода
function CalculateEANCheckSum(Str: string): char;
var
  iSumOdd, iSumEven, iSum, iDigit, i, iCheckSum: integer;
begin
  Result := #0;
  iSumOdd := 0;
  iSumEven := 0;
  for i := length(Str) downto 1 do
  begin
    if not TryStrToInt(Str[length(Str) - i + 1], iDigit) then
      exit;

    if (i mod 2 = 0) then // even
      iSumEven := iSumEven + iDigit
    else // odd
      iSumOdd := iSumOdd + iDigit;
  end;
  iSum := (iSumOdd * 3) + iSumEven;
  iCheckSum := (10 - (iSum mod 10));
  if ichecksum > 9 then
    Result := IntToStr(iCheckSum)[2]
  else
    Result := IntToStr(iCheckSum)[1];
end;

function GenEAN(pref, Num, suf: string): string;
var
  len: integer;
  I: integer;
begin
  len := 12;
  len := len - (Length(pref) + Length(Num) + Length(suf));
  if len < 0 then
    raise Exception.Create('Слишком длинный код');
  Result := pref + Num;
  for I := 0 to len - 1 do
    Result := Result + '0';
  Result := Result + suf;
  Result := Result + CalculateEANCheckSum(Result);
end;

function _wKurs(_val: string): double;
begin
  Result := 57.5;
end;

function _wRNDTO(const AValue: double; const Digits: TRoundToRange = -2): double;
var
  RV: double;
  DV: integer;
  _res: double;
begin
  if Digits > 0 then
  begin

    case Digits of
      1: DV := 1;
      2: DV := 10;
      3: DV := 100;
      4: DV := 1000;
      5: DV := 10000;
      6: DV := 100000;
      7: DV := 1000000;
      8: DV := 10000000;
      9: DV := 100000000;
      10: DV := 1000000000;
      else
        DV := 0;
    end;
    _res := ((Trunc(AValue) div DV) + 1) * DV;

  end
  else
  begin
    RV := IntPower(10, -Digits);
    if AValue < 0 then
      _res := Trunc((AValue * RV) - 0.5) / RV
    else
      _res := Trunc((AValue * RV) + 0.5) / RV;
  end;
  Result := _res;
end;

function _wRN(const AValue: double; const Digits: double = 0): double;
var
  _res: double;
begin
  _res := AValue;
  if AValue < 1 then
    _res := AValue;
  if (AValue > 1) and (AValue < 10) and (frac(AValue) > 0) then
    _res := _wRNDTO(AValue * 2, 1) / 2;
  if (AValue > 10) and (AValue < Digits) then
    _res := _wRNDTO(AValue, 0);
  if (AValue > 10) and (AValue > Digits) then
    _res := _wRNDTO(AValue, 2);

  Result := _res;
end;

function _wR5(r: real): double;
begin
  Result := _wRNDTO(r * 2, 1) / 2;
end;


{Новые функции работы с Edit}
  function GetValue(aEdit: TEdit):Currency;
  var
    _Result: String;
    N: Integer;
  begin
    _Result:= aEdit.Text;
    for N:= Length(_Result) downto 1 do
                  if not (_Result[N] in [DefaultFormatSettings.DecimalSeparator,'-','0'..'9']) then Delete(_Result, N, 1);

    TryStrToCurr(_Result,Result);
  end;

  function GetValue(aEdit: TLabeledEdit):Currency;
  var
    _Result: String;
    N: Integer;
  begin
    _Result:= aEdit.Text;
    for N:= Length(_Result) downto 1 do
                  if not (_Result[N] in [DefaultFormatSettings.DecimalSeparator,'-','0'..'9']) then Delete(_Result, N, 1);

    TryStrToCurr(_Result,Result);
  end;

  function FormatCurrValue(aValue: Currency):string;
  var
    _TSeparator: string;
  begin
    _TSeparator:= DefaultFormatSettings.ThousandSeparator;
    if aValue>=0 then
      Result:= FormatCurr('###'+_TSeparator+'###'+_TSeparator+'###'+_TSeparator+'##0.00',aValue) else
      Result:= FormatCurr('###'+_TSeparator+'###'+_TSeparator+'###'+_TSeparator+'##0.00',aValue);
  end;

  procedure FormatValue(aEdit: TEdit);
  var
    _Val: Currency;
  begin
    _Val:= GetValue(aEdit);
    aEdit.Text:= FormatCurrValue(_Val);
    if _Val>=0 then aEdit.Font.Color:= clDefault
    else
      aEdit.Font.Color:= clRed;
  end;

  procedure FormatValue(aEdit: TLabeledEdit);
  var
    _Val: Currency;
  begin
    _Val:= GetValue(aEdit);
    aEdit.Text:= FormatCurrValue(_Val);
    if _Val>=0 then aEdit.Font.Color:= clDefault
    else
      aEdit.Font.Color:= clRed;
  end;
{End Новые функции работы с Edit}

procedure lbxFill(lbx: TListBox; DS: TDataSource; _arr: array of variant);
var
  _Field0: String;
begin
  if DS.DataSet.Active then
    DS.DataSet.Close;
  DS.DataSet.Open;

  lbxClearData(lbx);

  DS.DataSet.First;
  while not DS.DataSet.EOF do
  begin
    if Length(_arr)>2 then
    begin
      _Field0:= DS.DataSet.FieldByName(_arr[0]).AsString;
      UTF8Delete(_Field0,100, Length(_Field0)-100);
      lbx.Items.AddObject(_Field0+
         ' | '+DS.DataSet.FieldByName(_arr[1]).Value, TwData.Create(DS.DataSet.FieldByName(_arr[2]).AsInteger))
    end
    else
      lbx.Items.AddObject(DS.DataSet.FieldByName(_arr[0]).Value, TwData.Create(DS.DataSet.FieldByName(_arr[1]).AsInteger));

    DS.DataSet.Next;
  end;

end;

procedure lbxFill(lbx: TListBox; _arr: array of string);
var
  i,ind : integer;
begin
  lbxClearData(lbx);

  for i:=0 to Length(_arr)-1 do
  begin
    ind:= i+1;
    lbx.Items.AddObject(_arr[i], TwData.Create(ind));
  end;
end;

function lbxSelectID(lbx: TListBox): integer;
begin
  Result := TwData(lbx.Items.Objects[lbx.ItemIndex]).Value;
end;

function lbxSelectIDs(lbx: TListBox): ArrayOfInteger;
var
  i, iResult: Integer;
begin
  SetLength(Result,lbx.Items.Count);
  iResult:= 0;
  for i:=0 to lbx.Items.Count-1 do
    if lbx.Selected[i] then
    begin
      Result[iResult]:= TwData(lbx.Items.Objects[i]).Value;
      Inc(iResult);
    end;
  SetLength(Result,iResult);
end;

procedure lbxClearData(lbx: TListBox);
var
  i: integer;
begin
  for i := 0 to lbx.Items.Count - 1 do
    TwData(lbx.Items.Objects[i]).Free;

  lbx.Items.Clear;
end;

function lbxItemIndexByID(lbx: TListBox; aID: integer): integer;
begin
  Result := 0;
  with lbx.Items do
  begin
    while (Result < Count) and (TwData(Objects[Result]).Value <> aID) do
      Result := Result + 1;
    if Result = Count then
      Result := -1;
  end;
end;

procedure Razdelitel(Sender: TObject; Dec: integer; ent: boolean);
//разделители
begin
  with Sender as TEdit do
  begin
    if ent = True then
      Text := ReplaceStr(Text, ThousandSeparator, '')
    else
      Text := FloatToStrF(StrToFloat(Text), ffNumber, 4, Dec);
  end;

end;

function EditValue(_Edit: TEdit): double;
var
  _Text: string;
begin
  _Text := _Edit.Text;
  _Text := ReplaceStr(_Text, ThousandSeparator, '');
  Result := StrToFloat(_Text);
end;

function ReplaceStr(_Str, _SearchStr, _ReplaceStr: string): string;
begin
  //while Pos(SearchStr, Str) <> 0 do
  //begin
  //  Insert(ReplaceStr, Str, Pos(SearchStr, Str));
  //  Delete(Str, Pos(SearchStr, Str), Length(SearchStr));
  //end;

  Result := UTF8StringReplace(_Str, _SearchStr, _ReplaceStr, [rfReplaceAll, rfIgnoreCase]);
end;

function FilterSimvol(Sender: TObject; var Key: char; minus: boolean): char;
  //фильтр лишних символов
var //цифровая маска
  vrPos, vrLength, vrSelStart: byte;
const
  I: byte = 1;
  //I+1 = количество знаков после запятой (в данном случае - 2 знака)
begin
  with Sender as TEdit do
  begin
    vrLength := Length(Text); //определяем длину текста
    vrPos := Pos(DecimalSeparator, Text);
    //проверяем наличие разделителя (запятой/точки)
    vrSelStart := SelStart; //определяем положение курсора
  end;
  if minus = True then
  begin
    case Key of

      '0'..'9', '-':
      begin
        //проверяем положение курсора и количество знаков после запятой
        if (vrPos > 0) and (vrLength - vrPos > I) and (vrSelStart >= vrPos) then
          Key := #0; //"погасить" клавишу
      end;
      ',', '.':
      begin
        //если запятая уже есть или запятую пытаются поставить перед
        //числом или никаких цифр в поле ввода еще нет
        if (vrPos > 0) or (vrSelStart = 0) or (vrLength = 0) then
          Key := #0 //"погасить" клавишу
        else
          //Key := #44; //всегда заменять точку на запятую
          Key := DecimalSeparator;
      end;
      #8: ; //позволить удаление знаков клавишей 'Back Space'
      else
        Key := #0; //"погасить" все остальные клавиши
    end;
  end
  else
  begin
    case Key of

      '0'..'9':
      begin
        //проверяем положение курсора и количество знаков после запятой
        if (vrPos > 0) and (vrLength - vrPos > I) and (vrSelStart >= vrPos) then
          Key := #0; //"погасить" клавишу
      end;
      ',', '.':
      begin
        //если запятая уже есть или запятую пытаются поставить перед
        //числом или никаких цифр в поле ввода еще нет
        if (vrPos > 0) or (vrSelStart = 0) or (vrLength = 0) then
          Key := #0 //"погасить" клавишу
        else
          //Key := #44; //всегда заменять точку на запятую
          Key := DecimalSeparator;
      end;
      #8: ; //позволить удаление знаков клавишей 'Back Space'
      else
        Key := #0; //"погасить" все остальные клавиши
    end;
  end;
  Result := Key;

end;

procedure CalcOpen(edt: TEdit; clc: TCalculatorDialog);
begin
  with edt do
  begin
    with clc do
    begin
      Text := ReplaceStr(Text, ThousandSeparator, '');
      Value := StrToFloat(Text);
      if Execute then
      begin
        Text := FloatToStrF(Value, ffNumber, 4, 2);
      end
      else
        Text := FloatToStrF(Value, ffNumber, 4, 2);
    end;
  end;
end;

procedure CheckClear(Sender: TObject; Dec: integer; txt: string);  //если пусто
begin
  with Sender as TEdit do
  begin
    if Text = '' then
    begin
      Text := '0';
      Razdelitel(Sender, Dec, False);
    end
    else
    begin
      if Text = '-' then
      begin
        Text := '-1';
        Razdelitel(Sender, Dec, False);
      end
      else
        exit;
    end;

  end;
end;

procedure cmbxFill(CMBX: TComboBox; DS: TDataSource; _arr: array of variant);
// заполнение Combobox по схеме текст+значение
var
  _Field0: String;
begin
  if DS.DataSet.Active then
    DS.DataSet.Close;
  DS.DataSet.Open;

  cmbxClearData(cmbx);

  DS.DataSet.First;
  while not DS.DataSet.EOF do
  begin
    if Length(_arr)>2 then
    begin
      _Field0:= DS.DataSet.FieldByName(_arr[0]).AsString;
      UTF8Delete(_Field0,100, Length(_Field0)-100);
      cmbx.Items.AddObject(_Field0+
         ' | '+DS.DataSet.FieldByName(_arr[1]).Value, TwData.Create(DS.DataSet.FieldByName(_arr[2]).AsInteger))
    end
    else
      cmbx.Items.AddObject(DS.DataSet.FieldByName(_arr[0]).Value, TwData.Create(DS.DataSet.FieldByName(_arr[1]).AsInteger));

    DS.DataSet.Next;
  end;
end;

function cmbxSelectID(cmbx: TCombobox): integer;
begin
  Result := TwData(cmbx.Items.Objects[cmbx.ItemIndex]).Value;
end;

function cmbxItemIndexByID(cmbx: TCombobox; aID: integer): integer;
  // выбор значения по ID
begin
  Result := 0;
  with cmbx.Items do
  begin
    while (Result < Count) and (TwData(Objects[Result]).Value <> aID) do
      Result := Result + 1;
    if Result = Count then
      Result := -1;
  end;
end;

procedure cmbxClearData(cmbx: TCombobox);
var
  i: integer;
begin
  for i := 0 to cmbx.Items.Count - 1 do
    TwData(cmbx.Items.Objects[i]).Free;

  cmbx.Items.Clear;
end;

//function Selectcmbx (cmbx:TCombobox; IDOdjects:integer):integer;
//var
//  ic:integer;
//begin

//with cmbx do begin
//    for ic:=0 to Items.Count-1 do begin
//          if integer(Items.Objects[ic])=IDOdjects
//      then begin
//           result:=ic;  exit;
//      end;
//    end;
//  end;
//end;

// преобразовывает HTMLEntries в тексте (&#1044;) в символы UTF8
// чуть правленная процедура из http://www.esperanto.mv.ru/UniRed/RUS/index.html (файл uniEStr.pas)
// и еще здесь http://www.freepascal.ru/forum/viewtopic.php?f=5&t=25591&p=127283#p127283
function HTMLEntrToUTF8(const S: WideString): WideString;
var
  W: WChar;
  i, j, n: integer;
  _PosSemicolon, _CodeLength: integer;
begin
  SetLength(Result, length(S));
  i := 1;
  j := 1;
  while i <= length(S) do
  begin
    W := WChar(S[i]);
    if (Copy(S, i, 2) = '&#') then
    begin
      i := i + 2;
      //detect code Length
      _PosSemicolon := UTF8Pos(';', S, i);
      _CodeLength := _PosSemicolon - i;

      if (_CodeLength >= 1) and (_CodeLength <= 4) then
      begin
        if TryStrToInt(Copy(S, i, _CodeLength), n)
        // can be mixed text
        then
          W := WChar(n)
        else
          W := '?'; // if no...

        i := i + _CodeLength + 1;
      end;

    end
    else
      Inc(i);
    Result[j] := W;
    Inc(j);
  end;
  SetLength(Result, j - 1);
end;

function CalcMD5File(_File: string): string;
var
  _md5Stream: TMD5Digest;
begin
  Result := '';
  if FileExists(_File) then
  begin
    _FileMD5 := TMemoryStream.Create;

    try
      _FileMD5.LoadFromFile(_File);
      _FileMD5.Position := 0;
      _md5Stream := MD5Buffer(_FileMD5.Memory^, _FileMD5.Size);
      Result := MD5Print(_md5Stream);
      _FileMD5.Free;
    except
      _FileMD5.Free;
    end;
  end;

end;

function CalcSumFromPercent(_Value, _Percent: double): double;
begin
  Result := _Value - (_Value * (1 - _Percent / 100));
end;

function CalcPercentFromSum(_Value, _Sum: double): double;
begin
  if _Value > 0 then
    Result := _Sum / _Value * 100
  else
    Result := 0;
end;

function SeparStrToDouble(_Value: string): double;
begin
  Result := StrToFloat(ReplaceStr(_Value, ThousandSeparator, ''));
end;

function CalcProbel(s: string; flag: boolean): integer;
  // flag = true - считаем слова, иначе - пробелы
var
  i, l, n: integer;
  // i - счётчик цикла
  // l - длина текущего слова
  // n - количество слов
begin
  if flag then
  begin //считаем слова
    try
      l := 0;
      n := 0;
      // проверяю есть ли пробел в конце.. если его там нет, то добавляю...
      if s[Length(s)] <> ' ' then
        s := s + ' ';
      // а вот, собственно, цикл, который проверяет кол-во слов... тут учтено, что могут стоять лишние пробелы
      for i := 1 to Length(s) do
        if s[i] <> ' ' then
          l := l + 1
        else
        if l <> 0 then
        begin
          l := 0;
          n := n + 1;
        end;
      // возвращаю результат
      Result := n;
    except
      Result := 0
    end;
  end
  else
  begin //считаем пробелы
    try
      n := 0;
      for i := 1 to length(s) do
        if s[i] = ' ' then
          Inc(n);
      // возвращаю результат
      Result := n;
    except
      Result := 1
    end;
  end;
end;

function RPosUTF8(c: char; const S: ansistring): SizeInt;

var
  I: SizeInt;
  p, p2: PChar;

begin
  I := UTF8Length(S);
  if I <> 0 then
  begin
    p := @s[i];
    p2 := @s[1];
    while (p2 <= p) and (p^ <> c) do
      Dec(p);
    i := p - p2 + 1;
  end;
  RPosUTF8 := i;
end;

function GetVersion: string;
var
  version: string;
  Info: TVersionInfo;
begin
  Info := TVersionInfo.Create;
  Info.Load(HINSTANCE);
  //[0] = Major version, [1] = Minor ver, [2] = Revision, [3] = Build Number
  version := IntToStr(Info.FixedInfo.FileVersion[0]) + '.' + IntToStr(
    Info.FixedInfo.FileVersion[1]) + '.' + IntToStr(Info.FixedInfo.FileVersion[2]) +
    '.' + IntToStr(Info.FixedInfo.FileVersion[3]);
  Result := version;
  Info.Free;
end;

function CheckVersionFile(aFileName: string): boolean;
  // проверяет версию, записанную в файле с текущей версии программы. Если версия в файле меньше текущей, то true.
var
  _File: TStringList;
  _Version: string;
begin
  _File := TStringList.Create;
  _File.LoadFromFile(aFileName);

  try
    _Version := Trim(_File.Strings[0]);
  finally
    FreeAndNil(_File);
  end;

  Result := CheckLessVersion(_Version, GetVersion);
end;

function CheckLessVersion(aStringVersion1: string;
  const aStringVersion2: string = ''): boolean;
var
  _Ver, _Ver2: TStringList;
  i: integer;
begin
  Result := False;
  _Ver := TStringList.Create;
  _Ver.Delimiter := '.';
  _Ver.DelimitedText := aStringVersion1;

  _Ver2 := TStringList.Create;
  _Ver2.Delimiter := '.';
  _Ver2.DelimitedText := aStringVersion2;
  try
    for i := 0 to _Ver.Count - 1 do
    begin
      if StrToInt(_Ver[i]) < StrToInt(_Ver2[i]) then
      begin
        Result := True;
        exit;
      end
      else
      if StrToInt(_Ver[i]) > StrToInt(_Ver2[i]) then
      begin
        Result := False;
        exit;
      end;

    end;
  finally
    _Ver.Free;
    _Ver2.Free;
  end;

end;

procedure DBGridClearOrderBy(aDBGrid: TDbGrid);
var
  i: integer;
begin
  for i := 0 to aDBGrid.Columns.Count - 1 do
  begin
    aDBGrid.Columns[i].Tag := 0;
    aDBGrid.Columns[i].Title.ImageIndex := -1;
  end;
end;

function BitTest(var aVal; aNumOfbits:integer):integer;
var
  s:set of byte absolute aVal;
i:integer;
begin
  Result:=0;
  for i:=0 to aNumOfBits-1 do
    if i in s then inc(Result);
end;

function GetDataString(_ACell: PCell; const aValueType: TValueType = vtDefault): string;
var
  _ValTmp: Double;
  N: Integer;
begin
  Result:='';
  if Assigned(_Acell) then
  begin
    case aValueType of
      vtDefault:
            begin
              case _ACell^.ContentType of
                cctNumber: Result := FloatToStr(_ACell^.NumberValue);
                cctDateTime: Result := DateToStr(_ACell^.DateTimeValue);
                else
                  Result := ReplaceStr(_ACell^.UTF8StringValue, #39, #39+#39);
              end
            end;
      vtNumber:
            begin
              Result := ReplaceStr(_ACell^.UTF8StringValue, ',',DefaultFormatSettings.DecimalSeparator);
              Result := ReplaceStr(Result, '.',DefaultFormatSettings.DecimalSeparator);

              for N:= Length(Result) downto 1 do
              if not (Result[N] in [DefaultFormatSettings.DecimalSeparator, '0'..'9']) then Delete(Result, N, 1);

              TryStrToFloat(Result,_ValTmp);
              Result:= FloatToStr(_ValTmp);
            end;
      vtString:
            begin
                case _ACell^.ContentType of
                  //WorkSheet.ReadCellFormat
                  cctNumber: Result := FloatToStr(_ACell^.Numbervalue);
                  cctUTF8String: Result :=  UTF8Trim(ReplaceStr(_ACell^.UTF8StringValue, #39, ''));
                end;
            end;
    end;
  end
  else
  Result:='';
end;

procedure WriteDataToCell(aField: TField; const ARow, ACol: Cardinal; var aWorkSheet: TsWorksheet);
var
  _Text: String;
begin
  _Text:='';
  case aField.DataType of
    ftSmallInt,
    ftLargeint,
    ftInteger,
    ftFloat,
    ftCurrency,
    ftBCD: aWorkSheet.WriteNumber(aRow,ACol,aField.AsFloat);
    ftDate,
    ftTime,
    ftDateTime: aWorkSheet.WriteDateTime(aRow,ACol,aField.AsDateTime);
    ftBoolean: if aField.AsBoolean then
        aWorkSheet.WriteText(aRow,ACol,'T')
      else
        aWorkSheet.WriteText(aRow,ACol,'F');
    ftString,
    ftWideString:
    begin
      if aField.AsString<>null then
        aWorkSheet.WriteText(aRow,ACol,aField.AsString)
      else
        aWorkSheet.WriteText(aRow,ACol,'');
    end
    else
      begin
        WriteStr(_Text,aField.DataType);
        aWorkSheet.WriteText(aRow,ACol,_Text);
      end;

      //Result := GetEnumName(aField.DataType);
  end;

end;

procedure SetColWidth(aWorksheet: TsWorksheet; aWidths: ArrayOfInteger);
var
  i: Integer;
begin
  for i:=0 to High(aWidths) do
      aWorksheet.WriteColWidth(i,aWidths[i],suChars,cwtCustom);

end;

procedure WriteValue(aWorksheet: TsWorksheet; aRow, aCol: integer; aField: TField; const aFontStyles: TsFontStyles; aCellColor: TsColor; aFontColor: TsColor;
  aBorders: boolean);
var
  _S: string;
begin

  case aField.DataType of
    ftSmallInt,
    ftLargeint,
    ftInteger          : aWorksheet.WriteNumber(aRow, aCol, aField.AsInteger, nfGeneral);
    ftFloat,
    ftBCD,
    ftCurrency         : aWorksheet.WriteNumber(aRow, aCol, aField.AsFloat, nfFixedTh);
    ftDateTime         : aWorksheet.WriteDateTime(aRow, aCol,aField.AsDateTime, nfShortDateTime);
    ftBoolean          : aWorksheet.WriteBoolValue(aRow, aCol,aField.AsBoolean);
    ftString,
    ftWideString       : aWorksheet.WriteText(aRow, aCol,aField.AsString)
    else
      begin
        WriteStr(_S, aField.DataType);
        aWorksheet.WriteText(aRow,aCol,_S);
      end;
  end;

  if aBorders then
    aWorksheet.WriteBorders(aRow,aCol,[cbNorth, cbWest, cbEast, cbSouth]);

  aWorksheet.WriteFontStyle(aRow,aCol,aFontStyles);
  aWorksheet.WriteFontColor(aRow,aCol,aFontColor);
  aWorksheet.WriteBackgroundColor(aRow,aCol,aCellColor);

end;

procedure WriteValue(aWorksheet: TsWorksheet; aRow, aCol: integer; aValue: variant; const aFontStyles: TsFontStyles; aCellColor: TsColor; aFontColor: TsColor;
  aBorders: boolean);
var
  _S: string;
begin

  case TVarData(aValue).VType of
    varSmallInt,
    varInteger,
    varint64     : aWorksheet.WriteNumber(aRow, aCol, aValue, nfGeneral);
    varSingle,
    varDouble,
    varCurrency  : aWorksheet.WriteNumber(aRow, aCol, aValue, nfFixedTh);
    varDate      : aWorksheet.WriteDateTime(aRow, aCol, aValue, nfShortDateTime);
    varBoolean   : aWorksheet.WriteBoolValue(aRow, aCol, aValue);
    varString    : aWorksheet.WriteText(aRow, aCol, aValue);
    else
      begin
        WriteStr(_S, TVarData(aValue).VType);
        aWorksheet.WriteText(aRow,aCol,_S);
      end;
  end;

  if aBorders then
    aWorksheet.WriteBorders(aRow,aCol,[cbNorth, cbWest, cbEast, cbSouth]);

  aWorksheet.WriteFontStyle(aRow,aCol,aFontStyles);
  aWorksheet.WriteFontColor(aRow,aCol,aFontColor);
  aWorksheet.WriteBackgroundColor(aRow,aCol,aCellColor);

end;

function GetSpreadSheetFormat(aFileName:string):TsSpreadsheetFormat;
begin
case UTF8LowerCase(ExtractFileExt(aFileName)) of
    '.xls': Result:= sfExcel8;
    '.xlsx': Result:= sfOOXML;
    '.ods': Result:= sfOpenDocument;
  end;
end;

function CompareCellValue(aWorksheet: TsWorksheet; aRow, aCol: integer; aValue: variant; aValueType: TValueType = vtDefault): boolean;
begin
  Result:=UTF8CompareStr(GetDataString(aWorksheet.GetCell(aRow, aCol), aValueType), VarToStr(aValue)) = 0;
end;

function MinValue(aArray: ArrayOfArrayVariant; aCol: integer): Double;
var
  i: Integer;
begin
  Result:= 0.0;
  Result:= aArray[0,aCol];

  for i:= 1 to High(aArray) do
      if aArray[i,aCol]< Result then
         Result:= aArray[i,aCol];
end;

function MinValue(aDS: TDataSet; aCol: integer): Double;
var
  i: Integer;
begin
  Result:= 0.0;

  Result:= aDS.Fields[aCol].AsFloat;

  for i:= 1 to aDS.RecordCount-1 do
      if aDS.Fields[aCol].AsFloat< Result then
         Result:= aDS.Fields[aCol].AsFloat;

end;

function getUTFSymbol(str:string;i:integer):string; //Функция достает нужный мне символ по номеру
var rez:string;
begin
rez:= UTF8Copy(str,i,1);
Result:=rez;
end;

end.
