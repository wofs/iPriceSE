unit wFormulaU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}
{$DEFINE WITHBASE}

interface

uses
  Classes, fpexprpars, LazUTF8, LCLType, SysUtils, Math, dateutils;

type

  TCharVariant = (cvVAR, cvFUNC, cvCONST, cvKURS, cvNONE);
  ArrayOfDouble = array of double;
  ArrayOfCurrency = array of Currency;
  { TFormula }

  TFormula = class(TFPExpressionParser)
  private
    fCalculateField: string;
    fCurrencyArray: ArrayOfCurrency;
    fVarArray: ArrayOfDouble;
    procedure EX_ABS(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_BINT(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_BRNDTO(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_DATE(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_INRANGE(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_KURS(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_MRNDUP(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_NOW(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_TIME(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    function fBANK_ROUND(aValue: double; aDIG: integer): double;
    function fBINT(aValue: double): double;
    function fBRNDTO(aValue, aDIG: double): double;
    procedure fFunctionTemplate(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure fIIF(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_MRNDTO(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_MRN(var Result: TFPExpressionResult; const Args: TExprParameterArray);
    procedure EX_MINT(var Result: TFPExpressionResult; const Args: TExprParameterArray);


    function FloatStrToBase(aValue: string): string;
    function FloatToBase(aValue: double): string;
    function fMINT(aValue: double): double;
    function fMRN(aValue: double; aLIMIT: integer): double;
    function fMRNDTO(aValue, aDIG: double): double;
    function fMRNDUP(aValue: Double; aDIG: double): Double;
    function FuncID(aName: string): integer;
    procedure GetVariants(aName: string; out i: integer; out aCharVariant: TCharVariant);
    function KursID(aName: string): integer;
    function ReadArgs(aArgs: TFPExpressionResult): variant;
    function ReplaceText(const S, OldPattern, NewPattern: string; aStart: integer): string;
    function ReplaceVar(const S, OldPattern: string; var NewPattern: string; aStart: integer): string;
    function VarID(aName: string): integer;
    function GetPosSemicolon(aText: string; aStart: integer): integer;

  public
    constructor Create(AOwner: TComponent); override;

    function Prepare(S: string; aCheck: boolean = True): string;
    function Calc(aFormula: string{$IFDEF WITHBASE}; const aBase: TObject=nil {$ENDIF}): Double;


    property VarArray: ArrayOfDouble read fVarArray write fVarArray;
    property CalculateField: string read fCalculateField write fCalculateField;
    property CurrencyArray:ArrayOfCurrency read fCurrencyArray write fCurrencyArray;

  end;


implementation

  {$IFDEF WITHBASE}
uses
  wBaseU;
  {$ENDIF}

const

  VarToFiels: array[0..32, 0..2] of string = (
    ('P1',  'CATALOG.PRICE', ''),
    ('P',  'PRICEPL', ''),
    ('P2', 'PRICEPL2', ''),
    ('P3', 'PRICEPL3', ''),
    ('P4', 'PRICEPL4', ''),
    ('P5', 'PRICEPL5', ''),
    ('P6', 'PRICEPL6', ''),
    ('P7', 'PRICEPL7', ''),
    ('P8', 'PRICEPL8', ''),
    ('P9', 'PRICEPL9', ''),
    ('P10', 'PRICEPL10', ''),
    ('P0', 'PLOUR.PRICECALC', ''),
    ('P02', 'PLOUR.PRICECALC2', ''),
    ('P03', 'PLOUR.PRICECALC3', ''),
    ('P04', 'PLOUR.PRICECALC4', ''),
    ('P05', 'PLOUR.PRICECALC5', ''),
    ('P06', 'PLOUR.PRICECALC6', ''),
    ('P07', 'PLOUR.PRICECALC7', ''),
    ('P08', 'PLOUR.PRICECALC8', ''),
    ('P09', 'PLOUR.PRICECALC9', ''),
    ('P010', 'PLOUR.PRICECALC10', ''),
    ('S',  '(PLOUR.STOCK+PLOUR.STOCK2+PLOUR.STOCK3+PLOUR.STOCK4+PLOUR.STOCK5)', ''),
    ('S1', 'PLOUR.STOCK', ''),
    ('S2', 'PLOUR.STOCK2', ''),
    ('S3', 'PLOUR.STOCK3', ''),
    ('S4', 'PLOUR.STOCK4', ''),
    ('S5', 'PLOUR.STOCK5', ''),
    ('N',  'P*(1+N/100)-P', 'PN'),
    ('M',  'M', 'PM'),
    ('D',  'D', 'PD'),
    ('C',  'C', 'PC'),
    ('K',  'K', 'PK'),
    ('PDATEDIF',  'ROUND(CURRENT_TIMESTAMP-PDATE)', '')
    );

  VarToFunctions: array[0..14, 0..3] of string = (
    ('CHOOSE', 'IIF', 'F', 'BFF'), // дублер IF для Fora
    ('ROUND', 'EX_MRNDTO', 'F', 'FF'),// дублер RNDTO для Fora

    ('IF', 'IIF', 'F', 'BFF'),
    ('RN', 'EX_MRN', 'F', 'FI'),
    ('INT', 'EX_MINT', 'I', 'F'),
    ('RNDTO', 'EX_MRNDTO', 'F', 'FF'),
    ('RNDUP', 'EX_MRNDUP', 'F', 'FF'),
    ('BINT', 'EX_BINT', 'I', 'F'),
    ('BRNDTO', 'EX_BRNDTO', 'F', 'FF'),
    ('ABS', 'EX_ABS', 'F', 'F'),
    ('INRANGE', 'EX_INRANGE', 'B', 'FII'),
    ('NOW', 'EX_NOW', 'D', 'I'),
    ('DATE', 'EX_DATE', 'F', 'I'),
    ('TIME', 'EX_TIME', 'F', 'I'),
    ('KURS', '', 'F', 'S')
    );

  VarToKurs: array[0..4] of string =(
  ('RUR'),
  ('USD'),
  ('EUR'),
  ('KZT'),
  ('UAH')
  );

  VarValuesArr = ['A'..'Z', 'a'..'z', #39];
  VarConstArr = ['0'..'9'];
  VarPrepareCAST ='CAST(%s AS NUMERIC(15,4))';
{ TFormula }

function TFormula.VarID(aName: string): integer;
begin
  for Result := 0 to High(VarToFiels) do
    if CompareText(aName, VarToFiels[Result, 0]) = 0 then
      Exit;
  Result := -1;
end;

function TFormula.KursID(aName: string): integer;
begin
  aName:= StringReplace(aName,' ','',[rfReplaceAll,rfIgnoreCase]);
  aName:= StringReplace(aName,#39,'',[rfReplaceAll,rfIgnoreCase]);
  //aName:= StringReplace(aName,'KURS(','',[rfReplaceAll,rfIgnoreCase]);
  //aName:= StringReplace(aName,')','',[rfReplaceAll,rfIgnoreCase]);

  for Result := 0 to High(VarToKurs) do
    if CompareText(aName, VarToKurs[Result]) = 0 then
      Exit;
  Result := -1;
end;

function TFormula.FuncID(aName: string): integer;
begin
  for Result := 0 to High(VarToFunctions) do
    if CompareText(aName, VarToFunctions[Result, 0]) = 0 then
      Exit;
  Result := -1;
end;

procedure TFormula.GetVariants(aName: string; out i: integer;
  out aCharVariant: TCharVariant);
begin
  i := VarID(aName);
  if i>-1 then
  begin
    aCharVariant := cvVAR;
    exit;
  end;

  i := FuncID(aName);
  if i>-1 then
  begin
    aCharVariant := cvFUNC;
    exit;
  end;

  i := KursID(aName);
  if i>-1 then
  begin
    aCharVariant := cvKURS;
    exit;
  end;

  if aName[1] in VarConstArr then
  begin
    i:=-2;
    aCharVariant := cvCONST;
    exit;
  end;

  i := -1;
  aCharVariant := cvNONE;
end;

function TFormula.GetPosSemicolon(aText: string; aStart: integer): integer;
const
  CharDelimiter = [' ', ',', '(', ')', '+', '-', '*', '/', '=', '<', '>'];
begin
  for Result := aStart+1 to UTF8Length(aText) do
    if (aText[Result] in CharDelimiter) then
      exit;

  Result := -1;
end;

procedure TFormula.fFunctionTemplate(var Result: TFPExpressionResult;
  const Args: TExprParameterArray);
begin
  // здесь можно назначить вызов функции из БД
  Result.ResFloat := 0; // заглушка
end;

function ARound(Number: Double; Prec: Byte=0): Double; //правильное арифметическое округление
var IntNumber, Stepen: LongInt;
begin
  Stepen:=Trunc(Power(10,Prec+1));
  IntNumber:=Trunc(Number*Stepen);
  if ((IntNumber mod 10)>=5) then
    Number:=(IntNumber div 10)+1
  else
    Number:=IntNumber div 10;
  Result:=Number*10/Stepen;
end;

function TFormula.fMRNDTO(aValue, aDIG: double): double;
var
  VORDER_ROUND: double;
  VAL, RND_DIGIT: Double;
begin
  if (aDIG = 0) or (aDIG < 0) then
    raise Exception.Create('[Ошибка!] RNDTO: Допустимые значения точности: 0.0001,0.001,0.01,0.1,1,10,100,1000 и тд.');

  VORDER_ROUND:= 0.0;
  VAL:= aValue;
  RND_DIGIT:= aDIG;

  VORDER_ROUND:= RND_DIGIT*100;

  Result:= ARound(VAL*100.0/VORDER_ROUND)*VORDER_ROUND/100.0;

end;

function TFormula.fBRNDTO(aValue, aDIG: double): double;
var
  VORDER_ROUND: double;
  VAL, RND_DIGIT: Double;
begin
  if (aDIG = 0) or (aDIG < 0) then
    raise Exception.Create('[Ошибка!] RNDTO: Допустимые значения точности: 0.0001,0.001,0.01,0.1,1,10,100,1000 и тд.');

  VORDER_ROUND:= 0.0;
  VAL:= aValue;
  RND_DIGIT:= aDIG;

  VORDER_ROUND:= RND_DIGIT*100;

  Result:= fBINT(VAL*100.0/VORDER_ROUND)*VORDER_ROUND/100.0;
end;

function TFormula.fMRNDUP(aValue : Double; aDIG:double) : Double;
var
  _Delitel: Double;
Begin
  if (aDIG = 0) or (aDIG < 0) then
    raise Exception.Create('[Ошибка!] RNDUP: Допустимые значения точности: 0.0001 .. 1 .. 1000000 (н-р 0.05 | 0.01 | 0.2 | 10 | 50 ...)');

  _Delitel:= 1/aDIG;
  result:=CEIL(aValue*_Delitel)/_Delitel;
end;

function TFormula.fMRN(aValue: double; aLIMIT: integer): double;
begin
   if (aValue<1) or (aValue=aLIMIT) then
   begin
     Result:= aValue;
     exit;
   end;

   if aValue<10 then
   begin
     Result:= fMRNDUP(aValue,0.5);
     exit;
   end;

   if aValue<aLIMIT then
   begin
     Result:= fMRNDUP(aValue,1);
     exit;
   end;

   if aValue>aLIMIT then
   begin
     Result:= fMRNDUP(aValue,10);
     exit;
   end;

end;

function TFormula.fMINT(aValue: double): double;
begin
  Result:= ARound(aValue);
end;

function TFormula.fBINT(aValue: double): double;
begin
  Result:= fBANK_ROUND(aValue,0);
end;

function TFormula.fBANK_ROUND(aValue: double; aDIG: integer): double;
var
  Factor: Double;
begin
  Factor := Exp(aDIG * Ln(10));
  aValue := aValue * Factor;

  Result := Round(aValue) / Factor;
end;

procedure TFormula.fIIF(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  If Args[0].resBoolean then
  begin
    Result.resfloat:=Args[1].resfloat
  end
  else
    Result.resfloat:=Args[2].resfloat;
end;

function TFormula.ReadArgs(aArgs: TFPExpressionResult):variant;
begin
case aArgs.ResultType of
    rtBoolean           : Result:= aArgs.ResBoolean;
    rtInteger           : Result:= aArgs.ResInteger;
    rtFloat             : Result:= aArgs.ResFloat;
    rtDateTime          : Result:= aArgs.ResDateTime;
    rtString            : Result:= aArgs.ResString;
  end;
end;

procedure TFormula.EX_MRNDTO(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  Result.resfloat:=fMRNDTO(Args[0].ResFloat,ReadArgs(Args[1]));
end;

procedure TFormula.EX_BRNDTO(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  Result.resfloat:=fBRNDTO(Args[0].ResFloat,ReadArgs(Args[1]));
end;

procedure TFormula.EX_MRN(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  Result.resfloat:= fMRN(Args[0].ResFloat, Args[1].ResInteger);
end;

procedure TFormula.EX_MRNDUP(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  Result.resfloat:= fMRNDUP(Args[0].ResFloat, ReadArgs(Args[1]));
end;

procedure TFormula.EX_MINT(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat:= fMINT(Args[0].ResFloat);
end;

procedure TFormula.EX_BINT(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat:= fBINT(Args[0].ResFloat);
end;

procedure TFormula.EX_ABS(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResFloat:= ABS(Args[0].ResFloat);
end;

procedure TFormula.EX_INRANGE(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  Result.resBoolean:= (Args[0].ResFloat>=Args[1].ResInteger) AND (Args[0].ResFloat<=Args[2].ResInteger);
end;

procedure TFormula.EX_NOW(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
  Result.ResDateTime:= IncDay(now,Args[0].ResInteger);
end;

procedure TFormula.EX_DATE(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
case Args[0].ResInteger of
  0: Result.ResFloat:= DayOf(now);
  1: Result.ResFloat:= MonthOf(now);
  2: Result.ResFloat:= YearOf(now);
  3: Result.ResFloat:= WeekOf(now);
  4: Result.ResFloat:= DayOfTheWeek(now);
  5: Result.ResFloat:= DayOfTheYear(now)
  else Result.ResFloat:= 0;
end;
end;

procedure TFormula.EX_TIME(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
case Args[0].ResInteger of
  0: Result.ResFloat:= HourOf(now);
  1: Result.ResFloat:= MinuteOf(now);
  2: Result.ResFloat:= SecondOf(now)
  else Result.ResFloat:= 0;
end;
end;

procedure TFormula.EX_KURS(var Result: TFPExpressionResult; const Args: TExprParameterArray);
begin
    if Length(fCurrencyArray)>0 then
    begin
      case Args[0].ResString of
        'RUR': Result.ResFloat:= fCurrencyArray[0];
        'USD': Result.ResFloat:= fCurrencyArray[1];
        'EUR': Result.ResFloat:= fCurrencyArray[2];
        'KZT': Result.ResFloat:= fCurrencyArray[3];
        'UAH': Result.ResFloat:= fCurrencyArray[4]
      else
        Result.ResFloat:= -1;
     end;
    end else
        Result.ResFloat:= -1;
end;

constructor TFormula.Create(AOwner: TComponent);
var
  i: integer;

begin
  inherited Create(aOwner);

  Self.Builtins := [bcMath]+[bcBoolean];
  with Self.Identifiers do
  begin
    Clear;
    for i := 0 to High(VarToFiels) do
        AddFloatVariable(VarToFiels[i, 0], 0);

    for i := 0 to High(VarToFunctions) do
      case VarToFunctions[i, 0] of
        'CHOOSE': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @fIIF); // для совместимости с Fora
        'ROUND': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_MRNDTO); // для совместимости с Fora

        'IF': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @fIIF);
        'RNDTO': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_MRNDTO);
        'RN': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_MRN);
        'INT': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_MINT);
        'BINT': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_BINT);
        'BRNDTO': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_BRNDTO);
        'ABS': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_ABS);
        'INRANGE': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_INRANGE);
        'RNDUP': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_MRNDUP);
        'NOW': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_NOW);
        'DATE': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_DATE);
        'TIME': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_TIME);
        'KURS': AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @EX_KURS);
        else
          AddFunction(VarToFunctions[i, 0], VarToFunctions[i, 2][1], VarToFunctions[i, 3], @fFunctionTemplate);
      end;

  end;

  fVarArray:=nil;

end;

function TFormula.ReplaceText(const S, OldPattern, NewPattern: string;
  aStart: integer): string;
begin
  Result := S;
  UTF8Delete(Result, aStart, UTF8Length(OldPattern));
  UTF8Insert(NewPattern, Result, aStart);
end;

function TFormula.ReplaceVar(const S, OldPattern: string; var NewPattern: string;
  aStart: integer): string;
var
  _VarArrayId: Integer;
  _VarValue: String;

begin
  Result := S;
  _VarArrayId:= VarID(OldPattern);

  if GetPosSemicolon(VarToFiels[_VarArrayId,1],1)>-1 then
    NewPattern  := '('+NewPattern+')';

  if Length(VarToFiels[_VarArrayId,2])>0 then
  begin
     NewPattern  := StringReplace(NewPattern, VarToFiels[_VarArrayId, 0], VarToFiels[_VarArrayId, 2], [rfIgnoreCase]);
     NewPattern  := Prepare(NewPattern, false);
  end;

  if Assigned(fVarArray) then
  begin
    _VarValue := FloatToBase(fVarArray[_VarArrayId]);
    if Length(VarToFiels[_VarArrayId, 2])>0 then
      NewPattern  := StringReplace(NewPattern, VarToFiels[_VarArrayId, 2],
        _VarValue, [rfIgnoreCase, rfReplaceAll])
      else
        NewPattern  := StringReplace(NewPattern, VarToFiels[_VarArrayId, 1],
          _VarValue, [rfIgnoreCase, rfReplaceAll]);
  end;

  UTF8Delete(Result, aStart, UTF8Length(OldPattern));
  UTF8Insert(NewPattern, Result, aStart);
end;

function TFormula.FloatToBase(aValue: double): string;
begin
  Result := FormatFloat('0.0###',aValue);
  Result:= StringReplace(Result,',','.',[]);
end;

function TFormula.FloatStrToBase(aValue: string): string;
var
  _Float: Double;
begin
  TryStrToFloat(StringReplace(aValue,'.',DecimalSeparator,[]),_Float);

  Result := FormatFloat('0.0###',_Float);
  Result:= StringReplace(Result,DecimalSeparator,'.',[]);
end;

function TFormula.Prepare(S: string; aCheck: boolean): string;
var
  _PosSemicolon, iText, i: integer;
  _FindedText, _VarValue, _ConstValue, _KursValue: string;
  _FindedChar:  char;
  _CharVariant: TCharVariant;

begin
  Result := '';
  S:= S+' ';
  try
    if aCheck then
    begin
      self.Expression:='';
      self.Expression:= S;
    end;

  i     := 0;
  iText := 1;
  _PosSemicolon := -1;
  _CharVariant := cvNONE;
  _VarValue := '';
  _ConstValue:= '';

  while iText<=UTF8Length(S) do
  begin
    _FindedChar := S[iText];

    if (_FindedChar in VarValuesArr) or (_FindedChar in VarConstArr) then
    begin
      _PosSemicolon := GetPosSemicolon(S, iText);

      if _PosSemicolon>-1 then
      begin
        _FindedText := UTF8Copy(S, iText, _PosSemicolon-iText);

        i := VarID(_FindedText);

        GetVariants(_FindedText, i, _CharVariant);
        case _CharVariant of
          cvVAR:
          begin
            _VarValue := VarToFiels[i, 1];
            _VarValue     :=Format(VarPrepareCAST,[_VarValue]);
            S     := ReplaceVar(S, _FindedText, _VarValue, iText);
            iText := iText+UTF8Length(_VarValue)-1;
          end;
          cvFUNC:
          begin
            S     := ReplaceText(S, _FindedText, VarToFunctions[i, 1], iText);
            iText := iText+UTF8Length(VarToFunctions[i, 1])-1;
          end;

          cvCONST:
          begin
            _ConstValue:= _FindedText;
            S     := ReplaceText(S, _FindedText, _ConstValue,iText);
            iText := iText+UTF8Length(_ConstValue)-1;
          end;

          cvKURS:
          begin
            if i=-1 then exit;
            _KursValue:= FloatStrToBase(FloatToStr(fCurrencyArray[i]));
            S     := ReplaceText(S, _FindedText, _KursValue,iText);
            iText := iText+UTF8Length(_KursValue)-1;
          end;

          cvNONE: iText := _PosSemicolon-1;
        end;
      end;
    end;

    Inc(iText);
  end; {while}

  Result := S;

  except
    raise;
  end;
end;

function TFormula.Calc(aFormula: string {$IFDEF WITHBASE}; const aBase: TObject = nil{$ENDIF}):Double;
var
  i: Integer;
  {$IFDEF WITHBASE}
  _Base: TwBase;
  {$ENDIF}
begin
  try
    {$IFDEF WITHBASE}
    _Base:= TwBase(aBase);

    if Assigned(_Base) then
    begin
     if Length(aFormula)>0 then
       Result:= _Base.SQLReadDS('select ('+Prepare(aFormula)+') as RES from RDB$DATABASE').DataSet.FieldByName('RES').AsFloat;

     exit;
    end;
    //{$ELSE}

    if Assigned(fVarArray) then
      for i := 0 to High(fVarArray) do
          self.Identifiers[i].AsFloat:= fVarArray[i];

       self.Expression:=aFormula;
       Result:=self.Evaluate.ResFloat;
     {$ENDIF}
  except
    raise;
  end;
end;


end.

