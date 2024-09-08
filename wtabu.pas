unit wTabU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, ComCtrls, fgl, Dialogs;
type

    { TwDataTab }

    TwDataTab = class
    private
      FValue : integer;
    public
      constructor Create(i : integer);
      property Value : integer read FValue write FValue;
    end;

    { TwTab }

    TwTab = class
    private
      FormName: string;
      wTabIndex: integer;

      _TabControl : TTabControl;
      _List: TList;
      function GetTabIndex: integer;
      function GetValueToString(Index: integer): string;
      procedure SetTabIndex(AValue: integer);
      function GetCount:integer;
      function GetText(Index: integer): string;
      procedure SetVisible(AValue: boolean);
      function GetVisible:boolean;
      function GetValue(Index: integer):integer;
      procedure SetValue(Index: integer; AValue: integer);

      procedure Log(_Text:string);
      procedure SetStatus(_Text:string); // вывод статуса

    public
      constructor Create(Sender:TObject; _Tab: TTabControl);
      destructor Destroy();

      function Add(_Text: string; AValue: integer): integer;
      function Del(_Index: integer): boolean;
      procedure Clear;

      property Visible : boolean read GetVisible write  SetVisible;
      property Count: integer read GetCount;
      property TabIndex: integer read GetTabIndex write SetTabIndex;
      property Text[Index:integer]: string read GetText;
      property Value[Index:integer]: integer read GetValue write SetValue;
      property ValueToString[Index:integer]: string read GetValueToString;

      property Index: integer read wTabIndex write wTabIndex;
      property TabControl: TTabControl read _TabControl write _TabControl;
      property ListValue: TList read _List write _List;
    end;


implementation
uses
  wLogU;


{ TwDataTab }

constructor TwDataTab.Create(i: integer);
begin
    fValue:=i;
end;

{ TwTab }

procedure TwTab.SetTabIndex(AValue: integer);
begin
  TabControl.TabIndex:=AValue;
end;

function TwTab.GetTabIndex: integer;
begin
  result:= TabControl.TabIndex;
end;

function TwTab.GetValueToString(Index: integer): string;
begin
  result:= IntToStr(Value[Index]);
end;

constructor TwTab.Create(Sender: TObject; _Tab: TTabControl);
begin
    try
      FormName:= TComponent(Sender).Name;
      TabControl:= _Tab;

      TabControl.Tabs.Clear;
      _List:= TList.Create;
    except
      on E: Exception do
        begin
          __Log.SaveLogError(E);
          wLog('wTab','Ошибка [Create]: "' + E.Message + '"');
          ShowMessage('Ошибка [Create]: "' + E.Message + '"');
          raise;
        end;
    end;
end;

destructor TwTab.Destroy();
begin
   Clear();
   _List.Free;
end;

function TwTab.Add(_Text: string; AValue: integer): integer;
begin
   TabControl.Tabs.Add(_Text);
   _List.Add(TwDataTab.Create(AValue));
   result:=_List.Count-1;
end;

function TwTab.Del(_Index: integer): boolean;
begin
    try
      TabControl.Tabs.Delete(_Index);
      TwDataTab(_List.Items[_Index]).Free;
      _List.Delete(_Index);
      result:=true;
    except
      result:=false;
    end;
end;

function TwTab.GetValue(Index: integer): integer;
begin
     result:=TwDataTab(_List.Items[Index]).Value;
end;

procedure TwTab.SetValue(Index: integer; AValue: integer);
begin
      TwDataTab(_List.Items[Index]).Value:=AValue;
end;

procedure TwTab.Log(_Text: string);
begin
  // здесь напишите процедуру ведения лог-файла
  // если вы не ведете лог-файл, то оставьте тело функции пустым
 wLog('['+FormName+']'+'['+TabControl.Name+'] '+'[wTab] ',_Text);
end;

procedure TwTab.SetStatus(_Text: string);
begin
  wStatus(FormName,_Text,true);
end;

procedure TwTab.SetVisible(AValue: boolean);
begin
  TabControl.Visible:=AValue;
end;

function TwTab.GetVisible: boolean;
begin
   result:= TabControl.Visible;
end;

function TwTab.GetCount: integer;
begin
  result:= TabControl.Tabs.Count;
end;

function TwTab.GetText(Index: integer): string;
begin
  result:=TabControl.Tabs[Index];
end;

procedure TwTab.Clear;
var
  i: integer;
begin
       TabControl.Tabs.Clear;

       for  i:= _List.Count-1 downto 0 do
            TwDataTab(_List.Items[i]).Free;

       _List.Clear;
end;

end.

