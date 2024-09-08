unit wLogU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}
{
-= Использование =-
1. __Log:=TwLog.Create;
2. выставить глобальный флаг __onLog := true;

wLog(_Who, _Text: string); - запись в лог (глобально доступная процедура)

__Log.Add('Main','Завершение приложения.'); // запись в лог
__Log.SaveLog(); // сохраняем в файл log.txt
FreeAndNil(__Log);  // освобождаем память

procedure wStatus(_FormID,_Text: string; _Panel0: boolean);  - работает в связке с wPlugin  выводит статус в главную форму или в форму плагина.


}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Dialogs;
type

  { TwLog }

  TwLog = class (TStringList)
    public
      procedure Add(who,s:string);
      function SaveLog():boolean;
      procedure SaveLogError(_Error: TObject);
  end;

  procedure wLog(_Who,_Text:string);
  procedure wStatus(_FormID,_Text: string; _Panel0: boolean);
  procedure wLicenseOK(AValue: boolean); //[wLicense]
var
  __Log: TwLog;
  __onLog: boolean; // флаг включения логирования

implementation
uses
  FmMainU, wPlugin, wBaseU;

procedure wLog(_Who, _Text: string);
begin
  if __onLog and (__Log<>nil) then
   begin
      __Log.Add(_Who,_Text);
   end;
end;

procedure wStatus(_FormID,_Text: string; _Panel0: boolean);
var
  _Plugin:TwPlugin;
  _Status: string;
begin
     _Plugin:= nil;

     _Plugin:= __wPluginGetForm(_FormID);

     if _Plugin <> nil then
          _Status:= _Plugin.SetStatus(_Text) else
          _Status:= _Text;

     if Length(_Status)>0 then
        begin
           if _Panel0 then FmMain.SetStatus(_Status,true) else FmMain.SetStatus(_Status,false);
        end;

     _Plugin:= nil;

     wLog(_FormID+' [Status]',_Text);
end;

procedure wLicenseOK(AValue: boolean);
begin
  FmMain.LicenseOK:=AValue;    //[wLicense]
end;

{ TwLog }

procedure TwLog.Add(who,s: string);
begin
     if (__Log <> nil) then
     begin
       if (Length(who) <> 0) then
          __Log.Add(DateTimeToStr(now())+' | '+who+' | '+s) else
          __Log.Add(DateTimeToStr(now())+' | '+s)
     end else
     ShowMessage('wLog - ошибка записи строки в лог: "[who] '+who+' [s] '+s+'" Eror: Объект не существует!');
end;

function TwLog.SaveLog: boolean;
begin
  result:=false;
  if __Log<> nil then
  begin
    try
      __Log.SaveToFile(PathLogFiles_Unsafe+'log.txt');
      result:=true;
    except
     result:=false;
    end;
  end;
end;

procedure TwLog.SaveLogError(_Error: TObject);
begin
  if __Log<> nil then
  begin
    try
      if Assigned(_Error) then
        __Log.Add(Exception(_Error).ClassName ,Exception(_Error).Message);
    finally
      __Log.SaveToFile(PathLogFiles_Unsafe+'log-crash.txt');
    end;
  end;
end;

end.

