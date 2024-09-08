unit wPlugin;

{$mode objfpc}{$H+}
{
(c) Degtyarev A. A.
License in file LICENSE.txt

Git-репозиторий: https://bitbucket.org/wofs/os_wplugin/src
Назначение:
Загрузка дополнительных форм как плагины в PageControl и их открепление при необходимости.
Открепление происходит при попытке "оторвать (сдвинуть мышью)" заголовок страницы, на которой находятся элементы плагина.
}

{
ВНИМАНИЕ!
---
Модуль переопределяет событие onClose главной формы приложения (точнее формы, указанной в процедуре __wPluginInit)!
Если вам необходимо использовать свой обработчик, то установите константу wChangeFormDestinationOnClose = false (в const этого модуля)
В этом случае, добавьте в свою процедуру закрытия главной формы (onClose) следующий код:

// [CODE]
for  i:=Plugin.Count-1 downto 0 do    // перебор всех загруженных плагинов
   begin
        Plugin[i].PluginUnload(i);    // выгрузка плагина из памяти
   end;

FreeAndNil(Plugin);
FreeAndNil(PluginList);
// [END CODE]

Это необходимо для корректного освобождения выделенной памяти.
-----------------
ЕСЛИ формы плагинов создаются в текущем приложении, где использован данный класс, то удалите формы плагинов из автосоздаваемых:
 ГлавноеМеню -> Проект -> Параметры Проекта -> Формы -> Автосоздаваемые формы.
}
{
Настройка:
  1. Добавляем модули плагинов и модуль  wPlugin в uses ГЛАВНОЙ ФОРМЫ приложения.
  2. Настраиваем список подгружаемых форм (плагинов. Для этого вызываем процедуру __wPluginSettings(TFormsName: array of TFormClass; DefaultPluginLoaded: byte);
  3. Инициализируем класс процедурой __wPluginInit(FmDestination: TForm; PControl: TPageControl; Owner: TWinControl);
  4. В созданной форме плагина, который хотим подгрузить обязательно наличие основного контрола Panel. Все остальные элементы должны быть дочерними для основного контрола Panel.
  5. Если требуется использовать для созданных вкладок изображения, то просто укажите источник в PageControl.Images (в инспекторе объектов или основном коде приложения) и изображения будут использованы автоматически. Номер изображения в ImageList будет соотвествовать индексу подгружаемого плагина (нумерация с нуля).
  6. При необходимости добавить свою статус-панель на открепленную форму плагина - установить флаг wStatusBar = true; (секция const).
}

{
--== Пример инициализации ==--
//Form1U
...
implementation
uses
    wPlugin, Form2U, Form3U, Form4U; // где Form2U, Form3U, Form4U - модули подгружаемых плагинов. Перед добавлением, проверьте, что они не указаны в основном uses главного модуля.
...
---|||||||||||----
Событие onCreate главной формы приложения:

__wPluginSettings([TForm2,TForm3,TForm4],3);
// TForm2, TForm3, TForm4 - классы подгружаемых форм массивом TFormClass (нумерация с нуля).TForm2 имеет индекс 0, TForm3 - 1,TForm4 - 2.
// количество подгружаемых плагинов не ограничено
// 3 - сколько плагинов из списка (отсчет сначала списка) подгружать /0 - не подгружать плагины по умолчанию.
// Если установлено false, то каждый плагин необходимо подгружать вручную, вызвав процедуру:
// Plugin.Add(TwPlugin.Create(0)); // где 0 - индекс плагина в массиве списка форм (TFormClass). В данном примере это будет TForm2
// Plugin.Add(TwPlugin.Create(1)); // где 1 - индекс плагина в массиве списка форм (TFormClass). В данном примере это будет TForm3
// так же, в этом режиме - будет создано меню, вызываемое правой кнопкой мыши на заголовке вкладок и которым можно будет их закрывать, выгружая плагин из памяти.

__wPluginInit(Form1,PageControl1,self);
// Form1 - главная форма приложения, на которой находится PageControl, который будет материнским для подгружаемых плагинов
// PageControl1 - материнский PageControl, куда страницами будут подгружены плагины
// self - компонент-родитель.
}
{
Дополнительно:
Если необходимо получить экземпляр плагина по имени формы, что можно воспользоваться функцией:
function __wPluginGetForm(_FormName: string): TwPlugin; // возвращает экземпляр плагина по имени формы

Если нужно выоводить статус в статус-панель плагина (установить const wStatusBar = true;), то используйте функцию
SetStatus из методов плагина.
При наличии статус-бара в форме-плагине запишет его в него и вернет пустую строку, если статус-бара нет, то вернет строку статуса назад.

}

interface

uses
  Classes, SysUtils, ComCtrls, Forms, ExtCtrls, Controls, fgl, Menus;

type

  { TwPlugin }

  TwPlugin = class
  private
    FormDestination: TForm;

    PluginIndex: integer; // Индекс плагина (номер в массиве загруженных)
    PluginListIndex: integer; // Индекс плагина в списке плагинов
    PluginPageIndex: integer; // Индекс страницы, в которую плагин загружен

    FormPlugin: TForm; // Форма донор (плагин)
    TFormPlugin: TFormClass; // ТФорма донор (класс формы плагина)
    FormPluginName: string; // имя формы, которое имеет плагин

    StatusBar: TStatusBar; // Статус-бар

    mX, mY: integer;  // координаты X,Y
    mStateClick: boolean; // флаг нажатия на кнопку мыши
    pcIndex: integer; // индекс страницы (служебное)

    function GetwStatusBar: boolean;
    procedure MenuUnpin();
    procedure MouseUpUnpin;

  public

    procedure InsertPluginToDestinationForm();
    // Внедрение контролов плагина в форму-получатель

    // прочие методы плагина
    procedure MouseDown(Button: TMouseButton; X, Y: integer);
    procedure MouseUp();
    procedure MouseMoveCancel();
    procedure MouseMove(X, Y: integer);
    procedure Unpin(X, Y: integer); // открепление встроенного контрола
    function SetStatus(_Text:string):string; // установить статус на статус-панели если wStatusBar = true - возвращает переменную _Text;
    procedure PluginClose(Sender: TObject);
    procedure PluginUnpin(Sender: TObject);
    // закрытие вкладки плагина с выгрузкой формы плагина из памяти
    procedure Unload();// выгружает плагин из памяти

    constructor Create(PlugIndex: integer);
    destructor Destroy();

    // переопределенные процедуры формы плагина
    procedure onMouseMove(Sender: TObject; Shift: TShiftState; X, Y: integer);
    procedure onMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: integer);
    procedure onCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure onWindowStateChange(Sender: TObject);

    // переопределение процедуры формы-получателя
    procedure PC_onMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: integer); // нажатие кнопки мыши в PageControl
    procedure PC_onMouseMove(Sender: TObject; Shift: TShiftState; X, Y: integer);
    // перемещение кнопки мыши по PageControl
    procedure PC_onMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: integer); // отжатие кнопки мыши в PageControl
    procedure FD_onClose(Sender: TObject; var CloseAction: TCloseAction);
    // закрытие формы-получателя
    procedure FormClose;

    property Index: integer read PluginIndex write PluginIndex;   // Индекс плагина (номер в массиве загруженных)
    property PageIndex: integer read PluginPageIndex write PluginPageIndex; // Индекс страницы, в которую плагин загружен
    property Name: string read FormPluginName write FormPluginName; // Индекс страницы, в которую плагин загружен
    property HaveStatusPanel: boolean read GetwStatusBar;
    property FormMain: TForm read FormDestination;

  end;

  TPluginLoadedList = specialize TFPGObjectList<TwPlugin>;  // типизированный список

procedure __wPluginSettings(TFormsName: array of TFormClass; DefaultPluginLoaded: byte);
procedure __wPluginInit(FmDestination: TForm; PControl: TPageControl;
  Owner: TWinControl); // инициализация плагина
function __wPluginGetForm(_FormName: string): TwPlugin;  // возвращает экземпляр плагина по имени формы

const
  wChangeFormDestinationOnClose = false;
  // true - изменить событие onClose главной формы приложения, false - не изменять.

  wStatusBar = true; // true - добавить свою статус-панель на открепленной форме.

  MouseMoviDelta = 5;
// дельта, при достижении которой будет принято решении об откреплении плагина от страницы

  ShowFormInTaskBar = true; // показывать ли кнопку окна в таскбаре

var
  // переменные, используемые в главной форме
  Plugin: TPluginLoadedList; // список загруженных плагинов
  PluginList: TStringList; // список форм-плагинов
  PluginAutoLoaded: boolean;// Флаг окончания автозагрузки плагинов

  AOwnerParent: TWinControl; // родитель создаваемого плагина
  FormDestination: TForm; // Форма-получатель элемента
  PageControl: TPageControl; //PageControl формы-получателя

  PCHeaderMenu: TPopupMenu;   // всплывающее меню PageControl
  PCHeaderMenuItemClose,PCHeaderMenuItemUnpin: TMenuItem;

  wDefaultPluginLoaded: byte;
  // флаг. Подгружать ли все плагины при загрузке приложения true / false

implementation

procedure __wPluginSettings(TFormsName: array of TFormClass; DefaultPluginLoaded: byte);
var
  i: integer;
begin
  PluginList := TStringList.Create; // создали общий список плагинов
 // PluginList.AddStrings(FormsName); // загружаем список из массива

  for i := 0 to Length(TFormsName) - 1 do
  begin
    RegisterClassAlias(TFormsName[i],'Plugin'+IntToStr(i)); // регистрируем классы из массива
    PluginList.Add('Plugin'+IntToStr(i));
  end;

  wDefaultPluginLoaded := DefaultPluginLoaded;
  // Флаг. Подгружать ли все плагины при загрузке приложения true / false

end;


procedure __wPluginInit(FmDestination: TForm; PControl: TPageControl;
  Owner: TWinControl);
var
  i: integer;
begin
    AOwnerParent := Owner;
    FormDestination := FmDestination;
    PageControl := PControl;

    PluginAutoLoaded:= false; // сбрасываем флаг окончания автозагрузки плагинов

    Plugin := TPluginLoadedList.Create();// создали список загруженных плагинов

    if (wDefaultPluginLoaded > 0) and (wDefaultPluginLoaded < PluginList.Count+1)then
    begin
      for i := 0 to wDefaultPluginLoaded-1 do
        // создаем экземпляры плагинов и заполняем массив
      begin
        Plugin.Add(TwPlugin.Create(i));
      end;
    end;
end;

function __wPluginGetForm(_FormName: string): TwPlugin;
var
  i:integer;
  _Name: string;
begin
  try
  result:= nil;
  if Plugin <> nil then
  begin

   for i:=0 to Plugin.Count-1 do
   begin
     _Name:= Plugin[i].Name;
     if _FormName = _Name then
     begin
       result:=Plugin[i];
       exit;
     end else result:= nil;

   end;
   end else
   result:= nil;
  except
        on E: Exception do
    begin
       raise;
    end;
  end;
end;

constructor TwPlugin.Create(PlugIndex: integer);
var
  TFormName:string;
begin
    PluginIndex := Plugin.Count;
    PluginListIndex:=PlugIndex;
    TFormName:='Plugin'+IntToStr(PluginListIndex);
    TFormPlugin := TFormClass(FindClass(TFormName));

    // переопределяем события формы плагина
    FormPlugin := TFormPlugin.Create(AOwnerParent);
    FormPlugin.Visible:= false;
    FormPlugin.onMouseMove := @onMouseMove;
    FormPlugin.onMouseUp := @onMouseUp;
    FormPlugin.onCloseQuery := @onCloseQuery;
    FormPlugin.onWindowStateChange := @onWindowStateChange;

    Name:= FormPlugin.Name;
    // переопределяем события формы назначения
    if PageControl.OnMouseMove = nil then
    begin
      if wChangeFormDestinationOnClose then
        FormDestination.OnClose := @FD_onClose;

      PageControl.OnMouseDown := @PC_onMouseDown;
      PageControl.OnMouseMove := @PC_onMouseMove;
      PageControl.OnMouseUp := @PC_onMouseUp;
    end;


      if (PageControl.PopupMenu = nil) then
      begin

       PCHeaderMenu := TPopupMenu.Create(PageControl);
       PCHeaderMenu.Images:= PageControl.Images;

       PCHeaderMenuItemClose := TMenuItem.Create(PCHeaderMenu);
       PCHeaderMenuItemClose.Caption := 'Закрыть вкладку';
       PCHeaderMenuItemClose.ImageIndex:=7;
       PCHeaderMenuItemClose.OnClick := @PluginClose;
       PCHeaderMenu.Items.Add(PCHeaderMenuItemClose);
       PageControl.PopupMenu := PCHeaderMenu;

       PCHeaderMenuItemUnpin := TMenuItem.Create(PCHeaderMenu);
       PCHeaderMenuItemUnpin.Caption := 'Открепить вкладку';
       PCHeaderMenuItemUnpin.ImageIndex:=6;
       PCHeaderMenuItemUnpin.OnClick := @PluginUnpin;
       PCHeaderMenu.Items.Add(PCHeaderMenuItemUnpin);

      end;

  InsertPluginToDestinationForm();
end;

destructor TwPlugin.Destroy();
begin

end;

procedure TwPlugin.onMouseMove(Sender: TObject; Shift: TShiftState; X, Y: integer);
begin
  MouseMove(X, Y);
end;

procedure TwPlugin.onMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: integer);
begin
  MouseUp();
end;

procedure TwPlugin.onCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  CanClose := False;
  InsertPluginToDestinationForm();
end;

procedure TwPlugin.onWindowStateChange(Sender: TObject);
begin
  if (FormPlugin.WindowState = wsMinimized) and (not ShowFormInTaskBar) then
       InsertPluginToDestinationForm();
end;

procedure TwPlugin.PC_onMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: integer);
var
  i: integer;
begin
  for i := 0 to Plugin.Count - 1 do
  begin
    if Plugin[i].PageIndex = (Sender as TPageControl).ActivePageIndex then
      Plugin[i].MouseDown(Button, X, Y);
  end;
end;

procedure TwPlugin.PC_onMouseMove(Sender: TObject; Shift: TShiftState; X, Y: integer);
var
  i: integer;
begin
  for i := 0 to Plugin.Count - 1 do
  begin
    if Plugin[i].PageIndex = (Sender as TPageControl).ActivePageIndex then
    begin
      Plugin[i].Unpin(X, Y);
    end;
  end;
end;

procedure TwPlugin.PC_onMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: integer);
var
  i: integer;
begin

  for i := 0 to Plugin.Count - 1 do
  begin
    if Plugin[i].PageIndex = (Sender as TPageControl).ActivePageIndex then
      Plugin[i].MouseMoveCancel();
  end;
end;

procedure TwPlugin.FD_onClose(Sender: TObject; var CloseAction: TCloseAction);
var
  i: integer;
begin
  for  i:=Plugin.Count-1 downto 0 do
  begin
       Plugin[i].Unload();  // выгружаем все плагины
  end;

  FreeAndNil(Plugin);
  FreeAndNil(PluginList);
end;

procedure TwPlugin.FormClose;
begin
  FormPlugin.Close;
end;

function TwPlugin.GetwStatusBar(): boolean;
begin
   result:= wStatusBar;
end;

procedure TwPlugin.InsertPluginToDestinationForm();
begin
  if FormPlugin = nil then
    exit;

    if ShowFormInTaskBar then
       FormPlugin.ShowInTaskBar:= stNever;

    if FormPlugin.Visible then FormPlugin.Visible := False;

    with FormDestination do
    begin

        PageControl.AddTabSheet;
        pcIndex := PageControl.PageCount - 1;
        PageControl.Page[pcIndex].Caption := FormPlugin.Caption;
        PageIndex:= pcIndex;  // устанавливаем новый индекс страницы

        if PluginList <> nil
        then
        begin
            if (not PluginAutoLoaded) and (PageControl.PageCount = PluginList.Count-(PluginList.Count-wDefaultPluginLoaded)) then
            begin
              PageControl.ActivePageIndex:=0;
              PluginAutoLoaded:=true;
            end else
                PageControl.ActivePageIndex := pcIndex;
        end;

        if PageControl.Images <> nil then
        begin
            PageControl.Page[pcIndex].ImageIndex := PluginListIndex;
        end;

        if (wStatusBar) and (StatusBar <> nil) then
        begin
          FreeAndNil(StatusBar);
        end;

      FormPlugin.Controls[0].Parent := PageControl.Pages[pcIndex];
      PageControl.Page[pcIndex].Enabled := True;
    end;
end;

procedure TwPlugin.MouseDown(Button: TMouseButton; X, Y: integer);
var
  i: integer;
  pnt: TPoint;
begin

    if Button = mbLeft then
    begin
      mX := X;
      mY := Y;
      mStateClick := True;
    end
    else
    begin      // если другая кнопка - активируем кликнутый таб
      pnt.x := X;
      pnt.y := Y;
      PageControl.ActivePageIndex := PageControl.IndexOfTabAt(pnt); //TabIndexAtClientPos(pnt);
      mStateClick := false;
    end;

end;

procedure TwPlugin.MouseUpUnpin;
var
  i: integer;
begin
  PageControl.Pages[pcIndex].Controls[0].Parent := FormPlugin;

  if ShowFormInTaskBar then
       FormPlugin.ShowInTaskBar:= stAlways;

    with FormDestination do
    begin

        PageControl.Page[pcIndex].Free;

        for i := 0 to Plugin.Count - 1 do
        begin
          if Plugin[i].PageIndex = pcIndex then
          begin
            Plugin[i].PageIndex:= -1;
          end
          else
            if Plugin[i].PageIndex > pcIndex  then
              Plugin[i].PageIndex:= Plugin[i].PageIndex - 1;

      end;

      if (wStatusBar) and (StatusBar = nil) then
      begin
           StatusBar:= TStatusBar.Create(FormPlugin);
           FormPlugin.InsertControl(StatusBar);
      end;

      mStateClick := False;
      PageControl.Cursor := crDefault;
      FormPlugin.Cursor := crDefault;
    end;
end;


procedure TwPlugin.MouseUp();
var
  i: integer;
begin
{$IFDEF WINDOWS}

MouseUpUnpin;

{$ENDIF}
end;

procedure TwPlugin.MouseMoveCancel();
begin
  mStateClick := False;
end;

procedure TwPlugin.MouseMove(X, Y: integer);
var
  pnt: TPoint;
begin
  if mStateClick then
  begin
    FormPlugin.Cursor := crSizeAll; // меняем курсор
    pnt := Mouse.CursorPos;
    FormPlugin.Top := pnt.Y - 150;
    FormPlugin.Left := pnt.X - 150;
  end;
end;

procedure TwPlugin.MenuUnpin();
var
  pnt: TPoint;
begin
  PageControl.cursor := crSizeAll; // меняем курсор

    pcIndex := PageIndex;

    with FormDestination do
    begin
      FormPlugin.WindowState := wsNormal;
      FormPlugin.Height := PageControl.Pages[pcIndex].Height;
      FormPlugin.Width := PageControl.Pages[pcIndex].Width;
      //FormPlugin.Caption := PageControl.Page[pcIndex].Caption;
      PageControl.Page[pcIndex].Enabled := False;

      pnt := Mouse.CursorPos;
      FormPlugin.Left := pnt.x;
      FormPlugin.Top := pnt.y;

      FormPlugin.Show;

      MouseUpUnpin;

    end;

end;

procedure TwPlugin.Unpin(X, Y: integer);
var
  pnt: TPoint;
  i: integer;
begin

    if mStateClick then
    begin
      PageControl.cursor := crSizeAll; // меняем курсор

      if ((X < mX - MouseMoviDelta) or (X > mX + MouseMoviDelta) or
        (Y < mY - MouseMoviDelta) or (Y > mY + MouseMoviDelta)) then
      begin
        mX := X;
        mY := Y;

        pcIndex := PageControl.ActivePageIndex;

        with FormDestination do
        begin
          FormPlugin.WindowState := wsNormal;
          FormPlugin.Height := PageControl.Pages[pcIndex].Height;
          FormPlugin.Width := PageControl.Pages[pcIndex].Width;
          //FormPlugin.Caption := PageControl.Page[pcIndex].Caption;
          PageControl.Page[pcIndex].Enabled := False;

          pnt := Mouse.CursorPos;
          FormPlugin.Left := pnt.x - 150;
          FormPlugin.Top := pnt.y - 150;

          FormPlugin.Show;

          // если мак или линукс - сразу после открепления отображаем окно.
          {$IFDEF UNIX}// mac & linux
                  MouseUpUnpin;
          {$ENDIF}
        end;
      end;
    end
    else
      PageControl.cursor := crDefault;
end;

function TwPlugin.SetStatus(_Text: string):string;
begin
    result:='';
    if (wStatusBar) and (StatusBar <> nil) then
    begin
         StatusBar.SimpleText:=_Text;
    end else
      result:=_Text;
end;

procedure TwPlugin.PluginClose(Sender: TObject);
var
  i: integer;
  _PageIndex: integer;
  _PluginPageIndex: integer;
  _PluginDeleteIndex: integer;
begin

  _PageIndex := PageControl.ActivePageIndex;

    for i := 0 to Plugin.Count - 1 do
    begin
      _PluginPageIndex := Plugin[i].PageIndex;

      if _PluginPageIndex = _PageIndex then
      begin
        Plugin[i].PageIndex:= -1;
        _PluginDeleteIndex := i;
      end
      else
      begin
        if (_PluginPageIndex > _PageIndex) then
          Plugin[i].PageIndex:= _PluginPageIndex - 1;
      end;
    end;
      Plugin[_PluginDeleteIndex].Unload();

    PageControl.Page[_PageIndex].Free;

end;

procedure TwPlugin.PluginUnpin(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to Plugin.Count - 1 do
  begin
    if Plugin[i].PageIndex = PageControl.ActivePageIndex then
    begin
      Plugin[i].MenuUnpin();
      break;
    end;
  end;
end;

procedure TwPlugin.Unload();
var
  i: integer;
begin

    if (wStatusBar) and (StatusBar <> nil) then
    begin
      FreeAndNil(StatusBar);
    end;

  FreeAndNil(Plugin[Index].FormPlugin);

  Plugin.Delete(Index);
  for i:=0 to Plugin.Count-1 do
    Plugin[i].Index:=i;

end;

end.
