unit wCustomClassThreadU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  wTProgressU, wTypesU
  ;

type

  {

      private: член может быть вызван/доступен только с помощью методов данного класса;
      public: член может быть вызван/доступен из любого другого места программы;
      protected: член может быть вызван/доступен из других классов в том же модуле и из производных классов, но не из внешних классов.
      published: переменная опубликована и будет доступна в Инспекторе объектов IDE.

  }
  TThreadExceptionEvent = procedure(
    Thread: TThread; E: Exception
  ) of object;

  { TwCustomThread }

  TwCustomThread = class(TThread)
    private
      fonEndThread: TNotifyEvent;
      fonStatusUpdate: TNotifyEvent;
      fonExceptionEvent: TThreadExceptionEvent;
      fOutStringArr: ArrayOfString;
      fStatus:   string;
      fResult: boolean;
      fStopForce: boolean;

    protected
      procedure SetStatus(aText: string);

    published

    public
      constructor Create(CreateSuspended: boolean);
      destructor Destroy(); override;

      property onEndThread: TNotifyEvent read fonEndThread write fonEndThread;
      property onStatusUpdate: TNotifyEvent read fonStatusUpdate write fonStatusUpdate;
      property onExceptionEvent: TThreadExceptionEvent read fonExceptionEvent write fonExceptionEvent;

      procedure Stop();

      property Result: boolean read fResult write fResult;
      property Status: string read fStatus write fStatus;
      property StopForce: boolean read fStopForce write fStopForce;
      property OutStringArr: ArrayOfString read fOutStringArr write fOutStringArr;
  end;

  TProgressBars = record
    Top: TProgressBarValues;
    Bottom: TProgressBarValues;
  end;

  { TwCustomThreadWithProgressBar }

  TwCustomThreadWithProgressBar = class(TwCustomThread)
    private
      fonProgressInit: TProgressEventInit;
      fonProgressStatus: TProgressEventStatus;
      fonProgressUpdate: TProgressEventUpdate;

    protected
      procedure ProgressInit(aProgressBar:TProgressBarName; aValue: integer);
      procedure ProgressUpdate(aProgressBar:TProgressBarName; const aValue: integer = -1);
      procedure ProgressStatus(aStatus:string);
    public
      property onProgressUpdate: TProgressEventUpdate read fonProgressUpdate write fonProgressUpdate;
      property onProgressInit: TProgressEventInit read fonProgressInit write fonProgressInit;
      property onProgressStatus: TProgressEventStatus read fonProgressStatus write fonProgressStatus;
  end;


implementation

{ TwCustomThreadWithProgressBar }

procedure TwCustomThreadWithProgressBar.ProgressInit(aProgressBar: TProgressBarName; aValue: integer);
begin
  if Assigned(onProgressInit) then onProgressInit(aProgressBar, aValue);
end;

procedure TwCustomThreadWithProgressBar.ProgressUpdate(aProgressBar: TProgressBarName; const aValue: integer);
begin
  if Assigned(onProgressUpdate) then onProgressUpdate(aProgressBar, aValue);
end;

procedure TwCustomThreadWithProgressBar.ProgressStatus(aStatus: string);
begin
  if Assigned(onProgressStatus) then onProgressStatus(aStatus);
end;

{ TwCustomThread }
procedure TwCustomThread.SetStatus(aText: string);
begin
  fStatus := aText;
  if Assigned(onStatusUpdate) then
     onStatusUpdate(Self);
end;

constructor TwCustomThread.Create(CreateSuspended: boolean);
begin
  fOutStringArr:= nil;
  FreeOnTerminate := True;
  inherited Create(CreateSuspended);
end;

destructor TwCustomThread.Destroy();
begin
  inherited Destroy();
end;

procedure TwCustomThread.Stop();
begin
  fStopForce:= true;
  fResult:= false;
end;

end.

