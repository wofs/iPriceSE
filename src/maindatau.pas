unit MainDataU;

{$mode objfpc}{$H+}

interface

uses
  Classes, db,
  IBDatabase, IBQuery, IBSQL;

type

  { TMainData }

  TMainData = class(TDataModule)
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  MainData: TMainData;

implementation

{$R *.lfm}

end.

