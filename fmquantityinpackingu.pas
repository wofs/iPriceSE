unit FmQuantityInPackingU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, SpinEx, Forms, Controls, Graphics, Dialogs,
  StdCtrls, ExtCtrls, Buttons;

type

  { TFmQuantityInPacking }

  TFmQuantityInPacking = class(TForm)
    btnCancel: TBitBtn;
    btnOK: TBitBtn;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Panel1: TPanel;
    spQuantInPackLeft: TSpinEditEx;
    spQuantInPackRight: TSpinEditEx;
  private

  public

  end;

var
  FmQuantityInPacking: TFmQuantityInPacking;

implementation

{$R *.lfm}

end.

