unit FmURLsU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  ExtCtrls, ComCtrls, Grids, Spin, Buttons,
  DOM, xmlread, XMLWrite,
  wGetU, wFuncU,
  wBaseU
  ;

type

  { TFmURLs }

  TFmURLs = class(TForm)
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnTest: TBitBtn;
    ePassword: TLabeledEdit;
    eUsername: TLabeledEdit;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    ImageList1: TImageList;
    Label1: TLabel;
    mBody: TMemo;
    mHeaders: TMemo;
    PageControl1: TPageControl;
    Panel1: TPanel;
    Panel2: TPanel;
    sgActions: TStringGrid;
    Splitter1: TSplitter;
    spTimeOut: TSpinEdit;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    ToolBar1: TToolBar;
    tbActAdd: TToolButton;
    tbActDel: TToolButton;
    procedure btnSaveClick(Sender: TObject);
    procedure btnTestClick(Sender: TObject);
    procedure ePasswordChange(Sender: TObject);
    procedure eUsernameChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure sgActionsClick(Sender: TObject);
    procedure spTimeOutChange(Sender: TObject);
    procedure tbActAddClick(Sender: TObject);
    procedure tbActDelClick(Sender: TObject);
  private
    fBase: TwBase;
    fFormatID: integer;
    fLocalPath: string;
    fXMLText: string;
    fSaved: boolean;
    fSgRow: integer;
    procedure FillGrid(aXMLDoc: TStream);
    procedure SaveXML(aXMLString: TStringStream);

  public
    property XMLText: string read fXMLText write fXMLText;
    property Base: TwBase read fBase write fBase;
    property FormatID: integer read fFormatID write fFormatID;
  end;

var
  FmURLs: TFmURLs;

implementation

{$R *.lfm}

{ TFmURLs }


procedure TFmURLs.FormShow(Sender: TObject);
var
  _XMLDoc: TStringStream;
begin
//  FmURLs.Caption:='Мастер загрызки прайс-листа из сети [ '+FormatName+' ]';

  _XMLDoc:= TStringStream.Create(XMLText);

  try
    FillGrid(_XMLdoc);
  finally
    _XMLDoc.Free;
  end;

  if sgActions.RowCount >1 then
    fSgRow:= sgActions.RowCount-1;
end;

procedure TFmURLs.sgActionsClick(Sender: TObject);
begin
  fSgRow:= sgActions.Row;
end;

procedure TFmURLs.spTimeOutChange(Sender: TObject);
begin
  fSaved:= false;
end;

procedure TFmURLs.tbActAddClick(Sender: TObject);
begin
  fSaved:= false;
  fSgRow:= sgActions.RowCount;
  sgActions.RowCount:= fSgRow+1;
  sgActions.Cells[0,fSgRow]:= IntTOStr(fSgRow);
  sgActions.Cells[1,fSgRow]:= '1';
  sgActions.Cells[2,fSgRow]:= '0';
end;

procedure TFmURLs.tbActDelClick(Sender: TObject);
var
  i: Integer;
begin
  fSaved:= false;
  if (sgActions.RowCount =1) or (fSgRow = 0) then exit;

  sgActions.DeleteRow(fsgRow);

  fSgRow:= fSgRow-1;

  for i:=1 to sgActions.RowCount-1 do
      sgActions.Cells[0,i]:= IntToStr(i);

end;

procedure TFmURLs.FormDestroy(Sender: TObject);
begin

end;

procedure TFmURLs.ePasswordChange(Sender: TObject);
begin
  fSaved:= false;
end;

procedure TFmURLs.btnSaveClick(Sender: TObject);
var
  _XMLString: TStringStream;
begin
  try
    _XMLString:= TStringStream.Create;

    try
      SaveXML(_XMLString);
      XMLText:= _XMLString.DataString;
    finally
      _XMLString.Free;
    end;

    fBase.SQLUpdate('FORMATS',['URL'],[XMLText],'ID='+IntToStr(FormatID),true);

    ShowMessage('Сохранение успешно завершено');
    fSaved:= true;
  except
    ShowMessage('Сохранение завершено с ошибкой!');
  end;
    Close();
end;

procedure TFmURLs.btnTestClick(Sender: TObject);
var
  _arr: ArrayArrayOfString;
  i: Integer;
  _XMLStream: TStringStream;
  wGet: TwGet;
begin
  _XMLStream:= TStringStream.Create;
  wGet:= TwGet.Create(self);

  try
    SaveXML(_XMLStream);
    _arr:= wGet.ExecuteXML(_XMLStream.DataString);
  finally
    _XMLStream.Free;
    wGet.Destroy;
  end;

  mBody.Clear;
  for i:=0 to High(_arr) do
  begin
    mHeaders.Append('');
    mBody.Append('');

    mHeaders.Append('########################## ACTION #'+ IntToStr(i+1)+' ##########################');
    mBody.Append('########################## ACTION #'+ IntToStr(i+1)+' ##########################');

    mHeaders.Append(_arr[i,0]);
    mBody.Append(_arr[i,1]);
  end;

end;

procedure TFmURLs.eUsernameChange(Sender: TObject);
begin
  fSaved:= false;
end;

procedure TFmURLs.FormCloseQuery(Sender: TObject; var CanClose: boolean);
begin
  if not fSaved then
    begin
        if MessageDlg('Закрыть без сохранения?',mtConfirmation, mbOKCancel, 0) = mrOK
         then
           ModalResult:= mrCancel
         else
           CanClose:= false;
    end else
        ModalResult:= mrOK;
end;

procedure TFmURLs.FormCreate(Sender: TObject);
begin
  fSaved:= true;
end;

procedure TFmURLs.SaveXML(aXMLString:TStringStream);
var
  xdoc: TXMLDocument;
  i, k: Integer;
  NodeActions: TDomNode;
  parentNode, RootNode: TDOMElement;
begin

  xdoc:=nil;
  try
    xdoc := TXMLDocument.create;

    //Создаём корневой узел
    RootNode := xdoc.CreateElement('FORMAT');
    Xdoc.Appendchild(RootNode);                           // Добавляем корневой узел в документ

    //Создаём родительский узел
    RootNode:= xdoc.DocumentElement;
    parentNode := xdoc.CreateElement('AUTH');
    TDOMElement(parentNode).SetAttribute('username', eUsername.Text);       // создаём атрибуты родительского узла
    TDOMElement(parentNode).SetAttribute('password', ePassword.Text);       // создаём атрибуты родительского узла
    TDOMElement(parentNode).SetAttribute('timeout', IntToStr(spTimeOut.Value));       // создаём атрибуты родительского узла

    RootNode.Appendchild(parentNode);

    RootNode:= xdoc.DocumentElement;
    parentNode := xdoc.CreateElement('ACTIONS');
    NodeActions:= RootNode.Appendchild(parentNode);

    for i:=1 to sgActions.RowCount-1 do begin

      parentNode := xdoc.CreateElement('ACTION');
      TDOMElement(parentNode).SetAttribute('NUMBER',IntToStr(i));
      TDOMElement(parentNode).SetAttribute('ON',sgActions.Cells[1,i]);
      TDOMElement(parentNode).SetAttribute('POST',sgActions.Cells[2,i]);
      TDOMElement(parentNode).SetAttribute('URL',sgActions.Cells[3,i]);
      TDOMElement(parentNode).SetAttribute('PARAMS',sgActions.Cells[4,i]);
      TDOMElement(parentNode).SetAttribute('LOCALFILENAME',sgActions.Cells[5,i]);

      NodeActions.Appendchild(parentNode);

    end;

  aXMLString.Size:= 0;

  WriteXMLFile(xdoc,aXMLString);

  //mBody.Append(XMLText);
  finally
    xdoc.Free;
  end;
end;

procedure TFmURLs.FillGrid(aXMLDoc:TStream);
var
  xdoc: TXMLDocument;                      // переменная документа
  Node: TDOMNode; // переменная узла документа
  i, k: integer;
  fTimeOut: Longint;
begin
  //if not FileExists(FileListBox1.FileName) then begin exit; end;
  if aXMLdoc.Size = 0 then
  begin
      spTimeOut.Value:= 6000;
      exit;
  end;

  xdoc:=nil;
  try
    ReadXMLFile(xdoc,aXMLDoc);
    Node:= xdoc.FindNode('FORMAT');
    Node:= Node.FirstChild;
    eUsername.Text:= TDOMElement(Node).GetAttribute('username');
    ePassword.Text:= TDOMElement(Node).GetAttribute('password');
    TryStrToInt(TDOMElement(Node).GetAttribute('timeout'),fTimeOut);
    spTimeOut.Value:= fTimeOut;
    Node:= Node.NextSibling;

    if Assigned (Node) then
        with Node.ChildNodes do
          begin
            sgActions.RowCount:= Count+1;
            try
              for i:=0 to Count-1 do begin
                if Assigned(Item[i]) then
                begin
                    if Item[i].HasAttributes then
                    for k:=0 to Item[i].Attributes.Length-1 do
                    begin
                      case Item[i].Attributes[k].NodeName of
                        'NUMBER': sgActions.Cells[0,i+1]:= Item[i].Attributes[k].NodeValue;
                        'ON': sgActions.Cells[1,i+1]:= Item[i].Attributes[k].NodeValue;
                        'POST': sgActions.Cells[2,i+1]:= Item[i].Attributes[k].NodeValue;
                        'URL': sgActions.Cells[3,i+1]:= Item[i].Attributes[k].NodeValue;
                        'PARAMS': sgActions.Cells[4,i+1]:= Item[i].Attributes[k].NodeValue;
                        'LOCALFILENAME': sgActions.Cells[5,i+1]:= Item[i].Attributes[k].NodeValue;
                      end;
                    end;

                end;
              end;

            finally
              Free;
            end;
          end;

  finally
    xdoc.Free;
  end;
end;

end.

