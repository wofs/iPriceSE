unit wGetU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fphttpclient, LConvEncoding, LazUTF8, Forms, gvector,LazFileUtils,
  DOM, xmlread, XMLWrite, opensslsockets, wLogU
  ;
type

  { TwGet }
  ArrayOfString = array of string;
  ArrayArrayOfString = array of array of string;

  TwActionsRecord = Record
    fOn: boolean;
    fPost: boolean;
    fURL: string;
    fParams: string;
    fUser: string;
    fPassword: string;
    fLocalFileName: string;
    fTimeOut: Integer;
  end;

  TwActions = specialize TVector<TwActionsRecord>;

  TwGet = class
    private
      fOwner: TComponent;
      fActions: TwActions;
      fWebGet: TFPHTTPClient;
      //fPostData: TStringList;
      //fHTML: string;
      fLocalPath: string;

      function DecodeText(aResponseHeaders, aBody: string): string;
      function GetLocalFileName(aURL: string): string;
      function SaveFile(aResponseHeaders, aLocalFileName: string;
        fResponseStream: TMemoryStream): string;
    public
      constructor Create(aOwner: TComponent);
      destructor Destroy();
      function GetPath(const aSubDir: string=''; const aPathDelimiter: boolean=true
        ): string;
      function GetFile(aURL, aLocalFileName: string): ArrayOfString;
      function PostForm(aURL: string; aPostData: TStrings; const aLocalFileName: string=''): ArrayOfString;
      function PostForm(aURL,aPostData: string): string;

      function LoadFile(): ArrayArrayOfString;
      function ExecuteXML(aXMLDoc: String): ArrayArrayOfString;
      function CreateAction(aOn, aPost: boolean; aURL, aParams, aUser,
        aPassword, aLOcalFileName: string; aTimeOut: integer): TwActionsRecord;

      property Actions: TwActions read fActions write fActions;
      property LocalPath: string read fLocalPath;
      function GetKursValute(): string;
  end;

implementation

{ TwGet }

function TwGet.CreateAction(aOn, aPost: boolean; aURL, aParams, aUser,
  aPassword, aLOcalFileName: string; aTimeOut: integer): TwActionsRecord;
begin
   result.fOn:= aOn;
   result.fPost:= aPost;
   result.fURL:= aURL;
   result.fParams:= aParams;
   result.fUser:= aUser;
   result.fPassword:= aPassword;
   result.fLocalFileName:= aLocalFileName;
   result.fTimeOut:= aTimeOut;
end;

function TwGet.GetKursValute(): string;
begin
 try
   Result:= fWebGet.SimpleGet('http://www.cbr.ru/scripts/XML_daily.asp');
 except
   raise;
 end;
end;

function TwGet.GetLocalFileName(aURL: string):string;

  function ClearFileName(aFileName:string):string;
  var
    _Pos: PtrInt;
  begin
    _Pos:= UTF8Pos('?',aFileName);
    if _Pos>0 then
         Result:= UTF8Copy(aFileName,1,_Pos-1)
    else Result:= aFileName;
  end;

var
  _Ext: String;
begin

  Result:= ClearFileName(ExtractFileName(aURL));
  _Ext:='';
  _Ext:= UTF8LOwerCase(ExtractFileExt(Result));

    case _Ext of
      '.php': Result:='';
      else
          Result:= Result;
    end;


end;

constructor TwGet.Create(aOwner: TComponent);
begin
    fOwner:= aOwner;
    fWebGet:= TFPHTTPClient.Create(fOwner);
    fWebGet.AddHeader('User-Agent','Mozilla/5.0 (compatible; fpweb)');
    fActions:= TwActions.Create;
    fWebGet.IOTimeout:= 6000;

    fLocalPath:= GetPath('prices');
end;

destructor TwGet.Destroy();
begin
  if Assigned(fWebGet) then fWebGet.Free;
  if Assigned(fActions) then fActions.Free;
end;

function TwGet.GetPath(const aSubDir: string = ''; const aPathDelimiter:boolean = true): string;
begin
  Result:= ExtractFileDir(Application.ExeName);
  if Length(aSubDir)>0 then Result:= includeTrailingPathDelimiter(Result)+aSubDir;
  if aPathDelimiter then Result:= includeTrailingPathDelimiter(Result);
end;


function TwGet.GetFile(aURL, aLocalFileName: string): ArrayOfString;
var
  fResponseStream: TMemoryStream;
  fHTML: String;
begin
  if Length(aLocalFileName) = 0 then aLocalFileName:= GetLocalFileName(aURL);

  fResponseStream:= TMemoryStream.Create;
  SetLength(Result,2);
  try
    try
      fWebGet.Get(aURL,fResponseStream);
    except
      on E: EHTTPClient do
      begin
        if not fWebGet.IsRedirect(fWebGet.ResponseStatusCode) then fHTML:= E.Message
         else fHTML:='Redirect...'+'|'+IntToStr(fWebGet.ResponseStatusCode)+'|'+fWebGet.ResponseStatusText;
      end;
    end;

    if Length(fWebGet.ResponseHeaders.Text)>0 then
    Result[0]:= fWebGet.ResponseHeaders.Text else
    Result[0]:= fHTML;

    if fResponseStream.Size>0 then
        fHTML:= SaveFile(Result[0],aLocalFileName,fResponseStream);

    Result[1]:= fHTML;

  finally
    fResponseStream.Free;
  end;


end;

function TwGet.SaveFile(aResponseHeaders, aLocalFileName: string; fResponseStream: TMemoryStream):string;
var
  _Pos, i: Integer;
  b: byte;
  _EndFileName: PtrInt;
  fFileName: String;
begin
  if not DirectoryExistsUTF8(fLocalPath) then ForceDirectoriesUTF8(fLocalPath);

  _Pos:= UTF8Pos('filename=',aResponseHeaders);
  try
    if _Pos>0 then
     begin
       if Length(aLocalFileName)=0 then
        begin
          _Pos:= _Pos+Length('filename=');
          _EndFileName:= UTF8Pos(LineEnding,aResponseHeaders,_Pos);
          fFileName:=UTF8Copy(aResponseHeaders,_Pos,_EndFileName-_Pos);
          fFileName:= StringReplace(fFileName,#34,'',[rfReplaceAll]);
        end else
          fFileName:= aLocalFileName;

       if FileExists(fLocalPath+fFileName) then DeleteFile(fLocalPath+fFileName);

       fResponseStream.SaveToFile(fLocalPath+fFileName);
       Result:='File loaded in '+fLocalPath+fFileName;
     end else
     begin
        if Length(ExtractFileName(aLocalFileName))=0 then
         begin
            fResponseStream.Seek(0,0);
            for i:=0 to fResponseStream.Size do
             begin
              fResponseStream.Read(b,1);
              Result:=Result+chr(b);
             end;
         end else
         begin
          fResponseStream.SaveToFile(fLocalPath+aLocalFileName);
          Result:='File loaded in '+fLocalPath+aLocalFileName;
         end;
     end;
  except
    Result:='FALSE';
  end;
end;

function TwGet.PostForm(aURL: string; aPostData: TStrings; const aLocalFileName: string = ''): ArrayOfString;
var
  fHTML: String;
  fResponseStream: TMemoryStream;

begin
  //if Length(aLocalFileName) = 0 then aLocalFileName:= PrepareFileName(aURL);


  fResponseStream:= TMemoryStream.Create;

  try
    SetLength(Result,2);
    fHTML:= '';

    try
      fWebGet.FormPost(aURL,aPostData,fResponseStream);
    except
      on E: EHTTPClient do
      begin
        if not fWebGet.IsRedirect(fWebGet.ResponseStatusCode) then fHTML:= E.Message else fHTML:='Redirect...'+'|'+IntToStr(fWebGet.ResponseStatusCode)+'|'+fWebGet.ResponseStatusText;
      end;
    end;

    Result[0]:= fWebGet.ResponseHeaders.Text;

    if fResponseStream.Size>0 then
        fHTML:= SaveFile(Result[0],aLocalFileName,fResponseStream);

    Result[1]:= fHTML;

  finally

    fResponseStream.Free;
  end;
end;

function TwGet.PostForm(aURL, aPostData: string): string;
begin
   try
     Result:=fWebGet.FormPost(aURL,aPostData);
   except
      on E: Exception do
      begin
        Result:='Error: '+E.Message;
      end;
   end;
end;

function TwGet.DecodeText(aResponseHeaders, aBody: string):string;
var
  _Pos, _EndFileName: PtrInt;
  aCharset: String;
begin
  _Pos:= UTF8Pos('charset=',aResponseHeaders);
  if _Pos>0 then
  begin
   _Pos:= _Pos+Length('charset=');
   _EndFileName:= UTF8Pos(LineEnding,aResponseHeaders,_Pos);
   aCharset:=UTF8LowerCase(UTF8Copy(aResponseHeaders,_Pos,_EndFileName-_Pos));
  end;

    case aCharset of
      'cp1251': Result:= CP1251ToUTF8(aBody);//body
      else
          Result:= aBody;//body
    end;
end;

function TwGet.LoadFile(): ArrayArrayOfString;
var
  i, _Pos: Integer;
  ResponseResult: ArrayOfString;
  PostData: TStringList;
  _EndFileName: PtrInt;
  aCharset: String;
begin
  try
    if Actions.IsEmpty then
       begin
         SetLength(Result,1,1);
         Result[0,0]:='FALSE';
         exit;
       end;
    SetLength(Result,Actions.Size,2);
    PostData:= TStringList.Create;
    ResponseResult:= nil;
    fWebGet.Free;
    fWebGet:= TFPHTTPClient.Create(fOwner);
    fWebGet.AddHeader('User-Agent','Mozilla/5.0 (compatible; fpweb)');
    try
      for i:=0 to Actions.Size-1 do
        if Actions.Items[i].fOn then
           begin
             fWebGet.IOTimeout:= Actions.Items[i].fTimeOut;
             if Actions.Items[i].fPost then
             begin
              PostData.Clear;
              PostData.Delimiter:='&';
              PostData.DelimitedText:=Actions.Items[i].fParams;

              if FileExists(Actions.Items[i].fLocalFileName) then DeleteFile(Actions.Items[i].fLocalFileName);

              ResponseResult:= PostForm(Actions.Items[i].fURL,PostData,Actions.Items[i].fLocalFileName);
              Result[i,0]:= ResponseResult[0]; //Headers
              Result[i,1]:= DecodeText(ResponseResult[0],ResponseResult[1]);//body
             end else
             begin
              fWebGet.UserName:= Actions.Items[i].fUser;
              fWebGet.Password:= Actions.Items[i].fPassword;
              ResponseResult:= GetFile(Actions.Items[i].fURL,Actions.Items[i].fLocalFileName);

              Result[i,0]:= ResponseResult[0];// Headers
              Result[i,1]:= DecodeText(ResponseResult[0],ResponseResult[1]);//body
             end;
           end;
    finally
      ResponseResult:= nil;
      PostData.Free;
    end;
  except
     on E: Exception do
     begin
       Result[0,0]:='Error: '+E.Message;
     end;
  end;
end;

function TwGet.ExecuteXML(aXMLDoc: String): ArrayArrayOfString;
var
  xdoc: TXMLDocument;                      // переменная документа
  Node: TDOMNode; // переменная узла документа
  i, k: integer;
  _Username, _Password: String;
  fOn, fPost: Boolean;
  _TimeOut: Longint;
  _XMLString: TStringStream;
begin
  //if not FileExists(FileListBox1.FileName) then begin exit; end;

  Actions.Clear;
  xdoc:=nil;
  _XMLString:= TStringStream.Create(aXMLDoc);

  try
    ReadXMLFile(xdoc,_XMLString);

    Node:= xdoc.FindNode('FORMAT');
    Node:= Node.FirstChild;
    _Username:= TDOMElement(Node).GetAttribute('username');
    _Password:= TDOMElement(Node).GetAttribute('password');
    TryStrToInt(TDOMElement(Node).GetAttribute('timeout'),_TimeOut);

    Node:= Node.NextSibling;

    if Assigned (Node) then
        with Node.ChildNodes do
          begin
            try
              for i:=0 to Count-1 do begin
                if Assigned(Item[i]) then
                begin
                    //if Item[i].HasAttributes then
                  if TDOMElement(Item[i]).GetAttribute('ON') = '1' then fOn:= true else fOn:= false;
                  if TDOMElement(Item[i]).GetAttribute('POST') = '1' then fPost:= true else fPost:= false;
                  Actions.PushBack(CreateAction(fOn,fPost,TDOMElement(Item[i]).GetAttribute('URL'),TDOMElement(Item[i]).GetAttribute('PARAMS'),_Username, _Password,TDOMElement(Item[i]).GetAttribute('LOCALFILENAME'),_TimeOut));
                end;
              end;

            finally
              Free;
            end;
          end;

    Result:= LoadFile();
  finally
    Node.Free;
    xdoc.Free;
    _XMLString.Free;
  end;
end;

end.

