unit wZipperU;

//
//   Degtyarev Alexander(c)2017 z-lib license
//   used Zipper
//

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LazUTF8, Forms, Dialogs,
  Zipper, LConvEncoding,LazFileUtils, FileUtil, wFuncU
  ;
type
    ArrayOfString = array of string;
    { TwZipper }

    TwZipper = class
    private
      FDecodeFileName: boolean;
      fFileListInZip: TStringList;
      function EndPathCP866ToUTF8(AText: string): string;

    public
      constructor Create ();
      destructor Destroy (); override;

      function ReadFileList(aZipFile: string): TStringList;
      function ExtractOneFile(aZipFile, aFileExtract, aUnpackPatch: string): string;
      procedure ExtractAllFiles(aZipFile, aUnpackPatch: string);
      procedure PackAllFiles(aZipFile, aRelativeDirectory: string);
      procedure PackFiles(aZipFile, RelativeDirectory, aMaskFiles: string);

      function ParseComboFileName(aComboFileName: string): ArrayOfString;
      function GetUnPackPath(const aDirectoryName: string='tmp'): string;

      property DecodeFileName: boolean read FDecodeFileName write FDecodeFileName default false;
    end;



implementation

function TwZipper.EndPathCP866ToUTF8(AText:string):string;
var
  c,i:integer;
  s,s1,s2,chr:string;
begin
  s:='';
  c:=UTF8Length(AText);
  for i:=c downto 1 do
  begin
       chr:=UTF8Copy(AText,i,1);
       if ((not(chr='/')) and (not(chr='\')))or(i=c) then
       begin
            s:=UTF8Copy(AText,i,1)+s;
       end
       else begin
            s:=UTF8Copy(AText,i,1)+s;
            break;
       end;
  end;
  dec(i);
  s1:=UTF8Copy(AText,1,i);
  s2:=CP866ToUTF8(s);
  Result:=s1+s2;
end;

function TwZipper.ReadFileList(aZipFile: string):TStringList;
var
   UnZipper: TUnZipper;
   i: Integer;
begin
  UnZipper      :=TUnZipper.Create;
  try
     UnZipper.FileName   := aZipFile;
     UnZipper.Examine;
     fFileListInZip.Clear;
     for i:=UnZipper.Entries.Count-1 downto 0 do
     begin
          if FDecodeFileName then
         fFileListInZip.Add(UTF8ToSys(EndPathCP866ToUTF8(UnZipper.Entries.Entries[i].ArchiveFileName))) else
         fFileListInZip.Add(UnZipper.Entries.Entries[i].ArchiveFileName);
     end;

  finally
     UnZipper.Free;
  end;

  Result:= fFileListInZip;
end;

function TwZipper.ExtractOneFile(aZipFile, aFileExtract, aUnpackPatch: string):string;
var
  UnZipper: TUnZipper; //PasZLib
  _Files: TStringList;
  _ArchiveFileName, _NewDiskFileName, _DiskFileName: string;
  i: Integer;
begin
  try
    UnZipper      :=TUnZipper.Create;
    try
       UnZipper.FileName   := aZipFile;
       UnZipper.OutputPath := aUnpackPatch;
       UnZipper.Examine;
       _Files:= TStringList.Create;
       if FDecodeFileName then
       _Files.Add(UTF8ToCP866(aFileExtract)) else
       _Files.Add(aFileExtract);

       UnZipper.UnZipFiles(_Files);
       Result:= aUnpackPatch+aFileExtract;

       for i:=UnZipper.Entries.Count-1 downto 0 do
       begin
            if FDecodeFileName then
          _ArchiveFileName:= UTF8ToSys(EndPathCP866ToUTF8(UnZipper.Entries.Entries[i].ArchiveFileName)) else
          _ArchiveFileName:= UnZipper.Entries.Entries[i].ArchiveFileName;

          _NewDiskFileName:= SysUtils.IncludeTrailingPathDelimiter(aUnpackPatch)+_ArchiveFileName;
          _DiskFileName:= SysUtils.IncludeTrailingPathDelimiter(aUnpackPatch)+UnZipper.Entries.Entries[i].DiskFileName;
            if FileExists(_DiskFileName) then
               RenameFile(_DiskFileName, _NewDiskFileName);

            //else
            //     if DirectoryExistsUTF8(_DiskFileName) then
            //     begin
            //       _DiskFileName:=SysUtils.IncludeTrailingPathDelimiter(_DiskFileName);
            //       _NewDiskFileName:=SysUtils.IncludeTrailingPathDelimiter(_NewDiskFileName);
            //       RenameFile(_DiskFileName, _NewDiskFileName);
            //     end;
       end;

       //UnZipper.UnZipAllFiles;
    finally
       _Files.Free;
       UnZipper.Free;
    end;
  except
     raise;
  end;

end;

procedure TwZipper.ExtractAllFiles(aZipFile, aUnpackPatch: string);
var
  UnZipper: TUnZipper; //PasZLib
begin

  try
    UnZipper      :=TUnZipper.Create;
    try
       UnZipper.FileName   := aZipFile;
       UnZipper.OutputPath := aUnpackPatch;
       UnZipper.Examine;
       UnZipper.UnZipAllFiles;
    finally
       UnZipper.Free;
    end;
  except
     raise;
  end;

end;

procedure TwZipper.PackAllFiles(aZipFile, aRelativeDirectory: string);
var
  AZipper: TZipper;
  ZEntries : TZipFileEntries;
  aFileList: TStringList;
  szPathEntry: String;
  i: SizeInt;
begin
  AZipper := TZipper.Create;
  try
    try
      AZipper.Filename := aZipFile;
      //aRelativeDirectory:='C:\MyFolder\MyFolder\';
      AZipper.Clear;
      ZEntries := TZipFileEntries.Create(TZipFileEntry);
      // Verify valid directory
      If DirPathExists(aRelativeDirectory) then
      begin
        i:=RPosUTF8(PathDelim,ChompPathDelim(aRelativeDirectory));
        szPathEntry:=LeftStr(aRelativeDirectory,i);

        // Use the FileUtils.FindAllFiles function to get everything (files and folders) recursively
        aFileList:=TstringList.Create;
        try
          FindAllFiles(aFileList, aRelativeDirectory);
          for i:=0 to aFileList.Count -1 do
          begin
            // Make sure the aRelativeDirectory files are not in the root of the ZipFile
            ZEntries.AddFileEntry(aFileList[i],CreateRelativePath(aFileList[i],szPathEntry));
          end;
        finally
          aFileList.Free;
        end;
      end;
      if (ZEntries.Count > 0) then
        AZipper.ZipFiles(ZEntries);
      except
        On E: EZipError do
          E.CreateFmt('Пакет не может быть создан %sПричина: %s', [LineEnding, E.Message])
      end;
  finally
    FreeAndNil(ZEntries);
    AZipper.Free;
  end;
end;

procedure TwZipper.PackFiles(aZipFile, RelativeDirectory, aMaskFiles: string);
var
  AZipper: TZipper;
  ZEntries : TZipFileEntries;
  aFileList: TStringList;
  szPathEntry, aRelativeDirectory: String;
  i: SizeInt;
begin
  AZipper := TZipper.Create;
  try
    try
      AZipper.Filename := aZipFile;
      //aRelativeDirectory:='C:\MyFolder\MyFolder\';
      aRelativeDirectory:= RelativeDirectory;
      AZipper.Clear;
      ZEntries := TZipFileEntries.Create(TZipFileEntry);
      // Verify valid directory
      If DirPathExists(aRelativeDirectory) then
      begin
        i:=RPosUTF8(PathDelim,ChompPathDelim(aRelativeDirectory));
        szPathEntry:=LeftStr(aRelativeDirectory,i);

        // Use the FileUtils.FindAllFiles function to get everything (files and folders) recursively
        aFileList:=TstringList.Create;
        try
          FindAllFiles(aFileList, aRelativeDirectory,aMaskFiles);
          for i:=0 to aFileList.Count -1 do
          begin
            // Make sure the aRelativeDirectory files are not in the root of the ZipFile
            ZEntries.AddFileEntry(aFileList[i],CreateRelativePath(aFileList[i],szPathEntry));
          end;
        finally
          aFileList.Free;
        end;
      end;
      if (ZEntries.Count > 0) then
      begin
        AZipper.ZipFiles(ZEntries);
        AZipper.SaveToFile(RelativeDirectory+aZipFile);
        DeleteFile(aZipFile);
      end;
      except
        On E: EZipError do
          E.CreateFmt('Пакет не может быть создан %sПричина: %s', [LineEnding, E.Message])
      end;
  finally
    FreeAndNil(ZEntries);
    AZipper.Free;
  end;

end;

function TwZipper.ParseComboFileName(aComboFileName: string):ArrayOfString;
var
  _Pos: integer;
begin
  SetLength(Result,2);
  _Pos:= UTF8Pos('|',aComboFileName);
  if _Pos>0 then
  begin
   Result[0]:= UTF8Copy(aComboFileName,1,_Pos-1); // path to zip
   Result[1]:= UTF8Copy(aComboFileName,_Pos+1,Length(aComboFileName)); // filename in zip
  end else
  begin
   Result:=nil;
  end;
end;

function TwZipper.GetUnPackPath(const aDirectoryName: string = 'tmp'):string;
begin
  Result:= includeTrailingPathDelimiter(ExtractFileDir(Application.ExeName));
  Result:= Result+aDirectoryName;
  if not DirectoryExistsUTF8(Result) then ForceDirectoriesUTF8(Result);
end;

constructor TwZipper.Create();
begin
  fFileListInZip:= TStringList.Create;
end;

destructor TwZipper.Destroy();
begin
  fFileListInZip.Free;
  inherited Destroy;
end;

end.

