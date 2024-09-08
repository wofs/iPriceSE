unit wDBTreeU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,ExtCtrls, StdCtrls, ComCtrls, Dialogs, Controls, Forms, Menus,
  wBaseU, wLogU, wTypesU
  ;

type

  { TTreeData }

  TTreeData = class
  private
    fID: integer;
    FLastItem: boolean;
    fTimeStamp: TDateTime;
    fIdFormat: integer;
  public
    constructor Create(aID: integer; const aTimeStamp: TDateTime = 0; const aIdFormat: integer = 0; aLastItem: boolean = false);
    property Value : integer read fID;
    property TimeStamp : TDateTime read fTimeStamp;
    property IdFormat : integer read fIdFormat;
    property LastItem : boolean read FLastItem;
  end;

  TwTree_FOT_Data = record
     IdFormat: integer;
     IdOwner: integer;
     TimeStamp: TDateTime;
     LastItem: boolean;
  end;

  TwTree_FOT_Data_Arr = array of TwTree_FOT_Data;

  { TwDBTree }

  TwDBTree = class
    private
      fEventBlock: boolean;
      fFirstFillTree: boolean;
      FonTransactionControl: TNotifyEvent;
      FShowChildrenItems: boolean;
      fTreeView: TTreeView;
      fCommit: boolean;

      fMultiSelect:boolean; // разрешает мультивыбор
      fTableName:string;
      fOrderBy:string;
      fTreeReadOnly: boolean;  // флаг
      fExpanded: boolean;
      fMoveTheNode: Boolean;
      fSourceNodeData: integer;
      fReceiverNodeData: integer;

      fBase: TwBase;

      fPopupMenu: TPopupMenu;   // всплывающее меню
      fPopupMenuAdd, fPopupMenuRename, fPopupMenuDelete: TMenuItem;
      fWhere: string;
      fWhereRoot: string;
      fWhereTime: string;
      fOwnerFieldValueArr: ArrayOfVariant;
      fParentDeletedNode: Integer;

      fOldValueString: String; // старое значение имени нода, для переименования
      fDbNoErrorFlag: Boolean; // флаг отсутсвтвия ошибок для переименования
      ssCtrlDown: Boolean;

      procedure onKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure Log(aText: string);
      procedure DataTreeClear;

      function ItemAdd(aidParentValue: integer; aNameValue: string): integer;
      function ItemRename(aNode: TTreeNode; var aText: string): boolean;
      function ItemDelete(aidNodeValue: integer): boolean;

      procedure onEdited(Sender: TObject; aNode: TTreeNode; var aText: string);
      procedure onEditingEnd(Sender: TObject; aNode: TTreeNode; aCancel: Boolean);
      procedure onChange(Sender: TObject; Node: TTreeNode);
      procedure onClick(Sender: TObject);
      procedure onDragDrop(Sender, Source: TObject; X, Y: Integer);
      procedure onDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
      function DragDrop(_Sender, _Source: TObject; _X, _Y: Integer):boolean;
      procedure OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

      procedure SetMultiSelect(AValue: boolean);

      function NodeFromDataInt(aNodeDataInt:integer):TTreeNode;
      function WhoParent(aID: integer): integer;// кто родитель?
      function HaveChildren(aID: integer): boolean;// есть детки?

    protected
      procedure FillBaseRecurse(const _dbArr: ArrayOfArrayVariant);
    public
      constructor Create(aBase: TwBase; aTreeView: TTreeView; aTableName, aOrderBy:string; aOwnerFieldValueArr: ArrayOfVariant);
      destructor Destroy();

      procedure Fill();
      procedure Fill(aTimeStamps: TwTree_FOT_Data_Arr);

      function SelectedItems(aAllData: boolean): TwTree_FOT_Data_Arr; // массив выбранных узлов
      function SelectedItems: ArrayOfInteger; // массив выбранных узлов

      function BreadCrumbs(aID:integer):string; // выводит родителей строкой

      procedure FindNodeWithDataInt(aNodeDataInt:integer);
      function GetNodeData(aNodeIndex: integer):integer;

      procedure onTreePopupMenuAdd(Sender: TObject);
      procedure onTreePopupMenuRename(Sender: TObject);
      procedure onTreePopupMenuDelete(Sender: TObject);

      property MultiSelect: boolean write SetMultiSelect;
      property Where: string write fWhere;
      property WhereTime: string write fWhereTime;
      property WhereRoot: string write fWhereRoot;
      property SetOwner: variant write fOwnerFieldValueArr[1];
      property Tree: TTreeView read fTreeView;
      property FirstFillTree: boolean read fFirstFillTree write fFirstFillTree;
      property PopupMenu: TPopupMenu read fPopupMenu write fPopupMenu;

      property Expanded: boolean read fExpanded write fExpanded;
      property ShowChildrenItems: boolean read FShowChildrenItems write FShowChildrenItems;
      property OrderBy: string read fOrderBy write fOrderBy;
      property MoveTheNode: boolean read fMoveTheNode write fMoveTheNode;
      property ReceiverNodeData: integer read fReceiverNodeData write fReceiverNodeData;
      property EventBlock: boolean read fEventBlock write fEventBlock;
      property onTransactionControl: TNotifyEvent read FonTransactionControl write FonTransactionControl;

  end;

implementation

{ TTreeData }

constructor TTreeData.Create(aID: integer; const aTimeStamp: TDateTime; const aIdFormat: integer; aLastItem: boolean);
begin
  fID:= aID;
  fTimeStamp:= aTimeStamp;
  fIdFormat:= aIdFormat;
  fLastItem:= aLastItem;
end;



{ TwDBTree }

procedure TwDBTree.Log(aText: string);
begin
    if __onLog and Assigned(__Log) then
     wLog('[Tree] ', aText);
end;

procedure TwDBTree.onKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  ssCtrlDown:=  (ssCtrl in Shift);
end;

procedure TwDBTree.DataTreeClear;
var
  i: Integer;
begin
  try
    with fTreeView do
    begin

       for  i:=Items.Count-1 downto 0 do
       begin
            TTreeData(Items[i].Data).Free;
       end;
    end;
  except
        on E: Exception do
    begin
      Log('Ошибка [DataTreeClear]: "' + E.Message + '"');
      __Log.SaveLogError(E);
      raise;
    end;
  end;
end;

function TwDBTree.ItemAdd(aidParentValue: integer; aNameValue: string): integer;
begin
 try
  result:=-1;
    if fTreeView.Items.Count = 0 then aidParentValue:=0;

   try
    fEventBlock:= true;
     if Length(fOwnerFieldValueArr)>0 then
      result:= fBase.SQLInsert(fTableName,['IDPARENT','NAME',fOwnerFieldValueArr[0],'FTIMESTAMP'],[aidParentValue,aNameValue,fOwnerFieldValueArr[1],now],fCommit)
     else
      result:= fBase.SQLInsert(fTableName,['IDPARENT','NAME','FTIMESTAMP'],[aidParentValue,aNameValue,now],fCommit)

  finally
    fEventBlock:= false;
  end;

  except
        on E: Exception do
    begin
      Log('Ошибка [ItemAdd]: "' + E.Message + '"');
      __Log.SaveLogError(E);
      result:=-1;
      raise;
    end;
  end;
end;

function TwDBTree.ItemRename(aNode: TTreeNode; var aText: string): boolean;
begin
 try
     result:=false;
     try
       EventBlock:= true;
       result:= fBase.SQLUpdate(fTableName,['NAME'],[aText],'ID='+IntToStr(TTreeData(aNode.Data).Value),fCommit);
     finally
       EventBlock:= false;
     end;

     if fTreeView.SortType <> stNone then
        fTreeView.AlphaSort;
  except
        on E: Exception do
    begin
      Log('Ошибка [ItemRename]: "' + E.Message + '"');
      __Log.SaveLogError(E);
      result:=false;
      raise;
    end;
  end;
end;

function TwDBTree.ItemDelete(aidNodeValue: integer): boolean;
begin
 try
   result:=false;
   fParentDeletedNode:= WhoParent(aidNodeValue);
 try
   fEventBlock:= true;
   result:= fBase.SQLDelete(fTableName,'ID = '+IntToStr(aidNodeValue),fCommit);
 finally
   fEventBlock:= false;
 end;
 except
   on E: Exception do
   begin
     Log('Ошибка [ItemDelete]: "' + E.Message + '"');
     ShowMessage('Ошибка при удалении узла! Возможно есть связанные списки! Дальнейшее удаление отменено.');
     result:=false;
     raise;
   end;
 end;
end;

procedure TwDBTree.onTreePopupMenuAdd(Sender: TObject);
var
  _IDParent, _NewID: Integer;
begin
 try
 Log('Добавляем новый узел...');

  if Length(fTreeView.Selected.Text) = 0 then exit;

   if fTreeView.Items.Count = 0 then _IDParent:=0 else _IDParent:= TTreeData(fTreeView.Selected.Data).Value;

   _NewID:=ItemAdd(_IDParent,'Новая запись');

   if Assigned(onTransactionControl) then onTransactionControl(self);

   Fill();
   FindNodeWithDataInt(_NewID);

 Log('Добавление нового узла успешно завершено.');
 except
   on E: Exception do
   begin
       Log('Ошибка [ItemAdd]: "' + E.Message + '"');
       Log('Сбой добавления узла.');
       __Log.SaveLogError(E);
    end;
 end;
end;

procedure TwDBTree.onTreePopupMenuRename(Sender: TObject);
begin
    if (fTreeVIew.Selected.Level <> 0) and (fTreeVIew.SelectionCount=1) then
       begin
         fTreeView.ReadOnly:=false;
         fTreeVIew.Selected.EditText;
       end;
end;

procedure TwDBTree.onTreePopupMenuDelete(Sender: TObject);
var
  _selCount: Cardinal;
  _arr: ArrayOfInteger;
  i: Integer;
begin
 try

   if (fTreeVIew.Selected.Level = 0) then exit;

       _selCount:= fTreeVIew.SelectionCount;

       if _selCount>1 then
          begin
            if MessageDlg('Выбрано несколько узлов дерева удалить их все? ВНИМАНИЕ! Связанные позиции тоже будут удалены!',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;
          end else
            if MessageDlg('Удалить узел дерева? ВНИМАНИЕ! Связанные позиции тоже будут удалены! ['+fTreeView.Selected.Text+'] ',mtConfirmation, mbOKCancel, 0) = mrCancel then exit;

       SetLength(_arr, _selCount);
       Log('Заполняем массив выбранными узлами.');
       _arr:= SelectedItems;
       Log('Количество выбранных узлов: '+IntToStr(Length(_arr)));
       Log('Список выбранных узлов:');
       for i:=0 to Length(_arr)-1 do
           Log(IntToStr(_arr[i]));
       Log('Конец списка.');

       screen.Cursor:= crSQLWait;


       for i:=0 to Length(_arr)-1 do
       begin
           Log('Удаляем узел:'+IntToStr(_arr[i]));

           if ItemDelete(_arr[i]) then
             begin
               Log('Удаление узла успешно завершено.');
             end else
             begin
               Log('Сбой при удалении узла!');
               exit;
             end;
       end;

   if Assigned(onTransactionControl) then onTransactionControl(self);


   Fill();
   if _selCount>1 then
      begin
        fTreeView.Items[0].Selected:=true;
        fTreeView.SetFocus;
      end
     else
        FindNodeWithDataInt(fParentDeletedNode);

 screen.Cursor:= crDefault;
 except
   on E: Exception do
   begin
       screen.Cursor:= crDefault;
       Log('Ошибка [_onTreePopupMenuDelete]: "' + E.Message + '"');
       Log('Сбой при удалении узла!');
       __Log.SaveLogError(E);
    end;
 end;
end;

procedure TwDBTree.onEdited(Sender: TObject; aNode: TTreeNode; var aText: string);
begin
 fOldValueString:=aNode.Text;
    Log('Переименовываем узел дерева... ['+fOldValueString+'->'+aText+']');
    if fOldValueString = aText then
       begin
         Log('Переименование узла отменено - названия идентичны.');
         fDbNoErrorFlag:=true;
         exit;
       end;
    fDbNoErrorFlag:= ItemRename(aNode,aText);
end;

procedure TwDBTree.onEditingEnd(Sender: TObject; aNode: TTreeNode; aCancel: Boolean);
var
  _NodeDataInt: integer;
begin
    if fDbNoErrorFlag then
       begin
         Log('Переименование узла успешно завершено.');
         _NodeDataInt:=TTreeData(aNode.Data).Value;
         Fill();
         FindNodeWithDataInt(_NodeDataInt);
         if Assigned(onTransactionControl) then onTransactionControl(self);
       end else
       begin
         Log('Сбой БД при переименовании узла дерева.');

         aNode.Text:=fOldValueString;
       end;
    fTreeView.ReadOnly:=fTreeReadOnly;
    fOldValueString:='';
end;

procedure TwDBTree.onChange(Sender: TObject; Node: TTreeNode);
begin
 with TTreeView(Sender) do
  begin
    if Selected.Level =0 then
       ReadOnly:=true else
       ReadOnly:=fTreeReadOnly;
  end;
end;

procedure TwDBTree.onClick(Sender: TObject);
begin
    if (fTreeView.Items.Count =0) or (fTreeView.Items[0].Count=0) then exit;
    //if fTreeView.SelectionCount>1 then
    //SetStatus('Множественный выбор') else
    //SetStatus('Выбрано: '+fTreeView.Selected.Text);
end;

procedure TwDBTree.onDragDrop(Sender, Source: TObject; X, Y: Integer);
begin
 if not ssCtrlDown and (Source is TTreeView) then exit;
 fMoveTheNode:=true;
 Log('Перемещение...');
  if DragDrop(Sender,Source,X,Y) then // перемещение узла
    begin
      Log('[_onDragDrop] '+'Перемещение узла успешно завершено.');
    end else
    begin
      Log('[_onDragDrop] '+'Сбой БД при перемещении узла дерева.');
    end;

 fMoveTheNode:=false;
end;

procedure TwDBTree.onDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
var
  Node: TTreeNode;
begin
   Accept:=true;
  if not (Source is TTreeView) then
    begin

   Node:= fTreeView.GetNodeAt(X, Y);
   if Node = nil then exit;
   if  fTreeView.Selected<>Node then
     begin
        fTreeView.ClearSelection(true);
        fTreeView.Selected:=Node;
     end;
    end;
end;

function TwDBTree.DragDrop(_Sender, _Source: TObject; _X, _Y: Integer): boolean;
var
   AnItem: TTreeNode;
   AttachMode: TNodeAttachMode;
   HT: THitTests;

begin
   result:=false;
   if (_Source is TTreeView) then
   begin
       if (fTreeView.Selected = nil) or (fTreeView.Selected.Level=0) then begin result:=false; exit; end;
       HT := fTreeView.GetHitTestInfoAt(_X, _Y) ;
       AnItem := fTreeView.GetNodeAt(_X, _Y) ;
       if AnItem = nil then
       begin
            result:=false; exit;
       end;

       fReceiverNodeData:=TTreeData(AnItem.Data).Value;

       fSourceNodeData:=TTreeData(fTreeView.Selected.Data).Value;

       if fSourceNodeData = fReceiverNodeData then begin result:=false; exit; end; // предотвращаем перемещение самого в себя

       Log('Перемещаем узел '+IntTOStr(fSourceNodeData)+' -> '+IntToStr(fReceiverNodeData));

       if WhoParent(fReceiverNodeData) = fSourceNodeData then
          begin
               result:=false;
               exit;
          end;

      try
          fBase.SQLUpdate(fTableName,['IDPARENT'],[fReceiverNodeData],'ID='+IntToStr(fSourceNodeData),fCommit);
       except
             on E: Exception do
         begin
           Log('Ошибка [DragDrop]: "' + E.Message + '"');
           result:=false;
           raise;
           exit;
         end;
       end;

       if (HT - [htOnItem, htOnIcon, htNowhere, htOnIndent] <> HT) then
       begin
         if (htOnItem in HT) or
            (htOnIcon in HT) then
             AttachMode := naAddChild
         else if htNowhere in HT then
             AttachMode := naAdd
         else if htOnIndent in HT then
             AttachMode := naInsert;
         fTreeView.Selected.MoveTo(AnItem, AttachMode) ;
       end;

       // выделяем перемещенный узел
       FindNodeWithDataInt(fSourceNodeData);
       result:=true;

       if Assigned(onTransactionControl) then onTransactionControl(self);

   end else
   begin
     // драг из другого элемента
      AnItem := fTreeView.GetNodeAt(_X, _Y) ;
      if AnItem = nil then
      begin
           result:=false; exit;
      end;
      fReceiverNodeData:=TTreeData(AnItem.Data).Value;
   end;
end;

procedure TwDBTree.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 ssCtrlDown:= (ssCtrl in Shift);
end;

procedure TwDBTree.SetMultiSelect(AValue: boolean);
begin
  fMultiSelect:= AValue;
  fTreeView.MultiSelect:= fMultiSelect;
end;

function TwDBTree.NodeFromDataInt(aNodeDataInt: integer): TTreeNode;
var
  _Noddy: TTreeNode;
  _Searching: Boolean;
begin
 try
 result:=nil;
 if fTreeView.Items.Count = 0 then exit;
     _Noddy := fTreeView.Items[0];
     _Searching := true;
     while (_Searching) and (_Noddy <> nil) do
     begin
       if TTreeData(_Noddy.Data).Value = aNodeDataInt then
       begin
         _Searching := False;
         result:=_Noddy
       end
       else
       begin
         _Noddy := _Noddy.GetNext
       end;
     end;
 except
       on E: Exception do
   begin
     Log('Ошибка [NodeFromDataInt]: "' + E.Message + '"');
     __Log.SaveLogError(E);
     raise;
   end;
 end;
end;

function TwDBTree.WhoParent(aID: integer): integer;
var
  _dbArr: ArrayOfArrayVariant;
begin
 try
    result:=-1;
    _dbArr:=fBase.SQLReadArr(fTableName,['IDPARENT'],'ID='+IntToStr(aID),fOrderBy);
    if Length(_dbArr)>0 then result:= integer(_dbArr[0,0]) else
    result:=-1;
    _dbArr:=nil;
except
      on E: Exception do
  begin
    result:=-1;
    raise;
  end;
end;
end;

function TwDBTree.HaveChildren(aID: integer): boolean;
begin
 result:= true;
  try
    if Length(fBase.SQLReadArr(fTableName,['ID'],'IDPARENT='+IntToStr(aID),fOrderBy))>0 then result:= true else result:=false;
 except
       on E: Exception do
   begin
     result:=true;
     raise;
   end;
 end;
end;

procedure TwDBTree.FillBaseRecurse(const _dbArr: ArrayOfArrayVariant);
var
  iTree: Integer;
  i: Integer;
  _Node: TTreeNode;
begin
  for i := 0 to High(_dbArr) do // перебираем в цикле
  begin
    iTree := 0;
    while iTree < Tree.Items.Count do
    begin
      //Log(IntTOStr(i)+' | '+IntToStr(iTree)+' | '+string(_dbArr[i, 2])+' | '+IntToStr(TTreeData(Tree.Items[iTree].Data).Value)+' | '+string(_dbArr[i, 1]));

      if TTreeData(Tree.Items[iTree].Data).Value = integer(_dbArr[i, 1]) then
      begin
        _Node:= nil;
        _Node:= NodeFromDataInt(integer(_dbArr[i, 0]));
        if not Assigned(_Node) then
          Tree.Items.AddChildObject(Tree.Items.Item[iTree],
          string(_dbArr[i, 2]), TTreeData.Create((integer(_dbArr[i, 0]))));  // добавляем дочерние объекты
        break;
      end else
        Inc(iTree);
    end;
  end;
end;

function TwDBTree.BreadCrumbs(aID: integer): string;
var
  _Node: TTreeNode;
begin
 result:='';
 try
   _Node:=NodeFromDataInt(aID);

   while(_Node<>nil) do begin
     result:='\'+_Node.Text+result;
     _Node:=_Node.Parent;
   end;
 except
    raise;
 end;
end;

procedure TwDBTree.FindNodeWithDataInt(aNodeDataInt: integer);
var
  _Noddy: TTreeNode;
  _Searching: Boolean;
begin
 try
     _Noddy := fTreeView.Items[0];

     _Searching := true;
     while (_Searching) and (_Noddy <> nil) do
     begin
       if TTreeData(_Noddy.Data).Value = aNodeDataInt then
       begin
         _Searching := False;
         fTreeView.ClearSelection(true);
         fTreeView.Selected := _Noddy;
  //       fTreeView.SetFocus;
       end
       else
       begin
         _Noddy := _Noddy.GetNext
       end;
     end;
 except
       on E: Exception do
   begin
     Log('Ошибка [FindNodeWithDataInt]: "' + E.Message + '"');
     __Log.SaveLogError(E);
     raise;
   end;
 end;
end;

function TwDBTree.GetNodeData(aNodeIndex: integer): integer;
begin
  Result:= TTreeData(fTreeView.Items[aNodeIndex].Data).Value;
end;

constructor TwDBTree.Create(aBase: TwBase; aTreeView: TTreeView; aTableName, aOrderBy: string; aOwnerFieldValueArr: ArrayOfVariant);
begin
  fFirstFillTree:= true;
  fBase:= aBase;

  if fBase.LongTransaction then fCommit:= false else fCommit:= true;

  fTreeView:= aTreeView;
  fTableName:= aTableName;
  fOrderBy:= aOrderBy;
  fOwnerFieldValueArr:= aOwnerFieldValueArr;
  FShowChildrenItems:= true;
  EventBlock:= false;

    if (fTreeView.PopupMenu = nil) and not fTreeView.ReadOnly then
    begin
       fPopupMenu:= TPopupMenu.Create(fTreeView);

       fPopupMenuAdd:= TMenuItem.Create(fPopupMenu);
       fPopupMenuAdd.Caption:= 'Добавить';
       fPopupMenuAdd.OnClick:=@onTreePopupMenuAdd;
       fPopupMenuAdd.ShortCut:=45;  //VK_INSERT in LCLType
       fPopupMenu.Items.Add(fPopupMenuAdd);

       fPopupMenuRename:= TMenuItem.Create(fPopupMenu);
       fPopupMenuRename.Caption:= 'Переименовать';
       fPopupMenuRename.OnClick:=@onTreePopupMenuRename;
       fPopupMenuRename.ShortCut:=113;   //VK_F2 in LCLType
       fPopupMenu.Items.Add(fPopupMenuRename);

       fPopupMenuDelete:= TMenuItem.Create(fPopupMenu);
       fPopupMenuDelete.Caption:= 'Удалить';
       fPopupMenuDelete.OnClick:=@onTreePopupMenuDelete;
       fPopupMenuDelete.ShortCut:=46; //VK_DELETE in LCLType
       fPopupMenu.Items.Add(fPopupMenuDelete);

       fTreeView.PopupMenu:=fPopupMenu;
    end;

   if not fTreeView.ReadOnly then
   begin
     if not Assigned(fTreeView.OnEdited) then
          fTreeView.OnEdited:= @onEdited;

     if not Assigned(fTreeView.OnEditingEnd) then
          fTreeView.OnEditingEnd:= @OnEditingEnd;

     if not Assigned(fTreeView.OnChange) then
          fTreeView.OnChange:= @onChange;

     if not Assigned(fTreeView.OnClick) then
          fTreeView.OnClick:= @onClick;

     if not Assigned(fTreeView.OnDragDrop) then
          fTreeView.OnDragDrop:= @onDragDrop;

     if not Assigned(fTreeView.onDragOver) then
          fTreeView.onDragOver:= @onDragOver;

     if not Assigned(fTreeView.OnKeyDown) then
       fTreeView.OnKeyDown:= @OnKeyDown;

     if not Assigned(fTreeView.OnKeyUp) then
       fTreeView.OnKeyUp:= @onKeyUp;
   end;

   fTreeView.ReadOnly:= true;
   fTreeReadOnly:= true;
end;

destructor TwDBTree.Destroy();
begin
  DataTreeClear;
end;

procedure TwDBTree.Fill();
var
  _Where, _WhereRoot: String;
  _dbArr: ArrayOfArrayVariant;
  _Node: TTreeNode;
begin
  // Корневой узел (Root), должен быть первым в выборке Query
     screen.Cursor:= crSQLWait;

     try
       Log('Заполняем дерево... ['+fTreeView.Name+']');

      if Length(fWhere)>0 then _Where:=fWhere
      else
       _Where:=fBase.PrepareIDOwnerWhere(_Where, fOwnerFieldValueArr); // подготавливаем дополнительную фильтрацию по владельцу

      _WhereRoot:='';
      if Length(fWhereRoot)>0 then
      _WhereRoot:= Format(fWhereRoot,[integer(fOwnerFieldValueArr[1])]);

      _dbArr:=fBase.SQLReadArr(fTableName,['ID','IDPARENT','NAME'],_Where+fWhereTime+_WhereRoot,fOrderBy);

     if Assigned(_dbArr) then
     begin
       with fTreeView do
       begin
         BeginUpdate;

         DataTreeClear;
         Items.Clear;

         Items.AddObject(nil, _dbArr[0, 2], TTreeData.Create((integer(_dbArr[0, 0]))));
         // добавляем корневой объект

         // первый проход
         FillBaseRecurse(_dbArr);

         // второй проход
         FillBaseRecurse(_dbArr);

        // if Expanded then
        Items[0].Expand(fExpanded);

        Items[0].Selected:=true;

         if fTreeView.SortType <> stNone then
               fTreeView.AlphaSort;

         EndUpdate;
       end;
      _dbArr:= nil;
      Log('Заполнение успешно завершено');
      screen.Cursor:= crDefault;
     end;
     except
           on E: Exception do
       begin
          screen.Cursor:= crDefault;
          Log('Ошибка [FillTree]: "' + E.Message + '"');
          __Log.SaveLogError(E);
         raise;
       end;
   end;
end;

procedure TwDBTree.Fill(aTimeStamps: TwTree_FOT_Data_Arr);
 var
   _Where, _WhereRoot, _Date: String;
   _dbArrOwners, _dbArrFormats: ArrayOfArrayVariant;
   iOwners, iFormats, iTree, iTimeStamps: Integer;
   _NoddyOwner, _NoddyFormat, _NoddyDateTime, _NoddyDate: TTreeNode;
 begin
   // Корневой узел (Root), должен быть первым в выборке Query
      try
        Log('Заполняем дерево... ['+fTreeView.Name+']');

       if Length(fWhere)>0 then _Where:=fWhere
       else
        _Where:=fBase.PrepareIDOwnerWhere(_Where, fOwnerFieldValueArr); // подготавливаем дополнительную фильтрацию по владельцу

       _WhereRoot:='';
       if Length(fWhereRoot)>0 then
       _WhereRoot:= Format(fWhereRoot,[integer(fOwnerFieldValueArr[1])]);

       _dbArrOwners:=fBase.SQLReadArr(fTableName,['ID','IDPARENT','NAME'],_Where+fWhereTime+_WhereRoot,fOrderBy);
       _dbArrFormats:=fBase.SQLReadArr('FORMATS',['ID','IDOWNER','NAME','IDFMTS_CATEGORY'],'','IDOWNER,NAME');


      if Assigned(_dbArrOwners) then
      begin
        with fTreeView do
        begin
          BeginUpdate;

          DataTreeClear;
          Items.Clear;

          Items.AddObject(nil, _dbArrOwners[0, 2], TTreeData.Create((integer(_dbArrOwners[0, 0]))));
          // добавляем корневой объект


          for iOwners := 0 to Length(_dbArrOwners) - 1 do // перебираем в цикле
          begin

            iTree := 0;
            while iTree < Items.Count do
            begin

              if TTreeData(Items[iTree].Data).Value = integer(_dbArrOwners[iOwners, 1]) then
              begin
                _NoddyOwner:= Items.AddChildObject(
                         Items.Item[iTree],
                         string(_dbArrOwners[iOwners, 2]),
                         TTreeData.Create(
                           integer(_dbArrOwners[iOwners, 0])
                         )
                       );
                 _NoddyOwner.ImageIndex:=0;

                     for iFormats:=0 to High(_dbArrFormats) do
                         begin
                             if (integer(_dbArrOwners[iOwners, 0]) = integer(_dbArrFormats[iFormats, 1])) and (integer(_dbArrFormats[iFormats, 3]) = 1) then
                             begin
                                _NoddyFormat:= Items.AddChildObject(
                                 _NoddyOwner,
                                 _dbArrFormats[iFormats, 2],
                                 TTreeData.Create(
                                   integer(_dbArrFormats[iFormats, 1]),
                                   0,
                                   _dbArrFormats[iFormats, 0]
                               )
                               );
                               _NoddyFormat.ImageIndex:=1;
                                _Date:='';
                                for iTimeStamps:=0 to High(aTimeStamps) do
                                    begin
                                      if aTimeStamps[iTimeStamps].IdFormat = integer(_dbArrFormats[iFormats, 0]) then
                                      begin
                                        if _Date <> FormatDateTime('yyyy mmm',aTimeStamps[iTimeStamps].TimeStamp) then
                                        begin
                                          _Date:= FormatDateTime('yyyy mmm',aTimeStamps[iTimeStamps].TimeStamp);
                                          _NoddyDate:= Items.AddChildObject(
                                            _NoddyFormat,
                                            _Date,
                                            TTreeData.Create(
                                              integer(aTimeStamps[iTimeStamps].IdOwner),
                                              0,
                                              -1
                                              )
                                              );
                                          _NoddyDate.ImageIndex:=2;
                                        end;

                                        if _Date = FormatDateTime('yyyy mmm',aTimeStamps[iTimeStamps].TimeStamp) then
                                              _NoddyDateTime:= Items.AddChildObject(
                                                _NoddyDate,
                                                DateTimeToStr(aTimeStamps[iTimeStamps].TimeStamp),
                                                TTreeData.Create(
                                                  integer(aTimeStamps[iTimeStamps].IdOwner),
                                                  aTimeStamps[iTimeStamps].TimeStamp,
                                                  0,
                                                  true
                                                  )
                                                  );

                                        _NoddyDateTime.ImageIndex:=3;
                                      end;
                                    end;
                             end;
                         end;

              end;

              Inc(iTree);
            end;

          end;
         // if Expanded then
         Items[0].Expand(fExpanded);
         Items[0].Selected:=true;

          EndUpdate;
        end;
       _dbArrOwners:= nil;
       _dbArrFormats:= nil;
       Log('Заполнение успешно завершено');
      end;
      except
            on E: Exception do
        begin
           Log('Ошибка [FillTree]: "' + E.Message + '"');
           __Log.SaveLogError(E);
          raise;
        end;
    end;
end;

function TwDBTree.SelectedItems: ArrayOfInteger;
var
  _NodesList: TList;
  i: Integer;

    procedure ShowChilds(aNode:TTreeNode);
    begin
     aNode := aNode.GetFirstChild;
     while Assigned(aNode) do
     begin
       _NodesList.Add(TTreeData.Create(TTreeData(aNode.Data).Value));
       ShowChilds(aNode);
       aNode := aNode.GetNextChild(aNode)
     end;
    end;

begin

 try
  result:=nil;
 if fTreeView.SelectionCount = 0 then
   begin
     exit;
   end;
    _NodesList:= TList.Create;

   for i:=0 to fTreeView.Items.Count-1 do
   begin

     if  fTreeView.Items[i].Selected then
     begin
       _NodesList.Add(TTreeData.Create(TTreeData(fTreeView.Items[i].Data).Value));
       if ShowChildrenItems then ShowChilds(fTreeView.Items[i]);
     end;

   end;

   SetLength(result,_NodesList.Count);

   for i:=0 to _NodesList.Count-1 do begin
       result[i]:= TTreeData(_NodesList.Items[i]).Value;
       TTreeData(_NodesList.Items[i]).Free;
   end;

   FreeAndNil(_NodesList);

   except
   FreeAndNil(_NodesList);
     raise;
   end;

end;

function TwDBTree.SelectedItems(aAllData: boolean): TwTree_FOT_Data_Arr;
var
  _NodesList: TList;
  i: Integer;
  _arr: ArrayOfInteger;

    procedure ShowChilds(aNode:TTreeNode);
    begin
     aNode := aNode.GetFirstChild;
     while Assigned(aNode) do
     begin
       _NodesList.Add(TTreeData.Create(TTreeData(aNode.Data).Value,TTreeData(aNode.Data).TimeStamp,TTreeData(aNode.Data).IdFormat,TTreeData(aNode.Data).LastItem));
       ShowChilds(aNode);
       aNode := aNode.GetNextChild(aNode)
     end;
    end;
begin
 if not aAllData then exit;


 try
  result:=nil;
 if fTreeView.SelectionCount = 0 then
   begin
     exit;
   end;
    _NodesList:= TList.Create;

   try
     for i:=0 to fTreeView.Items.Count-1 do
     begin
       if  fTreeView.Items[i].Selected then
       begin
         _NodesList.Add(TTreeData.Create(TTreeData(fTreeView.Items[i].Data).Value,TTreeData(fTreeView.Items[i].Data).TimeStamp,TTreeData(fTreeView.Items[i].Data).IdFormat,TTreeData(fTreeView.Items[i].Data).LastItem));
         ShowChilds(fTreeView.Items[i]);
       end;

     end;

     SetLength(result,_NodesList.Count);

     for i:=0 to _NodesList.Count-1 do begin
         result[i].IdOwner:= TTreeData(_NodesList.Items[i]).Value; //ID
         result[i].TimeStamp:= TTreeData(_NodesList.Items[i]).fTimeStamp;  //TEXTVALUE
         result[i].IdFormat:= TTreeData(_NodesList.Items[i]).IdFormat; //idFormat
         result[i].LastItem:= TTreeData(_NodesList.Items[i]).LastItem; //LastItem
         TTreeData(_NodesList.Items[i]).Free;
     end;
   finally
     FreeAndNil(_NodesList);
   end;

   except
     raise;
   end;

end;

end.

