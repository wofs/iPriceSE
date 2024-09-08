unit pkgFmCatalogU;

// 2017 (c) Degtyarev A. A.
{
-= ЛИЦЕНЗИЯ =-
2017 (c) Дегтярев Александр Александрович
Земля, Россия.
wofs2@rambler.ru

 Исходный код и результат компиляции нераздельно принадлежат автору.
 Запрещена модификация, распространение исходного когда, а так же сборка приложения.
 Равно как и использование, а так же и распространение результирующего приложения без явного на то согласия автора.
}

{$mode objfpc}{$H+}

interface

uses
  SysUtils, Dialogs, Forms, Controls, ComCtrls, ExtCtrls, StdCtrls, DBGrids,
  Classes, Graphics, Grids, Menus, FmNomenclatureEditU, FmNomenclatureEditMassU, db,

  wLogU, wDBaseU, wDBTreeU, wDBGridU, wFuncU, wFormulaU;

type

  { TFmCatalog }

  TFmCatalog = class(TForm)
    cbPriceField: TComboBox;
    DBGridCatalogPrice: TDBGrid;
    edPriceSearch: TEdit;
    edMatchSearch: TEdit;
    gbGroupPrice: TGroupBox;
    gbGroupOwner: TGroupBox;
    ILtabs: TImageList;
    ILtoolbars: TImageList;
    ImageList16: TImageList;
    lbPriceSearch: TLabel;
    lbMatchSearch: TLabel;
    mNomCopy: TMenuItem;
    mNomSplit: TMenuItem;
    mNomGoToGroup: TMenuItem;
    mNomEdit: TMenuItem;
    mNomDelete: TMenuItem;
    mNomSootv: TMenuItem;
    mNomAdd: TMenuItem;
    pMain: TPanel;
    pcCatalog: TPageControl;
    mPriceGrid: TPopupMenu;
    pPriceSearch: TPanel;
    pMatchGroup: TPanel;
    pMatchList: TPanel;
    pMarchSearch: TPanel;
    pPriceGroup: TPanel;
    pPrice: TPanel;
    SpltPrice: TSplitter;
    SpltMatch: TSplitter;
    StringGrid5: TStringGrid;
    TabNomenclature: TTabSheet;
    TabMatching: TTabSheet;
    tbPrice: TToolBar;
    tbMatch: TToolBar;
    tbMatchBtnAdd: TToolButton;
    tbMatchBtnEdit: TToolButton;
    tbMatchBtnDelete: TToolButton;
    tbMatchBtnMatch: TToolButton;
    tbCatalogBtnAdd: TToolButton;
    tbCatalogBtnEdit: TToolButton;
    tbCatalogBtnDelete: TToolButton;
    tbCatalogBtnMatch: TToolButton;
    tbCatalogBtnCopy: TToolButton;
    tbCatalogBtnGoToGroup: TToolButton;
    ToolButton3: TToolButton;
    TreeGroupPrice: TTreeView;
    TreeGroupOwner: TTreeView;

    procedure cbPriceFieldChange(Sender: TObject);
    procedure edPriceSearchChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure pcCatalogChange(Sender: TObject);
    procedure tbCatalogBtnAddClick(Sender: TObject);
    procedure tbCatalogBtnDeleteClick(Sender: TObject);
    procedure tbCatalogBtnEditClick(Sender: TObject);
    procedure tbCatalogBtnGoToGroupClick(Sender: TObject);
    procedure tbCatalogBtnCopyClick(Sender: TObject);
    procedure TreeGroupOwnerChange(Sender: TObject; Node: TTreeNode);

  private
    FormIDent: string;
    OwnerID: integer;     // ID выбранного контрагента
    OwnerMainID: integer; // ID основного контрагента (к которому привязан каталог)
    _DBase: TwDBase;
    _DBGrid: TwDBGrid;
    _DBTreeGroupPrice, _DBTreeGroupOwner: TwDBTreeView;

    property wFormID: string read FormIDent write FormIDent;

  public
    procedure SetStatus(_Text:string);

    { public declarations }
  end;

var
  FmCatalog: TFmCatalog;
  IDint:integer;

implementation

{$R *.lfm}

{ TODO: Отображение соответствий }
{ TODO: Добавление товаров каталога }
{ TODO: Добавление соответствий }
{ TODO: Настройка формата }
{ TODO: Импорт прайс-листов }
{ TODO: Автоматический поиск соответствий }
{ TODO: Ценообразование }

{ TFmCatalog }

procedure TFmCatalog.SetStatus(_Text: string);
begin
     wStatus(wFormID,_Text,true);
end;

procedure TFmCatalog.FormCreate(Sender: TObject);
begin
      wFormID:=Self.Name;
      OwnerID:=0; // инициализируем.
      wLog('Catalog','Инициализация плагина... ['+wFormID+']');

      try

//        _wFormula: TwFormula;

        //DBase
        DBase.Add(TwDBase.Create(Sender)); // инициализация соединения с БД
        _DBase:= __DBase(wFormID);

        // создаем формулу для расчета цен
        wFormula.Add(TwFormula.Create(wFormID,(Sender as TComponent)));
        _DBase.DBFormula:= __wFormula(wFormID); // предаем ее на попечение БД

        OwnerMainID:= _DBase.ReadSetting(1); // считываем настройки - текущий основной прайс-лист

        cmbxFill(cbPriceField,_DBase.SQLReadDS('PRICEFIELD',['NAME','ID'],'FCLOSE=0','PRIORITY',Self),['NAME','ID']);

        //__wFormula(wFormID).Formula:= string(_DBase.SQLReadArr('PRICEFIELD',['FORMULA'],'ID='+IntToStr(cmbxSelectID(cbPriceField)),'')[0,0]);
        _DBase.DBFormula.Prepare(string(_DBase.SQLReadArr('PRICEFIELD',['FORMULA'],'ID='+IntToStr(cmbxSelectID(cbPriceField)),'')[0,0]));

        //DBTree
         DBTree.Add(TwDBTreeView.Create(wFormID, TreeGroupOwner, nil, true,'OWNER','IDPARENT,NAME',[])); // инициализация дерева
         _DBTreeGroupOwner:= __DBTree(wFormID,TreeGroupOwner);

         TreeGroupOwner.PopupMenu.Items.Clear;
         TreeGroupOwner.DragMode:=dmManual;
         with TreeGroupOwner.PopupMenu do
         begin
           Images:= ImageList16;
         end;

         DBTree.Add(TwDBTreeView.Create(wFormID, TreeGroupPrice, DBGridCatalogPrice, true,'GROUP-CATALOG','IDPARENT,NAME',['IDOWNER',OwnerMainID])); //
         _DBTreeGroupPrice:= __DBTree(wFormID,TreeGroupPrice);

         with TreeGroupPrice.PopupMenu do
         begin
           Images:= ImageList16;
           Items[0].ImageIndex:= 0;
           Items[1].ImageIndex:= 1;
           Items[2].ImageIndex:= 2;
         end;

         //DBGrid
         DBGrid.Add(TwDBGrid.Create(wFormID, DBGridCatalogPrice,TreeGroupPrice,'CATALOG',['ID','IDGROUPNOMENCLATURE','NAME','UNIT','PRICE','PN','PM','PD','PC','PK','STOCK','SCOD','LABEL'],['IDOWNER',OwnerMainID],['IDGROUPNOMENCLATURE','NAME'])); // инициализация DBGrid
         _DBGrid:= __DBGrid(wFormID,DBGridCatalogPrice);
         _DBase.CalculateField:='CPRICE';

         _DBTreeGroupPrice.FillTree(true); // заполнение дерева
         _DBTreeGroupOwner.FillTree(true); // заполнение дерева

         //__DBase(wFormID).CalculateField:='CPRICE';
         //__DBGrid(wFormID,DBGridCatalogPrice).FillGrid([]); // заполнение грида


        wLog('Catalog','Инициализация плагина успешно завершена.');
      except
        on E: Exception do
        begin
            SetStatus('Сбой инициализации плагина.');
            wLog('Catalog','Ошибка [FmCreate]: "' + E.Message + '"');
            wLog('Catalog','Сбой инициализации плагина.');
            ShowMessage('Ошибка [FmCreate]: "' + E.Message + '"');

         end;
      end;
end;

procedure TFmCatalog.cbPriceFieldChange(Sender: TObject);
var
  _BookMark: TBookMark;
begin
  _DBase.DBFormula.Prepare(string(_DBase.SQLReadArr('PRICEFIELD',['FORMULA'],'ID='+IntToStr(cmbxSelectID(cbPriceField)),'')[0,0]));

  with DBGridCatalogPrice.DataSource.DataSet do begin
     _BookMark:= Bookmark;
      close;
      open;
     Bookmark:= _BookMark;
  end;


end;

procedure TFmCatalog.edPriceSearchChange(Sender: TObject);
var
  _Edit:TEdit;
begin
     _Edit:=(Sender as TEdit);

  if  Length(_Edit.Text)>0 then _Edit.Color:=clSkyBlue else _Edit.Color:=clDefault;

  __DBGrid(wFormID,DBGridCatalogPrice).FilterString:=_Edit.Text; // указываем строку для поиска
  __DBGrid(wFormID,DBGridCatalogPrice).FillGrid([]); // заполнение грида
end;

procedure TFmCatalog.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  CloseAction := caFree;
end;

procedure TFmCatalog.FormDestroy(Sender: TObject);
var
  i: integer;
begin
      try
       wLog('Catalog','Выгрузка плагина...');

        // выгружаем подгруженные DBTree
    _DBTreeGroupOwner.Unload();
    _DBTreeGroupPrice.Unload();
        // выгружаем подгруженные DBGrid
    _DBGrid.Unload();

        // выгружаем DBase
    _DBase.Unload();

    cmbxClearData(cbPriceField);  // очищаем объекты комбобокса


      wLog('Catalog','Выгрузка плагина успешно завершена.');

      except
        on E: Exception do
        begin
            SetStatus('Сбой выгрузки плагина: Каталог.');
            wLog('Catalog','Ошибка [FmDestroy]: "' + E.Message + '"');
            wLog('Catalog','Сбой выгрузки плагина.');
            ShowMessage('Ошибка [FmDestroy]: "' + E.Message + '"');
         end;
      end;
end;

procedure TFmCatalog.pcCatalogChange(Sender: TObject);
begin
try
  if TreeGroupOwner.Showing then __DBTree(wFormID,TreeGroupOwner).FillTree(true); // заполнение дерева
  if TreeGroupPrice.Showing then __DBTree(wFormID,TreeGroupPrice).FillTree(true); // заполнение дерева
except
      on E: Exception do
  begin
     wLog('Catalog','Ошибка [pcCatalogChange]: "' + E.Message + '"');
     ShowMessage('Ошибка [pcCatalogChange]: "' + E.Message + '"');
  end;
end;
end;

procedure TFmCatalog.tbCatalogBtnAddClick(Sender: TObject);
var
  _Form: TFmNomenclatureEdit;
  _Tree: TwDBTreeVIew;
  _GridDataset: TDataSet;

  _ParentID, _oldParentID: integer;
  _ParentName: string;
  _SelectedID: integer;
  _SelectedName: string;
  _Scod, _Label, _Edini: string;
  _TimeStamp: integer;

  _P, _N, _M, _D, _C, _K, _S: Double;
begin
  _SelectedID:= 0;

       _Tree:= _DBTreeGroupPrice;

       _TimeStamp:= DateToUnix(Now());
       _ParentID:= _Tree.SelectedItems[0];
       _oldParentID:= _ParentID;

       _SelectedID:= _DBase.SQLInsert('CATALOG',['IDGROUPNOMENCLATURE','NAME','IDOWNER','FTIMESTAMP'],[_ParentID,'Новая позиция',OwnerMainID,_TimeStamp],false);


       _SelectedName:= 'Новая позиция';//_GridDataset.FieldByName('Name').AsString;

       _ParentName:= _Tree.BreadCrumbs(_ParentID);

  _Form:= TFmNomenclatureEdit.Create(Self);
  with _Form do begin
      kName.Text:=_SelectedName;
      edGroup.Text:=_ParentName;
      edGroup.Tag:=_ParentID;
      kNumber.Text:=IntTOStr(_SelectedID);
      kScod.Text:=GenEAN(__SCODPREFIX, '', IntToStr(_SelectedID));

      _Form.Caption:= '[Номенклатура] -= Создание =-';
      try
        ShowModal;
      finally

      if ModalResult  = mrOK then
         begin
           _SelectedName:= kName.Text;
           _ParentID:= edGroup.Tag;
           _Edini:= kEdini.Text;
           _Scod:= kScod.Text;
           _Label:= kArticul.Text;

           _P:= EditValue(e_PRICE);
           _N:= EditValue(e_PN);
           _M:= EditValue(e_PM);
           _D:= EditValue(e_PD);
           _C:= EditValue(e_PC);
           _K:= EditValue(e_PK);
           _S:= 0;
           _TimeStamp:= DateToUnix(Now());

         if
           not _DBase.SQLUpdate('CATALOG',['IDGROUPNOMENCLATURE','NAME','UNIT','PRICE','PN','PM','PD','PC','PK','STOCK','SCOD','LABEL','FTIMESTAMP'],[_ParentID,_SelectedName,_Edini,_P,_N,_M,_D,_C,_K,_S,_Scod,_Label,_TimeStamp],'ID='+IntToStr(_SelectedID))
         then
           begin
             SetStatus('Добавление новой позиции завершено с ошибкой.');
             wLog('Catalog','Добавление новой позиции завершено с ошибкой.');
           end else
           begin
             _GridDataset:= DBGridCatalogPrice.DataSource.DataSet;

            if _oldParentID <> _ParentID then
              begin
              try
                _Tree.FindNodeWithDataInt(_ParentID);
              finally

                if _GridDataset.RecordCount>0 then
                   _GridDataset.Locate('ID',_SelectedID,[]);
              end;
              end else
              begin
                  _GridDataset.Close;
                  _GridDataset.Open;
               if _GridDataset.RecordCount>0 then
                  _GridDataset.Locate('ID',_SelectedID,[]);
              end;
           end;

         end else
         begin
             _DBase.SQLTransactionEnd(false);
         end;

       _Form.Free;

     end;

  end;
end;

procedure TFmCatalog.tbCatalogBtnDeleteClick(Sender: TObject);
var
  _GridDataset: TDataSet;
  _ID, i: integer;
  _BookMark: TBookMark;
  _arr: ArrayOfInteger;
  _SelCount: integer;
begin
     _GridDataset:= DBGridCatalogPrice.DataSource.DataSet;
     _SelCount:= DBGridCatalogPrice.SelectedRows.Count;

  if _SelCount > 1 then
    begin
    if MessageDlg('Удалить несколько позиций ('+IntToStr(_SelCount)+') ? При удалении позиции так же будут удалены связанные соответствия!',mtConfirmation, mbOKCancel, 0) = mrOK then
       begin
         _arr:= _DBGrid.SelectedRows;
         _BookMark:= _GridDataset.Bookmark;

         try
           for i:=0 to Length(_arr)-1 do
           begin
              _DBase.SQLDelete('CATALOG','ID='+IntToStr(_arr[i]),false);
           end;
         finally
            _DBase.SQLTransactionEnd(true);
            _arr:=nil;
            with _GridDataset do begin
                close;
                open;
                if RecordCount>0 then BookMark:= _BookMark;
            end;
            wLog('Catalog',IntTOStr(_SelCount)+' позиций успешно удалено.');
            SetStatus(IntTOStr(_SelCount)+' позиций успешно удалено.');
         end;

       end;
    end else
    begin
     if MessageDlg('Удалить позицию "'+_GridDataset.FieldByName('Name').AsString+'" ? При удалении позиции так же будут удалены связанные соответствия!',mtConfirmation, mbOKCancel, 0) = mrOK then
        begin
           _ID:=_GridDataset.FieldByName('ID').AsInteger;
           _BookMark:= _GridDataset.Bookmark;
           if _DBase.SQLDelete('CATALOG','ID='+IntToStr(_ID)) then
              wLog('Catalog','Позиция успешно удалена');
              SetStatus('Позиция успешно удалена');
           with _GridDataset do begin
               close;
               open;
               if RecordCount>0 then BookMark:= _BookMark;
           end;
        end;
    end;
end;

procedure TFmCatalog.tbCatalogBtnEditClick(Sender: TObject);
var
  _Form: TFmNomenclatureEdit;
  _FormMass: TFmNomenclatureEditMass;
  _Tree: TwDBTreeVIew;
  _GridDataset: TDataSet;
  _Target: TComponent;
  _ParentID, _oldParentID: integer;
  _ParentName: string;
  _SelectedID: integer;
  _TimeStamp: integer;
  _SelectedName: string;
  _SelectedRowsCount: integer;
  _Scod, _Label, _Unit: string;
  _P, _N, _M, _D, _C, _K, _S : Double;
  i,ic: integer;

  _arr: ArrayOfInteger;
begin
  _SelectedID:= 0;
  _GridDataset:= DBGridCatalogPrice.DataSource.DataSet;
  _SelectedRowsCount:= DBGridCatalogPrice.SelectedRows.Count;
  _TimeStamp:= DateToUnix(Now());

  if _GridDataset.RecordCount>0 then
     begin
       _SelectedID:= _GridDataset.FieldByName('ID').AsInteger;
       _SelectedName:= _GridDataset.FieldByName('Name').AsString;
       _Tree:=  _DBTreeGroupPrice;
       _ParentID:= _GridDataset.FieldByName('IDGROUPNOMENCLATURE').AsInteger;
       _oldParentID:= _ParentID;
       _ParentName:= _Tree.BreadCrumbs(_ParentID);
     end;

  if _SelectedRowsCount>1 then
     begin
// множественный выбор

   _FormMass:= TFmNomenclatureEditMass.Create(Self);
   _FormMass.gbGroup.Tag:= _ParentID;
   _FormMass.l_edGroupText.Caption:= _ParentName;
   _FormMass.cbMain.Caption:='Изменить позиций: '+IntToStr(_SelectedRowsCount);

   try
     _FormMass.ShowModal;
   finally

         // получаем ID выбранных записей
         _arr:=nil;
         _arr:= _DBGrid.SelectedRows;

         try
           if _FormMass.cbUnit.Checked then
            begin
                 _Unit:= (_FormMass.gbUnit.Controls[1] as TComboBox).Text;

                   try
                     for i:=0 to Length(_arr)-1 do
                         _DBase.SQLUpdate('CATALOG',['UNIT','FTIMESTAMP'],[_Unit,_TimeStamp],'ID='+IntToStr(_arr[i]),false);

                   except
                    raise;
                   end;
            end;

           if _FormMass.cbGroup.Checked then
            begin
                 _ParentID:= (_FormMass.gbGroup.Controls[2] as TEdit).Tag;

                   try
                     for i:=0 to Length(_arr)-1 do
                         _DBase.SQLUpdate('CATALOG',['IDGROUPNOMENCLATURE','FTIMESTAMP'],[_ParentID,_TimeStamp],'ID='+IntToStr(_arr[i]),false);
                   except
                    raise;
                   end;
            end;

           if _FormMass.cbPrice.Checked then
            begin
// перебор компонентов
//wLog('debug ComponentCount',IntToStr(_Target.ComponentCount));
//for i:=0 to _Target.ComponentCount-1 do
//wLog('debug',' i='+IntTOStr(i)+' ComponentName='+ _Target.Components[i].Name);

              _Target:= _FormMass.FindComponent('FmNomenclatureEdit');

              for i:=0 to _Target.ComponentCount-1 do
                if (_Target.Components[i] is TEdit) then
                 begin
                   if ((_Target.Components[i] as TEdit).Name= 'e_PRICE') then
                     _P:= EditValue(_Target.Components[i] as TEdit);

                   if ((_Target.Components[i] as TEdit).Name= 'e_PN') then
                     _N:= EditValue(_Target.Components[i] as TEdit);

                   if ((_Target.Components[i] as TEdit).Name= 'e_PM') then
                     _M:= EditValue(_Target.Components[i] as TEdit);

                   if ((_Target.Components[i] as TEdit).Name= 'e_PD') then
                     _D:= EditValue(_Target.Components[i] as TEdit);

                   if ((_Target.Components[i] as TEdit).Name= 'e_PC') then
                     _C:= EditValue(_Target.Components[i] as TEdit);

                   if ((_Target.Components[i] as TEdit).Name= 'e_PK') then
                     _K:= EditValue(_Target.Components[i] as TEdit);
                 end;
                _Target:=nil;

                 try
                   for i:=0 to Length(_arr)-1 do
                   _DBase.SQLUpdate('CATALOG',['PRICE','PN','PM','PD','PC','PK','FTIMESTAMP'],[_P,_N,_M,_D,_C,_K,_TimeStamp],'ID='+IntToStr(_arr[i]),false)
                 except
                   _arr:=nil;
                    raise;
                 end;
            end;

           _arr:=nil;
           _DBase.SQLTransactionEnd(true);

// общий except операции группового изменения
         except
          _DBase.SQLTransactionEnd(false);
          SetStatus('Ошибка группового изменения записей. Операция отменена.');
         end;

     _FormMass.Free;

         if _oldParentID <> _ParentID then
           begin
               try
                 _Tree.FindNodeWithDataInt(_ParentID);
               finally
                 _GridDataset:= DBGridCatalogPrice.DataSource.DataSet;
                 if _GridDataset.RecordCount>0 then
                    _GridDataset.Locate('ID',_SelectedID,[]);
               end;

           end
         else
           begin
             _GridDataset:= DBGridCatalogPrice.DataSource.DataSet;
             _GridDataset.Close;
             _GridDataset.Open;
             if _GridDataset.RecordCount>0 then
                _GridDataset.Locate('ID',_SelectedID,[]);
           end;
   end;

// множественный выбор
     end else
     begin

      _Form:= TFmNomenclatureEdit.Create(Self);

        with _Form do begin
          kName.Text:=_SelectedName;
          edGroup.Text:=_ParentName;
          edGroup.Tag:=_ParentID;
          kEdini.Text:=_GridDataset.FieldByName('UNIT').AsString;

          _P:= _GridDataset.FieldByName('PRICE').AsFloat;
          e_PRICE.Text:=_GridDataset.FieldByName('PRICE').AsString;
          Razdelitel(e_PRICE,2,false);
      //'PN','PM','PD','PC','PK'
          _N:= _GridDataset.FieldByName('PN').AsFloat;
          e_PN.Text:=_GridDataset.FieldByName('PN').AsString;
          Razdelitel(e_PN,2,false);

          _M:= _GridDataset.FieldByName('PM').AsFloat;
          e_PM.Text:=_GridDataset.FieldByName('PM').AsString;
          Razdelitel(e_PM,2,false);

          _D:= _GridDataset.FieldByName('PD').AsFloat;
          e_PD.Text:=_GridDataset.FieldByName('PD').AsString;
          Razdelitel(e_PD,2,false);

          _C:= _GridDataset.FieldByName('PC').AsFloat;
          e_PC.Text:=_GridDataset.FieldByName('PC').AsString;
          Razdelitel(e_PC,2,false);

          _K:= _GridDataset.FieldByName('PK').AsFloat;
          e_PK.Text:=_GridDataset.FieldByName('PK').AsString;
          Razdelitel(e_PK,2,false);

          kScod.Text:=_GridDataset.FieldByName('SCOD').AsString;
          kNumber.Text:= IntToStr(_SelectedID); //_GridDataset.FieldByName('ID').AsString;
          kArticul.Text:=_GridDataset.FieldByName('LABEL').AsString;
          end;
          _Form.Caption:= '[Номенклатура] -= Редактирование =-';

          try
            _Form.ShowModal;
          finally

          if _Form.ModalResult  = mrOK then
             begin
               with _Form do
               begin
                 _SelectedName:= kName.Text;
                 _ParentID:= edGroup.Tag;
                 _Unit:= kEdini.Text;
                 _Scod:= kScod.Text;
                 _Label:= kArticul.Text;

                 _P:= EditValue(e_PRICE);
                 _N:= EditValue(e_PN);
                 _M:= EditValue(e_PM);
                 _D:= EditValue(e_PD);
                 _C:= EditValue(e_PC);
                 _K:= EditValue(e_PK);
                // _S:= 0;
               end;

             if
               not _DBase.SQLUpdate('CATALOG',['IDGROUPNOMENCLATURE','NAME','UNIT','PRICE','PN','PM','PD','PC','PK','SCOD','LABEL','FTIMESTAMP'],[_ParentID,_SelectedName,_Unit,_P,_N,_M,_D,_C,_K,_Scod,_Label,_TimeStamp],'ID='+IntToStr(_SelectedID))
             then
               begin
                 SetStatus('Изменение позиции завершено с ошибкой.');
                 wLog('Catalog','Изменение позиции завершено с ошибкой.');
               end else
               begin
                 if _oldParentID <> _ParentID then
                   begin
                       try
                         _Tree.FindNodeWithDataInt(_ParentID);
                       finally
                         _GridDataset:= DBGridCatalogPrice.DataSource.DataSet;
                         if _GridDataset.RecordCount>0 then
                            _GridDataset.Locate('ID',_SelectedID,[]);
                       end;

                   end
                 else
                   begin
                     _GridDataset:= DBGridCatalogPrice.DataSource.DataSet;
                     _GridDataset.Close;
                     _GridDataset.Open;
                     if _GridDataset.RecordCount>0 then
                        _GridDataset.Locate('ID',_SelectedID,[]);
                   end;
               end;

             end;

           _Form.Free;

         end;
     ////
   end;
  end;

procedure TFmCatalog.tbCatalogBtnGoToGroupClick(Sender: TObject);
var
  _Tree: TwDBTreeVIew;
  _GridDataset: TDataSet;
  _ParentID: integer;
  _ID: integer;
begin
  _GridDataset:= DBGridCatalogPrice.DataSource.DataSet;
  _ParentID:= _GridDataset.FieldByName('IDGROUPNOMENCLATURE').AsInteger;
  _ID:= _GridDataset.FieldByName('ID').AsInteger;
  _Tree:=  _DBTreeGroupPrice;
  try
  _Tree.FindNodeWithDataInt(_ParentID);
  finally

  if _GridDataset.RecordCount>0 then
    _GridDataset.Locate('ID',_ID,[]);
  end;
end;

procedure TFmCatalog.tbCatalogBtnCopyClick(Sender: TObject);
var
  _Form: TFmNomenclatureEdit;
  _Tree: TwDBTreeVIew;
  _GridDataset: TDataSet;

  _ParentID, _oldParentID: integer;
  _ParentName: string;
  _SelectedID: integer;
  _SelectedName: string;
  _Scod, _Label, _Edini: string;
  _TimeStamp: integer;

  _P, _N, _M, _D, _C, _K, _S: Double;
begin

if DBGridCatalogPrice.SelectedRows.Count>1 then
   begin
        ShowMessage('Для копирования выберите одну позицию!');
        exit;
   end;

_SelectedID:= 0;
_GridDataset:= DBGridCatalogPrice.DataSource.DataSet;

if _GridDataset.RecordCount>0 then
   begin
     _SelectedName:= _GridDataset.FieldByName('Name').AsString;
     _Tree:=  _DBTreeGroupPrice;
     _ParentID:= _GridDataset.FieldByName('IDGROUPNOMENCLATURE').AsInteger;
     _oldParentID:= _ParentID;
     _ParentName:= _Tree.BreadCrumbs(_ParentID);
   end;

       _Tree:=  __DBTree(wFormID,TreeGroupPrice);

       _TimeStamp:= DateToUnix(Now());

       _SelectedID:= _DBase.SQLInsert('CATALOG',['IDGROUPNOMENCLATURE','NAME','IDOWNER','FTIMESTAMP'],[_ParentID,'Новая позиция',OwnerMainID,_TimeStamp],false);

  _Form:= TFmNomenclatureEdit.Create(Self);
  with _Form do begin
      kName.Text:=_SelectedName;
      edGroup.Text:=_ParentName;
      edGroup.Tag:=_ParentID;
      kNumber.Text:=IntTOStr(_SelectedID);

  kEdini.Text:=_GridDataset.FieldByName('UNIT').AsString;

  _P:= _GridDataset.FieldByName('PRICE').AsFloat;
  e_PRICE.Text:=_GridDataset.FieldByName('PRICE').AsString;
  Razdelitel(e_PRICE,2,false);
//'PN','PM','PD','PC','PK'
  _N:= _GridDataset.FieldByName('PN').AsFloat;
  e_PN.Text:=_GridDataset.FieldByName('PN').AsString;
  Razdelitel(e_PN,2,false);

  _M:= _GridDataset.FieldByName('PM').AsFloat;
  e_PM.Text:=_GridDataset.FieldByName('PM').AsString;
  Razdelitel(e_PM,2,false);

  _D:= _GridDataset.FieldByName('PD').AsFloat;
  e_PD.Text:=_GridDataset.FieldByName('PD').AsString;
  Razdelitel(e_PD,2,false);

  _C:= _GridDataset.FieldByName('PC').AsFloat;
  e_PC.Text:=_GridDataset.FieldByName('PC').AsString;
  Razdelitel(e_PC,2,false);

  _K:= _GridDataset.FieldByName('PK').AsFloat;
  e_PK.Text:=_GridDataset.FieldByName('PK').AsString;
  Razdelitel(e_PK,2,false);

  kScod.Text:=GenEAN(__SCODPREFIX, '', IntToStr(_SelectedID));

  kArticul.Text:='';

  _Form.Caption:= '[Номенклатура] -= Копирование =-';

      try
        ShowModal;
      finally

      if ModalResult  = mrOK then
         begin
           _SelectedName:= kName.Text;
           _ParentID:= edGroup.Tag;
           _Edini:= kEdini.Text;
           _Scod:= kScod.Text;
           _Label:= kArticul.Text;

           _P:= EditValue(e_PRICE);
           _N:= EditValue(e_PN);
           _M:= EditValue(e_PM);
           _D:= EditValue(e_PD);
           _C:= EditValue(e_PC);
           _K:= EditValue(e_PK);
           _S:= 0;
           _TimeStamp:= DateToUnix(Now());

         if
           not _DBase.SQLUpdate('CATALOG',['IDGROUPNOMENCLATURE','NAME','UNIT','PRICE','PN','PM','PD','PC','PK','STOCK','SCOD','LABEL','FTIMESTAMP'],[_ParentID,_SelectedName,_Edini,_P,_N,_M,_D,_C,_K,_S,_Scod,_Label,_TimeStamp],'ID='+IntToStr(_SelectedID))
         then
           begin
             SetStatus('Добавление новой позиции завершено с ошибкой.');
             wLog('Catalog','Добавление новой позиции завершено с ошибкой.');
           end else
           begin
             _GridDataset:= DBGridCatalogPrice.DataSource.DataSet;

            if _oldParentID <> _ParentID then
              begin
              try
                _Tree.FindNodeWithDataInt(_ParentID);
              finally

                if _GridDataset.RecordCount>0 then
                   _GridDataset.Locate('ID',_SelectedID,[]);
              end;
              end else
              begin
                  _GridDataset.Close;
                  _GridDataset.Open;
               if _GridDataset.RecordCount>0 then
                  _GridDataset.Locate('ID',_SelectedID,[]);
              end;
           end;

         end else
         begin
             _DBase.SQLTransactionEnd(false);
         end;

       _Form.Free;

     end;

  end;
end;

procedure TFmCatalog.TreeGroupOwnerChange(Sender: TObject; Node: TTreeNode);
var
  _Tree: TTreeView;
begin
     _Tree:= (Sender as TTreeVIew);
     OwnerID:= TTreeData(_Tree.Selected.Data).Value;
end;

end.

