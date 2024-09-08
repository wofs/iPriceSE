unit UtilsU;

{
(c) Degtyarev A. A.
License in file LICENSE.txt
}

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  db,
  wBaseU, wFuncU, wDBTreeU, wLogU, wTypesU,
  Forms, Menus, Dialogs, Controls, StdCtrls, Math
  ;

procedure ImportUpdatePRICECALC(aBase:TwBase; aKursAndPercent, aFormatId: string);

procedure PricesLoadNomenclatureEditMassForm(Sender:Tobject; aBase:TwBase; aDBTree: TwDBTree);

procedure MatchingAutoFind(aMode: integer; aBase: TwBase; ArrSearchPosition: ArrayOfArrayVariant; aMatchLevel,aIdMainOwner:integer);

procedure InsertMaching(aBase: TwBase; aCatalogOwner: Integer; aCatalogVendorcode: string; aIdPLOwner, aIDPrice: Integer; aQuantInPacked: Double; aIDTMP: Integer);
procedure InsertMaching(aBase:TwBase; aIDCatalog:Integer; aQuantInPacked: Double; aIdPLOwner, aIDTMP: Integer; aVendorCode: string);

implementation
uses
  FmListSelectU, FmNomenclatureEditMassU;

procedure ImportUpdatePRICECALC(aBase:TwBase; aKursAndPercent, aFormatId: string);
var
  aSQL: String;
begin
  aKursAndPercent:= StringReplace(FloatToStr(RoundTo(StrToFloat(aKursAndPercent), -2)),',','.',[rfReplaceAll]);
  aSQL:= 'update %s SET PRICECALC=  CAST((PRICE*'+aKursAndPercent+') as numeric(15,2)),'
  +'PRICECALC2= CAST((PRICE2*'+aKursAndPercent+') as numeric(15,2)),'
  +'PRICECALC3= CAST((PRICE3*'+aKursAndPercent+') as numeric(15,2)),'
  +'PRICECALC4= CAST((PRICE4*'+aKursAndPercent+') as numeric(15,2)),'
  +'PRICECALC5= CAST((PRICE5*'+aKursAndPercent+') as numeric(15,2)),'
  +'PRICECALC6= CAST((PRICE6*'+aKursAndPercent+') as numeric(15,2)),'
  +'PRICECALC7= CAST((PRICE7*'+aKursAndPercent+') as numeric(15,2)),'
  +'PRICECALC8= CAST((PRICE8*'+aKursAndPercent+') as numeric(15,2)),'
  +'PRICECALC9= CAST((PRICE9*'+aKursAndPercent+') as numeric(15,2)),'
  +'PRICECALC10= CAST((PRICE10*'+aKursAndPercent+') as numeric(15,2)) '
  +'WHERE IDFORMATS='+aFormatID+'';
  aBase.SQLUpdate(Format(aSQL,['PL_VERSIONS']),false);
  aBase.SQLUpdate(Format(aSQL,['PL_ITEMS']),false);
end;

procedure PricesLoadNomenclatureEditMassForm(Sender:Tobject; aBase:TwBase; aDBTree: TwDBTree); //TFmNomenclatureEditMass
var
  _Form: TFmNomenclatureEditMass;
  i: Integer;
  _arr, _arrGroupPrice: ArrayOfArrayVariant;
  _SelectedItems: ArrayOfInteger;
  _ParentID: integer;
  _ParentName, _IdMainOwner: string;
  _Target: TComponent;
  _PRICE, _N, _M, _D, _C, _K: Double;
begin

  _Form:= TFmNomenclatureEditMass.Create(TComponent(Sender));

  _IdMainOwner:= aBase.ReadSettingByName('setDefaultOwner');

  _Form.Caption:= 'Групповое добавление позиций в каталог.';
  _ParentID:= aBase.SQLReadArr('CATALOG_GROUP',['ID'],'IDPARENT=0','ID')[0,0];
  _ParentName:= aBase.SQLReadArr('select * from GETPARENTS_GROUP_CATALOG('+IntToStr(_ParentID)+');')[0,0];

  _Form.Base:= aBase;
  _Form.gbGroup.Tag:= _ParentID;
  _Form.l_edGroupText.Caption:= _ParentName;

  _Form.cbUnit.Checked:= true;
  _Form.cbUnit.Enabled:= false;

  _Form.cbGroup.Checked:= true;
  _Form.cbGroup.Enabled:= false;

  _Form.cbPrice.Checked:= true;
  _Form.cbPrice.Enabled:= false;

  _Form.cbAll.Checked:= true;
  _Form.gbChange.Enabled:= false;

  _Form.cbUnselect.Enabled:= false;
  try
    try

      _Form.ShowModal;

      if _Form.ModalResult = mrOK then
        begin

          if _Form.cbGroup.Checked then
            _ParentID:= (_Form.gbGroup.Controls[2] as TEdit).Tag;

          _Target:= _Form.FindComponent('FmNomenclatureEdit');

          for i:=0 to _Target.ComponentCount-1 do
            if (_Target.Components[i] is TEdit) then
             begin
               if ((_Target.Components[i] as TEdit).Name= 'e_PRICE1') then
                 _PRICE:= EditValue(_Target.Components[i] as TEdit);

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

          //if _ParentID = 0 then
          //    _arr:= aBase.SQLReadArr('CATALOG_GROUP',['ID'],'IDOWNER='+_IdMainOwner+' AND IDPARENT=0','') else
          _arr:= aBase.SQLReadArr('CATALOG_GROUP',['ID'],'ID='+IntToStr(_ParentID),'');

          if Assigned(_arr) then
          begin
            aDBTree.ShowChildrenItems:= false;
            _SelectedItems:= aDBTree.SelectedItems;
            aDBTree.ShowChildrenItems:= true;
            Screen.Cursor:= crSQLWait;
            for i:=0 to High(_SelectedItems) do
                begin
                  _arrGroupPrice:= aBase.SQLReadArr('PL_GROUP',['ID','NAME'],'ID='+IntToStr(_SelectedItems[i]),'');
                  if Assigned(_arrGroupPrice) then
                       aBase.SQLUpdate('EXECUTE PROCEDURE PL_GROUP_TO_CATALOG('+_IdMainOwner+','+IntToStr(_arr[0,0])+','+IntToStr(_arrGroupPrice[0,0])+','+QuotedStr(VarToStr(_arrGroupPrice[0,1]))+','+FloatToStr(_PRICE)+','+FloatToStr(_N)+','+FloatToStr(_M)+','+FloatToStr(_D)+','+FloatToStr(_C)+','+FloatToStr(_K)+');');
                end;

            ShowMessage('Добавление выбраных категории прайс-листа в каталог завершено.');
          end;
        end;

    finally
      _Form.Free;
      Screen.Cursor:= crDefault;
    end;
  except
    on E: Exception do
      begin
         wLog('Prices','Ошибка [AddGroupToCatalog] "' + E.Message + '"');
         ShowMessage('Ошибка [AddGroupToCatalog]: "' + E.Message + '"');
      end;
  end;
end;

procedure MatchingAutoFind(aMode: integer; aBase: TwBase; ArrSearchPosition: ArrayOfArrayVariant; aMatchLevel,aIdMainOwner:integer);
var
  _indistinctmatching: String;
  _arr, _arrScodCtg: ArrayOfArrayVariant;
  iScod, i: Integer;
begin
  for i:=0 to High(ArrSearchPosition) do begin
    case aMode of
      0:
        begin
         case aMatchLevel of
           1: _indistinctmatching:= ' indistinctmatching4('+QuotedStr(string(ArrSearchPosition[i,1]))+',CTG.NAME) ';
           2: _indistinctmatching:= ' indistinctmatchingmanual('+QuotedStr(string(ArrSearchPosition[i,1]))+',CTG.NAME,'+QuotedStr(IntToStr(CalcProbel(string(ArrSearchPosition[i,1]),false)))+') ';
         end;

         aBase.SQLUpdate('insert into W_TMP_TBL_NEUTRALSEARCH '
         +' select CTG.IDOWNER,CTG.ID,CTG.NAME '
         +' ,'+_indistinctmatching+' as WCS,'+string(ArrSearchPosition[i,0])+' '
         +' from "CATALOG" CTG',false);

        end;
      1:
        begin
           _arr:=nil;

           if Length(ArrSearchPosition)=1 then
              _arr:= aBase.SQLReadArr('PL_SCODS',['SCOD'],'IDPL_ITEMS='+string(ArrSearchPosition[i,0]),'SCOD') else
              if Length(ArrSearchPosition[i,1])>0 then
              begin
                SetLength(_arr,1,1);
                _arr[0,0]:= ArrSearchPosition[i,1];
              end;

           if Assigned(_arr) then
           for iScod:=0 to High(_arr) do
           begin
               _arrScodCtg:=nil;
               _arrScodCtg:= aBase.SQLReadArr('SELECT IDCTG_ITEMS FROM CTG_CHECK_SCOD('+IntToStr(aIdMainOwner)+','+QuotedStr(_arr[iScod,0])+')');
              if Assigned(_arrScodCtg) and (_arrScodCtg[0,0]<>null) then
                 aBase.SQLUpdate('insert into W_TMP_TBL_NEUTRALSEARCH'
                 +' select CTG.IDOWNER,CTG.ID,CTG.NAME '
                 +' ,95 as WCS,'+string(ArrSearchPosition[i,0])+' '
                 +' from "CATALOG" CTG'
                 +' where CTG.ID='+string(_arrScodCtg[0,0])
                 ,false);
           end;
          end;
      2:
        begin
         aBase.SQLUpdate('insert into W_TMP_TBL_NEUTRALSEARCH '
         +' select CTG.IDOWNER,CTG.ID,CTG.NAME '
         +' ,95 as WCS,'+string(ArrSearchPosition[i,0])+' '
         +' from "CATALOG" CTG'
         +' where CTG.LABEL='+QuotedStr(ArrSearchPosition[i,1])+' AND CTG.LABEL <>'''' AND CTG.LABEL IS NOT NULL'
         ,false);

         //aBase.SQLUpdate('insert into W_TMP_TBL_NEUTRALSEARCH '
         //+' select CTG.IDOWNER,CTG.ID,CTG.NAME '
         //+' ,80 as WCS,'+string(ArrSearchPosition[i,0])+' '
         //+' from "CATALOG" CTG'
         //+' where CTG.LABEL LIKE '+QuotedStr('%'+ArrSearchPosition[i,1]+'%')
         //,false);
        end;
    end;
    Application.ProcessMessages;
  end;
end;

procedure InsertMaching(aBase: TwBase; aCatalogOwner: Integer; aCatalogVendorcode: string; aIdPLOwner, aIDPrice: Integer; aQuantInPacked: Double; aIDTMP: Integer);
var
  _arr: ArrayOfArrayVariant;
  aIDCatalog, aIDMatching: Integer;
begin

  try
    _arr:= aBase.SQLReadArr('CATALOG',
           ['ID'],
           'IDOWNER='+IntToStr(aCatalogOwner)+' AND VENDORCODE='+QuotedStr(aCatalogVendorcode),'');

    if Assigned(_arr) then
      aIDCatalog:= _arr[0,0]
    else
      aIDCatalog:= 0;

    aIDMatching:= aBase.SQLInsert('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],
                                            [aIdPLOwner,aIDCatalog,aIDPrice,aQuantInPacked,now(),integer(1)],'IDOWNER, IDPL_ITEMS',false);
    aBase.SQLUpdate('W_TMP_ORDERS_IMPORT',['MTHID'],[aIDMatching],'ID='+IntToStr(aIDTMP),false);

  except
    raise;
  end;
end;

procedure InsertMaching(aBase: TwBase; aIDCatalog: Integer; aQuantInPacked: Double; aIdPLOwner, aIDTMP: Integer; aVendorCode: string);
var
  _arr: ArrayOfArrayVariant;
  _IDPrice, _PriceRoot, _IDCatalog, _IDMatching: Integer;
begin

  try

    //_SelectedVendorCode:= kVendorCode.Text;
    _arr:= aBase.SQLReadArr('PL_ITEMS',
           ['ID'],
           'IDOWNER='+IntToStr(aIdPLOwner)+' AND VENDORCODE='+QuotedStr(aVendorCode),'');
    if Assigned(_arr) then
      _IDPrice:= _arr[0,0]
    else
      _IDPrice:= 0;

    //_IDOwner:= IdMainOwner;

    if _IDPrice = 0 then
    begin
      //aVendorCode:= _DataSet.FieldByName('ORDVENDORCODE').AsString;
      _PriceRoot:= aBase.SQLReadArr('PL_GROUP',['ID'],'IDOWNER='+IntToStr(aIdPLOwner),'')[0,0];
      _IdPrice:= aBase.SQLInsert('PL_ITEMS',['IDPL_GROUP','IDOWNER','IDFORMATS','NAME','UNIT','LABEL','REMARK','FURL','FURLPICTURE','VENDORCODE','FTIMESTAMP'],
                                [_PriceRoot, aIdPLOwner,integer(0),'Создано процедурой занесения накладной','','','','','',aVendorCode,now],'IDOWNER,VENDORCODE',false);
    end;
    _IDCatalog:= aIDCatalog;

    _IDMatching:= aBase.SQLInsert('CATALOG_MATCHING',['IDOWNER','IDCATALOG','IDPL_ITEMS','QUANTITYINPACKING','FTIMESTAMP','IDUSER'],
                                            [aIdPLOwner,_IDCatalog,_IDPrice,aQuantInPacked,now(),integer(1)],'IDOWNER, IDPL_ITEMS',false);
    aBase.SQLUpdate('W_TMP_ORDERS_IMPORT',['MTHID'],[_IDMatching],'ID='+IntToStr(aIDTMP),false);

  except
    raise;
  end;
end;

end.

