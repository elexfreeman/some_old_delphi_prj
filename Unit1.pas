unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Grids, ComCtrls, ComObj, Buttons, jpeg,DataB;

type
  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    Import01: TEdit;
    Panel3: TPanel;
    Memo1: TMemo;
    dlgSave1: TSaveDialog;
    pgc1: TPageControl;
    ts1: TTabSheet;
    strngrd1: TStringGrid;
    pnl1: TPanel;
    btn1: TButton;
    btn2: TButton;
    ts2: TTabSheet;
    pnl2: TPanel;
    btn4: TButton;
    btn5: TButton;
    strngrd2: TStringGrid;
    ts3: TTabSheet;
    pnl3: TPanel;
    btn6: TButton;
    btn7: TButton;
    strngrd3: TStringGrid;
    ts4: TTabSheet;
    pnl4: TPanel;
    btn8: TButton;
    btn9: TButton;
    strngrd4: TStringGrid;
    ts5: TTabSheet;
    pnl5: TPanel;
    btn10: TButton;
    btn11: TButton;
    strngrd5: TStringGrid;
    ts6: TTabSheet;
    pnl6: TPanel;
    btn13: TButton;
    strngrd6: TStringGrid;
    chk1: TCheckBox;
    SaveDialog1: TSaveDialog;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    lbl1: TLabel;
    lbl2: TLabel;
    edtDogovorE: TEdit;
    lbl3: TLabel;
    lbl4: TLabel;
    edtEnergoOrg: TEdit;
    lbl5: TLabel;
    edtSetOrg: TEdit;
    edtPotreb: TEdit;
    edtSvedenia: TEdit;
    pnl7: TPanel;
    pnl8: TPanel;
    pb1: TProgressBar;
    lbl6: TLabel;
    btn12: TBitBtn;
    ts7: TTabSheet;
    pnl9: TPanel;
    btn17: TButton;
    btn18: TButton;
    btn19: TButton;
    strngrd7: TStringGrid;
    ts8: TTabSheet;
    pnl10: TPanel;
    btn14: TButton;
    btn15: TButton;
    btn16: TButton;
    strngrd8: TStringGrid;
    img1: TImage;
    btn20: TBitBtn;
    btnSaveDB7: TBitBtn;
    ts9: TTabSheet;
    Memo2: TMemo;
    btn3: TBitBtn;
    btnSaveDB4: TBitBtn;
    btnSaveDBSklad: TBitBtn;
    btnSaveDBABK: TBitBtn;
    btnSaveDBbank: TBitBtn;
    btnSaveDB14: TBitBtn;
    btnSaveDB_q1: TBitBtn;
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btn3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btn4Click(Sender: TObject);
    procedure btn6Click(Sender: TObject);
    procedure btn8Click(Sender: TObject);
    procedure btn10Click(Sender: TObject);
    procedure btn12Click(Sender: TObject);
    procedure btn13Click(Sender: TObject);
    procedure btn5Click(Sender: TObject);
    procedure btn7Click(Sender: TObject);
    procedure btn9Click(Sender: TObject);
    procedure btn11Click(Sender: TObject);
    procedure btnExel_Z7_Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure btnSaveExelClick(Sender: TObject);
    procedure btn17Click(Sender: TObject);
    procedure btn14Click(Sender: TObject);
    procedure chk1Click(Sender: TObject);
    procedure btn20Click(Sender: TObject);
    procedure btnSaveDB7Click(Sender: TObject);
    procedure btnSaveDB4Click(Sender: TObject);
    procedure btnSaveDBSkladClick(Sender: TObject);
    procedure btnSaveDBABKClick(Sender: TObject);
    procedure btnSaveDBbankClick(Sender: TObject);
    procedure btnSaveDB14Click(Sender: TObject);
    procedure btnSaveDB_q1Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;
 function convertZPT(s:string):string;
var
  Form1: TForm1;


implementation
uses SQLite3, SQLiteTable3;

{$R *.dfm}

 function convertZPT(s:string):string;
 var i:Integer;
 s1:string;
 begin
    s1:='';
    for i:=1 to length(s) do
     if s[i]=',' then s1:=s1+'.' else s1:=s1+s[i];
    convertZPT:= s1;
 end;
procedure TForm1.btnExel_Z7_Click(Sender: TObject);
var Exel:Variant;
i,j,k:integer;
begin
 if Form1.SaveDialog1.Execute then
     begin

        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');

        Exel:= CreateOleObject('Excel.Application');
 //Exel.Workbooks.Add;
//form1.Caption:=(ExtractFilePath(Application.ExeName));
 Exel.Workbooks.Open(ExtractFilePath(Application.ExeName)+'book.xlsx');
 Exel.ActiveWorkBook.Sheets.Item[1].Activate;
 //Exel.Range['A2'] := 'Hellow';


     for i:=1 to 31 do
     begin

       k:=50;
       repeat
            Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,k-50+3]:=strngrd1.Cells[k,i];
             k:=k+1;
       until k>50+25;
       end;


 //Exel.ActiveWorkBook.Sheets.Item[2].Activate;
 //Exel.Range['A2'] := 'Hellow';

// Exel.ActiveWorkBook.Sheets.Item[9].Activate;
// Exel.Range['B2'] := 'Hellow 101 kbbkbkbkkb';
   Exel.ActiveWorkBook.SaveAs(SaveDialog1.FileName+'.xlsx');
 Exel.ActiveWorkBook.Close;
Exel.Application.Quit;
     end;


end;


//Ячейка 7
procedure TForm1.btn1Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin

  if Form1.OpenDialog1.Execute then
   begin
      Form1.Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
      Form1.Import01.Text:=OpenDialog1.FileName;
      Memo2.Lines.Clear;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          date1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              date1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
            end;
            //выделяем мощность
            if ((s[j]='P')and(s[j+1]='4')and(s[j+2]='8'))
            then
            begin
               k:=j+4;
               while (s[k]<>';') do
               begin
                 data1:=data1+s[k];
                 k:=k+1;
               end;

            end;

          end;
            Memo2.Lines.Add(date1+','+data1);
        end;
      end;
       for i:=1 to Memo2.Lines.Count do
       begin
         s:=Memo2.Lines[i];
         s1:='';
         for j:=1 to Length(s) do
         begin
            if s[j]=',' then s1:=s1+';' else
            if s[j]='.' then s1:=s1+','else s1:=s1+s[j];
         end;
         Memo2.Lines[i]:=s1;
       end;
       end;

      k:=0;
      strngrd1.RowCount:=Memo2.Lines.Count;
      strngrd1.colCount:=49+28;
      for i := 1 to 48 do
       strngrd1.Cells[i,0]:=IntToStr(i);
      for i := 49 to 49+28 do
       strngrd1.Cells[i,0]:=IntToStr(i-49);
      //запоняем днными
      for i:=1 to Memo2.Lines.Count do
      begin
         s1:=Memo2.Lines[i];
         s2:='';
         if Length(s1)>9 then
         begin
           s3:='';
           for j:=9 to Length(s1) do
           begin
             if s1[j]<>';' then
             s2:=s2+s1[j]
             else
             begin
               strngrd1.Cells[k,i]:=s2;
               k:=k+1;
               s2:='';
               end;
           end;
           strngrd1.Cells[k,i]:=s2;
           s2:='';
           k:=0;
         end;
      end;
            //запоняем даты
            k:=1;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              data1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
              strngrd1.Cells[0,k]:=data1;
              k:=k+1;
            end;
          end;
        end;
       end;

///////Складываем значения 1+2 3+4

     for i:=1 to 31 do
     begin
       j:=1;
       k:=50;
       repeat
            strngrd1.Cells[k,i] := FloatToStr(

            (StrToFloat(strngrd1.Cells[j,i])+StrToFloat(strngrd1.Cells[j+1,i]))*3600

             );
             j:=j+2;
             k:=k+1;
       until j>48;
       end;

 Form1.btn3Click(Sender);
end;

procedure TForm1.btn2Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin
  Memo2.Lines.Clear;
  for i:=1 to 31 do
  begin
     s:='';
     for j:=0 to 24+50+1 do
     begin
          s:=s+strngrd1.Cells[j,i]+';';
     end;
  memo2.Lines.Add(s);
  end;
  if Form1.dlgSave1.Execute then
     begin
        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');
     end;

end   ;

procedure TForm1.Button1Click(Sender: TObject);
var
   i,j,k:Integer;
   s1,s2,s3,s,data1:string;

begin

      k:=0;
      strngrd1.RowCount:=Memo2.Lines.Count;
      strngrd1.colCount:=49;
      for i := 1 to 48 do
       strngrd1.Cells[i,0]:=IntToStr(i);

      //запоняем днными
      for i:=1 to Memo2.Lines.Count do
      begin
         s1:=Memo2.Lines[i];
         s2:='';
         if Length(s1)>9 then
         begin
           s3:='';
           for j:=9 to Length(s1) do
           begin
             if s1[j]<>';' then
             s2:=s2+s1[j]
             else
             begin
               strngrd1.Cells[k,i]:=s2;
               k:=k+1;
               s2:='';
               end;
           end;
           strngrd1.Cells[k,i]:=s2;
           s2:='';
           k:=0;
         end;
      end;
            //запоняем даты
            k:=1;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              data1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
              strngrd1.Cells[0,k]:=data1;
              k:=k+1;
            end;
          end;
        end;
       end;

end;

//Расчитать
procedure TForm1.btn3Click(Sender: TObject);
var i,j,k:Integer;
summa_7, summa_7_f,summa_4, summa_4_f,summa_sk, summa_sk_f,summa_abk, summa_abk_f,summa_b, summa_b_f
, summa_1, summa_1_f, summa_14, summa_14_f: real;
begin
     for i:=0 to 31 do
       for j:=0 to 26 do
            strngrd6.Cells[j,i] := strngrd1.Cells[j,i];
     // расчитываем суммы во всеех ячейках

       summa_7:=0;
       summa_4:=0;
       summa_sk:=0;
       summa_abk:=0;
       summa_b:=0;
       summa_1:=0;
       summa_14:=0;
     for i:=1 to 31 do
     begin
       j:=1;
       k:=50;
       repeat
            summa_7:=summa_7+strtofloat(strngrd1.Cells[k,i]);
            summa_4:=summa_4+strtofloat(strngrd2.Cells[k,i]);
            summa_sk:=summa_sk+strtofloat(strngrd3.Cells[k,i]);
            summa_abk:=summa_abk+strtofloat(strngrd4.Cells[k,i]);
            summa_b:=summa_b+strtofloat(strngrd5.Cells[k,i]);
            summa_1:=summa_1+strtofloat(strngrd7.Cells[k,i]);
            summa_14:=summa_14+strtofloat(strngrd8.Cells[k,i]);
             j:=j+2;
             k:=k+1;
       until j>48;
       strngrd1.Cells[74,i]:=FloatToStr(summa_7);
       strngrd2.Cells[74,i]:=FloatToStr(summa_4);
       strngrd3.Cells[74,i]:=FloatToStr(summa_sk);
       strngrd4.Cells[74,i]:=FloatToStr(summa_abk);
       strngrd5.Cells[74,i]:=FloatToStr(summa_b);
       strngrd7.Cells[74,i]:=FloatToStr(summa_1);
       strngrd8.Cells[74,i]:=FloatToStr(summa_14);
       summa_7_f:=summa_7_f+summa_7;
       summa_4_f:=summa_4_f+summa_4;
       summa_sk_f:=summa_sk_f+summa_sk;
       summa_abk_f:=summa_abk_f+summa_abk;
       summa_1_f:=summa_1_f+summa_1;
       summa_b_f:=summa_b_f+summa_b;
       summa_14_f:=summa_14_f+summa_14;
       summa_7:=0;
       summa_4:=0;
       summa_sk:=0;
       summa_abk:=0;
       summa_b:=0;
       summa_1:=0;
       summa_14:=0;

     end;
     strngrd1.Cells[74,0]:=FloatToStr(summa_7_f);;
     strngrd2.Cells[74,0]:=FloatToStr(summa_4_f);
     strngrd3.Cells[74,0]:=FloatToStr(summa_sk_f);
     strngrd4.Cells[74,0]:=FloatToStr(summa_abk_f);
     strngrd5.Cells[74,0]:=FloatToStr(summa_b_f);
     strngrd7.Cells[74,0]:=FloatToStr(summa_1_f);
     strngrd8.Cells[74,0]:=FloatToStr(summa_14_f);

   // Считаем общию сумму
   summa_7:=0;
   if chk1.Checked then
     begin
       summa_7:=0;
        for i:=1 to 31 do
          begin

             k:=1;
             for j:=49 to 50+25 do
               begin
                  strngrd6.Cells[j-49,i]:=FloatToStr(
                     StrToFloat(strngrd1.Cells[j,i])+
                     StrToFloat(strngrd2.Cells[j,i])+
                     StrToFloat(strngrd7.Cells[j,i])+
                     StrToFloat(strngrd8.Cells[j,i])+
                     StrToFloat(strngrd3.Cells[j,i])  -
                     StrToFloat(strngrd4.Cells[j,i])  -
                     StrToFloat(strngrd5.Cells[j,i])

             );
               end;

             summa_7:=summa_7+

             StrToFloat(strngrd1.Cells[74,i])+
             StrToFloat(strngrd2.Cells[74,i])+
             StrToFloat(strngrd7.Cells[74,i])+
             StrToFloat(strngrd8.Cells[74,i])+
             StrToFloat(strngrd3.Cells[74,i])  -
             StrToFloat(strngrd4.Cells[74,i])  -
             StrToFloat(strngrd5.Cells[74,i])

             ;

          end;
     end
      else
      begin
        summa_7:=0;
            for i:=1 to 31 do
                begin

                   k:=1;
                   for j:=49 to 50+25 do
                     begin
                        strngrd6.Cells[j-49,i]:=FloatToStr(

                           StrToFloat(strngrd1.Cells[j,i])+
                           StrToFloat(strngrd8.Cells[j,i])+
                           StrToFloat(strngrd7.Cells[j,i])+
                           StrToFloat(strngrd2.Cells[j,i])  -
                          // StrToFloat(strngrd3.Cells[j,i])  -
                           StrToFloat(strngrd4.Cells[j,i])  -
                           StrToFloat(strngrd5.Cells[j,i])

                   );
                     end;

                   summa_7:=summa_7+

                   StrToFloat(strngrd1.Cells[74,i])+
                   StrToFloat(strngrd7.Cells[74,i])+
                   StrToFloat(strngrd8.Cells[74,i])+
                   StrToFloat(strngrd2.Cells[74,i]) -
                   //StrToFloat(strngrd3.Cells[74,i])  -
                   StrToFloat(strngrd4.Cells[74,i])  -
                   StrToFloat(strngrd5.Cells[74,i])

                   ;

                end;
      end;
      strngrd6.Cells[25,0]:=FloatToStr( summa_7);

       //выставляем даты
       for i:=0 to 31 do
            strngrd6.Cells[0,i] := strngrd1.Cells[0,i];


end;

procedure TForm1.FormCreate(Sender: TObject);
var
  i,j:Integer;
begin
     strngrd1.RowCount:=32;
     strngrd1.colCount:=49+27;
     strngrd2.RowCount:=32;
     strngrd2.colCount:=49+27;
     strngrd3.RowCount:=32;
     strngrd3.colCount:=49+27;
     strngrd4.RowCount:=32;
     strngrd4.colCount:=49+27;
     strngrd5.RowCount:=32;
     strngrd5.colCount:=49+27;
     strngrd7.RowCount:=32;
     strngrd7.colCount:=49+27;
     strngrd8.RowCount:=32;
     strngrd8.colCount:=49+27;


     strngrd6.RowCount:=32;
     strngrd6.colCount:=26;

     for i:=1 to 31 do

     for j:=1 to 49+27 do
     begin
          strngrd1.Cells[j,i] := '0';
          strngrd2.Cells[j,i] := '0';
          strngrd3.Cells[j,i] := '0';
          strngrd4.Cells[j,i] := '0';
          strngrd5.Cells[j,i] := '0';
          strngrd6.Cells[j,i] := '0';
          strngrd7.Cells[j,i] := '0';
          strngrd8.Cells[j,i] := '0';
     end;


end;

//Ячейка 4
procedure TForm1.btn4Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin

  if Form1.OpenDialog1.Execute then
   begin
      Form1.Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
      Form1.Import01.Text:=OpenDialog1.FileName;
      Form1.Memo2.Lines.Clear;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          date1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              date1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
            end;
            //выделяем мощность
            if ((s[j]='P')and(s[j+1]='4')and(s[j+2]='8'))
            then
            begin
               k:=j+4;
               while (s[k]<>';') do
               begin
                 data1:=data1+s[k];
                 k:=k+1;
               end;

            end;

          end;
            Memo2.Lines.Add(date1+','+data1);
        end;
      end;
       for i:=1 to Memo2.Lines.Count do
       begin
         s:=Memo2.Lines[i];
         s1:='';
         for j:=1 to Length(s) do
         begin
            if s[j]=',' then s1:=s1+';' else
            if s[j]='.' then s1:=s1+','else s1:=s1+s[j];
         end;
         Memo2.Lines[i]:=s1;
       end;
       end;

      k:=0;
      strngrd2.RowCount:=Memo2.Lines.Count;
      strngrd2.colCount:=49+28;
      for i := 1 to 48 do
       strngrd2.Cells[i,0]:=IntToStr(i);

       for i := 49 to 49+28 do
       strngrd2.Cells[i,0]:=IntToStr(i-49);
      //запоняем днными
      for i:=1 to Memo2.Lines.Count do
      begin
         s1:=Memo2.Lines[i];
         s2:='';
         if Length(s1)>9 then
         begin
           s3:='';
           for j:=9 to Length(s1) do
           begin
             if s1[j]<>';' then
             s2:=s2+s1[j]
             else
             begin
               strngrd2.Cells[k,i]:=s2;
               k:=k+1;
               s2:='';
               end;
           end;
           strngrd2.Cells[k,i]:=s2;
           s2:='';
           k:=0;
         end;
      end;
            //запоняем даты
            k:=1;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              data1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
              strngrd2.Cells[0,k]:=data1;
              k:=k+1;
            end;
          end;
        end;
       end;

       ///////Складываем значения 1+2 3+4

     for i:=1 to 31 do
     begin
       j:=1;
       k:=50;
       repeat
            strngrd2.Cells[k,i] := FloatToStr(

            (StrToFloat(strngrd2.Cells[j,i])+StrToFloat(strngrd2.Cells[j+1,i]))*3600

             );
             j:=j+2;
             k:=k+1;
       until j>48;
       end;

  Form1.btn3Click(Sender);
end;

//Склад ГО
procedure TForm1.btn6Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin

  if Form1.OpenDialog1.Execute then
   begin
      Form1.Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
      Form1.Import01.Text:=OpenDialog1.FileName;
      Form1.Memo2.Lines.Clear;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          date1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              date1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
            end;
            //выделяем мощность
            if ((s[j]='P')and(s[j+1]='4')and(s[j+2]='8'))
            then
            begin
               k:=j+4;
               while (s[k]<>';') do
               begin
                 data1:=data1+s[k];
                 k:=k+1;
               end;

            end;

          end;
            Memo2.Lines.Add(date1+','+data1);
        end;
      end;
       for i:=1 to Memo2.Lines.Count do
       begin
         s:=Memo2.Lines[i];
         s1:='';
         for j:=1 to Length(s) do
         begin
            if s[j]=',' then s1:=s1+';' else
            if s[j]='.' then s1:=s1+','else s1:=s1+s[j];
         end;
         Memo2.Lines[i]:=s1;
       end;
       end;

      k:=0;
      strngrd3.RowCount:=Memo2.Lines.Count;
      strngrd3.colCount:=49+48;
      for i := 1 to 48 do
       strngrd3.Cells[i,0]:=IntToStr(i);
      for i := 49 to 49+28 do
       strngrd3.Cells[i,0]:=IntToStr(i-49);
      //запоняем днными
      for i:=1 to Memo2.Lines.Count do
      begin
         s1:=Memo2.Lines[i];
         s2:='';
         if Length(s1)>9 then
         begin
           s3:='';
           for j:=9 to Length(s1) do
           begin
             if s1[j]<>';' then
             s2:=s2+s1[j]
             else
             begin
               strngrd3.Cells[k,i]:=s2;
               k:=k+1;
               s2:='';
               end;
           end;
           strngrd3.Cells[k,i]:=s2;
           s2:='';
           k:=0;
         end;
      end;
            //запоняем даты
            k:=1;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              data1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
              strngrd3.Cells[0,k]:=data1;
              k:=k+1;
            end;
          end;
        end;
       end;


       ///////Складываем значения 1+2 3+4

     for i:=1 to 31 do
     begin
       j:=1;
       k:=50;
       repeat
            strngrd3.Cells[k,i] := FloatToStr(

            (StrToFloat(strngrd3.Cells[j,i])+StrToFloat(strngrd3.Cells[j+1,i]))*20

             );
             j:=j+2;
             k:=k+1;
       until j>48;
       end;
   Form1.btn3Click(Sender);
end;


//АБК
procedure TForm1.btn8Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin

  if Form1.OpenDialog1.Execute then
   begin
      Form1.Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
      Form1.Import01.Text:=OpenDialog1.FileName;
      Form1.Memo2.Lines.Clear;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          date1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              date1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
            end;
            //выделяем мощность
            if ((s[j]='P')and(s[j+1]='4')and(s[j+2]='8'))
            then
            begin
               k:=j+4;
               while (s[k]<>';') do
               begin
                 data1:=data1+s[k];
                 k:=k+1;
               end;

            end;

          end;
            Memo2.Lines.Add(date1+','+data1);
        end;
      end;
       for i:=1 to Memo2.Lines.Count do
       begin
         s:=Memo2.Lines[i];
         s1:='';
         for j:=1 to Length(s) do
         begin
            if s[j]=',' then s1:=s1+';' else
            if s[j]='.' then s1:=s1+','else s1:=s1+s[j];
         end;
         Memo2.Lines[i]:=s1;
       end;
       end;

      k:=0;
      strngrd4.RowCount:=Memo2.Lines.Count;
      strngrd4.colCount:=49+28;
      for i := 1 to 48 do
       strngrd4.Cells[i,0]:=IntToStr(i);
      for i := 49 to 49+28 do
       strngrd1.Cells[i,0]:=IntToStr(i-49);
      //запоняем днными
      for i:=1 to Memo2.Lines.Count do
      begin
         s1:=Memo2.Lines[i];
         s2:='';
         if Length(s1)>9 then
         begin
           s3:='';
           for j:=9 to Length(s1) do
           begin
             if s1[j]<>';' then
             s2:=s2+s1[j]
             else
             begin
               strngrd4.Cells[k,i]:=s2;
               k:=k+1;
               s2:='';
               end;
           end;
           strngrd4.Cells[k,i]:=s2;
           s2:='';
           k:=0;
         end;
      end;
            //запоняем даты
            k:=1;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              data1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
              strngrd4.Cells[0,k]:=data1;
              k:=k+1;
            end;
          end;
        end;
       end;

       ///////Складываем значения 1+2 3+4

     for i:=1 to 31 do
     begin
       j:=1;
       k:=50;
       repeat
            strngrd4.Cells[k,i] := FloatToStr(

            (StrToFloat(strngrd4.Cells[j,i])+StrToFloat(strngrd4.Cells[j+1,i]))*80

             );
             j:=j+2;
             k:=k+1;
       until j>48;
       end;
  Form1.btn3Click(Sender);
end;


//БАнк
procedure TForm1.btn10Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin

  if Form1.OpenDialog1.Execute then
   begin
      Form1.Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
      Form1.Import01.Text:=OpenDialog1.FileName;
      Form1.Memo2.Lines.Clear;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          date1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              date1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
            end;
            //выделяем мощность
            if ((s[j]='P')and(s[j+1]='4')and(s[j+2]='8'))
            then
            begin
               k:=j+4;
               while (s[k]<>';') do
               begin
                 data1:=data1+s[k];
                 k:=k+1;
               end;

            end;

          end;
            Memo2.Lines.Add(date1+','+data1);
        end;
      end;
       for i:=1 to Memo2.Lines.Count do
       begin
         s:=Memo2.Lines[i];
         s1:='';
         for j:=1 to Length(s) do
         begin
            if s[j]=',' then s1:=s1+';' else
            if s[j]='.' then s1:=s1+','else s1:=s1+s[j];
         end;
         Memo2.Lines[i]:=s1;
       end;
       end;

      k:=0;
      strngrd5.RowCount:=Memo2.Lines.Count;
      strngrd5.colCount:=49+28;
      for i := 1 to 48 do
       strngrd5.Cells[i,0]:=IntToStr(i);
      for i := 49 to 49+28 do
       strngrd1.Cells[i,0]:=IntToStr(i-49);
      //запоняем днными
      for i:=1 to Memo2.Lines.Count do
      begin
         s1:=Memo2.Lines[i];
         s2:='';
         if Length(s1)>9 then
         begin
           s3:='';
           for j:=9 to Length(s1) do
           begin
             if s1[j]<>';' then
             s2:=s2+s1[j]
             else
             begin
               strngrd5.Cells[k,i]:=s2;
               k:=k+1;
               s2:='';
               end;
           end;
           strngrd5.Cells[k,i]:=s2;
           s2:='';
           k:=0;
         end;
      end;
            //запоняем даты
            k:=1;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              data1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
              strngrd5.Cells[0,k]:=data1;
              k:=k+1;
            end;
          end;
        end;
       end;

       
///////Складываем значения 1+2 3+4

     for i:=1 to 31 do
     begin
       j:=1;
       k:=50;
       repeat
            strngrd5.Cells[k,i] := FloatToStr(

            (StrToFloat(strngrd5.Cells[j,i])+StrToFloat(strngrd5.Cells[j+1,i]))

             );
             j:=j+2;
             k:=k+1;
       until j>48;
       end;

 Form1.btn3Click(Sender);
end;

//Общая
procedure TForm1.btn12Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin

  if Form1.OpenDialog1.Execute then
   begin
      Form1.Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
      Form1.Import01.Text:=OpenDialog1.FileName;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          date1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              date1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
            end;
            //выделяем мощность
            if ((s[j]='P')and(s[j+1]='4')and(s[j+2]='8'))
            then
            begin
               k:=j+4;
               while (s[k]<>';') do
               begin
                 data1:=data1+s[k];
                 k:=k+1;
               end;

            end;

          end;
            Memo2.Lines.Add(date1+','+data1);
        end;
      end;
       for i:=1 to Memo2.Lines.Count do
       begin
         s:=Memo2.Lines[i];
         s1:='';
         for j:=1 to Length(s) do
         begin
            if s[j]=',' then s1:=s1+';' else
            if s[j]='.' then s1:=s1+','else s1:=s1+s[j];
         end;
         Memo2.Lines[i]:=s1;
       end;
       end;

      k:=0;
      strngrd1.RowCount:=Memo2.Lines.Count;
      strngrd1.colCount:=49;
      for i := 1 to 48 do
       strngrd1.Cells[i,0]:=IntToStr(i);

      //запоняем днными
      for i:=1 to Memo2.Lines.Count do
      begin
         s1:=Memo2.Lines[i];
         s2:='';
         if Length(s1)>9 then
         begin
           s3:='';
           for j:=9 to Length(s1) do
           begin
             if s1[j]<>';' then
             s2:=s2+s1[j]
             else
             begin
               strngrd1.Cells[k,i]:=s2;
               k:=k+1;
               s2:='';
               end;
           end;
           strngrd1.Cells[k,i]:=s2;
           s2:='';
           k:=0;
         end;
      end;
            //запоняем даты
            k:=1;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              data1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
              strngrd1.Cells[0,k]:=data1;
              k:=k+1;
            end;
          end;
        end;
       end;

end;

procedure TForm1.btn13Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin
  Memo2.Lines.Clear;
  for i:=1 to 31 do
  begin
     s:='';
     for j:=0 to 49 do
     begin   
          s:=s+strngrd6.Cells[j,i]+';';
     end;
  memo2.Lines.Add(s);
  end;
  if Form1.dlgSave1.Execute then
     begin
        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');
     end;

end   ;





procedure TForm1.btn5Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin
  Memo2.Lines.Clear;
  for i:=1 to 31 do
  begin
     s:='';
     for j:=0 to 24+50+1 do
     begin
          s:=s+strngrd2.Cells[j,i]+';';
     end;
  memo2.Lines.Add(s);
  end;
  if Form1.dlgSave1.Execute then
     begin
        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');
     end;

end   ;


procedure TForm1.btn7Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin
  Memo2.Lines.Clear;
  for i:=1 to 31 do
  begin
     s:='';
     for j:=0 to 24+50+1 do
     begin
          s:=s+strngrd3.Cells[j,i]+';';
     end;
  memo2.Lines.Add(s);
  end;
  if Form1.dlgSave1.Execute then
     begin
        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');
     end;

end   ;


procedure TForm1.btn9Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin
  Memo2.Lines.Clear;
  for i:=1 to 31 do
  begin
     s:='';
     for j:=0 to 24+50+1 do
     begin
          s:=s+strngrd4.Cells[j,i]+';';
     end;
  memo2.Lines.Add(s);
  end;
  if Form1.dlgSave1.Execute then
     begin
        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');
     end;

end   ;


procedure TForm1.btn11Click(Sender: TObject);
var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin
  Memo2.Lines.Clear;
  for i:=1 to 31 do
  begin
     s:='';
     for j:=0 to 24+50+1 do
     begin   
          s:=s+strngrd5.Cells[j,i]+';';
     end;
  memo2.Lines.Add(s);
  end;
  if Form1.dlgSave1.Execute then
     begin
        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');
     end;

end   ;


procedure TForm1.Button2Click(Sender: TObject);
var Exel:Variant;
i,j,k:integer;
begin
 if Form1.SaveDialog1.Execute then
     begin

        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');

        Exel:= CreateOleObject('Excel.Application');
 //Exel.Workbooks.Add;
//form1.Caption:=(ExtractFilePath(Application.ExeName));
 Exel.Workbooks.Open(ExtractFilePath(Application.ExeName)+'book.xlsx');
 Exel.ActiveWorkBook.Sheets.Item[1].Activate;
 //Exel.Range['A2'] := 'Hellow';


     for i:=1 to 31 do
     begin

       k:=50;
       repeat
            Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,k-50+3]:=strngrd2.Cells[k,i];
             k:=k+1;
       until k>50+25;
       end;


 //Exel.ActiveWorkBook.Sheets.Item[2].Activate;
 //Exel.Range['A2'] := 'Hellow';

// Exel.ActiveWorkBook.Sheets.Item[9].Activate;
// Exel.Range['B2'] := 'Hellow 101 kbbkbkbkkb';
   Exel.ActiveWorkBook.SaveAs(SaveDialog1.FileName+'.xlsx');
 Exel.ActiveWorkBook.Close;
Exel.Application.Quit;
     end;


end;
procedure TForm1.Button3Click(Sender: TObject);
var Exel:Variant;
i,j,k:integer;
begin
 if Form1.SaveDialog1.Execute then
     begin

        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');

        Exel:= CreateOleObject('Excel.Application');
 //Exel.Workbooks.Add;
//form1.Caption:=(ExtractFilePath(Application.ExeName));
 Exel.Workbooks.Open(ExtractFilePath(Application.ExeName)+'book.xlsx');
 Exel.ActiveWorkBook.Sheets.Item[1].Activate;
 //Exel.Range['A2'] := 'Hellow';


     for i:=1 to 31 do
     begin

       k:=50;
       repeat
            Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,k-50+3]:=strngrd3.Cells[k,i];
             k:=k+1;
       until k>50+25;
       end;


 //Exel.ActiveWorkBook.Sheets.Item[2].Activate;
 //Exel.Range['A2'] := 'Hellow';

// Exel.ActiveWorkBook.Sheets.Item[9].Activate;
// Exel.Range['B2'] := 'Hellow 101 kbbkbkbkkb';
   Exel.ActiveWorkBook.SaveAs(SaveDialog1.FileName+'.xlsx');
 Exel.ActiveWorkBook.Close;
Exel.Application.Quit;
     end;


end;

procedure TForm1.Button4Click(Sender: TObject);
var Exel:Variant;
i,j,k:integer;
begin
 if Form1.SaveDialog1.Execute then
     begin

        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');

        Exel:= CreateOleObject('Excel.Application');
 //Exel.Workbooks.Add;
//form1.Caption:=(ExtractFilePath(Application.ExeName));
 Exel.Workbooks.Open(ExtractFilePath(Application.ExeName)+'book.xlsx');
 Exel.ActiveWorkBook.Sheets.Item[1].Activate;
 //Exel.Range['A2'] := 'Hellow';


     for i:=1 to 31 do
     begin

       k:=50;
       repeat
            Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,k-50+3]:=strngrd4.Cells[k,i];
             k:=k+1;
       until k>50+25;
       end;


 //Exel.ActiveWorkBook.Sheets.Item[2].Activate;
 //Exel.Range['A2'] := 'Hellow';

// Exel.ActiveWorkBook.Sheets.Item[9].Activate;
// Exel.Range['B2'] := 'Hellow 101 kbbkbkbkkb';
   Exel.ActiveWorkBook.SaveAs(SaveDialog1.FileName+'.xlsx');
 Exel.ActiveWorkBook.Close;
Exel.Application.Quit;
     end;


end;

procedure TForm1.Button5Click(Sender: TObject);
var Exel:Variant;
i,j,k:integer;
begin
 if Form1.SaveDialog1.Execute then
     begin

        Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');

        Exel:= CreateOleObject('Excel.Application');
 //Exel.Workbooks.Add;
//form1.Caption:=(ExtractFilePath(Application.ExeName));
 Exel.Workbooks.Open(ExtractFilePath(Application.ExeName)+'book.xlsx');
 Exel.ActiveWorkBook.Sheets.Item[1].Activate;
 //Exel.Range['A2'] := 'Hellow';


     for i:=1 to 31 do
     begin

       k:=50;
       repeat
            Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,k-50+3]:=strngrd5.Cells[k,i];
             k:=k+1;
       until k>50+25;
       end;


 //Exel.ActiveWorkBook.Sheets.Item[2].Activate;
 //Exel.Range['A2'] := 'Hellow';

// Exel.ActiveWorkBook.Sheets.Item[9].Activate;
// Exel.Range['B2'] := 'Hellow 101 kbbkbkbkkb';
   Exel.ActiveWorkBook.SaveAs(SaveDialog1.FileName+'.xlsx');
 Exel.ActiveWorkBook.Close;
Exel.Application.Quit;
     end;


end;
procedure TForm1.Button6Click(Sender: TObject);
var Exel:Variant;
i,j,k,b:integer;
begin
 if Form1.SaveDialog1.Execute then
     begin
       Form1.Memo2.Lines.SaveToFile(dlgSave1.FileName+'.csv');
       Exel:= CreateOleObject('Excel.Application');
       //Exel.Workbooks.Add;
      //form1.Caption:=(ExtractFilePath(Application.ExeName));
       Exel.Workbooks.Open(ExtractFilePath(Application.ExeName)+'book.xlsx');
       Exel.ActiveWorkBook.Sheets.Item[1].Activate;
       //Exel.Range['A2'] := 'Hellow';
       for i:=1 to 31 do
       begin

             k:=1;
             repeat
                  Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,k+2]:=strngrd6.Cells[k,i];
                   k:=k+1;
             until k>25;
        end;


       //Exel.ActiveWorkBook.Sheets.Item[2].Activate;
       //Exel.Range['A2'] := 'Hellow';

      // Exel.ActiveWorkBook.Sheets.Item[9].Activate;
      // Exel.Range['B2'] := 'Hellow 101 kbbkbkbkkb';
       Exel.ActiveWorkBook.SaveAs(SaveDialog1.FileName+'.xlsx');
       Exel.ActiveWorkBook.Close;
       Exel.Application.Quit;
     end;


end;

procedure TForm1.btnSaveExelClick(Sender: TObject);
var Exel:Variant;
i,j,k,k1:integer;
begin
 if Form1.SaveDialog1.Execute then
     begin
           Exel:= CreateOleObject('Excel.Application');
           //Exel.Workbooks.Add;
          //form1.Caption:=(ExtractFilePath(Application.ExeName));
           Exel.Workbooks.Open(ExtractFilePath(Application.ExeName)+'book.xlsx');
           //Главная//////////////////////////////////////////////
           //Активируем лист1
           Exel.ActiveWorkBook.Sheets.Item[1].Activate;
           //Вывод шапки
           Exel.Range['A5'] := '"Потребитель"  '+edtPotreb.Text;
           Exel.Range['A6'] := '"Договор энергоснабжения  "  '+edtDogovorE.Text;
           Exel.Range['A8'] := 'Сетевая организация:  '+edtEnergoOrg.Text;
           Exel.Range['A3'] := edtSvedenia.Text;
            // Вывод  Таблицы общей
           for i:=1 to 31 do
               begin
                 k:=1;k1:=1;
                 repeat
                      if k=3 then k:=k+1;
                      //Заносим в ячейку данные преобразовав запятую в точку
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,k+2]:=convertZPT(strngrd6.Cells[k1,i]);
                      //Пропускаем чтобы не зависала программа
                      Application.ProcessMessages;
                       lbl6.Caption:='Сохранение Главная';
                       k:=k+1;k1:=k1+1;
                 until k>26;
                 //Увеличиваем бар
                  pb1.Position:=i;
                 end;
           pb1.Position:=0;
           lbl6.Caption:='';
            //Ячейка4///////////////////
            Exel.ActiveWorkBook.Sheets.Item[2].Activate;
           for i:=1 to 31 do
               begin
                 k:=1;k1:=1;
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd2.Cells[k1+49,i]);
                      Application.ProcessMessages;
                       lbl6.Caption:='Сохранение Ячейки4';
                       k:=k+1;k1:=k1+1;
                 until k>26;
                  pb1.Position:=i;
                 end;
           pb1.Position:=0;
           lbl6.Caption:='';
        ///////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                  //Ячейка7///////////////////
            Exel.ActiveWorkBook.Sheets.Item[3].Activate;
           for i:=1 to 31 do
               begin
                 k:=1;k1:=1;
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd1.Cells[k1+49,i]);
                      Application.ProcessMessages;
                       lbl6.Caption:='Сохранение Ячейки7';
                       k:=k+1;k1:=k1+1;
                 until k>26;
                  pb1.Position:=i;
                 end;
           pb1.Position:=0;
           lbl6.Caption:='';
        ///////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


          //CкладГО///////////////////
            Exel.ActiveWorkBook.Sheets.Item[4].Activate;
           for i:=1 to 31 do
               begin
                 k:=1;k1:=1;
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd3.Cells[k1+49,i]);
                      Application.ProcessMessages;
                       lbl6.Caption:='Сохранение Ячейки5 CкладГО';
                       k:=k+1;k1:=k1+1;
                 until k>26;
                  pb1.Position:=i;
                 end;
           pb1.Position:=0;
           lbl6.Caption:='';
        ///////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                  //банк///////////////////
            Exel.ActiveWorkBook.Sheets.Item[7].Activate;
           for i:=1 to 31 do
               begin
                 k:=1;k1:=1;
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd5.Cells[k1+49,i]);
                      Application.ProcessMessages;
                       lbl6.Caption:='Сохранение банк';
                       k:=k+1;k1:=k1+1;
                 until k>26;
                  pb1.Position:=i;
                 end;
           pb1.Position:=0;
           lbl6.Caption:='';
        ///////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
          //АБК///////////////////
            Exel.ActiveWorkBook.Sheets.Item[8].Activate;
           for i:=1 to 31 do
               begin
                 k:=1;k1:=1;
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd4.Cells[k1+49,i]);
                      Application.ProcessMessages;
                       lbl6.Caption:='Сохранение АБК';
                       k:=k+1;k1:=k1+1;
                 until k>26;
                  pb1.Position:=i;
                 end;
           pb1.Position:=0;
           lbl6.Caption:='';
        ///////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        //Ячейка1///////////////////
            Exel.ActiveWorkBook.Sheets.Item[5].Activate;
           for i:=1 to 31 do
               begin
                 k:=1;k1:=1;
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd7.Cells[k1+49,i]);
                      Application.ProcessMessages;
                       lbl6.Caption:='Сохранение Ячейка1';
                       k:=k+1;k1:=k1+1;
                 until k>26;
                  pb1.Position:=i;
                 end;
           pb1.Position:=0;
           lbl6.Caption:='';
        ///////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        //Ячейка14///////////////////
            Exel.ActiveWorkBook.Sheets.Item[6].Activate;
           for i:=1 to 31 do
               begin
                 k:=1;k1:=1;
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd8.Cells[k1+49,i]);
                      Application.ProcessMessages;
                       lbl6.Caption:='Сохранение Ячейка14';
                       k:=k+1;k1:=k1+1;
                 until k>26;
                  pb1.Position:=i;
                 end;
           pb1.Position:=0;
           lbl6.Caption:='';
        ///////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
 //Exel.ActiveWorkBook.Sheets.Item[2].Activate;
 //Exel.Range['A2'] := 'Hellow';

// Exel.ActiveWorkBook.Sheets.Item[9].Activate;
// Exel.Range['B2'] := 'Hellow 101 kbbkbkbkkb';
         Exel.ActiveWorkBook.SaveAs(SaveDialog1.FileName+'.xlsx');
        Exel.ActiveWorkBook.Close;
        Exel.Application.Quit;
        Application.MessageBox('Результаты сохранены.','',MB_OK+MB_ICONWARNING);
     end;


end;
procedure TForm1.btn17Click(Sender: TObject);
 var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin

  if Form1.OpenDialog1.Execute then
   begin
      Form1.Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
      Form1.Import01.Text:=OpenDialog1.FileName;
      Memo2.Lines.Clear;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          date1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              date1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
            end;
            //выделяем мощность
            if ((s[j]='P')and(s[j+1]='4')and(s[j+2]='8'))
            then
            begin
               k:=j+4;
               while (s[k]<>';') do
               begin
                 data1:=data1+s[k];
                 k:=k+1;
               end;

            end;

          end;
            Memo2.Lines.Add(date1+','+data1);
        end;
      end;
       for i:=1 to Memo2.Lines.Count do
       begin
         s:=Memo2.Lines[i];
         s1:='';
         for j:=1 to Length(s) do
         begin
            if s[j]=',' then s1:=s1+';' else
            if s[j]='.' then s1:=s1+','else s1:=s1+s[j];
         end;
         Memo2.Lines[i]:=s1;
       end;
       end;

      k:=0;
      strngrd7.RowCount:=Memo2.Lines.Count;
      strngrd7.colCount:=49+28;
      for i := 1 to 48 do
       strngrd7.Cells[i,0]:=IntToStr(i);
      for i := 49 to 49+28 do
       strngrd7.Cells[i,0]:=IntToStr(i-49);
      //запоняем днными
      for i:=1 to Memo2.Lines.Count do
      begin
         s1:=Memo2.Lines[i];
         s2:='';
         if Length(s1)>9 then
         begin
           s3:='';
           for j:=9 to Length(s1) do
           begin
             if s1[j]<>';' then
             s2:=s2+s1[j]
             else
             begin
               strngrd7.Cells[k,i]:=s2;
               k:=k+1;
               s2:='';
               end;
           end;
           strngrd7.Cells[k,i]:=s2;
           s2:='';
           k:=0;
         end;
      end;
            //запоняем даты
            k:=1;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              data1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
              strngrd7.Cells[0,k]:=data1;
              k:=k+1;
            end;
          end;
        end;
       end;

///////Складываем значения 1+2 3+4

     for i:=1 to 31 do
     begin
       j:=1;
       k:=50;
       repeat
            strngrd7.Cells[k,i] := FloatToStr(

            (StrToFloat(strngrd7.Cells[j,i])+StrToFloat(strngrd7.Cells[j+1,i]))*2400

             );
             j:=j+2;
             k:=k+1;
       until j>48;
       end;

 Form1.btn3Click(Sender);
end;

procedure TForm1.btn14Click(Sender: TObject);
 var
  i,j,k:Integer;
  date1, data1,s,s1,s2,s3:string;
begin

  if Form1.OpenDialog1.Execute then
   begin
      Form1.Memo1.Lines.LoadFromFile(OpenDialog1.FileName);
      Form1.Import01.Text:=OpenDialog1.FileName;
      Memo2.Lines.Clear;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          date1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              date1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
            end;
            //выделяем мощность
            if ((s[j]='P')and(s[j+1]='4')and(s[j+2]='8'))
            then
            begin
               k:=j+4;
               while (s[k]<>';') do
               begin
                 data1:=data1+s[k];
                 k:=k+1;
               end;

            end;

          end;
            Memo2.Lines.Add(date1+','+data1);
        end;
      end;
       for i:=1 to Memo2.Lines.Count do
       begin
         s:=Memo2.Lines[i];
         s1:='';
         for j:=1 to Length(s) do
         begin
            if s[j]=',' then s1:=s1+';' else
            if s[j]='.' then s1:=s1+','else s1:=s1+s[j];
         end;
         Memo2.Lines[i]:=s1;
       end;
       end;

      k:=0;
      strngrd8.RowCount:=Memo2.Lines.Count;
      strngrd8.colCount:=49+28;
      for i := 1 to 48 do
       strngrd8.Cells[i,0]:=IntToStr(i);
      for i := 49 to 49+28 do
       strngrd8.Cells[i,0]:=IntToStr(i-49);
      //запоняем днными
      for i:=1 to Memo2.Lines.Count do
      begin
         s1:=Memo2.Lines[i];
         s2:='';
         if Length(s1)>9 then
         begin
           s3:='';
           for j:=9 to Length(s1) do
           begin
             if s1[j]<>';' then
             s2:=s2+s1[j]
             else
             begin
               strngrd8.Cells[k,i]:=s2;
               k:=k+1;
               s2:='';
               end;
           end;
           strngrd8.Cells[k,i]:=s2;
           s2:='';
           k:=0;
         end;
      end;
            //запоняем даты
            k:=1;
      for i:=1 to Memo1.Lines.Count do
      begin
        if Length(Memo1.Lines[i])>3 then
        begin
          //выделяем дату
          data1:='';
          j:=1;
          s:= Memo1.Lines[i];
          for j:=1 to Length(s) do
          begin
            if ((s[j]='D')and(s[j+1]='A')and(s[j+2]='T'))
            then
            begin
              data1:= s[j+9]+s[j+10]+'/'+s[j+7]+s[j+8]+'/'+ s[j+5]+s[j+6];
              strngrd8.Cells[0,k]:=data1;
              k:=k+1;
            end;
          end;
        end;
       end;

///////Складываем значения 1+2 3+4

     for i:=1 to 31 do
     begin
       j:=1;
       k:=50;
       repeat
            strngrd8.Cells[k,i] := FloatToStr(

            (StrToFloat(strngrd8.Cells[j,i])+StrToFloat(strngrd8.Cells[j+1,i]))*2400

             );
             j:=j+2;
             k:=k+1;
       until j>48;
       end;

 Form1.btn3Click(Sender);
end;

procedure TForm1.chk1Click(Sender: TObject);
begin
Form1.btn3Click(Sender);
end;

procedure TForm1.btn20Click(Sender: TObject);
begin
Application.ProcessMessages;
  DataBase.Show;
end;

procedure TForm1.btnSaveDB7Click(Sender: TObject);
var
    sldb1:TSQLiteDatabase;
    i,k:integer;
    sltb:TSQLiteTable;
begin
       //Открываем базу данных
       sldb1 := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'ElDB.db');
       //sldb.BeginTransaction;
       //Ячейка7///////////////////
       //Цикл по дням
       i:=1;
       for i:=1 to strngrd1.RowCount-1 do
           begin
             k:=1;
             //Цикл по часам
             repeat
                  //Удаляем прошлые записи
                  sldb1.ExecSQL('DELETE FROM val WHERE (ch=1)AND(hour1="'+inttostr(k)+'")AND(date1="'+inttostr(DateTimeToUnix(ConvertDate(strngrd1.Cells[0,i])))+'")');
                  //Вставляем запись
                  sldb1.ExecSQL('insert into val (ch, date1,hour1,value1) values(1, "'+inttostr(DateTimeToUnix(ConvertDate(strngrd1.Cells[0,i])))+'","'+inttostr(k)+'","'+strngrd1.Cells[k+49,i]+'")');
                  //Пропускаем процесс чтобы не зависало
                  Application.ProcessMessages;
                  //Выводим сообщение
                  lbl6.Caption:='Сохранение Ячейки7';
                  k:=k+1;
             until k>24;
             pb1.Position:=i;
           end;
       pb1.Position:=0;
       lbl6.Caption:='';
       //Освобождаем ффай БД
       sldb1.Free;
end;

procedure TForm1.btnSaveDB4Click(Sender: TObject);
var
    sldb1:TSQLiteDatabase;
    i,k:integer;
    sltb:TSQLiteTable;
begin
       //Открываем базу данных
       sldb1 := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'ElDB.db');
       //sldb.BeginTransaction;
       //Ячейка4///////////////////
       //Цикл по дням
       for i:=1 to strngrd2.RowCount-1 do
           begin
             k:=1;
             //Цикл по часам
             repeat
                  //Удаляем прошлые записи
                  sldb1.ExecSQL('DELETE FROM val WHERE (ch=2)AND(hour1="'+inttostr(k)+'")AND(date1="'+inttostr(DateTimeToUnix(ConvertDate(strngrd2.Cells[0,i])))+'")');
                  //Вставляем запись
                  sldb1.ExecSQL('insert into val (ch, date1,hour1,value1) values(2, "'+inttostr(DateTimeToUnix(ConvertDate(strngrd2.Cells[0,i])))+'","'+inttostr(k)+'","'+strngrd2.Cells[k+49,i]+'")');
                  //Пропускаем процесс чтобы не зависало
                  Application.ProcessMessages;
                  //Выводим сообщение
                  lbl6.Caption:='Сохранение Ячейки4';
                  k:=k+1;
             until k>24;
             pb1.Position:=i;
           end;
       pb1.Position:=0;
       lbl6.Caption:='';
       //Освобождаем ффай БД
       sldb1.Free;
end;

procedure TForm1.btnSaveDBSkladClick(Sender: TObject);
var
    sldb1:TSQLiteDatabase;
    i,k:integer;
    sltb:TSQLiteTable;
begin
       //Открываем базу данных
       sldb1 := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'ElDB.db');
       //sldb.BeginTransaction;
       //Склад ГО///////////////////
       //Цикл по дням
       for i:=1 to strngrd3.RowCount-1 do
           begin
             k:=1;
             //Цикл по часам
             repeat
                  //Удаляем прошлые записи
                  sldb1.ExecSQL('DELETE FROM val WHERE (ch=3)AND(hour1="'+inttostr(k)+'")AND(date1="'+inttostr(DateTimeToUnix(ConvertDate(strngrd3.Cells[0,i])))+'")');
                  //Вставляем запись
                  sldb1.ExecSQL('insert into val (ch, date1,hour1,value1) values(3, "'+inttostr(DateTimeToUnix(ConvertDate(strngrd3.Cells[0,i])))+'","'+inttostr(k)+'","'+strngrd3.Cells[k+49,i]+'")');
                  //Пропускаем процесс чтобы не зависало
                  Application.ProcessMessages;
                  //Выводим сообщение
                  lbl6.Caption:='Сохранение СкладГО';
                  k:=k+1;
             until k>24;
             pb1.Position:=i;
           end;
       pb1.Position:=0;
       lbl6.Caption:='';
       //Освобождаем ффай БД
       sldb1.Free;
end;

procedure TForm1.btnSaveDBABKClick(Sender: TObject);
var
    sldb1:TSQLiteDatabase;
    i,k:integer;
    sltb:TSQLiteTable;
begin
       //Открываем базу данных
       sldb1 := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'ElDB.db');
       //sldb.BeginTransaction;
       //Склад АБК///////////////////
       //Цикл по дням
       for i:=1 to strngrd4.RowCount-1 do
           begin
             k:=1;
             //Цикл по часам
             repeat
                  //Удаляем прошлые записи
                  sldb1.ExecSQL('DELETE FROM val WHERE (ch=4)AND(hour1="'+inttostr(k)+'")AND(date1="'+inttostr(DateTimeToUnix(ConvertDate(strngrd4.Cells[0,i])))+'")');
                  //Вставляем запись
                  sldb1.ExecSQL('insert into val (ch, date1,hour1,value1) values(4, "'+inttostr(DateTimeToUnix(ConvertDate(strngrd4.Cells[0,i])))+'","'+inttostr(k)+'","'+strngrd4.Cells[k+49,i]+'")');
                  //Пропускаем процесс чтобы не зависало
                  Application.ProcessMessages;
                  //Выводим сообщение
                  lbl6.Caption:='Сохранение АБК';
                  k:=k+1;
             until k>24;
             pb1.Position:=i;
           end;
       pb1.Position:=0;
       lbl6.Caption:='';
       //Освобождаем ффай БД
       sldb1.Free;
end;

procedure TForm1.btnSaveDBbankClick(Sender: TObject);
var
    sldb1:TSQLiteDatabase;
    i,k:integer;
    sltb:TSQLiteTable;
begin
       //Открываем базу данных
       sldb1 := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'ElDB.db');
       //sldb.BeginTransaction;
       //Склад БАНК///////////////////
       //Цикл по дням
       for i:=1 to strngrd5.RowCount-1 do
           begin
             k:=1;
             //Цикл по часам
             repeat
                  //Удаляем прошлые записи
                  sldb1.ExecSQL('DELETE FROM val WHERE (ch=5)AND(hour1="'+inttostr(k)+'")AND(date1="'+inttostr(DateTimeToUnix(ConvertDate(strngrd5.Cells[0,i])))+'")');
                  //Вставляем запись
                  sldb1.ExecSQL('insert into val (ch, date1,hour1,value1) values(5, "'+inttostr(DateTimeToUnix(ConvertDate(strngrd5.Cells[0,i])))+'","'+inttostr(k)+'","'+strngrd5.Cells[k+49,i]+'")');
                  //Пропускаем процесс чтобы не зависало
                  Application.ProcessMessages;
                  //Выводим сообщение
                  lbl6.Caption:='Сохранение БАНК';
                  k:=k+1;
             until k>24;
             pb1.Position:=i;
           end;
       pb1.Position:=0;
       lbl6.Caption:='';
       //Освобождаем ффай БД
       sldb1.Free;
end;

procedure TForm1.btnSaveDB14Click(Sender: TObject);
var
    sldb1:TSQLiteDatabase;
    i,k:integer;
    sltb:TSQLiteTable;
begin
       //Открываем базу данных
       sldb1 := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'ElDB.db');
       //sldb.BeginTransaction;
       //Ячейка 14///////////////////
       //Цикл по дням
       for i:=1 to strngrd8.RowCount-1 do
           begin
             k:=1;
             //Цикл по часам
             repeat
                  //Удаляем прошлые записи
                  sldb1.ExecSQL('DELETE FROM val WHERE (ch=7)AND(hour1="'+inttostr(k)+'")AND(date1="'+inttostr(DateTimeToUnix(ConvertDate(strngrd8.Cells[0,i])))+'")');
                  //Вставляем запись
                  sldb1.ExecSQL('insert into val (ch, date1,hour1,value1) values(7, "'+inttostr(DateTimeToUnix(ConvertDate(strngrd8.Cells[0,i])))+'","'+inttostr(k)+'","'+strngrd8.Cells[k+49,i]+'")');
                  //Пропускаем процесс чтобы не зависало
                  Application.ProcessMessages;
                  //Выводим сообщение
                  lbl6.Caption:='Сохранение Ячейка 14';
                  k:=k+1;
             until k>24;
             pb1.Position:=i;
           end;
       pb1.Position:=0;
       lbl6.Caption:='';
       //Освобождаем ффай БД
       sldb1.Free;
end;

procedure TForm1.btnSaveDB_q1Click(Sender: TObject);
var
    sldb1:TSQLiteDatabase;
    i,k:integer;
    sltb:TSQLiteTable;
begin
       //Открываем базу данных
       sldb1 := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'ElDB.db');
       //Ячейка 14///////////////////
       //Цикл по дням
       for i:=1 to strngrd7.RowCount-1 do
           begin
             k:=1;
             //Цикл по часам
             repeat
                  //Удаляем прошлые записи
                  sldb1.ExecSQL('DELETE FROM val WHERE (ch=6)AND(hour1="'+inttostr(k)+'")AND(date1="'+inttostr(DateTimeToUnix(ConvertDate(strngrd7.Cells[0,i])))+'")');
                  //Вставляем запись
                  sldb1.ExecSQL('insert into val (ch, date1,hour1,value1) values(6, "'+inttostr(DateTimeToUnix(ConvertDate(strngrd7.Cells[0,i])))+'","'+inttostr(k)+'","'+strngrd7.Cells[k+49,i]+'")');
                  //Пропускаем процесс чтобы не зависало
                  Application.ProcessMessages;
                  //Выводим сообщение
                  lbl6.Caption:='Сохранение Ячейка 14';
                  k:=k+1;
             until k>24;
             pb1.Position:=i;
           end;
       pb1.Position:=0;
       lbl6.Caption:='';
       //Освобождаем ффай БД
       sldb1.Free;
end;

end.
