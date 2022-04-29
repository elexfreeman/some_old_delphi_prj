unit DataB;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, Menus, ComCtrls, ExtCtrls, StdCtrls, Buttons,ComObj;

type
  TDataBase = class(TForm)
    pnl1: TPanel;
    pnl2: TPanel;
    pgc1: TPageControl;
    ts1: TTabSheet;
    ts2: TTabSheet;
    ts3: TTabSheet;
    ts4: TTabSheet;
    ts5: TTabSheet;
    ts6: TTabSheet;
    ts7: TTabSheet;
    ts8: TTabSheet;
    dtp1: TDateTimePicker;
    dtp2: TDateTimePicker;
    mm1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N6: TMenuItem;
    strngrd1: TStringGrid;
    strngrd2: TStringGrid;
    strngrd3: TStringGrid;
    strngrd4: TStringGrid;
    strngrd5: TStringGrid;
    strngrd6: TStringGrid;
    strngrd7: TStringGrid;
    strngrd8: TStringGrid;
    Memo: TMemo;
    btn1: TButton;
    btn2: TButton;
    pb1: TProgressBar;
    chk1: TCheckBox;
    chk2: TCheckBox;
    chk3: TCheckBox;
    chk4: TCheckBox;
    chk5: TCheckBox;
    chk6: TCheckBox;
    chk7: TCheckBox;
    btn12: TBitBtn;
    dlgSave1: TSaveDialog;
    lbl6: TLabel;
    dlgSave2: TSaveDialog;
    dlgOpen1: TOpenDialog;
    procedure N4Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure chk1Click(Sender: TObject);
    procedure chk2Click(Sender: TObject);
    procedure chk3Click(Sender: TObject);
    procedure chk4Click(Sender: TObject);
    procedure chk5Click(Sender: TObject);
    procedure chk6Click(Sender: TObject);
    procedure chk7Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btn12Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;



function DateTimeToUnix(ConvDate: TDateTime): Longint;
function UnixToDateTime(USec: Longint): TDateTime;
function ConvertDate(s:string):tdate;
var
  DataBase: TDataBase;
  y_7: array of array of real;
  y_4: array of array of real;
  y_sklad: array of array of real;
  y_abk: array of array of real;
  y_bank: array of array of real;
  y_1: array of array of real;
  y_14: array of array of real;
  y_summa: array of array of real;
  maxRow:integer;

implementation
uses SQLite3, SQLiteTable3;

{$R *.dfm}


const
   // Sets UnixStartDate to TDateTime of 01/01/1970 
  UnixStartDate: TDateTime = 25569.0;

var  minDate,maxDate,tempDate:TDate;
     minDate1,maxDate1:Integer;
     dbname:string;
function convertZPT(s:string):string;
 var i:Integer;
 s1:string;
 begin
    s1:='';
    for i:=1 to length(s) do
     if s[i]=',' then s1:=s1+'.' else s1:=s1+s[i];
    convertZPT:= s1;
 end;

 function DateTimeToUnix(ConvDate: TDateTime): Longint;
 begin
   //example: DateTimeToUnix(now); 
  Result := Round((ConvDate - UnixStartDate) * 86400);
 end;

 function UnixToDateTime(USec: Longint): TDateTime;
 begin
   //Example: UnixToDateTime(1003187418);
  Result := (Usec / 86400) + UnixStartDate;
 end;

function IncDate(USec: Longint): Longint;
 begin
   //Example: UnixToDateTime(1003187418);
  Result := (Usec + 86400);
 end;

  function convertTCK(s:string):string;
 var i:Integer;
 s1:string;
 begin
    s1:='';
    for i:=1 to length(s) do
     if s[i]='.' then s1:=s1+',' else s1:=s1+s[i];
    convertTCK:= s1;
 end;

procedure ShowSumm ;
var
sldb:TSQLiteDatabase;
sltb:TSQLiteTable;
   i,j:integer;
begin
   sldb := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'ElDB.db');
   //select sum(value1) as summa, date1, hour1  from val where ch=1 group by date1 order by hour1
   //Показать Ячейку1
   Application.ProcessMessages;
      try
          sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE ch = 1 order By date1');
          try
            j:=1;
            //strngrd1.RowCount:=sltb.Count div 24;
           // strngrd1.colCount:=25;
            for i := 0 to sltb.Count - 1 do
              begin
              //  strngrd1.Cells[strtoint(sltb.FieldAsString(3)),GetDay(sltb.FieldAsString(2))]:=sltb.FieldAsString(4);
              //  strngrd1.Cells[0,GetDay(sltb.FieldAsString(2))]:=sltb.FieldAsString(2);
                Application.ProcessMessages;
                sltb.Next;
                j:=j+1;
                if j>24 then j:=1;
              end;
          finally
          end;
        finally
        end;
     Application.ProcessMessages;

end;

function GetMonth(s:string):Integer;
begin
  DataBase.Memo.Lines.Add(s);
  GetMonth:=0;
end;

function GetDay(s:string):Integer;
var  i:Integer;
temp:string;
begin
  temp:='';
  GetDay:=1;
  if Length(s)>0 then
  begin
    i:=1;
    repeat
      temp:=temp+s[i];
      i:=i+1;
    until (i>Length(s))or(s[i]='/');
    GetDay:=StrToInt(temp);
  end;

end;

function ConvertDate(s:string):tdate;
var i:Integer;
temp:string;
begin
   temp:='';
   for i:=1 to Length(s) do
    if s[i]='/' then temp:=temp+'.' else temp:=temp+s[i];
ConvertDate:=StrToDate(temp);

end;

procedure Summa_1;
var i,j,k:Integer;
summa_7, summa_7_f,summa_4, summa_4_f,summa_sk, summa_sk_f,summa_abk, summa_abk_f,summa_b, summa_b_f
, summa_1, summa_1_f, summa_14, summa_14_f: real;
begin
  with DataBase do
  begin
 tempDate:=minDate;
 i:=1;

 for j:=1 to 25 do
   begin
    strngrd1.Cells[j,0]:=IntToStr(j);
    strngrd2.Cells[j,0]:=IntToStr(j);
    strngrd3.Cells[j,0]:=IntToStr(j);
    strngrd4.Cells[j,0]:=IntToStr(j);
    strngrd5.Cells[j,0]:=IntToStr(j);
    strngrd6.Cells[j,0]:=IntToStr(j);
    strngrd7.Cells[j,0]:=IntToStr(j);
    strngrd8.Cells[j,0]:=IntToStr(j);
    if j=25 then strngrd8.Cells[j,0]:='Сумма';
    Application.ProcessMessages;
   end;
   i:=1;
 tempDate:=minDate;
 repeat
   for j:=1 to 24 do
   begin
    if chk1.Checked then  strngrd8.Cells[j,i]:=FloatToStr(strToFloat(strngrd8.Cells[j,i])+strToFloat(strngrd1.Cells[j,i]));
    if chk2.Checked then  strngrd8.Cells[j,i]:=FloatToStr(strToFloat(strngrd8.Cells[j,i])+strToFloat(strngrd2.Cells[j,i]));
    if chk3.Checked then  strngrd8.Cells[j,i]:=FloatToStr(strToFloat(strngrd8.Cells[j,i])+strToFloat(strngrd3.Cells[j,i]));
    if chk4.Checked then  strngrd8.Cells[j,i]:=FloatToStr(strToFloat(strngrd8.Cells[j,i])-strToFloat(strngrd4.Cells[j,i]));
    if chk5.Checked then  strngrd8.Cells[j,i]:=FloatToStr(strToFloat(strngrd8.Cells[j,i])-strToFloat(strngrd5.Cells[j,i]));
    if chk6.Checked then  strngrd8.Cells[j,i]:=FloatToStr(strToFloat(strngrd8.Cells[j,i])+strToFloat(strngrd7.Cells[j,i]));
    if chk7.Checked then  strngrd8.Cells[j,i]:=FloatToStr(strToFloat(strngrd8.Cells[j,i])+strToFloat(strngrd6.Cells[j,i]));
    Application.ProcessMessages;
   end;

   strngrd8.Cells[0,i]:=strngrd1.Cells[0,i];
   i:=i+1;
   tempDate:=tempDate+1;
 until tempDate=maxDate;


end;
end;

procedure SummaTable;
var
  i,j,k:Integer;
begin
    i:=1;
 tempDate:=minDate;
with DataBase do
begin
  tempDate:=minDate;
 i:=1;
   repeat
      for j:=1 to 25 do
       begin
           y_summa[j,i]:=0;
       end;
       i:=i+1;
       tempDate:=tempDate+1;
       Application.ProcessMessages;
    until tempDate>=maxDate;

  tempdate:=minDate;
  k:=1;
 repeat
   for j:=1 to 24 do
   begin
    if chk1.Checked then  y_summa[j,k]:=y_summa[j,k]+y_7[j,k];
    if chk2.Checked then  y_summa[j,k]:=y_summa[j,k]+y_4[j,k];
    if chk3.Checked then  y_summa[j,k]:=y_summa[j,k]+y_sklad[j,k];
    if chk4.Checked then  y_summa[j,k]:=y_summa[j,k]-y_abk[j,k];
    if chk5.Checked then  y_summa[j,k]:=y_summa[j,k]-y_bank[j,k];
    if chk6.Checked then  y_summa[j,k]:=y_summa[j,k]+y_1[j,k];
    if chk7.Checked then  y_summa[j,k]:=y_summa[j,k]+y_14[j,k];
    Application.ProcessMessages;
    y_summa[25,k]:=y_summa[25,k]+y_summa[j,k];
   end;

   //strngrd8.Cells[0,k]:=strngrd1.Cells[0,k];
   tempdate:=tempDate+1;
   k:=k+1;
 until tempdate>=maxDate;

  for j:=1 to 25 do
   begin
    strngrd1.Cells[j,0]:=IntToStr(j);
    strngrd2.Cells[j,0]:=IntToStr(j);
    strngrd3.Cells[j,0]:=IntToStr(j);
    strngrd4.Cells[j,0]:=IntToStr(j);
    strngrd5.Cells[j,0]:=IntToStr(j);
    strngrd6.Cells[j,0]:=IntToStr(j);
    strngrd7.Cells[j,0]:=IntToStr(j);
    strngrd8.Cells[j,0]:=IntToStr(j);
    if j=25 then strngrd8.Cells[j,0]:='Сумма';
    Application.ProcessMessages;
   end;
   strngrd8.Cells[0,0]:='Дата/Часы';
end;
end;

procedure readDB(all:Boolean;nDate,kDate:string;db:string );
var sltb:TSQLiteTable;
   i,j:integer;
   sldb:TSQLiteDatabase;
   temp_e:Integer;
   M_day,m_month,m_year:Word;
begin
with DataBase do
begin
       pb1.Position:=0;
       sldb := TSQLiteDatabase.Create(db);

      if all then
      begin
         //Вычисляем макс кол-во строк ////////////////////////////////////
         try
                sltb := sldb.GetTable('select min(date1), max(date1) from val');
                  minDate:=UnixToDateTime(sltb.FieldAsInteger(0));
                  maxDate:=UnixToDateTime(sltb.FieldAsInteger(1));
         finally
         end;
       end
       else
       begin
         minDate:=StrToDate(nDate);
         maxDate:=StrToDate(kDate);

       end;
       Application.ProcessMessages;
       tempDate:=minDate;
       maxRow:=1;
       repeat
         Inc(maxRow);
         tempDate:=tempDate+1;
        until tempDate>=maxDate;

     /////////////////////////////////////////////////////////////////////
    //ЗАполняем нулями все
   SetLength(y_7,26,maxRow+1);
    SetLength(y_4,26,maxRow+1);
    SetLength(y_sklad,26,maxRow+1);
    SetLength(y_abk,26,maxRow+1);
    SetLength(y_bank,26,maxRow+1);
   SetLength(y_1,26,maxRow+1);
    SetLength(y_14,26,maxRow+1);
    SetLength(y_summa,26,maxRow+1);
      strngrd1.RowCount:=maxRow;
      strngrd1.colCount:=25;
      strngrd2.RowCount:=maxRow;
      strngrd2.colCount:=25;
      strngrd3.RowCount:=maxRow;
      strngrd3.colCount:=25;
      strngrd4.RowCount:=maxRow;
      strngrd4.colCount:=25;
      strngrd5.RowCount:=maxRow;
      strngrd5.colCount:=25;
      strngrd6.RowCount:=maxRow;
      strngrd6.colCount:=25;
      strngrd7.RowCount:=maxRow;
      strngrd7.colCount:=25;

      strngrd8.colCount:=26;
      strngrd8.RowCount:=maxRow;
     tempDate:=minDate;
     for i:=1 to maxRow do
     begin
       strngrd1.Cells[0,i]:=DateToStr(tempDate);
       strngrd2.Cells[0,i]:=DateToStr(tempDate);
       strngrd3.Cells[0,i]:=DateToStr(tempDate);
       strngrd4.Cells[0,i]:=DateToStr(tempDate);
       strngrd5.Cells[0,i]:=DateToStr(tempDate);
       strngrd6.Cells[0,i]:=DateToStr(tempDate);
       strngrd7.Cells[0,i]:=DateToStr(tempDate);
       strngrd8.Cells[0,i]:=DateToStr(tempDate);
       for j:=1 to 24 do
       begin
         y_7[j,i]:=0;
         y_4[j,i]:=0;
         y_abk[j,i]:=0;
         y_bank[j,i]:=0;
         y_sklad[j,i]:=0;
         y_1[j,i]:=0;
         y_14[j,i]:=0;
         y_summa[j,i]:=0;
       end;
       tempDate:=tempDate+1;
     end;
     pb1.Position:=1;
   //Показать Ячейку7 ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
       // memo.Lines.Clear;
        //memo.lines.Add(inttostr(datetimetounix(StrToDate('01.01.2011'))));
       // memo.lines.Add(inttostr(1295308800));
        //memo.lines.Add(datetostr(StrToDate(01.01.2011)));
        repeat

         Application.ProcessMessages;
          try
               //  Memo.Lines.Add('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 1)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
                 sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 1)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    //tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));

                    DecodeDate(tempDate,m_year,m_day,m_month);
                    y_7[StrToInt(sltb.FieldAsString(3)),temp_e]:=StrToFloat(convertTCK(sltb.FieldAsString(4)));
                    //y_7[StrToInt(sltb.FieldAsString(3)),temp_e]:=sltb.FieldAsDouble(4);
                    //strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
           // Memo.Lines.Add(DateToStr(tempDate)+inttostr(datetimeToUnix(tempDate)));

         Application.ProcessMessages;
          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=2;
     //=============================================================================
  //Показать Ячейку4 ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
         Application.ProcessMessages;
          try
              
                 sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 2)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
             try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    y_4[StrToInt(sltb.FieldAsString(3)),temp_e]:=StrToFloat(convertTCK(sltb.FieldAsString(4)));
                    //strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;
          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=3;
     //=============================================================================
  //Показать Ячейку sklad ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
         Application.ProcessMessages;
          try                
                 sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 3)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    y_sklad[StrToInt(sltb.FieldAsString(3)),temp_e]:=StrToFloat(convertTCK(sltb.FieldAsString(4)));
                    //strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;
          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=4;
     //=============================================================================
       //Показать Ячейку abk++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
         Application.ProcessMessages;
          try
              
                 sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 4)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
               try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    y_abk[StrToInt(sltb.FieldAsString(3)),temp_e]:=StrToFloat(convertTCK(sltb.FieldAsString(4)));
                    //strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;
          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=5;
     //=============================================================================
  //Показать Ячейку bank ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
         Application.ProcessMessages;
          try
              
                 sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 5)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    y_bank[StrToInt(sltb.FieldAsString(3)),temp_e]:=StrToFloat(convertTCK(sltb.FieldAsString(4)));
                    //strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;
          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=6;
     //=============================================================================
  //Показать Ячейку1 ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
         Application.ProcessMessages;
          try
              
                 sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 6)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')') ;
               try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    y_1[StrToInt(sltb.FieldAsString(3)),temp_e]:=StrToFloat(convertTCK(sltb.FieldAsString(4)));
                    //strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;
          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=7;
     //=============================================================================
  //Показать Ячейку 14++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
         Application.ProcessMessages;
          try

                 sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 7)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')') ;
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    y_14[StrToInt(sltb.FieldAsString(3)),temp_e]:=StrToFloat(convertTCK(sltb.FieldAsString(4)));
                    //strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;
          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=8;
     //=============================================================================
sldb.Free;
end;
end;

procedure ShowTable;
var i,j:Integer;
begin
with DataBase do
begin
   for i:=1 to maxRow do
   for j:=1 to 25 do
   begin
   strngrd1.Cells[j,i]:=floattostr(y_7[j,i]);
   strngrd2.Cells[j,i]:=floattostr(y_4[j,i]);
   strngrd3.Cells[j,i]:=floattostr(y_sklad[j,i]);
   strngrd4.Cells[j,i]:=floattostr(y_abk[j,i]);
   strngrd5.Cells[j,i]:=floattostr(y_bank[j,i]);
   strngrd6.Cells[j,i]:=floattostr(y_14[j,i]);
   strngrd7.Cells[j,i]:=floattostr(y_1[j,i]);
   strngrd8.Cells[j,i]:=floattostr(y_summa[j,i]);
   end;
    pb1.Position:=9;
end;
end;


procedure ShowAll;
var sltb:TSQLiteTable;
   i,j:integer;
   sldb:TSQLiteDatabase;
   maxRow,temp_e:Integer;
   M_day,m_month,m_year:Word;

begin
with DataBase do
begin
       pb1.Position:=0;
       sldb := TSQLiteDatabase.Create(ExtractFilePath(Application.ExeName)+'ElDB.db');
       //Вычисляем макс кол-во строк ////////////////////////////////////
       try
              sltb := sldb.GetTable('select min(date1), max(date1) from val');
                minDate:=UnixToDateTime(sltb.FieldAsInteger(0));
                maxDate:=UnixToDateTime(sltb.FieldAsInteger(1));
       finally
       end;

       Application.ProcessMessages;
       tempDate:=minDate;
       maxRow:=1;
       repeat
         Inc(maxRow);
         tempDate:=tempDate+1;
        until tempDate>=maxDate;

     /////////////////////////////////////////////////////////////////////

    //ЗАполняем нулями все
      strngrd1.RowCount:=maxRow;
      strngrd1.colCount:=25;
      strngrd2.RowCount:=maxRow;
      strngrd2.colCount:=25;
      strngrd3.RowCount:=maxRow;
      strngrd3.colCount:=25;
      strngrd4.RowCount:=maxRow;
      strngrd4.colCount:=25;
      strngrd5.RowCount:=maxRow;
      strngrd5.colCount:=25;
      strngrd6.RowCount:=maxRow;
      strngrd6.colCount:=25;
      strngrd7.RowCount:=maxRow;
      strngrd7.colCount:=25;

      strngrd8.colCount:=26;
      strngrd8.RowCount:=maxRow;
     tempDate:=minDate;
     for i:=1 to maxRow do
     begin
       strngrd1.Cells[0,i]:=DateToStr(tempDate);
       strngrd2.Cells[0,i]:=DateToStr(tempDate);
       strngrd3.Cells[0,i]:=DateToStr(tempDate);
       strngrd4.Cells[0,i]:=DateToStr(tempDate);
       strngrd5.Cells[0,i]:=DateToStr(tempDate);
       strngrd6.Cells[0,i]:=DateToStr(tempDate);
       strngrd7.Cells[0,i]:=DateToStr(tempDate);
       strngrd8.Cells[0,i]:=DateToStr(tempDate);
       for j:=1 to 26 do
       begin
         strngrd1.Cells[j,i]:='0';
         strngrd2.Cells[j,i]:='0';
         strngrd3.Cells[j,i]:='0';
         strngrd4.Cells[j,i]:='0';
         strngrd5.Cells[j,i]:='0';
         strngrd6.Cells[j,i]:='0';
         strngrd7.Cells[j,i]:='0';
        // strngrd8.Cells[j,i]:='0';
       end;
       tempDate:=tempDate+1;
     end;
     pb1.Position:=1;

    //Показать Ячейку7 ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
         Application.ProcessMessages;
          try
              sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 1)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              //Memo.Lines.Add('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 1)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')') ;
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=sltb.FieldAsString(4);
                    strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd1.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;
          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=2;
     //=============================================================================
    //Показать Ячейку4 ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
        Application.ProcessMessages;
          try
              sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 2)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              //Memo.Lines.Add('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 1)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')') ;
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    strngrd2.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=sltb.FieldAsString(4);
                    strngrd2.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd2.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;

          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=3;
     //=============================================================================
    //Показать SkladGO ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
        Application.ProcessMessages;
          try
              sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 3)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              //Memo.Lines.Add('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 1)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')') ;
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    strngrd3.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=sltb.FieldAsString(4);
                    strngrd3.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd3.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;

          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=4;
     //=============================================================================
     //Показать ABK ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
       Application.ProcessMessages;
          try
              sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 4)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              //Memo.Lines.Add('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 1)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')') ;
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    strngrd4.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=sltb.FieldAsString(4);
                    strngrd4.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd4.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;

          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=5;
     //=============================================================================
     //Показать BANAK ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
       Application.ProcessMessages;
          try
              sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 5)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              //Memo.Lines.Add('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 1)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')') ;
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    strngrd5.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=sltb.FieldAsString(4);
                    strngrd5.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd5.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;

          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=6;
     //=============================================================================
     //Показать 1 ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
       Application.ProcessMessages;
          try
              sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 6)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              //Memo.Lines.Add('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 1)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')') ;
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    strngrd7.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=sltb.FieldAsString(4);
                    strngrd7.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd7.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;

          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=7;
     //=============================================================================
     //Показать 14 ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        tempDate:=minDate;
        temp_e:=1;
        repeat
       Application.ProcessMessages;
          try
              sltb := sldb.GetTable('select id,ch, date1, hour1 ,value1 from val WHERE (ch = 7)AND(date1='+inttostr(DateTimeToUnix(tempDate))+')');
              try
                for i := 0 to sltb.Count - 1 do
                  begin
                    tempDate:= UnixToDateTime(StrToInt(sltb.FieldAsString(2)));
                    DecodeDate(tempDate,m_year,m_day,m_month);
                    strngrd6.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=sltb.FieldAsString(4);
                    strngrd6.Cells[strtoint(sltb.FieldAsString(3)),temp_e]:=convertTCK(strngrd6.Cells[strtoint(sltb.FieldAsString(3)),temp_e]);
                    Application.ProcessMessages;
                    sltb.Next;
                  end;
                  temp_e:=temp_e+1;
              finally
              end;
            finally
            end;
         Application.ProcessMessages;

          tempDate:=tempDate+1;
        until tempDate>=maxDate;
        pb1.Position:=8;
     //=============================================================================

    // Высчитываем общею таблицу
     sltb.Free;



     pb1.Position:=9;
     pb1.Position:=0;
 end;

 Summa_1;
 end;




procedure TDataBase.N4Click(Sender: TObject);
var
sldb:TSQLiteDatabase;
sltb:TSQLiteTable;
   i:integer;
   sql1:string;
begin
dbname:=ExtractFilePath(Application.ExeName)+'ElDB.db';
dbname:=ExtractShortPathName(dbname);
  if not DeleteFile(dbname) then ShowMessage('Ошибка удаления');
  sldb := TSQLiteDatabase.Create(dbname);

  try
    if not sldb.TableExists('ch') then
       begin
       sldb.ExecSQL('CREATE TABLE ch (id INTEGER PRIMARY KEY, caption TEXT)');
       sldb.ExecSQL('insert into ch (caption) values("Ячейка7")');
       sldb.ExecSQL('insert into ch (caption) values("Ячейка4")');
       sldb.ExecSQL('insert into ch (caption) values("СкладГО")');
       sldb.ExecSQL('insert into ch (caption) values("АБК")');
       sldb.ExecSQL('insert into ch (caption) values("Банк")');
       sldb.ExecSQL('insert into ch (caption) values("Ячейка1")');
       sldb.ExecSQL('insert into ch (caption) values("Ячейка14")');
       sldb.ExecSQL('CREATE TABLE val (id INTEGER PRIMARY KEY, ch integer, date1 integer, hour1 integer, value1 double)');
        Memo.Lines.BeginUpdate;
        try
          Memo.clear;
          sltb := sldb.GetTable('select id,caption from ch');
          try
          for i := 0 to sltb.Count - 1 do
            begin
              memo.Lines.add(sltb.FieldAsString(0)+', '+sltb.FieldAsString(1));
              sltb.Next;
            end;
          finally
            sltb.Free;
          end;
        finally
          Memo.Lines.EndUpdate;
        end;
       ShowMessage('База данных "'+dbname+'" Создана');
       end;
  except
    ShowMessage('При создании базы произошла ошибка.');
    //Application.terminate;
  end;
   sldb.Free;
end;

procedure TDataBase.btn1Click(Sender: TObject);
var sltb:TSQLiteTable;
   i:integer;
   sldb:TSQLiteDatabase;
   date1:TDateTime;
   k:TTimeStamp;
begin
  readDB(false,DateToStr(dtp1.Date),DateToStr(dtp2.Date),dbname);
  SummaTable;
  ShowTable;

end;

procedure TDataBase.btn2Click(Sender: TObject);
begin
  readDB(true,'12.12.2012','12.12.2012',dbname);

  SummaTable;
  ShowTable;
end;

procedure TDataBase.FormShow(Sender: TObject);
begin
  Application.ProcessMessages;
//DataBase.btn2Click(Sender);
end;

procedure TDataBase.chk1Click(Sender: TObject);
begin

  SummaTable;
  ShowTable;
end;

procedure TDataBase.chk2Click(Sender: TObject);
begin

  SummaTable;
  ShowTable;;
end;

procedure TDataBase.chk3Click(Sender: TObject);
begin

  SummaTable;
  ShowTable;;
end;

procedure TDataBase.chk4Click(Sender: TObject);
begin

  SummaTable;
  ShowTable;;
end;

procedure TDataBase.chk5Click(Sender: TObject);
begin

  SummaTable;
  ShowTable;;
end;

procedure TDataBase.chk6Click(Sender: TObject);
begin

  SummaTable;
  ShowTable;;
end;

procedure TDataBase.chk7Click(Sender: TObject);
begin

  SummaTable;
  ShowTable;;
end;

procedure TDataBase.FormDestroy(Sender: TObject);
begin
Finalize(y_7);
Finalize(y_4);
Finalize(y_sklad);
Finalize(y_abk);
Finalize(y_bank);
Finalize(y_1);
Finalize(y_14);
Finalize(y_summa);

end;

procedure TDataBase.btn12Click(Sender: TObject);
var Exel:Variant;
i,j,k,k1:integer;
begin
 if dlgSave1.Execute then
     begin
           Exel:= CreateOleObject('Excel.Application');
           //Exel.Workbooks.Add;
          //form1.Caption:=(ExtractFilePath(Application.ExeName));
           Exel.Workbooks.Open(ExtractFilePath(Application.ExeName)+'bookALL.xlsx');
           //Главная//////////////////////////////////////////////
           //Активируем лист1
           Exel.ActiveWorkBook.Sheets.Item[1].Activate;
           //Вывод шапки
           //Exel.Range['A5'] := '"Потребитель"  '+edtPotreb.Text;
           //Exel.Range['A6'] := '"Договор энергоснабжения  "  '+edtDogovorE.Text;
           //Exel.Range['A8'] := 'Сетевая организация:  '+edtEnergoOrg.Text;
          // Exel.Range['A3'] := edtSvedenia.Text;
            // Вывод  Таблицы общей
           for i:=1 to maxRow do
               begin
                 k:=1;k1:=1;
                 Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,1]:=strngrd6.Cells[0,i];
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
           for i:=1 to maxRow do
               begin
                 k:=1;k1:=1;
                 Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,1]:=strngrd6.Cells[0,i];
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd2.Cells[k1,i]);
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
           for i:=1 to maxRow do
               begin
                 k:=1;k1:=1;
                 Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,1]:=strngrd6.Cells[0,i];
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd1.Cells[k1,i]);
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
           for i:=1 to maxRow do
               begin
                 k:=1;k1:=1;
                 Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,1]:=strngrd6.Cells[0,i];
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd3.Cells[k1,i]);
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
           for i:=1 to maxRow do
               begin
                 k:=1;k1:=1;
                 Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,1]:=strngrd6.Cells[0,i];
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd5.Cells[k1,i]);
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
           for i:=1 to maxRow do
               begin
                 k:=1;k1:=1;
                 Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,1]:=strngrd6.Cells[0,i];
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd4.Cells[k1,i]);
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
           for i:=1 to maxRow do
               begin
                 k:=1;k1:=1;
                 Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,1]:=strngrd6.Cells[0,i];
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd7.Cells[k1,i]);
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
           for i:=1 to maxRow do
               begin
                 k:=1;k1:=1;
                 Exel.ActiveWorkBook.ActiveSheet.Cells[i+14,1]:=strngrd6.Cells[0,i];
                 repeat
                      if k=3 then k:=k+1;
                      Exel.ActiveWorkBook.ActiveSheet.Cells[i+4,k+1]:=convertZPT(strngrd8.Cells[k1,i]);
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
         Exel.ActiveWorkBook.SaveAs(dlgSave1.FileName+'.xlsx');
        Exel.ActiveWorkBook.Close;
        Exel.Application.Quit;
        Application.MessageBox('Результаты сохранены.','',MB_OK+MB_ICONWARNING);
     end;
end;

procedure TDataBase.N2Click(Sender: TObject);
begin
if dlgSave2.Execute then
     begin
       CopyFile(PAnsiChar(dbname), PAnsiChar(dlgSave2.Filename+'.db'), true);
      //dbname:=dlg
     end;

end;

procedure TDataBase.FormCreate(Sender: TObject);
begin
dbname:=ExtractFilePath(Application.ExeName)+'ElDB.db';
dbname:=ExtractShortPathName(dbname);
//ShowMessage(dbname);
end;

procedure TDataBase.N3Click(Sender: TObject);
begin
   if dlgOpen1.Execute then
     begin
      dbname:=dlgOpen1.FileName;
      dbname:=ExtractShortPathName(dbname);
      //ShowMessage(dbname);
        readDB(true,'12.12.2012','12.12.2012',dbname);
        SummaTable;
        ShowTable;
     end;
end;

end.
