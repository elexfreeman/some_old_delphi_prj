program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  DataB in 'DataB.pas' {DataBase},
  SQLite3 in 'SQLite3.pas',
  SQLiteTable3 in 'SQLiteTable3.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Конвертер';
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TDataBase, DataBase);
  Application.Run;
end.
