program bbxt2;

uses
  Forms,
  gl in 'gl.pas' {Form1};

{$R *.res}

begin
if CreateMutex then                 //创建句柄，判断此应用程序是否在运行
  begin
  Application.Initialize;
  Application.HelpFile := 'C:\Documents and Settings\xx\My Documents\My Pictures\3_coffee break\coffeepot.ico';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end else
  begin
    DestroyMutex;                     //释放句柄
  end;
end.
