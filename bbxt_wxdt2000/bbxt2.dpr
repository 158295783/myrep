program bbxt2;

uses
  Forms,
  gl in 'gl.pas' {Form1};

{$R *.res}

begin
if CreateMutex then                 //����������жϴ�Ӧ�ó����Ƿ�������
  begin
  Application.Initialize;
  Application.HelpFile := 'C:\Documents and Settings\xx\My Documents\My Pictures\3_coffee break\coffeepot.ico';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end else
  begin
    DestroyMutex;                     //�ͷž��
  end;
end.
