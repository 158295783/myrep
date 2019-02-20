unit gl;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, ExcelXP, OleServer, DB, DBClient,DateUtils,
  Grids, DBGrids, MConnect, ComCtrls,StrUtils,SHELLAPI, tlhelp32,
  Menus, ADODB, XLSSheetData5, XLSReadWriteII5,Xc12DataStyleSheet5;

type
  TForm1 = class(TForm)
    DataSource5: TDataSource;

    DataSource1: TDataSource;
    ADODataSet3: TADODataSet;
    DataSource3: TDataSource;
    ADOConnection1: TADOConnection;
    ADODataSet5: TADODataSet;
   // ADODataSet1: TADODataSet;
    ADOQuery1: TADOQuery;
    ClientDataSet2: TClientDataSet;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Panel2: TPanel;
    Panel1: TPanel;
    GroupBox2: TGroupBox;
    DBGrid5: TDBGrid;
    GroupBox4: TGroupBox;
    DBGrid1: TDBGrid;
    Panel3: TPanel;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label1: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label7: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    DateTimePicker1: TDateTimePicker;
    hour: TEdit;
    UpDown1: TUpDown;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    BitBtn2: TBitBtn;
    CheckBox1: TCheckBox;
    BitBtn3: TBitBtn;
    ComboBox3: TComboBox;
    qs: TEdit;
    UpDown2: TUpDown;
    qf: TEdit;
    UpDown3: TUpDown;
    CheckBox2: TCheckBox;
    Edit1: TEdit;
    BitBtn1: TBitBtn;
    GroupBox3: TGroupBox;
    Memo2: TMemo;
    GroupBox5: TGroupBox;
    Button1: TButton;
    GroupBox6: TGroupBox;
    Button2: TButton;
    DateTimePicker2: TDateTimePicker;
    Label11: TLabel;
    Edit2: TEdit;
    Label12: TLabel;
    CheckBox3: TCheckBox;
    ProgressBar1: TProgressBar;
    Timer1: TTimer;
    GroupBox7: TGroupBox;
    Button3: TButton;
    Button4: TButton;
    GroupBox8: TGroupBox;
    Button5: TButton;
    XLS: TXLSReadWriteII5;
    Button6: TButton;
    xls2: TXLSReadWriteII5;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    ExcelApplication2: TExcelApplication;
    ExcelWorkbook2: TExcelWorkbook;
    ExcelWorksheet2: TExcelWorksheet;
    Memo1: TMemo;
    ADODataSet1: TADODataSet;
    Timer2: TTimer;
    Button11: TButton;
    procedure FormCreate(Sender: TObject);
    procedure DBGrid5DblClick(Sender: TObject);
    procedure UpDown1Click(Sender: TObject; Button: TUDBtnType);
    procedure ComboBox1Change(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
   // procedure ExcelApplication1WorkbookBeforeClose(ASender: TObject;const Wb: _Workbook; var Cancel: WordBool);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure DBGrid5CellClick(Column: TColumn);
    //procedure BitBtn5Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure UpDown2Click(Sender: TObject; Button: TUDBtnType);
    procedure UpDown3Click(Sender: TObject; Button: TUDBtnType);
    procedure DBGrid5KeyPress(Sender: TObject; var Key: Char);
    procedure CheckBox2Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
  
    //procedure ExcelWorksheet1Change(ASender: TObject;const Target: ExcelRange);
  private
    { Private declarations }
  public
    { Public declarations }
    function MYftpDownLoad(fn1,fn2:string):boolean;

  end;
type
TMessageGrid = array of array of string;
tdnsd=array[1..12,1..48] of byte;
tdnsdstr=array[1..12,1..4] of string;
tarray=array[1..3,1..8]of single;
tbeen=array[1..3,1..48] of byte;

var
  Form1: TForm1;
  irow:smallint;
  jcol:smallint;
  bbsty:byte;
   bend:tbeen;
    load_data:Boolean;
  zero:shortstring;
  dnsd:tdnsd;
  dnsdstr:tdnsdstr;
  bb_year,bb_mon,bb_day:word;
  Msgs: TMessageGrid;
  tcx:tarray;
  ycname,evename,dnname:string;
  bb_valid:smallint;
 // done:boolean;
//const
  rptdir:string;
procedure cre_view(irow,jcol:smallint);forward;
//function get_clh(s:shortstring):word;forward;
function CreateMutex: Boolean;     // 全项目公用函数
procedure DestroyMutex;            // 全项目公用函数
implementation
var Mutex: hWnd;
{$R *.dfm}
procedure EndProcess(AFileName: string); forward;
function FindProcess(AFileName: string): boolean;forward;
function GetDaysOfMon(y,m:integer):smallint;forward;
function get_minutes(year,mon,lb:integer):shortstring;forward;
procedure ch2gno(var ch,xh:longint;tclh:byte);forward;
function do_chuli(s:shortstring):shortstring;forward;
 function do_yinyong(s:string):string;forward;
function get_kgfhcs(ch,xh:longint;mon,day,lb:smallint):integer; forward;
function get_kgfhcs_mon(ch,xh:longint;mon,lb:smallint):integer; forward;

function get_monDN(ch,xh:longint;mon,day,lb:smallint):real;forward;
function get_DN_Value(ch,xh:longint;mon,day,ho,mi:smallint):real;forward;
function get_dayDN(ch,xh:longint;mon,day,lb:smallint):real;forward;

function get_cos_day(ch,xh:longint;mon,day,slct:smallint):integer; forward;
function get_cos_day2(ch,xh:longint;mon,day:smallint):single; forward;
function get_cos_mon(ch,xh:longint;mon,slct:smallint):integer; forward;
function get_cos_mon2(ch,xh:longint;mon:smallint):single; forward;

function get_allyxsj(ch,xh:longint;mon,day,slct:smallint):integer; forward;
function get_allyxsj2(ch,xh:longint;mon,day,slct:smallint):single; forward;
function get_monyxsj(ch,xh:longint;mon,slct:smallint):integer; forward;
function get_monyxsj2(ch,xh:longint;mon,slct:smallint):single; forward;
function get_yc_DayTJ2(ch,xh:longint;mon,day:smallint; slct:byte):real;forward;
function get_yc_Value(ch,xh:longint;mon,day,ho,mi:smallint):real;forward;
function get_yc_mon_cz(ch,xh:longint;mon,day:smallint):single;forward;
function get_yc_MonTJ2(ch,xh:longint;mon:smallint; slct:byte):real;forward;
function get_yc_MonTJ(ch,xh:longint;mon:smallint; slct:byte):string;forward;
function get_yc_MonFhl(ch,xh:longint;mon:smallint):real;forward;
function get_yc_HourMax(ch,xh:longint;mon,day,ho:smallint):real;forward;
function get_yc_DayMin2(ch,xh:longint;mon,day:smallint):string;forward;
function get_yc_DayMin(ch,xh:longint;mon,day:smallint):real;forward;
function get_yc_DayMax2(ch,xh:longint;mon,day:smallint):string;forward;
function get_yc_DayMax(ch,xh:longint;mon,day:smallint):real;forward;
function get_yc_DayAvg(ch,xh:longint;mon,day:smallint):real;forward;
function get_yc_DayFhl(ch,xh:longint;mon,day:smallint):real;forward;
procedure DestroyMutex;
begin
  if Mutex <> 0 then CloseHandle(Mutex);
end;
function CreateMutex: Boolean;
var
  PrevInstHandle: THandle;
  AppTitle: PChar;
begin
  AppTitle := StrAlloc(100);
  StrPCopy(AppTitle, Application.Title);
  Result := True;
  Mutex := Windows.CreateMutex(nil, False, AppTitle);
  if (GetLastError = ERROR_ALREADY_EXISTS) or (Mutex = 0) then begin
    Result := False;
    SetWindowText(Application.Handle,'');
    PrevInstHandle := FindWindow(nil, AppTitle);
    if PrevInstHandle <> 0 then begin
      if IsIconic(PrevInstHandle) then
        ShowWindow(PrevInstHandle, SW_RESTORE)
      else
        BringWindowToTop(PrevInstHandle);
      SetForegroundWindow(PrevInstHandle);
    end;
    if Mutex <> 0 then Mutex := 0;
  end;
  StrDispose(AppTitle);
end;
function FindProcess(AFileName: string): boolean;
var
  hSnapshot: THandle;//用于获得进程列表
  lppe: TProcessEntry32;//用于查找进程
  Found: Boolean;//用于判断进程遍历是否完成
  KillHandle: THandle;//用于杀死进程
begin
  Result :=False;
  hSnapshot := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);//获得系统进程列表
  lppe.dwSize := SizeOf(TProcessEntry32);//在调用Process32First API之前，需要初始化lppe记录的大小
  Found := Process32First(hSnapshot, lppe);//将进程列表的第一个进程信息读入ppe记录中
  while Found do
  begin
    if ((UpperCase(ExtractFileName(lppe.szExeFile))=UpperCase(AFileName)) or (UpperCase(lppe.szExeFile )=UpperCase(AFileName))) then
    begin
      {if MsShow('发现打开Excel,是否将其关闭?',2)=6 then
      begin
      //由于我的操作系统是xp，所以在调用TerminateProcess API之前
      //我必须先获得关闭进程的权限,如果操作系统是NT以下可以直接中止进程
      KillHandle := OpenProcess(PROCESS_TERMINATE, False, lppe.th32ProcessID);
      TerminateProcess(KillHandle, 0);//强制关闭进程
      CloseHandle(KillHandle);
      end;}
      Result :=True;
    end;
    Found := Process32Next(hSnapshot, lppe);//将进程列表的下一个进程信息读入lppe记录中
  end;
end;
function split(s,s1:string):TStringList;
begin
    Result:=TStringList.Create;
        while Pos(s1,s)>0 do
        begin
            Result.Add(Copy(s,1,Pos(s1,s)-1));
             Delete(s,1,Pos(s1,s));
        end;
    Result.Add(s);
end;
procedure EndProcess(AFileName: string);
const
  PROCESS_TERMINATE = $0001;
var
  ContinueLoop: BOOL;
  FSnapShotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  FSnapShotHandle := CreateToolhelp32SnapShot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := SizeOf(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle, FProcessEntry32);
  while integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile))=UpperCase(AFileName)) or (UpperCase(FProcessEntry32.szExeFile )=UpperCase(AFileName))) then
    TerminateProcess(OpenProcess(PROCESS_TERMINATE, BOOL(0),FProcessEntry32.th32ProcessID), 0);
    ContinueLoop := Process32Next(FSnapshotHandle, FProcessEntry32);
  end;
end;
procedure bendin(var dnsd:tdnsd);
var
  i,j,p:byte;
begin
  for i:=1 to 3 do  begin
    for j:=1 to 24 do begin
      bend[i,j]:=0;
    end;
  end;
  j:=1;
  p:=1;
  for i:=1 to 48 do begin
    if dnsd[bb_mon,i]<>dnsd[bb_mon,i+1] then begin
      bend[dnsd[bb_mon,i]+1,p]:=j;
      p:=p+1;
      bend[dnsd[bb_mon,i]+1,p]:=i;
      p:=p+1;
      j:=i+1;
    end;
  end;
  i:=i-1;
  bend[dnsd[bb_mon,i]+1,p]:=j;
  p:=p+1;
  //showmessage(inttostr(i));
  bend[dnsd[bb_mon,i]+1,p]:=i;
end;


function TForm1.MYftpDownLoad(fn1,fn2:string):boolean;
var
  errs:string;
begin
  result:=false;
  fn1:=uppercase(fn1);
  fn2:=uppercase(fn2);
  screen.Cursor:=crHourglass;
  //DCOMConnection1.AppServer.LoadFile(fn1,fn2);
  //errs:=DCOMConnection1.AppServer.SQLerrs;
  if errs<>'' then begin
    ShowMessage('MYftpDownLoad: '+errs);
    result:=false;
  end else begin
    result:=true;
  end;
  screen.Cursor:=crDefault;

end;
function GetDaysOfMon(y,m:integer):smallint;
begin
  case m of
    1,3,5,7,8,10,12:result:=31;
    4,6,9,11:result:=30;
    2:begin
      if ((y mod 4=0) and (y mod 100<>0)) or (y mod 400=0) then
        result:=29
      else
        result:=28;
    end;
    else result:=0;
  end;
end;
function numtodate(tt:string):string;
var
  ss:string;
begin
  ss:='';
  if length(tt)=8 then begin
    ss:=MidStr(tt,3,2)+' '+MidStr(tt,5,2)+':'+MidStr(tt,7,2);
  end;
  if length(tt)=7 then begin
    ss:=MidStr(tt,2,2)+' '+MidStr(tt,4,2)+':'+MidStr(tt,6,2);
  end;
  //form1.Edit2.Text:=ss;
  result:=ss;
end;

function numtodated(tt:string):string;
var
  ss:string;
begin
  ss:='';
  if length(tt)=8 then begin
    ss:=MidStr(tt,5,2)+':'+MidStr(tt,7,2);
  end;
  if length(tt)=7 then begin
    ss:=MidStr(tt,4,2)+':'+MidStr(tt,6,2);
  end;
  //form1.Edit2.Text:=ss;
  result:=' '+ss;
end;


function datetonum(mon,day,ho,mi:smallint):string;
var
  ss:string;
begin
  ss:=inttostr(mon);
  if day<10 then ss:=ss+'0'+inttostr(day)
  else ss:=ss+inttostr(day);
  if ho<10 then ss:=ss+'0'+inttostr(ho)
  else ss:=ss+inttostr(ho);
  if mi<10 then ss:=ss+'0'+inttostr(mi)
  else ss:=ss+inttostr(mi);
  result:=ss;
end;

function gettimes(mon,day,lb:smallint):string;    //时间段开始字符串
var
  tt:string;
begin
  tt:=inttostr(mon);
  if lb=0 then  tt:=tt+'000000'
  else begin
  if  day>=10 then
    tt:=tt+inttostr(day)+'0000'
  else
    tt:=tt+'0'+inttostr(day)+'0000';
  end;
  result:=tt;
end;

function gettimed(mon,day,lb:smallint):string;    //时间段结束字符串
var
  tt:string;
begin
  if day=getdaysofmon(bb_year,mon) then lb:=0;
  if lb=0 then begin
   tt:=inttostr(mon+1);
   tt:=tt+'010000';
  end else begin
  tt:=inttostr(mon);
  day:=day+1;
  if  day>=10 then
    tt:=tt+inttostr(day)+'0000'
  else
    tt:=tt+'0'+inttostr(day)+'0000';
  end;
  result:=tt;
end;
function gettimes_(mon,day,lb:smallint):string;    //时间段结束字符串  无锡大通
var
  tt:string;
begin
  tt:=inttostr(mon);
  if lb=0 then  tt:=tt+'000000'
  else begin
  if  day>=10 then
    tt:=tt+inttostr(day)+'2000'
  else
    tt:=tt+'0'+inttostr(day)+'2000';
  end;
  result:=tt;
end;

function gettimed_(mon,day,lb:smallint):string;    //时间段开始字符串  无锡大通
var
  tt:string;
begin
  if day=getdaysofmon(bb_year,mon) then lb:=0;
  if lb=0 then begin
   tt:=inttostr(mon+1);
   tt:=tt+'010000';
  end else begin
  tt:=inttostr(mon);
  day:=day-1;
  if  day>=10 then
    tt:=tt+inttostr(day)+'2000'
  else
    tt:=tt+'0'+inttostr(day)+'2000';
  end;
  result:=tt;
end;
function get_yc_DayAvg(ch,xh:longint;mon,day:smallint):real;
var
  dttm1:tdate;
  str1,str2:string;
begin
  dttm1:=encodedate(bb_year,mon,day);
 // if (dttm1>=date) then dttm1:=dttm1/0 else begin
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    str1:=gettimes(mon,day,1);
        str2:=gettimed(mon,day,1);
        commandtext:='select avg(abs(val'+inttostr(xh)+')) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2;
    open;
    result:=fields[0].AsFloat;
  end;
 // end;
end;
function get_yc_DayMin2(ch,xh:longint;mon,day:smallint):string;
var
  dttm1:tdate;
  str1, str2:string;
  val:real;
begin
  try val:=get_yc_DayMin(ch,xh,mon,day);except dttm1:=dttm1/0 ;end;
  dttm1:=encodedate(bb_year,mon,day);
 // if (dttm1>=date) then dttm1:=dttm1/0 else begin
    with form1.adodataset1 do begin
      if active then close;
    //sql.Clear;
      str1:=gettimes(mon,day,1);
        str2:=gettimed(mon,day,1);
         commandtext:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and abs(val'+inttostr(xh)+')='+floattostr(val)+' and savetime>='+str1+' and savetime<'+str2;
      open;
      result:=fields[0].asstring;
    end;
 // end;
end;
function get_yc_DayMax2(ch,xh:longint;mon,day:smallint):string;
var
  dttm1:tdate;
  str1, str2:string;
  val:real;
begin
  try val:=get_yc_Daymax(ch,xh,mon,day);except dttm1:=dttm1/0 ;end;
  dttm1:=encodedate(bb_year,mon,day);
  //if (dttm1>=date) then dttm1:=dttm1/0 else begin
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    str1:=gettimes(mon,day,1);
        str2:=gettimed(mon,day,1);
            commandtext:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and abs(val'+inttostr(xh)+')='+floattostr(val)+' and savetime>='+str1+' and savetime<'+str2;
    open;
    result:=fields[0].Asstring;
  end;
 // end;
end;
function get_yc_HourMax(ch,xh:longint;mon,day,ho:smallint):real;
var
  dttm1:tdatetime;
  str1:string;
begin
  dttm1:=encodedate(bb_year,mon,day)+encodetime(ho,0,0,0);
  if (dttm1>now-1/24) then dttm1:=dttm1/0 else begin
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    str1:= datetonum(mon,day,ho,0);
        commandtext:='select max(abs(val'+inttostr(xh)+')) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<100+'+str1;
    open;
    result:=fields[0].AsFloat;
  end;
  end;
end;
function get_yc_DayMin(ch,xh:longint;mon,day:smallint):real;
var
  dttm1:tdate;
  str1,str2:string;
begin

  dttm1:=encodedate(bb_year,mon,day);
  //if (dttm1>=date) then dttm1:=dttm1/0 else begin
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    str1:= datetonum(mon,day,0,0);
    str2:=' and abs(val'+inttostr(xh)+')>'+zero ;
    commandtext:='select min(abs(val'+inttostr(xh)+')) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<10000+'+str1+str2;
    open;
    result:=fields[0].AsFloat;
  end;
 // end;
end;
function get_yc_DayMax(ch,xh:longint;mon,day:smallint):real;
var
  dttm1:tdate;
  str1:string;
begin
  dttm1:=encodedate(bb_year,mon,day);
  //if (dttm1>=date) then dttm1:=dttm1/0 else begin
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    str1:= datetonum(mon,day,0,0);
    commandtext:='select max(abs(val'+inttostr(xh)+')) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<10000+'+str1;
    open;
    result:=fields[0].AsFloat;
  end;
 // end;
end;
function get_yc_DayFhl(ch,xh:longint;mon,day:smallint):real;
begin
try
if  get_yc_dayMax(ch,xh,mon,day)<>0  then
result:=get_yc_dayAvg(ch,xh,mon,day)/get_yc_dayMax(ch,xh,mon,day)
else result:=0;
except
   result:=0;
   end;
end;
function get_yc_MonFhl(ch,xh:longint;mon:smallint):real;
var
days:byte;
sumv:real;
cnt:byte;
i:byte;
begin
days:=getDaysOfMon(bb_year,mon);
cnt:=0;
sumv:=0;
for i:=1 to days do
try
 if  get_yc_DayFhl(ch,xh,mon,i)<>0 then begin
sumv:=sumv+get_yc_DayFhl(ch,xh,mon,i) ;
inc(cnt);
end;
 except
 //showmessage('ddddd');
 end;
//showmessage(floattostr(sumv/cnt));
if cnt<>0 then
  result:=sumv/cnt
else
  result:=0;
end;

function get_yc_Value(ch,xh:longint;mon,day,
    ho,mi:smallint):real;
var
  dttm1:tdatetime;
  ss:string;
begin
  dttm1:=encodedate(bb_year,mon,day)+encodetime(ho,mi,0,0);
  if (dttm1>now-5/60/24) then dttm1:=dttm1/0 else begin
  ss:=datetonum(mon,day,ho,mi);
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    //sql.Add('select val'+inttostr(xh)+' from dn2_yc_table where groupno='+inttostr(ch)+' and savetime=to_date('''+inttostr(bb_year)+'-'+inttostr(mon)+'-'+inttostr(day)+'-'+inttostr(ho)+'-'+inttostr(mi)+''',''yyyy-mm-dd-hh24-mi'')');
    commandtext:='select val'+inttostr(xh)+' from yc_table where groupno='+inttostr(ch)+' and savetime='+ss;
    //form1.Edit5.Text:=commandtext;
    open;
    result:=fields[0].AsFloat;
  end;
  end;
end;
procedure load_dnsd(var dnsd:tdnsd);
var
  i,j:byte;
  //dttm1:tdate;
  str1:string;
begin
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    str1:='select * from dnsds2 where tsidx between 1 and 12';
    commandtext:=str1;
    try
      open;
    except
      showmessage('adodataset6 ERR: '+commandText);
    end;
    while not eof  do begin
    for i:=1  to 48 do begin dnsd[fieldbyname('tsidx').AsInteger,i]:=fieldbyname('ts'+inttostr(i)).asinteger end;
    next;
    end;
  end;

  for i:=1 to 12 do
  for j:=1 to 48 do
      dnsdstr[i,dnsd[i,j]+1]:=dnsdstr[i,dnsd[i,j]+1]+inttostr(j)+',';
  str1:='';
  for i:=1 to 12 do begin
  for j:=1 to 3 do
      dnsdstr[i,j]:=dnsdstr[i,j]+'0';
      dnsdstr[i,4]:=dnsdstr[i,3];
      dnsdstr[i,3]:=dnsdstr[i,2];
      dnsdstr[i,2]:=dnsdstr[i,1];
   str1:=str1+' '+dnsdstr[i,1]+' '+dnsdstr[i,2]+' '+dnsdstr[i,3];
  end;
  //showmessage(str1);
end;
procedure TForm1.FormCreate(Sender: TObject);
var
  ss:shortstring;
   Hour2,   Min2,   Sec2,   MSec2,md,lb:   Word;
      ToTal,totall:longint;
begin
DateTimePicker1.Date:=date;
      DecodeTime(now,   Hour2,   Min2,   Sec2,   MSec2);
      if Hour2>19 then begin
      ToTal:=24*60*60*1000-(Hour2)*60*60*1000-(Min2)*60*1000-(Sec2)*1000-MSec2;
      totall:=total+20*60*60*1000+5*60*1000;
      end else begin
       ToTall:=20*60*60*1000+5*60*1000-(Hour2)*60*60*1000-(Min2)*60*1000-(Sec2)*1000-MSec2;
        memo1.Lines.Add(inttostr(totall)+'---'+inttostr(Hour2));
      end;
      timer1.Interval:=totall;

 // DateTimePicker1.Date:=date;
  adoConnection1.Open;
  //DCOMConnection1.AppServer.opendbs;
  {with adodataset5 do begin
    if active then close;
    commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,chgtime from allfiles where filename like ''%.XLSM%'' order by filename desc';
    //params[0].Value:='''%.XLS%''';
    try
      open;
    except
      showmessage('ERR:'+adodataset5.CommandText);
    end;
  end; }
  load_data:=false;
  with adodataset3 do begin
    if active then close;
    commandtext:='select rtuno,name,ananum 遥测点数,dignum 遥信点数 from prtu where ananum<>0  order by rtuno';
    try
      open;
    except
      showmessage('ERR:'+adodataset1.CommandText);
    end;
  end;
  with  adodataset5 do begin
    if active then close;
    commandtext:='select filename,valid from rptlist where  filename like ''%.XLS%'' order by valid';
        //edit1.Text:=commandtext;
    try
      open;
    except
      showmessage('no such type tables!');
    end;
      first;
    while (not eof) do begin
      ss:=trim(fieldbyname('filename').asstring);
      ss:=copy(ss,1,length(ss)-5);
      combobox2.Items.Add(ss);
      next;
    end;
  end;
  rptdir:=ExtractFilePath(Application.Exename)+'report\';
  if combobox2.Items.count<>0 then
    combobox2.ItemIndex:=0;
  //windowstate:=wsMaximized;
  load_dnsd(dnsd);
  bb_valid:=0;
end;

procedure TForm1.DBGrid5DblClick(Sender: TObject);
var
  fn1,fn2:string;
begin
//下载文件
 { fn1:=trim(adodataset5.fieldbyname('filename').AsString);
  combobox2.Text:=fn1;
  BitBtn2Click(nil);}
end;

procedure TForm1.UpDown1Click(Sender: TObject; Button: TUDBtnType);
var
  i:integer;
  ss:shortstring;

begin
   if combobox1.ItemIndex<>2 then exit;
   i:=strtoint(hour.text);
   if Button = btnext then   begin
     i:=i+1;
     if i>23 then i:=0;
     if i<10 then ss:='0'+inttostr(i)
     else ss:=inttostr(i);
     hour.Text:=ss;
   end;
   if Button = btprev then  begin
     i:=i-1;
     if i<0 then
     //showmessage('')
     i:=23;
     if i<10 then ss:='0'+inttostr(i)
     else ss:=inttostr(i);
     hour.Text:=ss;
   end;
end;
procedure TForm1.ComboBox1Change(Sender: TObject);
var
  ss:shortstring;
begin
  combobox2.Items.Clear;
  combobox2.Text:='';
  case combobox1.ItemIndex of
    0:begin
       with adodataset5 do begin
         if active then close;
         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where  filename like ''%遥测%月%'' or filename like ''%电压%月%'' or filename like ''%电流%月%'' order by filename desc';
         try
           open;
         except
           showmessage('ERR:'+adodataset5.CommandText);
         end;
       end;
       label7.Caption:='月报取数:';
       label6.Caption:='月报取时:';
     with  adodataset1 do begin
        if active then close;
        commandtext:='select filename from rptlist where    ( filename like ''%遥测%月%'' or filename like ''%电压%月%'' or filename like ''%电流%月%'') order by filename';
        //edit1.Text:=commandtext;
        try
          open;
        except
          showmessage('no such type tables!');
        end;
        first;
        while (not eof) do begin
          ss:=trim(fieldbyname('filename').asstring);
          ss:=copy(ss,1,length(ss)-5);
          combobox2.Items.Add(ss);
          next;
        end;
      end;
    end;
    1:begin
       with adodataset5 do begin
         if active then close;
         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where   filename like ''%日%'' order by filename desc';
         try
           open;
         except
           showmessage('ERR:'+adodataset5.CommandText);
         end;
       end;
      with  adodataset1 do begin
        if active then close;
        commandtext:='select filename from rptlist where    filename like ''%日%'' order by filename';
        //edit1.Text:=commandtext;
        try
          open;
        except
          showmessage('no such type tables!');
        end;
        first;
        while (not eof) do begin
           ss:=trim(fieldbyname('filename').asstring);
          ss:=copy(ss,1,length(ss)-5);
          combobox2.Items.Add(ss);
          next;
        end;
      end;
       label7.Caption:='日报取数:';
       label6.Caption:='日报取时:';
       
    end;
    2:begin
      with adodataset5 do begin
         if active then close;
         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where filename like ''%时%'' order by filename desc';
         try
           open;
         except
           showmessage('ERR:'+adodataset5.CommandText);
         end;
       end;
      with  adodataset1 do begin
        if active then close;
        commandtext:='select filename from rptlist where    filename like ''%时%'' order by filename';
        //edit1.Text:=commandtext;
        try
          open;
        except
          showmessage('no such type tables!');
        end;
        first;
        while (not eof) do begin
           ss:=trim(fieldbyname('filename').asstring);
          ss:=copy(ss,1,length(ss)-4);
          combobox2.Items.Add(ss);
          next;
        end;
      end;
    end;
    3:begin
      with adodataset5 do begin
         if active then close;
         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where    filename like ''%电量月%'' or filename like ''%能耗月%'' order by filename desc';
         try
           open;
         except
           showmessage('ERR:'+adodataset5.CommandText);
         end;
       end;

      with  adodataset1 do begin
        if active then close;
        commandtext:='select filename from rptlist where   filename like ''%电量月%''  or filename like ''%能耗月%'' order by filename';
        //edit1.Text:=commandtext;
        try
          open;
        except
          showmessage('no such type tables!');
        end;
        first;
        while (not eof) do begin
           ss:=trim(fieldbyname('filename').asstring);
          ss:=copy(ss,1,length(ss)-5);
          combobox2.Items.Add(ss);
          next;
        end;
      end;
    end;
     4:begin
      with adodataset5 do begin
         if active then close;
         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where    filename like ''%电量日%'' or filename like ''%能耗日%'' order by filename desc';
         try
           open;
         except
           showmessage('ERR:'+adodataset5.CommandText);
         end;
       end;

      with  adodataset1 do begin
        if active then close;
        commandtext:='select filename from rptlist where   filename like ''%电量日%''  or filename like ''%能耗日%'' order by filename';
        //edit1.Text:=commandtext;
        try
          open;
        except
          showmessage('no such type tables!');
        end;
        first;
        while (not eof) do begin
           ss:=trim(fieldbyname('filename').asstring);
          ss:=copy(ss,1,length(ss)-5);
          combobox2.Items.Add(ss);
          next;
        end;
      end;
    end;
    5:begin
      with adodataset5 do begin
         if active then close;
         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where   filename like ''%特殊%'' order by filename desc';
         try
           open;
         except
           showmessage('ERR:'+adodataset5.CommandText);
         end;
       end;
      with  adodataset1 do begin
        if active then close;
        commandtext:='select filename from rptlist where    filename like ''%特殊%'' order by filename';
        //edit1.Text:=commandtext;
        try
          open;
        except
          showmessage('no such type tables!');
        end;
        first;
        while (not eof) do begin
           ss:=trim(fieldbyname('filename').asstring);
          ss:=copy(ss,1,length(ss)-5);
          combobox2.Items.Add(ss);
          next;
        end;
      end;
    end;
  end;   //case
  if combobox2.Items.count<>0 then
    combobox2.ItemIndex:=0;
  form1.DBGrid5CellClick(nil);
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
var
  qzw:Variant;
  s1,s2:word;
  i,j,gno,vno,blan,blans:integer;
  k:byte;
  ss,str1,sqls,str2,clh,lb,str3,str4,xlsS,xls_,xls__:string;
  fn:real;
  pmax,pmin,pavg,fmax,fmin,favg:real;
  dw:array[1..5] of string;
  dwidx:word;
  thedate,sdate,edate:tdate;
begin
  with adoconnection1 do begin
     if not connected then open;
 end;
  decodedate(DateTimePicker1.Date,bb_year,bb_mon,bb_day);
   ycname:='hyc'+inttostr(bb_year mod 10);
  evename:='eve_v'+inttostr(bb_year mod 10);
  dnname:='hdn'+inttostr(bb_year mod 10);

   //ycname:='yct';
   blan:=0;
   progressbar1.Position:=0;
  //showmessage(ycname+'  '+evename );
    bbsty:=combobox1.ItemIndex;
  case combobox1.ItemIndex of
    0:ss:='.XLSX';
    1:ss:='.XLSX';
    2:ss:='.XLSX';
    4:ss:='.XLSX';
    3:ss:='.XLSX';
  end;
   xlsS:= trim(combobox2.Text)+ss;
   xls_:= trim(combobox2.Text)+'_'+inttostr(bb_year*10000+bb_mon*100+bb_day)+ss;
    xls__:= trim(combobox2.Text)+'_'+inttostr(bb_year*100+bb_mon)+ss;
   if pos('电流',xlss)>0 then dwidx:=1;
   if pos('电压',xlss)>0 then dwidx:=2;
   if pos('负荷',xlss)>0 then dwidx:=3;
   if pos('能耗',xlss)>0 then dwidx:=4;
   if pos('电量',xlss)>0 then dwidx:=4;
   if pos('抄表',xlss)>0 then dwidx:=5;


   if pos('电能抄表',xlss)>0 then   ycname:='hdn'+inttostr(bb_year mod 10);
   dw[1]:='(A)';
   dw[2]:='(V)';
   dw[3]:='(kW)';
   dw[4]:='(kWh)';
   dw[5]:='(电流:A;电压:V;负荷:kW)';
  ss:=rptdir+trim(combobox2.Text)+ss;

 { if FindProcess('EXCEL.EXE') then
  begin
    //if MsShow('检测到打开了Excel,是否让其关闭?',2)<>6 then Exit;
    ExcelApplication1.DisplayAlerts[0]:=false;
    ExcelApplication1.Quit;
    ExcelWorksheet1.Disconnect;
    ExcelWorkbook1.Disconnect;
    ExcelApplication1.Disconnect;
    EndProcess('EXCEL.EXE');
  end; }

  memo1.Lines.Add(ss);
    try
       XLS.Filename :=  ss;
       xls.Read;
    except
     memo2.Lines.Add('请确认已下载Excel文件!'+ss);
     exit;
    end;
  irow:= XLS[0].LastRow;
  jcol:=XLS[0].LastCol;
  memo1.Lines.Add('irow:'+inttostr(irow));
  memo1.Lines.Add('jcol:'+inttostr(jcol));
  memo1.Lines.Add('bbsty:'+inttostr(bbsty));
 if bb_valid=0 then begin

  progressbar1.Max:=jcol;
   case combobox1.ItemIndex of
    0:begin
      for i:=4 to irow do begin
         try
         gno:=XLS.Sheets[0].AsInteger[0,i];
         except
         gno:=0;
         end;
         if gno=1 then blan:=i;

      end;
      memo1.Lines.Add(inttostr(blan));
    end;
    1:begin
      for i:=4 to irow do begin
         ss:=XLS.Sheets[0].Asstring[0,i];
         //showmessage(ss);
         if ss='0' then   begin
         blan:=i;
         //showmessage(ss);
         end;
      end;
    end;
     4:begin

         blan:=5;

    end;
    2:begin
      for i:=4 to irow do begin
         ss:=XLS.Sheets[0].Asstring[0,i];
         if ss=':00' then blan:=i;
      end;
    end;
    5:ss:='.XLST';
    3:begin
       for i:=4 to irow do begin
        try
         gno:=XLS.Sheets[0].AsInteger[0,i];
         except
         gno:=0;
         end;
         if gno=1 then blan:=i;
      end;
    end;
  end;
   memo1.Lines.Add('range:'+inttostr(blan));
   ss:='';
  if bbsty<>5 then begin
     lb:=XLS.Sheets[0].Asstring[0,1];     //*********************//

       SetLength(Msgs,jcol+1,3);
       for i:=1 to jcol do begin
         str3:= XLS.Sheets[0].Asstring[i,0];

         try
         gno:=XLS.Sheets[0].Asinteger[i,0];
         if length(str3)<>0 then begin
         // memo2.Lines.Add(inttostr(i)+':'+inttostr(gno))  ;
         vno:=gno mod 200;
         gno:=gno div 200;
         Msgs[i-1,0]:=inttostr(gno);
         Msgs[i-1,1]:=inttostr(vno);
          str1:=XLS.Sheets[0].Asstring[i,1];
         Msgs[i-1,2]:=str1;
         if pos(','+Msgs[i-1,0]+',',ss)=0 then
           ss:=ss+','+Msgs[i-1,0];
         end else begin
            Msgs[i-1,0]:='';
            Msgs[i-1,1]:='';
             Msgs[i-1,2]:='0';
         end;

         except

            Msgs[i-1,0]:='';
            Msgs[i-1,1]:='';
         end;
       end;


    s1:=0;
    s2:=0;
    if load_data=false then begin
     with form1.adoquery1 do begin
            if active then close;
              sql.Clear;
              if ((bbsty<>3) and (bbsty<>4)) then
                sql.add('truncate table yc_table')
              else
                sql.add('truncate table dn_table');
            try
             // execsql;
            except
              //showmessage('ddfd');
            end;
            //showmessage(commandtext);
            {if active then close;
                commandText:='commit' ;
            try
              execute;
            except
            end;}
     end;
     end;  //if load_data
  end;
 // showmessage(inttostr(bbsty));
  case  bbsty of

     1:begin
        IF load_data=false then begin
                str1:=gettimes(bb_mon,bb_day,1);
                str2:=gettimed(bb_mon,bb_day,1);
                sqls:='insert into yc_table  select * from '+ycname+' where savetime>=';
                sqls:=sqls+str1+' and savetime<'+str2;
                str4:=' and savetime >= ' +str1+' and savetime<'+str2;
                //showmessage(inttostr(pos(',',ss)));
                if pos(',',ss)=1 then
                  ss:=copy(ss,2,length(ss)-1);
                if length(ss)>1 then
                  sqls:=sqls+' and groupno in ('+ss+')'
                else
                  sqls:=sqls+' and groupno='+ss;
                  sqls:=sqls+' and chgtime is not null';
                   memo1.Lines.Add(sqls);
                  with form1.adoquery1 do begin
                    if active then close;
                    sql.Clear;
                    sql.Add(sqls);
                    try
                      execsql;
                    except
                    end;
                    //edit1.text:=sqls;
                  end; //with   **  do begin
          end;
          ycname:='yc_table';

        for j:=1 to jcol do begin
           progressbar1.Position:=j;
           for gno:=1 to 3 do
                  for vno:=1 to 8 do
                    tcx[gno,vno]:=0;
              str1:=XLS.Sheets[0].Asstring[j,0];
              if str1='' then  begin

              exit;
              end;
               str1:=' group by floor(mod(savetime,10000)/100)';
              str2:=' and abs(val'+Msgs[j-1,1]+')<='+Msgs[j-1,2];
              Combobox3.ItemIndex:=0;
              case ComboBox3.ItemIndex of
                 0: str3:='val'+Msgs[j-1,1];
                 1: str3:='max((val'+Msgs[j-1,1]+'))';
                 2: str3:='min((val'+Msgs[j-1,1]+'))';
                 3: str3:='avg((val'+Msgs[j-1,1]+'))';
              end;

                with adoquery1 do begin
                  if active then close;
                  sql.Clear;

                 sqls:='select floor(mod(savetime,10000)/100) time, '+str3+' value from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str4;
                 if ComboBox3.ItemIndex=0 then begin
                    sqls:=sqls+' and mod(savetime,100)='+inttostr(strtoint(qf.Text));
                 end else begin
                    sqls:=sqls+str1;
                 end;
                 memo1.lines.Add(sqls);
                 sql.Add(sqls);
                  open;
                  if  adoquery1.eof then  begin
                    i:=blan;
                    while(i<=irow)do  begin
                      XLS.Sheets[0].Asstring[j,i]:='';
                      i:=i+1;
                    end;
                    continue;
                  end else begin
                    i:=blan;
                    first;
                    while(not adoquery1.eof) do begin
                      gno:=adoquery1.fieldbyname('time').asinteger;
                      //clh:=midstr(clh,12,2);
                      i:=i+1;

                      XLS.Sheets[0].asfloat[j,gno+blan]:=adoquery1.fieldbyname('value').asfloat;
                      next;
                    end;
                  end;
                end;
                while(i<=irow)do  begin

                  clh:=XLS.Sheets[0].Asstring[j,i];
                  //showmessage(inttostr(pos(ss,'#')));
                  if pos('#',clh)<>1 then begin i:=i+1;continue;    end;
                  //showmessage(midstr(clh,2,2));
                  case  strtoint(midstr(clh,2,2)) of
                    1:begin
                     if s1=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max(val'+Msgs[j-1,1]+') ax,min(val'+Msgs[j-1,1]+') mn,avg(val'+Msgs[j-1,1]+') vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str1+str4;
                          //edit2.Text:=sql.Text;
                          sql.Add(sqls);
                          open;
                          pmax:=adoquery1.fieldbyname('ax').asfloat;
                          pmin:=adoquery1.fieldbyname('mn').asfloat;
                          pavg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=pmax;
                    end;
                    2:begin
                       if s1=0 then begin
                       with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max(val'+Msgs[j-1,1]+') ax,min(val'+Msgs[j-1,1]+') mn,avg(val'+Msgs[j-1,1]+') vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str1+str4;
                          //edit2.Text:=sql.Text;
                          sql.Add(sqls);
                          open;
                          pmax:=adoquery1.fieldbyname('ax').asfloat;
                          pmin:=adoquery1.fieldbyname('mn').asfloat;
                          pavg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=pmin;
                    end;
                    3:begin
                       if s1=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max(val'+Msgs[j-1,1]+') ax,min(val'+Msgs[j-1,1]+') mn,avg(val'+Msgs[j-1,1]+') vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str1+str4;
                          //edit2.Text:=sql.Text;
                          sql.Add(sqls);
                          open;
                          pmax:=adoquery1.fieldbyname('ax').asfloat;
                          pmin:=adoquery1.fieldbyname('mn').asfloat;
                          pavg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=pavg;
                    end;
                    4:begin
                      //showmessage(inttostr(s2));
                      if s2=0 then begin
                       with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max(val'+Msgs[j-1,1]+') ax,min(val'+Msgs[j-1,1]+') mn,avg(val'+Msgs[j-1,1]+') vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str1+str4;
                          //edit2.Text:=sql.Text;
                          sql.Add(sqls);
                          open;
                          pmax:=adoquery1.fieldbyname('ax').asfloat;
                          pmin:=adoquery1.fieldbyname('mn').asfloat;
                          pavg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                       with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select val'+Msgs[j-1,1]+' value from '+ycname+' where abs(val'+Msgs[j-1,1]+')='+floattostr(fmax)+' and groupno='+Msgs[j-1,0]+str4;
                          sql.Add(sqls);
                          open;
                        end;
                        fmax:=adoquery1.fieldbyname('value').asfloat;
                        XLS.Sheets[0].AsFloat[j,i]:=fmax;
                    end;
                    5:begin
                      if s2=0 then begin
                       with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max(val'+Msgs[j-1,1]+') ax,min(val'+Msgs[j-1,1]+') mn,avg(val'+Msgs[j-1,1]+') vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str1+str4;

                          sql.Add(sqls);
                          open;
                          pmax:=adoquery1.fieldbyname('ax').asfloat;
                          pmin:=adoquery1.fieldbyname('mn').asfloat;
                          pavg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                       with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select val'+Msgs[j-1,1]+' value from '+ycname+' where abs(val'+Msgs[j-1,1]+')='+floattostr(fmin)+' and groupno='+Msgs[j-1,0]+str4;
                          sql.Add(sqls);
                          open;
                        end;
                        fmin:=adoquery1.fieldbyname('value').asfloat;
                      XLS.Sheets[0].AsFloat[j,i]:=fmin;
                    end;
                    6:begin
                      if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max((val'+Msgs[j-1,1]+')) ax,min(abs(val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str4;
                           sql.Add(sqls);
                          open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=favg;
                    end;
                    7:begin
                      if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max((val'+Msgs[j-1,1]+')) ax,min(abs(val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2;
                           sql.Add(sqls);
                          open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                      if fmax<>0 then
                        XLS.Sheets[0].AsFloat[j,i]:=favg/fmax
                      else
                        XLS.Sheets[0].AsFloat[j,i]:=0;
                    end;

                    8:begin
                      if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max((val'+Msgs[j-1,1]+')) ax,min(abs(val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2;
                        memo1.lines.Add(sqls);
                         sql.Add(sqls);
                          open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                      with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select  min(savetime) time from '+ycname+' where groupno='+Msgs[j-1,0]+' and val'+Msgs[j-1,1]+'='+floattostr(fmax);
                          memo1.lines.Add(sqls);
                           sql.Add(sqls);
                          open;
                      end;
                      XLS.Sheets[0].asstring[j,i]:=numtodated(adoquery1.fieldbyname('time').asstring);
                    end;
                    9:begin
                       if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max((val'+Msgs[j-1,1]+')) ax,min(abs(val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2;
                           memo1.lines.Add(sqls);
                            sql.Add(sqls);
                          open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                       with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select  min(savetime) time from '+ycname+' where groupno='+Msgs[j-1,0]+' and val'+Msgs[j-1,1]+'='+floattostr(fmin);
                           memo1.lines.Add(sqls);
                            sql.Add(sqls);
                          open;
                      end;
                      XLS.Sheets[0].Asstring[j,i]:=numtodated(adoquery1.fieldbyname('time').asstring);
                    end;
                    10:begin

                      if tcx[1,4]=0 then begin
                        zero:='0';
                        tcx[2,4]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,4);
                        tcx[1,4]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=tcx[2,4]/1440.00;
                    end;
                    11:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[3,6]=0 then begin
                        tcx[3,1]:=strtoint(get_minutes(bb_year,bb_mon,1));
                        tcx[3,6]:=1;
                      end;
                      if tcx[1,1]=0 then begin
                        zero:='0';
                        tcx[2,1]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,1);
                        tcx[1,1]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=tcx[2,1]/(tcx[3,1]/tcx[3,4]);

                    end;
                    12:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[3,7]=0 then begin
                        tcx[3,2]:=strtoint(get_minutes(bb_year,bb_mon,2));
                        tcx[3,7]:=1;
                      end;
                      if tcx[1,2]=0 then begin
                        zero:='0';
                        tcx[2,2]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,2);
                        tcx[1,2]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=tcx[2,2]/(tcx[3,2]/tcx[3,4]);
                    end;
                    13:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[3,8]=0 then begin
                        tcx[3,3]:=strtoint(get_minutes(bb_year,bb_mon,3));
                        tcx[3,8]:=1;
                      end;
                      if tcx[1,3]=0 then begin
                        zero:='0';
                        tcx[2,3]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,3);
                        tcx[1,3]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=tcx[2,3]/(tcx[3,3]/tcx[3,4]);
                    end;
                    14:begin
                     if tcx[1,8]=0 then begin
                        zero:='0';
                        tcx[2,8]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,8);
                        tcx[1,8]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=tcx[2,8]/1440.0;
                    end;
                    15:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[3,6]=0 then begin
                        tcx[3,1]:=strtoint(get_minutes(bb_year,bb_mon,1));
                        tcx[3,6]:=1;
                      end;
                      if tcx[1,5]=0 then begin
                        zero:='0';
                        tcx[2,5]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,5);
                        tcx[1,5]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=tcx[2,5]/(tcx[3,1]/tcx[3,4]);
                    end;
                    16:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[3,7]=0 then begin
                        tcx[3,2]:=strtoint(get_minutes(bb_year,bb_mon,2));
                        tcx[3,7]:=1;
                      end;
                      if tcx[1,6]=0 then begin
                        zero:='0';
                        tcx[2,6]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,6);
                        tcx[1,6]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=tcx[2,6]/(tcx[3,2]/tcx[3,4]);
                    end;
                    17:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[3,8]=0 then begin
                        tcx[3,3]:=strtoint(get_minutes(bb_year,bb_mon,3));
                        tcx[3,8]:=1;
                      end;
                      if tcx[1,7]=0 then begin
                        zero:='0';
                        tcx[2,7]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,7);
                        tcx[1,7]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=tcx[2,7]/(tcx[3,3]/tcx[3,4]);
                    end;
                    18:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[1,4]=0 then begin
                        zero:='0';
                        tcx[2,4]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,4);
                        tcx[1,4]:=1;
                      end;

                      if tcx[1,8]=0 then begin
                        zero:='0';
                        tcx[2,8]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,8);
                        tcx[1,8]:=1;
                      end;
                      XLS.Sheets[0].Asstring[j,i]:=formatfloat('0.000',100-((tcx[2,4]+tcx[2,8])/14.40))+'%';
                    end;
                    19:begin
                      if tcx[1,4]=0 then begin
                        zero:='0';
                        tcx[2,4]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,4);
                        tcx[1,4]:=1;
                      end;
                        XLS.Sheets[0].Asfloat[j,i]:=tcx[2,4];
                        //showmessage('dd');
                    end;
                     20:begin
                      if tcx[1,1]=0 then begin
                        zero:='0';
                        tcx[2,1]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,1);
                        tcx[1,1]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,1];
                    end;
                    21:begin
                      if tcx[1,2]=0 then begin
                        zero:='0';
                        tcx[2,2]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,2);
                        tcx[1,2]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,2];
                    end;
                    22:begin
                      if tcx[1,3]=0 then begin
                        zero:='0';
                        tcx[2,3]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,3);
                        tcx[1,3]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,3];
                    end;
                    23:begin
                      if tcx[1,8]=0 then begin
                        zero:='0';
                        tcx[2,8]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,8);
                        tcx[1,8]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,8];
                    end;
                    24:begin
                      if tcx[1,5]=0 then begin
                      zero:='0';
                      tcx[2,5]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,5);
                      tcx[1,5]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,5];
                    end;
                    25:begin
                      if tcx[1,6]=0 then begin
                      zero:='0';
                      tcx[2,6]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,6);
                      tcx[1,6]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,6];
                    end;
                    26:begin
                      if tcx[1,7]=0 then begin
                      zero:='0';
                      tcx[2,7]:=get_allyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,bb_day,7);
                      tcx[1,7]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,7];
                    end;
                  end; //case
                   i:=i+1;
                end;  //while
                s1:=0;
                s2:=0;
                XLS.Sheets[0].Asstring[j,0]:='';
                XLS.Sheets[0].Asstring[j,1]:='';
               // progressbar1.Position:=j;
              end; //  every point


      end;
       4:begin
        IF load_data=false then begin
                   if not ((bb_mon=1) and (bb_day=1)) then begin
                        str1:=gettimed_(bb_mon,bb_day,1);
                        str2:=gettimes_(bb_mon,bb_day,1);
                        sqls:='insert into dn_table  select * from dn_inc where savetime>=';
                        sqls:=sqls+str1+' and savetime<'+str2;
                        sqls:=sqls+' and floor(chgtime/10000)='+inttostr(bb_year);
                        str4:=' and savetime >= ' +str1+' and savetime<'+str2;
                    end else begin
                        str1:=gettimed_(bb_mon,bb_day,1);
                        str2:=gettimes_(bb_mon,bb_day,1);
                        sqls:='insert into dn_table  select * from dn_inc where savetime>=';
                        sqls:=sqls+str1+' or savetime<'+str2;
                        sqls:=sqls+' and floor(chgtime/10000)='+inttostr(bb_year);
                        str4:=' and savetime >= ' +str1+' and savetime<'+str2;
                    end;
                //showmessage(inttostr(pos(',',ss)));
                if pos(',',ss)=1 then
                  ss:=copy(ss,2,length(ss)-1);
                if length(ss)>1 then
                  sqls:=sqls+' and groupno in ('+ss+')'
                else
                  sqls:=sqls+' and groupno='+ss;
                  sqls:=sqls+' and floor(chgtime/10000)='+inttostr(bb_year);
                   memo1.Lines.Add(sqls);
                  with form1.adoquery1 do begin
                    if active then close;
                    sql.Clear;
                    sql.Add(sqls);
                    try
                      execsql;
                    except
                    end;
                    //edit1.text:=sqls;
                  end; //with   **  do begin
          end;
          ycname:='dn_table';

        for j:=1 to jcol do begin
           progressbar1.Position:=j;
           for gno:=1 to 3 do
                  for vno:=1 to 8 do
                    tcx[gno,vno]:=0;
              str1:=XLS.Sheets[0].Asstring[j,0];
              if str1='' then  begin

              continue;
              end;
               str1:=' group by floor(mod(savetime,10000)/100)';
              str2:=' and abs(val'+Msgs[j-1,1]+')<='+Msgs[j-1,2];
              Combobox3.ItemIndex:=0;
              case ComboBox3.ItemIndex of
                 0: str3:='val'+Msgs[j-1,1];
                 1: str3:='max((val'+Msgs[j-1,1]+'))';
                 2: str3:='min((val'+Msgs[j-1,1]+'))';
                 3: str3:='avg((val'+Msgs[j-1,1]+'))';
              end;

                with adoquery1 do begin
                  if active then close;
                  sql.Clear;

                 sqls:='select floor(mod(savetime,10000)/100) time, '+str3+' value from '+ycname+' where groupno='+Msgs[j-1,0]+' order by floor(mod(savetime,10000)/100)';
                
                 memo1.Lines.Add(sqls);
                 sql.Add(sqls);
                  open;
                  if  adoquery1.eof then  begin
                    i:=blan;
                    while(i<=irow)do  begin
                      XLS.Sheets[0].Asstring[j,i]:='';
                      i:=i+1;
                    end;
                    continue;
                  end else begin
                    i:=blan;
                    first;
                    while(not adoquery1.eof) do begin
                      gno:=adoquery1.fieldbyname('time').asinteger;
                      if (gno>19) then gno:=gno-20
                      else gno:=gno+4;
                      //clh:=midstr(clh,12,2);
                      i:=i+1;

                      XLS.Sheets[0].asfloat[j,gno+blan]:=adoquery1.fieldbyname('value').asfloat;
                      next;
                    end;
                  end;
                end;

                s1:=0;
                s2:=0;
                XLS.Sheets[0].Asstring[j,0]:='';
                XLS.Sheets[0].Asstring[j,1]:='';
               // progressbar1.Position:=j;
              end; //  every point


      end;
      0:begin
           // memo1.Lines.Add('0:begin');
           if load_data=false then begin
              str1:=gettimes(bb_mon,bb_day,0);
              str2:=gettimed(bb_mon,bb_day,0);
              sqls:='insert into yc_table  select * from '+ycname+' where savetime>=';
              sqls:=sqls+str1+' and savetime<'+str2;
              str4:=' and savetime>='+str1+' and savetime<'+str2;
              //showmessage(inttostr(pos(',',ss)));
              if pos(',',ss)=1 then
                ss:=copy(ss,2,length(ss)-1);
             if length(ss)>1 then
                sqls:=sqls+' and groupno in ('+ss+')'
              else
                sqls:=sqls+' and groupno='+ss;
                  sqls:=sqls+' and chgtime is not null';
                    memo1.Lines.Add(sqls);
                with form1.adoquery1 do begin
                  if active then close;
                  sql.Clear;
                  sql.Add(sqls);
                  //commandText:=sqls;
                   //edit1.Text:=sqls;
                  try
                   memo1.Lines.Add(sqls);
                    //execsql;
                  except
                    //edit1.Text:=sqls;
                  end;

                end; //with   **  do begin
            end ;//if load_data
            ycname:='yc_table';

        for j:=1 to jcol do begin
        Application.ProcessMessages;
         progressbar1.Position:=j;
           for gno:=1 to 3 do
                  for vno:=1 to 8 do
                    tcx[gno,vno]:=0;
                     str1:=XLS.Sheets[0].Asstring[j,0];
              if str1='' then  begin

              continue;
              end;
              sqls:='';
              str1:=' group by floor(mod(savetime,1000000)/10000)';
              str2:=' and abs(val'+Msgs[j-1,1]+')<='+Msgs[j-1,2];
              ComboBox3.ItemIndex:=3;
              case ComboBox3.ItemIndex of
                 0: str3:='val'+Msgs[j-1,1];
                 1: str3:='max((val'+Msgs[j-1,1]+'))';
                 2: str3:='min((val'+Msgs[j-1,1]+'))';
                 3: str3:='avg((val'+Msgs[j-1,1]+'))';
              end;
                with adoquery1 do begin
                  if active then close;
                  sql.Clear;
                  sqls:='select floor(mod(savetime,1000000)/10000) time, '+str3+' value from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str4;
                 if ComboBox3.ItemIndex=0 then begin
                    sqls:=sqls+' and mod(savetime,10000)='+qs.Text+qf.Text;
                 end else begin
                    sqls:=sqls+str1;
                 end;
                 sql.Add(sqls);
                 //edit1.Text:=commandtext;
                 //open;
                 open;
                  //form1.Edit1.Text:=commandtext;
                  if  adoquery1.eof then begin

                    i:=blan;
                     memo1.lines.Add(inttostr(i)+'---'+inttostr(irow));
                    while(i<=irow)do  begin
                      XLS.Sheets[0].Asstring[j,i]:='';
                      i:=i+1;
                    end;
                    continue;
                  end else begin
                   // i:=4;
                    first;
                    while(not adoquery1.eof) do begin
                      clh:=adoquery1.fieldbyname('time').asstring;
                      memo1.Lines.Add('---'+clh);
                     // clh:='22';
                     // i:=pos('-',clh);
                      //showmessage(inttostr(i));
                     // while (i>0) do begin
                     //  clh:=midstr(clh,i+1,length(clh)-1);
                       //showmessage(clh);
                      // i:=pos('-',clh);
                      // end;
                      gno:=strtoint(clh);
                      i:=gno+blan-1;
                      //i:=i+1;
                      XLS.Sheets[0].AsFloat[j,i]:=adoquery1.fieldbyname('value').asfloat;
                    //showmessage(ss);
                      next;
                    end;
                  end;
                end;
                 memo1.Lines.Add(inttostr(i)+'---'+inttostr(irow));
                while(i<irow)do  begin

                  i:=i+1;
                  clh:=XLS.Sheets[0].Asstring[j,i];
                  if pos('#',clh)<>1 then continue;
                  //showmessage(midstr(clh,2,2));
                  case  strtoint(midstr(clh,2,2)) of
                   1:begin
                    if s1=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max(avg((val'+Msgs[j-1,1]+'))) ax,min(avg((val'+Msgs[j-1,1]+'))) mn,avg(avg((val'+Msgs[j-1,1]+'))) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str1+str4;
                          adoquery1.sql.Add(sqls);
                          open ;
                          pmax:=adoquery1.fieldbyname('ax').asfloat;
                          pmin:=adoquery1.fieldbyname('mn').asfloat;
                          pavg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=pmax;
                    end;
                    2:begin
                      if s1=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max(avg((val'+Msgs[j-1,1]+'))) ax,min(avg((val'+Msgs[j-1,1]+'))) mn,avg(avg((val'+Msgs[j-1,1]+'))) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str1+str4;
                           adoquery1.sql.Add(sqls);
                          open;
                          pmax:=adoquery1.fieldbyname('ax').asfloat;
                          pmin:=adoquery1.fieldbyname('mn').asfloat;
                          pavg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=pmin;
                    end;
                    3:begin
                      if s1=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select max(avg((val'+Msgs[j-1,1]+'))) ax,min(avg((val'+Msgs[j-1,1]+'))) mn,avg(avg((val'+Msgs[j-1,1]+'))) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str1+str4;
                          //edit2.Text:=sql.Text;
                          sql.Add(sqls);
                          open;
                          pmax:=adoquery1.fieldbyname('ax').asfloat;
                          pmin:=adoquery1.fieldbyname('mn').asfloat;
                          pavg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=pavg;
                    end;
                    4:begin
                      if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                         sqls:='select max((val'+Msgs[j-1,1]+')) ax,min((val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str4;
                          //edit2.Text:=sql.Text;
                          sql.Add(sqls);
                         open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                     { with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select val'+Msgs[j-1,1]+' value from '+ycname+' where (val'+Msgs[j-1,1]+')='+floattostr(fmax)+' and groupno='+Msgs[j-1,0]+str2+str4;
                          memo1.Lines.Add(sqls);
                          sql.Add(sqls);
                         open;
                      end;
                      fmax:=adoquery1.fieldbyname('value').asfloat; }
                      XLS.Sheets[0].AsFloat[j,i]:=fmax;
                    end;
                    5:begin
                       if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                         sqls:='select max((val'+Msgs[j-1,1]+')) ax,min((val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str4;
                          //edit2.Text:=sql.Text;
                          sql.Add(sqls);
                         open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;

                     { with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                          sqls:='select val'+Msgs[j-1,1]+' value from '+ycname+' where (val'+Msgs[j-1,1]+')='+floattostr(fmin)+' and groupno='+Msgs[j-1,0]+str2+str4;
                           memo1.Lines.Add(sqls);
                           sql.Add(sqls);
                         open;
                      end;
                      fmin:=adoquery1.fieldbyname('value').asfloat;}
                      XLS.Sheets[0].AsFloat[j,i]:=fmin;
                    end;
                    6:begin
                       if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                         sqls:='select max((val'+Msgs[j-1,1]+')) ax,min((val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str4;
                          memo1.Lines.Add(sqls);
                          sql.Add(sqls);
                         open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=favg;
                    end;
                    7:begin
                      if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                         sqls:='select max((val'+Msgs[j-1,1]+')) ax,min((val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str4;
                          //edit2.Text:=sql.Text;
                          sql.Add(sqls);
                         open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                      if fmax<>0 then
                        XLS.Sheets[0].AsFloat[j,i]:=(favg/fmax)
                      else
                        XLS.Sheets[0].AsFloat[j,i]:=0;
                    end;

                    8:begin
                     if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                         sqls:='select max((val'+Msgs[j-1,1]+')) ax,min((val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str4;
                          memo1.Lines.Add(sqls);
                          sql.Add(sqls);
                        open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                      with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                           sqls:='select  min(savetime) time from '+ycname+' where groupno='+Msgs[j-1,0]+' and val'+Msgs[j-1,1]+'='+floattostr(fmax)+str4;
                           memo1.Lines.Add(sqls);
                           sql.Add(sqls);
                         open;
                      end;
                      XLS.Sheets[0].Asstring[j,i]:=numtodate(adoquery1.fieldbyname('time').asstring);
                    end;
                    9:begin
                       if s2=0 then begin
                        with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                         sqls:='select max((val'+Msgs[j-1,1]+')) ax,min((val'+Msgs[j-1,1]+')) mn,avg((val'+Msgs[j-1,1]+')) vg from '+ycname+' where groupno='+Msgs[j-1,0]+str2+str4;
                          memo1.Lines.Add(sqls);
                          sql.Add(sqls);
                          open;
                          fmax:=adoquery1.fieldbyname('ax').asfloat;
                          fmin:=adoquery1.fieldbyname('mn').asfloat;
                          favg:=adoquery1.fieldbyname('vg').asfloat;
                        end;
                        s2:=1;
                      end;
                       with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                           sqls:='select  min(savetime) time from '+ycname+' where groupno='+Msgs[j-1,0]+' and val'+Msgs[j-1,1]+'='+floattostr(fmin)+str4;
                          memo1.Lines.Add(sqls);
                          sql.Add(sqls);
                          open;
                      end;
                      XLS.Sheets[0].Asstring[j,i]:=numtodate(adoquery1.fieldbyname('time').asstring);
                    end;
                     10:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[1,4]=0 then begin
                        zero:=Msgs[j-1,2];
                        tcx[2,4]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,4);
                        tcx[1,4]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=(tcx[2,4]/(1440.0*tcx[3,4]));
                    end;
                    11:begin
                      if tcx[3,6]=0 then begin
                        tcx[3,1]:=strtoint(get_minutes(bb_year,bb_mon,1));
                        tcx[3,6]:=1;
                      end;
                      if tcx[1,1]=0 then begin
                        zero:=Msgs[j-1,2];
                        tcx[2,1]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,1);
                        tcx[1,1]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=(tcx[2,1]/tcx[3,1]);
                    end;
                    12:begin
                       if tcx[3,7]=0 then begin
                        tcx[3,2]:=strtoint(get_minutes(bb_year,bb_mon,2));
                        tcx[3,7]:=1;
                      end;
                      if tcx[1,2]=0 then begin
                        zero:=Msgs[j-1,2];
                        tcx[2,2]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,2);
                        tcx[1,2]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=(tcx[2,2]/tcx[3,2]);
                    end;
                    13:begin
                      if tcx[3,8]=0 then begin
                        tcx[3,3]:=strtoint(get_minutes(bb_year,bb_mon,3));
                        tcx[3,8]:=1;
                      end;
                      if tcx[1,3]=0 then begin
                        zero:=Msgs[j-1,2];
                        tcx[2,3]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,3);
                        tcx[1,3]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=(tcx[2,3]/tcx[3,3]);
                    end;
                    14:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[1,8]=0 then begin
                        zero:=Msgs[j-1,2];
                        tcx[2,8]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,8);
                        tcx[1,8]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=(tcx[2,8]/(1440.0*tcx[3,4]));
                    end;
                    15:begin
                      if tcx[3,6]=0 then begin
                        tcx[3,1]:=strtoint(get_minutes(bb_year,bb_mon,1));
                        tcx[3,6]:=1;
                      end;
                      if tcx[1,5]=0 then begin
                        zero:=Msgs[j-1,2];
                        tcx[2,5]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,5);
                        tcx[1,5]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=(tcx[2,5]/tcx[3,1]);
                    end;
                    16:begin
                      if tcx[3,7]=0 then begin
                        tcx[3,2]:=strtoint(get_minutes(bb_year,bb_mon,2));
                        tcx[3,7]:=1;
                      end;
                      if tcx[1,6]=0 then begin
                        zero:=Msgs[j-1,2];
                        tcx[2,6]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,6);
                        tcx[1,6]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=(tcx[2,6]/tcx[3,2]);
                    end;
                    17:begin
                      if tcx[3,8]=0 then begin
                        tcx[3,3]:=strtoint(get_minutes(bb_year,bb_mon,3));
                        tcx[3,8]:=1;
                      end;
                      if tcx[1,7]=0 then begin
                        zero:='0';
                        tcx[2,7]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,7);
                        tcx[1,7]:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=(tcx[2,7]/tcx[3,3]);
                    end;
                    18:begin
                      if tcx[3,5]=0 then begin
                        tcx[3,4]:=getdaysofmon(bb_year,bb_mon);
                        tcx[3,5]:=1;
                      end;
                      if tcx[1,4]=0 then begin
                        zero:='0';
                        tcx[2,4]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,4);
                        tcx[1,4]:=1;
                      end;

                      if tcx[1,8]=0 then begin
                        zero:=Msgs[j-1,2];
                        tcx[2,8]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,8);
                        tcx[1,8]:=1;
                      end;
                      XLS.Sheets[0].Asstring[j,i]:=formatfloat('0.000',100-((tcx[2,4]+tcx[2,8])/(14.40*tcx[3,4])))+'%';
                    end;
                    19:begin
                      if tcx[1,4]=0 then begin
                        zero:='0';
                        tcx[2,4]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,4);
                        tcx[1,4]:=1;
                      end;
                        XLS.Sheets[0].Asfloat[j,i]:=tcx[2,4];
                    end;
                    20:begin
                      if tcx[1,1]=0 then begin
                        zero:='0';
                        tcx[2,1]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,1);
                        tcx[1,1]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,1];
                    end;
                    21:begin
                      if tcx[1,2]=0 then begin
                        zero:='0';
                        tcx[2,2]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,2);
                        tcx[1,2]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,2];
                    end;
                    22:begin
                      if tcx[1,3]=0 then begin
                        zero:='0';
                        tcx[2,3]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,3);
                        tcx[1,3]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,3];
                    end;
                    23:begin
                      if tcx[1,8]=0 then begin
                        zero:='0';
                        tcx[2,8]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,8);
                        tcx[1,8]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,8];
                    end;
                    24:begin
                      if tcx[1,5]=0 then begin
                      zero:='0';
                      tcx[2,5]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,5);
                      tcx[1,5]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,5];
                    end;
                    25:begin
                      if tcx[1,6]=0 then begin
                      zero:='0';
                      tcx[2,6]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,6);
                      tcx[1,6]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,6];
                    end;
                    26:begin
                      if tcx[1,7]=0 then begin
                      zero:='0';
                      tcx[2,7]:=get_monyxsj(strtoint(Msgs[j-1,0]),strtoint(Msgs[j-1,1]),bb_mon,7);
                      tcx[1,7]:=1;
                      end;
                      XLS.Sheets[0].Asfloat[j,i]:=tcx[2,7];               end;

                  end; //case  tongji leixing
                 
                end;  //while
                s1:=0;
                s2:=0;
                XLS.Sheets[0].Asstring[j,0]:='';
                XLS.Sheets[0].Asstring[j,1]:='';
                //progressbar1.Position:=j;
             end; //  every point

      end;
      2:begin
       for i:=4 to 15 do  begin
          str1:=XLS.Sheets[0].Asstring[0,i];
         XLS.Sheets[0].Asstring[0,i]:=hour.Text+str1;
          end;
        memo1.lines.Clear;
        str1:=inttostr(bb_mon);
        str2:=inttostr(bb_mon);
        if bb_day<10 then   str1:=str1+'0'+inttostr(bb_day)
        else str1:=str1+inttostr(bb_day);

        if strtoint(hour.Text)<10 then   str1:=str1+'0'+inttostr(strtoint(hour.Text))
        else str1:=str1+hour.Text;
        str1:=str1+'00';

        if hour.Text='23' then begin bb_day:=bb_day+1;  hour.Text:='-1'; end;
         if bb_day<10 then   str2:=str2+'0'+inttostr(bb_day)
        else str2:=str2+inttostr(bb_day);

        if (strtoint(hour.Text)+1)<10 then   str2:=str2+'0'+inttostr(strtoint(hour.Text)+1)
        else str2:=str2+inttostr(strtoint(hour.Text)+1);
        str2:=str2+'00';


        sqls:='insert into yc_table  select * from '+ycname+' where savetime>=';
        sqls:=sqls+str1+' and savetime<'+str2;
        if pos(',',ss)=1 then
          ss:=copy(ss,2,length(ss)-1);
        if length(ss)>1 then
          sqls:=sqls+' and groupno in ('+ss+')'
        else
          sqls:=sqls+' and groupno='+ss;
          sqls:=sqls+' and chgtime is not null';
             memo1.lines.Add(sqls);
          with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
            sql.Add(sqls);

            try
              execsql;
            except
            end;

        for j:=1 to jcol do begin
             progressbar1.Position:=j;
            str1:= XLS.Sheets[0].Asstring[j,0];
              if str1='' then  begin

              exit;
              end;
              str2:=' and abs(val'+Msgs[j-1,1]+')>'+Msgs[j-1,2];
                with adodataset1 do begin
                  if active then close;
                  //sql.Clear;
                  commandText:='select savetime time,val'+Msgs[j-1,1]+' value from yc_table where groupno='+Msgs[j-1,0]+str2+' order by savetime ';
                  memo1.lines.Add(commandText);
                  open;
                  if  adodataset1.eof then  begin
                    i:=5;
                    while(i<=irow)do  begin
                      XLS.Sheets[0].Asstring[j,i]:='';
                      i:=i+1;
                    end;
                    continue;
                  end else begin
                    first;
                    while(not adodataset1.eof) do begin
                      clh:=adodataset1.fieldbyname('time').asstring;
                      if length(clh)<8 then clh:='0'+clh;
                      clh:=midstr(clh,7,2);
                      i:=(strtoint(clh) div 5)+5;
                      XLS.Sheets[0].AsFloat[j,i]:=(adodataset1.fieldbyname('value').asfloat);
                      next;
                    end;
                  end;
                end;
                while(i<=irow)do  begin
                  i:=i+1;
                  clh:=XLS.Sheets[0].Asstring[j,i];
                  //showmessage(inttostr(pos(ss,'#')));
                  if pos('#',clh)<>1 then continue;
                  case  strtoint(midstr(clh,2,2)) of
                    1:begin
                      if s1=0 then begin
                        with adodataset1 do begin
                          if active then close;
                          //sql.Clear;
                          commandText:='select max(abs(val'+Msgs[j-1,1]+')) ax,min(abs(val'+Msgs[j-1,1]+')) mn,avg(abs(val'+Msgs[j-1,1]+')) vg from yc_table where groupno='+Msgs[j-1,0]+str2;
                          //edit2.Text:=sql.Text;
                          open;
                          pmax:=adodataset1.fieldbyname('ax').asfloat;
                          pmin:=adodataset1.fieldbyname('mn').asfloat;
                          pavg:=adodataset1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      with adodataset1 do begin
                          if active then close;
                          //sql.Clear;
                          commandText:='select val'+Msgs[j-1,1]+' value from yc_table where abs(val'+Msgs[j-1,1]+')='+floattostr(pmax)+' and groupno='+Msgs[j-1,0]+str2;
                          //edit2.Text:=commandText;
                          open;
                      end;
                       pmax:=adodataset1.fieldbyname('value').asfloat;

                      XLS.Sheets[0].AsFloat[j,i]:=(pmax);
                    end;
                    2:begin
                      if s1=0 then begin
                        with adodataset1 do begin
                          if active then close;
                          //sql.Clear;
                          commandText:='select max(abs(val'+Msgs[j-1,1]+')) ax,min(abs(val'+Msgs[j-1,1]+')) mn,avg(abs(val'+Msgs[j-1,1]+')) vg from yc_table where groupno='+Msgs[j-1,0]+str2;
                          //edit2.Text:=sql.Text;
                          open;
                          pmax:=adodataset1.fieldbyname('ax').asfloat;
                          pmin:=adodataset1.fieldbyname('mn').asfloat;
                          pavg:=adodataset1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      with adodataset1 do begin
                          if active then close;
                          //sql.Clear;
                          commandText:='select val'+Msgs[j-1,1]+' value from yc_table where abs(val'+Msgs[j-1,1]+')='+floattostr(pmin)+' and groupno='+Msgs[j-1,0]+str2;
                          //edit2.Text:=sql.Text;
                          open;
                      end;
                       pmin:=adodataset1.fieldbyname('value').asfloat;

                      XLS.Sheets[0].AsFloat[j,i]:=(pmin);
                    end;
                    3:begin
                      if s1=0 then begin
                        with adodataset1 do begin
                          if active then close;
                          //sql.Clear;
                          commandText:='select max(abs(val'+Msgs[j-1,1]+')) ax,min(abs(val'+Msgs[j-1,1]+')) mn,avg(abs(val'+Msgs[j-1,1]+')) vg from yc_table groupno='+Msgs[j-1,0]+str2;
                          //edit2.Text:=sql.Text;
                          open;
                          pmax:=adodataset1.fieldbyname('ax').asfloat;
                          pmin:=adodataset1.fieldbyname('mn').asfloat;
                          pavg:=adodataset1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      XLS.Sheets[0].AsFloat[j,i]:=(pavg);
                    end;
                    4:begin
                      if s1=0 then begin
                        with adodataset1 do begin
                          if active then close;
                          //sql.Clear;
                          commandText:='select max(abs(val'+Msgs[j-1,1]+')) ax,min(abs(val'+Msgs[j-1,1]+')) mn,avg(abs(val'+Msgs[j-1,1]+')) vg from yc_table where groupno='+Msgs[j-1,0]+str2;
                          //edit2.Text:=sql.Text;
                          open;
                          pmax:=adodataset1.fieldbyname('ax').asfloat;
                          pmin:=adodataset1.fieldbyname('mn').asfloat;
                          pavg:=adodataset1.fieldbyname('vg').asfloat;
                        end;
                        s1:=1;
                      end;
                      if pmax<>0 then
                        XLS.Sheets[0].AsFloat[j,i]:=(pavg/pmax)
                      else
                        XLS.Sheets[0].AsFloat[j,i]:=0;
                    end;
                    5:begin
                      with adodataset1 do begin
                          if active then close;
                          //sql.Clear;
                          commandText:='select  min(savetime) time from yc_table where groupno='+Msgs[j-1,0]+' and  val'+Msgs[j-1,1]+'='+floattostr(pmax);
                          //edit2.Text:=sql.Text;
                          open;
                      end;
                      XLS.Sheets[0].Asstring[j,i]:=numtodate(adodataset1.fieldbyname('time').asstring);
                    end;
                    6:begin
                       with adodataset1 do begin
                          if active then close;
                          //sql.Clear;
                          commandText:='select  min(savetime) time from yc_table where groupno='+Msgs[j-1,0]+' and  val'+Msgs[j-1,1]+'='+floattostr(pmin);
                          //edit2.Text:=sql.Text;
                          open;
                      end;
                      XLS.Sheets[0].Asstring[j,i]:=numtodate(adodataset1.fieldbyname('time').asstring);
                    end;
                  end; //case
                end;  //while
                s1:=0;
                XLS.Sheets[0].Asstring[j,0]:='';
                 XLS.Sheets[0].Asstring[j,1]:='';

                //progressbar1.Position:=j;
              end; //  every point
        end; //with   **  do begin
      end;

      5:begin
        //screen.cursor:=crhourglass;
        cre_view(irow,jcol);
        for i:=0 to jcol do begin
          for j:=0 to irow do begin
          try
            ss:= XLS.Sheets[0].Asstring[i,j];
          except
            continue;
          end;
            if ss<>'' then begin
              if ss[1]='#' then begin
                try
                 XLS.Sheets[0].Asstring[i,j]:=do_chuli(ss);

                except

                  XLS.Sheets[0].Asstring[i,j]:='';
                 end;
              end;
            end;

          end;
        end;
        //screen.cursor:=crdefault;
      end;
      3:begin                 //电量表
          
          //showmessage(inttostr(jcol));
       // for j:=0 to 30 do  XLS.Sheets[0].AsString[0,blan+j]:='';
        if bb_day>24 then  begin    //下月
           gno:=getdaysofmon(bb_year,bb_mon);
           sdate:=encodedate(bb_year,bb_mon,25);
           edate:=incMonth(datetimepicker1.Date) ;
           edate:=encodedate(yearof(edate),monthof(edate),24);
        end;
        if bb_day<25 then begin      //当月
           thedate:=incMonth(datetimepicker1.Date,-1) ;
           sdate:=encodedate(yearof(thedate),monthof(thedate),25);
           gno:=getdaysofmon(yearof(thedate),monthof(thedate));
           edate:=encodedate(bb_year,bb_mon,24);
        end;
        vno:=25;
        while vno<=gno do begin
          // XLS.Sheets[0].AsString[0,blan+vno-25]:=inttostr(vno);
           vno:=vno+1;
        end;
         vno:=1;
        while vno<25 do begin
          // XLS.Sheets[0].AsString[0,blan+gno-25+vno]:=inttostr(vno);
           vno:=vno+1;
        end;
          sdate:=encodedate(bb_year,bb_mon,1);
           gno:=getdaysofmon(bb_year,bb_mon);
            edate:=encodedate(bb_year,bb_mon,gno);
        vno:=gno;
        for j:=1 to jcol do begin
         progressbar1.Position:=j;

                 //str1:=' and mod(a.savetime,10000)=0 and b.savetime=a.savetime+10000 ';
                if length(Msgs[j-1,1]) =0 then continue;
                blans:=strtoint(Msgs[j-1,1])+ strtoint(Msgs[j-1,0])*200;
                str1:=' and mod(a.savetime,10000)=0 and b.savetime=get_date('+inttostr(bb_year)+',a.savetime,1) ';
                //str1:=' and mod(a.savetime,10000)=0 and b.savetime=get_date(a.savetime,1) ';
                with adoquery1 do begin
                  if active then close;
                  sql.Clear;
                 //sqls:='select a.savetime time,b.val'+Msgs[j-1,1]+'-a.val'+Msgs[j-1,1]+' value from dn_table a,dn_table b where a.groupno='+Msgs[j-1,0]+' and b.groupno='+Msgs[j-1,0]+str1;
                 sqls:='select did time,st value from h_inc where did>='+inttostr(yearof(sdate)*10000+monthof(sdate)*100+1)+' and did<='+inttostr(yearof(edate)*10000+monthof(edate)*100+31)+' and kind=0 and saveno='+inttostr(blans);
                  sql.Add(sqls);
                  try
                  open;
                  except
                     //edit2.Text:=sqls;
                  end;
                  memo1.lines.Add(sqls);
                  if  adoquery1.eof then
                    i:=35
                  else begin
                    i:=blan;
                    first;
                    while(not adoquery1.eof) do begin
                      clh:=adoquery1.fieldbyname('time').asstring;

                        clh:=midstr(clh,length(clh)-1,2);

                      gno:=strtoint(clh);

                      try
                      fmax:=adoquery1.fieldbyname('value').asfloat;
                     // if fmax<0 then fmax:=0-fmax;
                      XLS.Sheets[0].AsFloat[j,gno+blan-1]:=(fmax);
                      except
                      //edit2.Text:='fuck';
                      end;
                      next;
                    end;
                  end;
                end;
                gno:=31;
                str1:=' and floor(a.savetime/10000)=round(b.savetime/10000) ';
                while(gno+blan-14<irow)do  begin
                  gno:=gno+1;
                  i:=gno+blan-1;
                  clh:=XLS.Sheets[0].Asstring[j,gno+blan-1];
                  if pos('#',clh)<>1 then continue;
                  fn:=0;
                  k:=1;
                 case  strtoint(midstr(clh,2,2)) of
                    1:begin
                         with adoquery1 do begin
                          if active then close;
                          sql.Clear;
                        sqls:='select sum(st) value from h_inc where did>='+inttostr(yearof(sdate)*10000+monthof(sdate)*100+1)+' and did<='+inttostr(yearof(edate)*10000+monthof(edate)*100+31)+' and kind=2 and saveno='+inttostr(blans);

                              sql.Add(sqls);
                               memo1.lines.Add(sqls);
                              open;
                              fn:=fn+adoquery1.fieldbyname('value').asfloat;
                              end;

                       edit1.Text:=sqls;
                       XLS.Sheets[0].AsFloat[j,i]:=(adoquery1.fieldbyname('value').asfloat);
                    end;
                    2:begin
                     with adoquery1 do begin
                          if active then close;
                           sql.Clear;
                      sqls:='select sum(st) value from h_inc where did>='+inttostr(yearof(sdate)*10000+monthof(sdate)*100+1)+' and did<='+inttostr(yearof(edate)*10000+monthof(edate)*100+31)+' and kind=3 and saveno='+inttostr(blans);

                               sql.Add(sqls);
                               memo1.lines.Add(sqls);
                              open;
                              fn:=fn+adoquery1.fieldbyname('value').asfloat;
                              end;

                       edit1.Text:=sqls;
                       XLS.Sheets[0].AsFloat[j,i]:=(adoquery1.fieldbyname('value').asfloat);

                    end;
                    3:begin
                     sqls:='select sum(st) value from h_inc where did>='+inttostr(yearof(sdate)*10000+monthof(sdate)*100+1)+' and did<='+inttostr(yearof(edate)*10000+monthof(edate)*100+31)+' and kind=4 and saveno='+inttostr(blans);

                    with adoquery1 do begin
                          if active then close;
                           sql.Clear;

                               sql.Add(sqls);
                               memo1.lines.Add(sqls);
                              open;
                              fn:=fn+adoquery1.fieldbyname('value').asfloat;
                              end;

                       edit1.Text:=sqls;
                       XLS.Sheets[0].AsFloat[j,i]:=(adoquery1.fieldbyname('value').asfloat);

                    end;
                    4:begin
                    sqls:='select max(val'+Msgs[j-1,1]+')-min(val'+Msgs[j-1,1]+') value from dn_table where groupno='+Msgs[j-1,0]+' and val'+Msgs[j-1,1]+' is not null and mod(savetime,10000)=0';                          //edit2.Text:=sql.Text;
                          sqls:='select sum(st) value from h_inc where did>='+inttostr(yearof(sdate)*10000+monthof(sdate)*100+1)+' and did<='+inttostr(yearof(edate)*10000+monthof(edate)*100+31)+' and kind=0 and saveno='+inttostr(blans);

                        with adoquery1 do begin
                          if active then close;
                            sql.Clear;
                             sql.Add(sqls);
                           memo1.lines.Add(sqls);
                          open;
                        end;
                       XLS.Sheets[0].AsFloat[j,i]:=(adoquery1.fieldbyname('value').asfloat);
                    end;
                  end; //case  tongji leixing
                end;  //while
                XLS.Sheets[0].Asstring[j,0]:='';

             end; //  every point
      end;
    end;    //case bbsty
   //showmessage(inttostr(blan)) ;
if bbsty<>5 then begin
   s1:=1;
   s2:=0;

   memo1.Lines.Add(inttostr(blan));
          for j:=1 to jcol  do begin
            str2:=XLS.Sheets[0].Asstring[j,blan-2];
            ss:=XLS.Sheets[0].Asstring[j+1,blan-2];
                    if  (str2=ss) then
                      s2:=s2+1
                    else begin
                       //qzw:=ExcelWorksheet1.Range[ExcelWorksheet1.Cells.Item[blan-2,s1],ExcelWorksheet1.Cells.Item[blan-2,s1+s2-1]];//单元格从A2到G2
                       //qzw.merge;
                      // XLS.Sheets[0].MergedCells.Add(s1,blan-2,s1+s2,blan-2);
                       if pos('总',str2)<>0 then    // qzw.Merge;//he
                          str2:=copy(str2,1,length(str2)-length('总'));
                       //qzw.value:=str2;
                      // qzw.HorizontalAlignment:=xlHAlignCenter;
                      s1:=j+1;
                      s2:=1;
                    end;
          end;
          s1:=1;
          s2:=0;
         for j:=1 to jcol  do begin
          str2:=XLS.Sheets[0].Asstring[j,blan-4];
          ss:=XLS.Sheets[0].Asstring[j+1,blan-4];
          if  (str2=ss) then
            s2:=s2+1
          else begin
            // qzw:=ExcelWorksheet1.Range[ExcelWorksheet1.Cells.Item[blan-3,s1],ExcelWorksheet1.Cells.Item[blan-3,s1+s2-1]];//单元格从A2到G2
            // qzw.merge;
           // XLS.Sheets[0].MergedCells.Add(s1,blan-3,s1+s2-1,blan-3);
             if pos('总',str2)<>0 then    // qzw.Merge;//he
                str2:=copy(str2,1,length(str2)-length('总'));
            // qzw.value:=str2;
             //qzw.HorizontalAlignment:=xlHAlignCenter;
            s1:=j+1;
            s2:=1;
          end;
        end;
  // showmessage(str2);

   Msgs[2,2]:=XLS.Sheets[0].Asstring[0,0];
  //Msgs[0,0]:=ExcelWorksheet1.Cells.Item[2,1];
  str1:=XLS.Sheets[0].Asstring[0,blan-3];
  case bbsty of
    0: begin
      case combobox3.ItemIndex of
      0: str1:='取点值';
      1:  str1:='日最大值';
      2:  str1:='日最小值';
      3:  str1:='日平均值';
      end;
    end;
    1:  begin
    case combobox3.ItemIndex of
      0: str1:='取点值';
      1:  str1:='时最大值';
      2:  str1:='时最小值';
      3:  str1:='时平均值';
      end;
    end;
   { 3: begin
       qzw:=excelworksheet1.rows;
       qzw.rows[2].insert;
    end; }
   end;
  XLS.Sheets[0].Asstring[0,blan-3]:=str1;
   case bbsty of
    0,3:  Msgs[1,1]:='报表时间：'+inttostr(bb_year)+'年'+inttostr(bb_mon)+'月';
    1,2,4:  Msgs[1,1]:='报表时间：'+inttostr(bb_year)+'年'+inttostr(bb_mon)+'月'+inttostr(bb_day)+'日';
   end;
    XLS.Sheets[0].Cell[0,0].FontName := '宋体';

   XLS.Sheets[0].Rows[0].HeightPt:=40;
   XLS.Sheets[0].Cell[0,0].FontSize := 14;
   XLS.Sheets[0].Cell[0,0].VertAlignment := cvaCenter;
   XLS.Sheets[0].Cell[0,0].HorizAlignment :=  chaCenter;
   XLS.Sheets[0].MergedCells.Add(0,0,jcol,0);


    XLS.Sheets[0].Asstring[0,1]:=Msgs[1,1];
    XLS.Sheets[0].MergedCells.Add(0,1,trunc(jcol div 2),1);
     XLS.Sheets[0].Cell[0,1].HorizAlignment :=  chaLeft;

      XLS.Sheets[0].Asstring[trunc(jcol div 2)+1,1]:='单位：'+dw[dwidx];
    XLS.Sheets[0].MergedCells.Add(trunc(jcol div 2)+1,1,jcol,1);
     XLS.Sheets[0].Cell[trunc(jcol div 2)+1,1].HorizAlignment :=  chaRight;
      end;
      end else begin       // bb_valid=1
          memo1.lines.Add(inttostr(jcol)+','+inttostr(irow)) ;
          // XLS.Sheets[0].Asstring[0,1]:= inttostr(bb_year)+'年'+inttostr(bb_mon)+'月';
          for j:=1 to jcol do begin
              for i:=1 to irow do begin

                    try
                      ss:= XLS.Sheets[0].Asstring[j,i];

                    except
                      continue;
                    end;
                      if ss<>'' then begin

                        if ss[1]='&' then begin
                          if (pos('日报表',ss)>0) then ss:=ss+','+inttostr(bb_year*10000+bb_mon*100+bb_day )
                          else ss:=ss+','+inttostr(bb_year*100+bb_mon) ;
                          try
                           XLS.Sheets[0].Asfloat[j,i]:=strtofloat(do_yinyong(ss));

                          except
                            memo1.lines.Add('eeeer:'+do_yinyong(ss));
                            XLS.Sheets[0].Asstring[j,i]:='';
                           end;
                        end;
                         if ss[1]='^' then begin

                          try
                           XLS.Sheets[0].Asstring[j,i]:=inttostr(bb_year)+'年'+inttostr(bb_mon)+'月'+inttostr(bb_day)+'日';

                          except
                            memo1.lines.Add('e2eeer');
                            XLS.Sheets[0].Asstring[j,i]:='';
                           end;
                        end;
                      end;




              end;

          end;



      end;
     xls.Calculate;
     if ((bbsty=3) and (bb_day>24)) then decodedate(edate,bb_year,bb_mon,bb_day);
     if ( (bbsty=1) or (bbsty=4)) then xlss:=xls_
     else
       xlss:=xls__;
     if not FileExists(rptdir+inttostr(combobox1.ItemIndex)+'\') then
  try
    begin
      CreateDir(rptdir+inttostr(combobox1.ItemIndex)+'\');
    end;
  except
    memo1.lines.Add('Cannot Create '+rptdir+inttostr(combobox1.ItemIndex));
  end;
   if not FileExists(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\') then
  try
    begin
      CreateDir(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\');
    end;
  except
    memo1.lines.Add('Cannot Create '+rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\');
  end;
  if not FileExists(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\') then
  try
    begin
      CreateDir(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\');
    end;
  except
    memo1.lines.Add('Cannot Create '+rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\');
  end;


  if ((bbsty<>1) and (bbsty<>4)) then begin
  try
  xls.SaveToFile(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+xlss);
   memo1.lines.Add(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+xlss);
 except
  // memo1.lines.Add(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_mon)+'\'+xls);
 end;
 end else begin
 if not FileExists(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+inttostr(bb_day)+'\') then
  try
    begin
      CreateDir(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+inttostr(bb_day)+'\');
    end;
  except
   memo1.lines.Add('Cannot Create '+rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+inttostr(bb_day)+'\');
  end;
 try
   xls.SaveToFile(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+inttostr(bb_day)+'\'+xlss);
    memo1.lines.Add(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+inttostr(bb_day)+'\'+xlss);
 except
  // memo1.lines.Add(rptdir+inttostr(combobox1.ItemIndex)+'\'+inttostr(bb_mon)+'\'+xls);
 end;

 
end;
{ if  checkbox1.Checked then begin

   with  ExcelWorksheet1.PageSetup do begin
     try
        CenterFooter := '第&P页';
        footermargin:=8;
        printtitlerows:='A1:M2';
        leftmargin:=24;
        rightmargin:=24;
        topmargin:=28;
        bottommARgin:=28;
        orientation:=2;
     except
     end;
   end;
   try
     ExcelWorksheet1.PrintOut;       //打印
   except
   end;
  end; }
  //done:=true;
  //ExcelWorksheet1.Protect;
end;


  function get_ch(s:shortstring):smallint;
  begin
    result:=strtoint(copy(s,7,3));
  end;

  function get_xh(s:shortstring):smallint;
  begin
    result:=strtoint(copy(s,11,4));
  end;

  function get_day(s:shortstring):word;
  begin
    result:=strtoint(copy(s,16,2));
  end;

  function get_hour(s:shortstring):word;
  begin
    result:=strtoint(copy(s,19,2));
  end;

  function get_minute(s:shortstring):word;
  begin
    result:=strtoint(copy(s,22,2));
  end;

  function get_mon(s:string):word;
  begin
    result:=strtoint(copy(s,25,2));
  end;
  function get_minutes(year,mon,lb:integer):shortstring;
var
  i,days, gf_tms,yh_tms,dg_tms:integer;
begin
  result:='0';
    days:=getdaysofmon(year,mon);
  dg_tms:=0;
  yh_tms:=0;
  gf_tms:=0;
  for i:=1 to 48 do begin
    case dnsd[mon,i] of
      4:begin {低谷}
        inc(dg_tms);
      end;
      3:begin {腰荷}
        inc(yh_tms);
      end;
      1..2:begin {高峰}
        inc(gf_tms);
      end;
    end;
  end;
  //dg_tms:=dg_tms*days;
  //yh_tms:=yh_tms*days;
  //gf_tms:=gf_tms*days;
  case lb of
    1:begin //高峰
      result:=inttostr(gf_tms*days*30);
    end;
    2:begin //腰荷
      result:=inttostr(yh_tms*days*30);
    end;
    3:begin //低谷
      result:=inttostr(dg_tms*days*30);
    end;
  end;
end;

function get_yc_MonTJ2(ch,xh:longint;mon:smallint; slct:byte):real;
var
  str1,str2,str3,str4,sqls:string;
begin
    if (slct=4) then begin
      result:=get_yc_MonTJ2(ch,xh,mon,3)/get_yc_MonTJ2(ch,xh,mon,1);
      exit;
    end;
    if (slct=22) then begin
      result:=get_yc_MonFhl(ch,xh,mon);
      exit;
    end;
  str1:=gettimes(mon,1,0);
  str2:=gettimed(mon,1,0);
  str3:=' and abs(val'+inttostr(xh)+')>'+zero ;

  str4:=' and (round(mod(savetime,10000)/100)*2+round(mod(mod(savetime,10000),100)/30)+1) in(';
  case slct of
   1:sqls:='select max(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2;
   2:sqls:='select min(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3;
   3:sqls:='select avg(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3;
   7:sqls:='select max(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str4+dnsdstr[mon,2]+')';
   8:sqls:='select min(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+str4+dnsdstr[mon,2]+')';
   9:sqls:='select avg(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+str4+dnsdstr[mon,2]+')';
   10:sqls:='select max(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str4+dnsdstr[mon,3]+')';
   11:sqls:='select min(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+str4+dnsdstr[mon,3]+')';
   12:sqls:='select avg(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+str4+dnsdstr[mon,3]+')';
   13:sqls:='select max(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str4+dnsdstr[mon,4]+')';
   14:sqls:='select min(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+str4+dnsdstr[mon,4]+')';
   15:sqls:='select avg(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+str4+dnsdstr[mon,4]+')';
   23:sqls:='select max(max(val'+inttostr(xh)+')-min(val'+inttostr(xh)+')) from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+'group by floor(savetime)';
  end;
  //showmessage(inttostr(length(sqls)));
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    //showmessage(inttostr(length(sql.Strings[0])));
    result:=fields[0].AsFloat;
  end;
end;
function get_yxd(ch,xh:longint;mon,day:smallint; slct:byte):integer;
var
  str1,str2,str3,str4,sqls:string;
begin
 str1:=gettimes(mon,day,1);
  str2:=gettimed(mon,day,1);
  str3:=gettimed(mon,day,0);
  str4:=' and val'+inttostr(xh)+' is not null ';
  case slct of
   1:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4;
   2:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str3+str4;
  end;

  with form1.adodataset1  do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;

    open;
    //showmessage(inttostr(length(sql.Strings[0])));
    result:=fields[0].AsInteger;
  end;
end;
function get_yc_DayTJ2(ch,xh:longint;mon,day:smallint; slct:byte):real;
var
  str1,str2,str3,str4,sqls:string;
begin

 str1:=gettimes(mon,day,1);
  str2:=gettimed(mon,day,1);
  str3:=' and abs(val'+inttostr(xh)+')> '+zero;
  str4:=' and (round(mod(savetime,10000)/100)*2+round(mod(mod(savetime,10000),100)/30)+1) in(';
  case slct of
   11:sqls:='select max(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str4+dnsdstr[mon,3]+')';
   13:sqls:='select min(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+str4+dnsdstr[mon,4]+')';
   7:sqls:='select max(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str4+dnsdstr[mon,2]+')';
   8:sqls:='select min(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+str4+dnsdstr[mon,2]+')';
   10:sqls:='select max(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str4+dnsdstr[mon,3]+')';
   14:sqls:='select min(val'+inttostr(xh)+') from yc_table where groupno='+inttostr(ch)+' and floor(savetime)>='+str1+' and floor(savetime)<'+str2+str3+str4+dnsdstr[mon,4]+')';
    end;
  //showmessage(inttostr(length(sqls)));
  with form1.adodataset1  do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    //showmessage(inttostr(length(sql.Strings[0])));
    result:=fields[0].AsFloat;
  end;
end;
function get_yc_DayTJ(ch,xh:longint;mon,day:smallint; slct:byte):string;
var

  str1,str2,str3,str4:string;
  sqls:string;
begin

  str1:=gettimes(mon,day,1);
  str2:=gettimed(mon,day,1);
  str3:=' and val'+inttostr(xh)+'=';
 str4:=' and (round(mod(savetime,10000)/100)*2+round(mod(mod(savetime,10000),100)/30)+1) in(';
  case slct of
   7:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,2]+')'+str3+floattostr(get_yc_dayTJ2(ch,xh,mon,day,7));
   8:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,2]+')'+str3+floattostr(get_yc_dayTJ2(ch,xh,mon,day,8));
   10:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,3]+')'+str3+floattostr(get_yc_dayTJ2(ch,xh,mon,day,10));
   11:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,3]+')'+str3+floattostr(get_yc_dayTJ2(ch,xh,mon,day,11));
   13:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,4]+')'+str3+floattostr(get_yc_dayTJ2(ch,xh,mon,day,13));
   14:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,4]+')'+str3+floattostr(get_yc_dayTJ2(ch,xh,mon,day,14));
  end;
  with form1.adodataset1 do begin
    if active then close;
   // sql.Clear;
    commandtext:=sqls;
   // showmessage(sql.Strings[0]);
    open;
    result:=fields[0].Asstring;
  end;
end;

function get_yc_MonTJ(ch,xh:longint;mon:smallint; slct:byte):string;
var

  str1,str2,str3,str4:string;
  sqls:string;
begin

  str1:=gettimes(mon,1,0);
  str2:=gettimed(mon,1,0);
  str3:=' and val'+inttostr(xh)+'=';
  str4:=' and (round(mod(savetime,10000)/100)*2+round(mod(mod(savetime,10000),100)/30)+1) in(';
  case slct of
   5:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+floattostr(get_yc_monTJ2(ch,xh,mon,1));
   6:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+floattostr(get_yc_monTJ2(ch,xh,mon,2));
   16:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,2]+')'+str3+floattostr(get_yc_monTJ2(ch,xh,mon,7));
   19:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,2]+')'+str3+floattostr(get_yc_monTJ2(ch,xh,mon,8));
   17:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,3]+')'+str3+floattostr(get_yc_monTJ2(ch,xh,mon,10));
   20:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,3]+')'+str3+floattostr(get_yc_monTJ2(ch,xh,mon,11));
   18:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,4]+')'+str3+floattostr(get_yc_monTJ2(ch,xh,mon,13));
   21:sqls:='select min(savetime) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,4]+')'+str3+floattostr(get_yc_monTJ2(ch,xh,mon,14));
  end;
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    //showmessage(sql.Strings[0]);
    open;
    result:=fields[0].Asstring;
  end;
end;

function get_kgfhcs(ch,xh:longint;mon,day,lb:smallint):integer;
var
  ymd1,ymd2:integer;
  str3,str4,sqls:string;
begin

  ymd1:=bb_year*10000+mon*100+day;
  ymd2:=1+ymd1;
  str3:=' and val'+inttostr(xh)+'=';
  str4:=' and (round(mod(savetime,10000)/100)*2+round(mod(mod(savetime,10000),100)/30)+1) in(';
  case lb of
   1:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''合闸'',''分闸'')';
   2:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''合闸'',''分闸'')';
   3:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''合闸''';
   4:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''分闸''';
  end;
  //showmessage(sqls);
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    result:=fields[0].asinteger;
  end;

end;
function get_ykcs(ch,xh:longint;mon,day,lb:smallint):integer;
var
  ymd1,ymd2:integer;
  str3,str4,sqls:string;
begin

  ymd1:=bb_year*10000+mon*100+day;
  ymd2:=1+ymd1;
 str3:='to_date('''+inttostr(bb_year)+'-'+inttostr(mon)+'-'+inttostr(day)+''',''yyyy-mm-dd'')';

  str4:='1+'+str3;
  case lb of
   1:sqls:='select count(*) from event_table where ch='+inttostr(ch)+'  and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''合闸成功'',''分闸成功'') and ms not like ''%操作:AVC%''';
   2:sqls:='select count(*) from event_table where ch='+inttostr(ch)+'  and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''上调'',''下调'') and ms not like ''%操作:AVC%''';
   3:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''信号复归''';
   4:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''上调'',''下调'',''合闸成功'',''分闸成功'') and ms not like ''%操作:AVC%''';

    5:sqls:='select count(*) from oprecord where ch='+inttostr(ch)+'  and optime>='+str3+' and optime<'+str4+' and opcase like ''%执行%开关%'' ';
   6:sqls:='select count(*) from oprecord where ch='+inttostr(ch)+'   and optime>='+str3+' and optime<'+str4+' and (opcase like ''%执行%上调%'' or opcase like ''%执行%下调%'' )';
   7:sqls:='select count(*) from oprecord where ch='+inttostr(ch)+'  and optime>='+str3+' and optime<'+str4+' and opcase like ''%执行%开关%''  and dm like ''%复归%''';
   8:sqls:='select count(*) from oprecord where ch='+inttostr(ch)+'  and optime>='+str3+' and optime<'+str4+' and opcase like ''%执行%''';

    end;
  //if lb=1 then form1.Edit5.Text:=sqls;
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    result:=fields[0].asinteger;
  end;


end;
function get_ykcs_mon(ch,xh:longint;mon,lb:smallint):integer;
var
  ymd1,ymd2:integer;
  str3,str4,sqls:string;
begin

  ymd1:=bb_year*10000+mon*100+1;
  ymd2:=getdaysofmon(bb_year,mon)+ymd1;
  str3:='to_date('''+inttostr(bb_year)+'-'+inttostr(mon)+'-1'',''yyyy-mm-dd'')';
  str4:=inttostr(getdaysofmon(bb_year,mon))+'+'+str3;

  case lb of
   1:sqls:='select count(*) from event_table where ch='+inttostr(ch)+'  and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''合闸成功'',''分闸成功'') and ms not like ''%操作:AVC%''';
   2:sqls:='select count(*) from event_table where ch='+inttostr(ch)+'  and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''上调'',''下调'') and ms not like ''%操作:AVC%''';
   3:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''信号复归''';
   4:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''上调'',''下调'',''合闸成功'',''分闸成功'') ms not like ''%操作:AVC%''';

    5:sqls:='select count(*) from oprecord where ch='+inttostr(ch)+'  and optime>='+str3+' and optime<'+str4+' and opcase like ''%执行%开关%'' ';
   6:sqls:='select count(*) from oprecord where ch='+inttostr(ch)+'   and optime>='+str3+' and optime<'+str4+'  and (opcase like ''%执行%上调%'' or opcase like ''%执行%下调%'' )';
   7:sqls:='select count(*) from oprecord where ch='+inttostr(ch)+'  and optime>='+str3+' and optime<'+str4+' and opcase like ''%执行%开关%''  and dm like ''%复归%''';
   8:sqls:='select count(*) from oprecord where ch='+inttostr(ch)+'  and optime>='+str3+' and optime<'+str4+' and opcase like ''%执行%''';
  end;
 //if lb=3 then form1.Edit5.Text:=sqls;
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    result:=fields[0].asinteger;
  end;
end;
function get_yk_moncgl(ch,xh:longint;mon,day,lb:smallint):real;
begin
case lb of
1:begin
  try
    if  get_ykcs(ch,xh,mon,day,5)<>0  then
      result:=get_ykcs(ch,xh,mon,day,1)/get_ykcs(ch,xh,mon,day,5)
    else result:=0;
  except
   result:=0;
  end;
end;
2:begin
  try
    if  get_ykcs(ch,xh,mon,day,6)<>0  then
      result:=get_ykcs(ch,xh,mon,day,2)/get_ykcs(ch,xh,mon,day,6)
    else result:=0;
  except
   result:=0;
  end;
end;
3:begin
  try
    if  get_ykcs(ch,xh,mon,day,7)<>0  then
      result:=get_ykcs(ch,xh,mon,day,3)/get_ykcs(ch,xh,mon,day,7)
    else result:=0;
  except
   result:=0;
  end;
end;
4:begin
  try
    if  get_ykcs(ch,xh,mon,day,8)<>0  then
      result:=get_ykcs(ch,xh,mon,day,4)/get_ykcs(ch,xh,mon,day,8)
    else result:=0;
  except
   result:=0;
  end;
end;
5:begin
  try
    if  get_ykcs_mon(ch,xh,mon,5)<>0  then
      result:=get_ykcs_mon(ch,xh,mon,1)/get_ykcs_mon(ch,xh,mon,5)
    else result:=0;
  except
   result:=0;
  end;
end;
6:begin
  try
    if  get_ykcs_mon(ch,xh,mon,6)<>0  then
      result:=get_ykcs_mon(ch,xh,mon,2)/get_ykcs_mon(ch,xh,mon,6)
    else result:=0;
  except
   result:=0;
  end;
end;
7:begin
  try
    if  get_ykcs_mon(ch,xh,mon,7)<>0  then
      result:=get_ykcs_mon(ch,xh,mon,3)/get_ykcs_mon(ch,xh,mon,7)
    else result:=0;
  except
   result:=0;
  end;
end;
8:begin
  try
    if  get_ykcs_mon(ch,xh,mon,8)<>0  then
      result:=get_ykcs_mon(ch,xh,mon,4)/get_ykcs_mon(ch,xh,mon,8)
    else result:=0;
  except
   result:=0;
  end;
end;
end;
end;

function get_eventcs(ch,xh:longint;mon,day,lb:smallint):integer;
var
  ymd1,ymd2:integer;
  str3,str4,sqls:string;
begin

  ymd1:=bb_year*10000+mon*100+day;
  ymd2:=1+ymd1;
  str3:=' and val'+inttostr(xh)+'=';
  str4:=' and (round((savetime-round(savetime))*48)+1) in(';
  case lb of
    1:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''事故跳闸''';
   2:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''事故跳闸''';
   3:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''中断''';
   4:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''事故跳闸''';
  end;
  //showmessage(sqls);
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    result:=fields[0].asinteger;
  end;


end;
function get_eventcs_mon(ch,xh:longint;mon,lb:smallint):integer;
var
  ymd1,ymd2:integer;
  str3,str4,sqls:string;
begin

  ymd1:=bb_year*10000+mon*100+1;
  ymd2:=getdaysofmon(bb_year,mon)+ymd1;
  str3:=' and val'+inttostr(xh)+'=';
  str4:=' and (round((savetime-round(savetime))*48)+1) in(';
  case lb of
   1:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''中断''';
   2:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''事故跳闸''';
   3:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''事故跳闸''';
   4:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''事故跳闸''';
  end;
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    result:=fields[0].asinteger;
  end;
end;
function get_kgfhcs_mon(ch,xh:longint;mon,lb:smallint):integer;
var
  ymd1,ymd2:integer;
  str3,str4,sqls:string;
begin

  ymd1:=bb_year*10000+mon*100+1;
  ymd2:=getdaysofmon(bb_year,mon)+ymd1;
  str3:=' and val'+inttostr(xh)+'=';
  str4:=' and (round((savetime-round(savetime))*48)+1) in(';
  case lb of
   1:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''合闸'',''分闸'')';
   2:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt in (''合闸'',''分闸'')';
   3:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''合闸''';
   4:sqls:='select count(*) from event_table where ch='+inttostr(ch)+' and xh='+inttostr(xh)+' and ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+' and zt=''分闸''';
  end;
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    result:=fields[0].asinteger;
  end;
end;
 procedure ch2gno(var ch,xh:longint;tclh:byte);
  begin
     with form1.adodataset1 do begin
         if active then close;
         //sql.Clear;
         case tclh of
         2: begin
           commandtext:='select saveno,zerorange from prtuana where rtuno='+inttostr(ch)+' and sn='+inttostr(xh);
           open;
           ch:=fields[0].AsInteger div 200;
           xh:=fields[0].AsInteger mod 200;
           zero:=fields[1].AsString;
         end;
         3:    ;
         4:begin
           commandtext:='select saveno from prtupul where rtuno='+inttostr(ch)+' and sn='+inttostr(xh);
           open;
           //showmessage(commandtext);
           ch:=fields[0].AsInteger div 200;
           xh:=fields[0].AsInteger mod 200;
         end;
         end;
         //showmessage(sql.text);

     end;
    // showmessage(inttostr(ch)+'_'+inttostr(xh));
  end;
function get_allyxsj(ch,xh:longint;mon,day,slct:smallint):integer;
var
  {1:高峰超上限，2:低谷超上限, 3:腰荷超上限, 4:日超上限时间}
  {5:日超上限率，6:日超下限率, 7:日合格率, 8:日超下限时间}

  yc_up,yc_dn,str1,str2,str3,str4,str5,sqls:string;
begin

  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:='select upperlimit,lowerlimit from prtuana where saveno='+inttostr(ch*200+xh);
    open;
    try
    yc_up:=fields[0].Asstring;yc_dn:=fields[1].AsString;
    except
    yc_up:='0';yc_dn:='0';
    end;
  end;
  //ch2gno(ch,xh,2);
  //showmessage(inttostr(ch)+'_'+inttostr(xh));
  str1:=gettimes(mon,day,1);

  str2:=gettimed(mon,day,1);
  str3:=' and val'+inttostr(xh);
  str4:=' and (round(mod(savetime,10000)/100)*2+round(mod(mod(savetime,10000),100)/30)+1) in(';
  str5:=' and abs(val'+inttostr(xh)+')>'+zero ;
  case slct of
   1:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,2]+')'+str3+'>'+yc_up;
   2:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,3]+')'+str3+'>'+yc_up;
   3:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,4]+')'+str3+'>'+yc_up;
   4:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+'>'+yc_up;
   5:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,2]+')'+str3+'<'+yc_dn+str5;
   6:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,3]+')'+str3+'<'+yc_dn+str5;
   7:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,4]+')'+str3+'<'+yc_dn+str5;
   8:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+'<'+yc_dn;
  end;

  with form1.adodataset1 do begin
    if active then close;
   //sql.Clear;
    commandtext:=sqls;
   //form1.Edit1.Text:=sqls;
   open;
    result:=fields[0].asinteger*5;
  end;
end;

function get_cos_day(ch,xh:longint;mon,day,slct:smallint):integer;
var


  yc_up,yc_dn,str1,str2,str3,str4,str5,sqls:string;
begin


  str1:=gettimes(mon,day,1);

  str2:=gettimed(mon,day,1);
  str3:=' and val'+inttostr(xh);
  str4:=' and (round(mod(savetime,10000)/100)*2+round(mod(mod(savetime,10000),100)/30)+1) in(';
  str5:=' and val'+inttostr(xh)+'>'+zero ;
  case slct of
   1:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+' in (1,2,-1,-2)';
   2:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+' in (1,2)';
  end;

  with form1.adodataset1 do begin
    if active then close;
   //sql.Clear;
    commandtext:=sqls;
   open;
    result:=fields[0].asinteger;
  end;
end;
function get_cos_day2(ch,xh:longint;mon,day:smallint):single;
var
  v1,v2:integer;
begin
   v1:=  get_cos_day(ch,xh,mon,day,2);
   v2:= get_cos_day(ch,xh,mon,day,1);
   if v2<>0 then begin
   result:=v1/v2;
  end  else   begin
    result:=0;
  end;

end;

function get_cos_mon(ch,xh:longint;mon,slct:smallint):integer;
var


  yc_up,yc_dn,str1,str2,str3,str4,str5,sqls:string;
begin


   str1:=gettimes(mon,1,0);

  str2:=gettimed(mon,1,0);
   str3:=' and val'+inttostr(xh);
  str4:=' and (round(mod(savetime,10000)/100)*2+round(mod(mod(savetime,10000),100)/30)+1) in(';
  str5:=' and abs(val'+inttostr(xh)+')>'+zero ;
  case slct of
   1:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+' in (1,2,-1,-2)';
   2:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+' in (1,2)';
  end;
  //form1.edit1.text :=sqls;
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    result:=fields[0].asinteger*5;
  end;
end;
function get_cos_mon2(ch,xh:longint;mon:smallint):single;
var
  v1,v2:integer;
begin
  v1:=  get_cos_mon(ch,xh,mon,2);
   v2:= get_cos_mon(ch,xh,mon,1);
   if v2<>0 then begin
   result:=v1/v2;
  end  else   begin
    result:=0;
  end;
  // result:=get_cos_mon(ch,xh,mon,1)/(get_cos_mon(ch,xh,mon,2));

end;

function get_allyxsj2(ch,xh:longint;mon,day,slct:smallint):single;
var
  days :integer;
begin
  days:=getdaysofmon(bb_year,mon);
  if (slct>=5)and(slct<=7) then
    days:=get_yxd(ch,xh,mon,day,1)*5;

  if days<>0 then begin
  case slct of
   1:result:=get_allyxsj(ch,xh,mon,day,1)/(strtoint(get_minutes(bb_year,mon,1))/days);
   2:result:=get_allyxsj(ch,xh,mon,day,2)/(strtoint(get_minutes(bb_year,mon,2))/days);
   3:result:=get_allyxsj(ch,xh,mon,day,3)/(strtoint(get_minutes(bb_year,mon,3))/days);
   5:result:=get_allyxsj(ch,xh,mon,day,4)/days;
   6:result:=get_allyxsj(ch,xh,mon,day,8)/days;
   8:result:=get_allyxsj(ch,xh,mon,day,5)/(strtoint(get_minutes(bb_year,mon,1))/days);
   9:result:=get_allyxsj(ch,xh,mon,day,6)/(strtoint(get_minutes(bb_year,mon,2))/days);
   10:result:=get_allyxsj(ch,xh,mon,day,7)/(strtoint(get_minutes(bb_year,mon,3))/days);
   7:result:=1-(get_allyxsj(ch,xh,mon,day,4)+get_allyxsj(ch,xh,mon,day,8))/days;

  end;
  end else result:=-1;
end;
function get_monyxsj(ch,xh:longint;mon,slct:smallint):integer;
var

  {1:高峰超上限，2:低谷超上限, 3:腰荷超上限, 4:日超上限时间}
  {5:日超上限率，6:日超下限率, 7:日合格率, 8:日超下限时间}

  //dttm1:tdate;
  yc_up,yc_dn,str1,str2,str3,str4,str5,sqls:string;
begin
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:='select upperlimit,lowerlimit from prtuana where saveno='+inttostr(ch*200+xh);
    open;
   try
    yc_up:=fields[0].Asstring;yc_dn:=fields[1].AsString;
    except
    yc_up:='0';yc_dn:='0';
    end;
  end;

   str1:=gettimes(mon,1,0);

  str2:=gettimed(mon,1,0);
   str3:=' and val'+inttostr(xh);
  str4:=' and (round(mod(savetime,10000)/100)*2+round(mod(mod(savetime,10000),100)/30)+1) in(';
  str5:=' and abs(val'+inttostr(xh)+')>'+zero ;
  case slct of
   1:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,2]+')'+str3+'>'+yc_up;
   2:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,3]+')'+str3+'>'+yc_up;
   3:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,4]+')'+str3+'>'+yc_up;
   4:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+'>'+yc_up;
   5:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,2]+')'+str3+'<'+yc_dn+str5;
   6:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,3]+')'+str3+'<'+yc_dn+str5;
   7:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str4+dnsdstr[mon,4]+')'+str3+'<'+yc_dn+str5;
   8:sqls:='select count(*) from yc_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+'<'+yc_dn+str5;
  end;
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:=sqls;
    open;
    result:=fields[0].asinteger*5;
  end;
end;
function get_monyxsj2(ch,xh:longint;mon,slct:smallint):single;
var
  days :integer;
begin
  days:=getdaysofmon(bb_year,mon);
  if (slct>=5)and(slct<=7) then
    days:=get_yxd(ch,xh,mon,1,2)*5;
    if days<>0 then begin
  case slct of
   1:result:=get_monyxsj(ch,xh,mon,1)/strtoint(get_minutes(bb_year,mon,1));
   2:result:=get_monyxsj(ch,xh,mon,2)/strtoint(get_minutes(bb_year,mon,2));
   3:result:=get_monyxsj(ch,xh,mon,3)/strtoint(get_minutes(bb_year,mon,3));
   5:result:=get_monyxsj(ch,xh,mon,4)/(days);
   6:result:=get_monyxsj(ch,xh,mon,8)/(days);
   8:result:=get_monyxsj(ch,xh,mon,5)/strtoint(get_minutes(bb_year,mon,1));
   9:result:=get_monyxsj(ch,xh,mon,6)/strtoint(get_minutes(bb_year,mon,2));
   10:result:=get_monyxsj(ch,xh,mon,7)/strtoint(get_minutes(bb_year,mon,3));
   7:result:=1-(get_monyxsj(ch,xh,mon,8)+get_monyxsj(ch,xh,mon,4))/(days);
  end;
  end else result:=-1;
  //showmessage(inttostr(get_monyxsj(ch,xh,mon,8))+'_'+inttostr(get_monyxsj(ch,xh,mon,4))+'_'+inttostr(days));
end;
function get_cos(ch,xh:longint;mon,day,slct:smallint):single;
var
  cun,hcun :integer;
  str1,str2,str3:string;
begin
  cun:=0;
  hcun:=0;
  case slct of
  1:begin
    str1:=gettimes(mon,day,1);
   str2:=gettimed(mon,day,1);
  end;
  2:begin
    str1:=gettimes(mon,1,0);
   str2:=gettimed(mon,1,0);
  end;
  end;
   str3:=' and val'+inttostr(xh);
   with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:='select count(*) from yc_table where  groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+'=50';
    open;

    try
    hcun:=fields[0].Asinteger;
    except
    hcun:=0;
    end;
  end;
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:='select count(*) from yc_table where  groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+' in (50,100)';
   // form1.edit1.Text:='select count(*) from yc_table where  groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<'+str2+str3+'in (50,100)';

    open;

    try
    cun:=fields[0].Asinteger;
    except
    cun:=0;
    end;
  end;
  if ((hcun=0) and (cun=0)) then result:=0
  else
  result:=hcun/cun;
 end;

function get_dayDN(ch,xh:longint;mon,day,lb:smallint):real;
var
  dttm1:tdate;
  str1,str2,str4,sqls:string;
  fn:real;
  g,k:smallint;
begin
  dttm1:=encodedate(bb_year,mon,day);
  if (dttm1>date) then dttm1:=dttm1/0 else begin
  str1:=gettimes(mon,day,1);
  str2:=gettimed(mon,day,1);
  str4:=' and a.savetime>='+str1+' and a.savetime<='+str2+' and b.savetime>='+str1+' and b.savetime<='+str2;
  fn:=0;
  if lb<4 then begin
  k:=1;
  while k<49 do begin
    if (bend[lb,k]>0) and (bend[lb,k+1]>0) then begin
      g:=bend[lb,k]-1;
      with form1.adodataset1 do begin
        if active then close;
        if bend[lb,k+1]=48 then
          sqls:='select b.val'+inttostr(xh)+'-a.val'+inttostr(xh)+' value from dn_table a,dn_table b where a.groupno='+inttostr(ch)+' and b.groupno='+inttostr(ch)+' and (round(mod(a.savetime,10000)/100)*2*30+mod(a.savetime,100))='+inttostr(g)+'*30 and mod(b.savetime,100)=0 and get_date('+inttostr(bb_year)+',a.savetime,1)=b.savetime '+str4
        else
          sqls:='select b.val'+inttostr(xh)+'-a.val'+inttostr(xh)+' value from dn_table a,dn_table b where a.groupno='+inttostr(ch)+' and b.groupno='+inttostr(ch)+' and (round(mod(a.savetime,10000)/100)*2*30+mod(a.savetime,100))='+inttostr(g)+'*30 and (round(mod(b.savetime,10000)/100)*2*30+mod(b.savetime,100))='+inttostr(bend[lb,k+1])+'*30'+str4;
        commandtext:=sqls;

        open;
        fn:=fn+fieldbyname('value').asfloat;
      end;
    end;
    k:=k+2;
  end;
  end else begin
     with form1.adodataset1 do begin
      if active then close;
      commandText:='select max(val'+inttostr(xh)+')-min(val'+inttostr(xh)+') value from dn_table where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<='+str2;

      open;
      fn:=form1.adodataset1.fieldbyname('value').asfloat;
      end;
  end;
    result:=fn;
  end;

end;


function get_monDN2(ch,xh:longint;mon,day,lb:smallint):real;
var
  dttm1:tdate;
  str1,str2,str3,str4,sqls:string;
  fn:real;
   g,k:smallint;
begin

     str1:=gettimes(mon,day,1);

  str2:='12319999' ;
  str4:=' and a.savetime>='+str1+' and a.savetime<='+str2+' and b.savetime>='+str1+' and b.savetime<='+str2;
  str3:=' and floor(a.savetime/10000)=round(b.savetime/10000)';
  fn:=0;
  if lb<4 then begin
  k:=1;
  while k<49 do begin
    if (bend[lb,k]>0) and (bend[lb,k+1]>0) then begin
       g:=bend[lb,k]-1;
      with form1.adodataset1 do begin
        if active then close;
        if bend[lb,k+1]=48 then
          sqls:='select sum(b.val'+inttostr(xh)+'-a.val'+inttostr(xh)+') value from hdn'+inttostr(bb_year mod 10)+' a,hdn'+inttostr(bb_year mod 10)+' b where a.groupno='+inttostr(ch)+' and b.groupno='+inttostr(ch)+' and (round(mod(a.savetime,10000)/100)*2+round(mod(a.savetime,100)/30))='+inttostr(g)+' and mod(b.savetime,100)=0 and get_date('+inttostr(bb_year)+',a.savetime,1)=b.savetime'+str4
        else
          sqls:='select sum(b.val'+inttostr(xh)+'-a.val'+inttostr(xh)+') value from hdn'+inttostr(bb_year mod 10)+' a,hdn'+inttostr(bb_year mod 10)+' b where a.groupno='+inttostr(ch)+' and b.groupno='+inttostr(ch)+' and (round(mod(a.savetime,10000)/100)*2+round(mod(a.savetime,100)/30))='+inttostr(g)+' and (round(mod(b.savetime,10000)/100)*2+round(mod(b.savetime,100)/30))='+inttostr(bend[lb,k+1])+str4+str3;
        commandtext:=sqls;
        open;
        fn:=fn+fieldbyname('value').asfloat;
      end;
    end;
    k:=k+2;
  end;
  end else begin
     with form1.adodataset1 do begin
      if active then close;
      commandText:='select max(val'+inttostr(xh)+')-min(val'+inttostr(xh)+') value from hdn'+inttostr(bb_year mod 10)+' where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<='+str2;
      open;
      fn:=form1.adodataset1.fieldbyname('value').asfloat;
      end;
  end;
    result:=fn;



end;


function get_monDN(ch,xh:longint;mon,day,lb:smallint):real;
var
  dttm1:tdate;
  str1,str2,str3,str4,sqls:string;
  fn:real;
   g,k:smallint;
begin
 if mon=1 then
     str1:='1010000'
  else
     str1:=gettimes(mon-1,day,1);

  str2:=gettimes(mon,day,1) ;
  str4:=' and a.savetime>='+str1+' and a.savetime<='+str2+' and b.savetime>='+str1+' and b.savetime<='+str2;
  str3:=' and floor(a.savetime/10000)=round(b.savetime/10000)';
  fn:=0;
  if lb<4 then begin
  k:=1;
  while k<49 do begin
    if (bend[lb,k]>0) and (bend[lb,k+1]>0) then begin
       g:=bend[lb,k]-1;
      with form1.adodataset1 do begin
        if active then close;
        if bend[lb,k+1]=48 then
          sqls:='select sum(b.val'+inttostr(xh)+'-a.val'+inttostr(xh)+') value from his_dn a,his_dn b where a.groupno='+inttostr(ch)+' and b.groupno='+inttostr(ch)+' and (round(mod(a.savetime,10000)/100)*2*30+mod(a.savetime,100))='+inttostr(g)+'*30 and mod(b.savetime,100)=0 and get_date('+inttostr(bb_year)+',a.savetime,1)=b.savetime'+str4
        else
          sqls:='select sum(b.val'+inttostr(xh)+'-a.val'+inttostr(xh)+') value from his_dn a,his_dn b where a.groupno='+inttostr(ch)+' and b.groupno='+inttostr(ch)+' and (round(mod(a.savetime,10000)/100)*2*30+mod(a.savetime,100))='+inttostr(g)+'*30 and (round(mod(b.savetime,10000)/100)*2*30+mod(b.savetime,100))='+inttostr(bend[lb,k+1])+'*30 '+str4+str3;
        commandtext:=sqls;
        open;
        fn:=fn+fieldbyname('value').asfloat;
      end;
    end;
    k:=k+2;
  end;
  end else begin
     with form1.adodataset1 do begin
      if active then close;
      commandText:='select max(val'+inttostr(xh)+')-min(val'+inttostr(xh)+') value from his_dn where groupno='+inttostr(ch)+' and savetime>='+str1+' and savetime<='+str2;
      open;
      fn:=form1.adodataset1.fieldbyname('value').asfloat;
      end;
  end;
  if mon=1 then fn:=fn+get_monDN2(ch,xh,12,day,lb);
    result:=fn;



end;




function get_DN_Value(ch,xh:longint;mon,day,ho,mi:smallint):real;
var
  dttm1:tdatetime;
begin
  dttm1:=encodedate(bb_year,mon,day)+encodetime(ho,mi,0,0);
  if (dttm1>now-5/60/24) then dttm1:=dttm1/0 else begin
  with form1.adodataset1 do begin
    if active then close;
    //sql.Clear;
    commandtext:='select val'+inttostr(xh)+' from dn_table where groupno='+inttostr(ch)+' and savetime='+datetonum(mon,day,ho,mi);;
    open;
    result:=fields[0].AsFloat;
  end;
  end;
end;
function get_yc_mon_cz(ch,xh:longint;mon,day:smallint):single;
var {月差值}
  ycv1,ycv2:single;
  y2,m2,d2:word;

begin
  ycv1:=get_yc_value(ch,xh,mon,1,0,0);
  decodedate(encodedate(bb_year,mon,day)+getdaysofmon(bb_year,mon),y2,m2,d2);
  ycv2:=get_yc_value(ch,xh,m2,1,0,0);
  result:=ycv2-ycv1;
end;

function letter2int(s:string):integer;
var
   ss:string;
   vv,len,i:integer;
begin
  vv:=0;
  ss:='ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  len:=length(s);
  if len>1 then begin
  for i:=1 to (len) do begin
    vv:=vv*26+pos(s[i],ss);
  end;
  end else begin
    vv:=pos(s,ss);
  end;
   result:=vv;
end;
function SplitString(pString:Pchar;psubString:PChar):TStringList;
var
   nSize,SubStringSize:DWord;
   intI,intJ,intK:DWORD;
   ts:TStringList;
   curChar:Char;
   strString:string;
   strsearchSubStr:string;
begin
   nSize:=strLen(pString);
   SubStringSize:=strLen(PSubString);
   ts:=TStringList.Create;
   strstring:='';
   inti:=0;
   while intI<=(nSize-1) do
   begin
      if (nsize-inti)>= substringSize then
      begin
          if ((PString+intI)^=pSubString^) then
          begin
             intk:=inti;
            strSearchSubStr:='';
            curchar:=(pstring+intk)^;
            strsearchSubStr:=strSearchSubStr+Curchar;
            intk:=intk+1;
            for intj:= 1 to SubStringSize-1  do
            begin
               if ((pString+intk)^=(PSubString+intj)^) then
               begin
                  curchar:=(pstring+intk)^;
                  intk:=intk+1;
                  strsearchSubStr:=strSearchSubStr+Curchar;
            end
            else begin
              inti:=intk;
              strString:=strString+strSearchSubStr;
              break; //不匹配 退出FOR
            end;
        end;
         if (intJ=substringSize) or (SubStringSize=1) then
         begin
            inti:=intk;
            ts.add(strstring);
            strstring:='';
         end;
      end
      else begin
         curChar:=(pString+inti)^;
         strstring:=strstring+curchar;
         inti:=inti+1;
      end;
      if inti=nsize then
      begin
         ts.Add(strString);
         strString:='';
       end;
     end
     else begin //将剩下的字符给作为一个字符串复制给字符串集合
        strString:=strstring+string(pString+inti);
        ts.Add(strstring);
        inti:=nsize;
     end;
   end;
   Result:=ts;

end;

{
function SplitString(const Source,ch:String):TStringList;
var
Temp:String;
I:Integer;
chLength:Integer;
begin
Result:=TStringList.Create;
//如果是空自符串则返回空列表
if Source='' then Exit;
Temp:=Source;
I:=Pos(ch,Source);
chLength := Length(ch);
while I<>0 do
begin
Result.Add(Copy(Temp,0,I-chLength+1));
Delete(Temp,1,I-1 + chLength);
I:=pos(ch,Temp);
end;
Result.add(Temp);
end;
  }


function do_yinyong(s:string):string;
var
   ss,s1,s2,s3:string;

   ir,jc:integer;
   ss1:TStringList;
begin
   ss:='';              //pos(',',s)业务发展部能耗月报表,P,39

   try
   s:=copy(s,2,length(s)-1);
    ss1:=split(s,',');
   form1.memo1.lines.AddStrings(ss1);
   s1:=ss1.Strings[0]+'_'+ss1.Strings[3];
    s2:=ss1.Strings[1];
     s3:=ss1.Strings[2];
  {
   s:= copy(s,pos(',',s)+1,length(s)-pos(',',s));

    s2:=copy(s,1,pos(',',s)-1);

   s:= copy(s,pos(',',s)+1,length(s)-pos(',',s));

    s3:=s;
    }
     form1.memo1.lines.Add(s1+'---'+s2+'---'+s3);

    except
     form1.memo1.lines.Add('fk');
    end;

   try
   jc:=letter2int(s2)-1;
   ir:=strtoint(s3)-1;
   except
   jc:=2;
   ir:=2;
   end;
   
   try
       //form1.memo1.lines.Add(rptdir+inttostr(form1.combobox1.ItemIndex)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+s1+'.xlsx');
       if pos('日报表',s1)>0 then
       form1.XLS2.Filename := rptdir+inttostr(1)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+inttostr(bb_day)+'\'+s1+'.xlsx'
       else
       form1.XLS2.Filename := rptdir+inttostr(0)+'\'+inttostr(bb_year)+'\'+inttostr(bb_mon)+'\'+s1+'.xlsx';

       form1.xls2.Read;
       form1.memo1.lines.Add('get_yinyong:'+s1+','+inttostr(jc)+','+inttostr(ir)+','+form1.xls2.Sheets[0].asstring[jc,ir]);
       ss:= form1.xls2.Sheets[0].asstring[jc,ir];
    except
     form1.memo1.lines.Add('请确认已下载Excel文件:'+s1+'---'+form1.XLS2.Filename);
     exit;
    end;

   ss1.Free;
   result:=ss;
end;
function do_chuli(s:shortstring):shortstring;
var
  ss:shortstring;
  clh,year,mon,day,ho,mi,sc,msc:word;
  ch,xh:longint;
  col,row :smallint;
  dttm:tdatetime;
  fn:single;
  procedure get_colrow(s:shortstring; var col,row:smallint);
  begin

    col:=strtoint(copy(s,16,3));
    row:=strtoint(copy(s,20,3));

  end;

  function check_date(yy,mm,dd:word):boolean;
  var
    dt:tdatetime;
  begin
    result:=true;
    try
      dt:=encodedate(yy,mm,dd);
    except
      //on error do
        result:=false;
    end;
    if dt>now then
       result:=false;
  end;
  function str2date(ymdhm:string):tdatetime;
var
  yy,mm,dd,hh,mi:word;i:byte;
  begin
    ymdhm:=trim(ymdhm);
    if length(ymdhm)>16  then  begin
      i:=pos('-',ymdhm);
      yy:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      i:=pos('-',ymdhm);
      mm:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      i:=pos(' ',ymdhm);
      dd:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      i:=pos(':',ymdhm);
      if pos('下午',ymdhm)<>0 then
        hh:=strtoint(copy(ymdhm,6,2))+12
      else
        hh:=strtoint(copy(ymdhm,6,2));
      delete(ymdhm,1,i);
      i:=pos(':',ymdhm);
      mi:=strtoint(copy(ymdhm,1,2));
    end else if length(ymdhm)=11  then  begin
      i:=pos('-',ymdhm);
      mm:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      i:=pos(' ',ymdhm);
      dd:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      i:=pos(':',ymdhm);
      hh:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      mi:=strtoint(copy(ymdhm,1,length(s)));
      yy:=bb_year;
    end else if length(ymdhm)=8  then  begin

      i:=pos(' ',ymdhm);
      dd:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      i:=pos(':',ymdhm);
      hh:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      mi:=strtoint(copy(ymdhm,1,length(s)));
      yy:=bb_year;
      mm:=bb_mon;
    end else if length(ymdhm)=16  then begin
      i:=pos('-',ymdhm);
      yy:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      i:=pos('-',ymdhm);
      mm:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      i:=pos(' ',ymdhm);
      dd:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      i:=pos(':',ymdhm);
      hh:=strtoint(copy(ymdhm,1,i-1));
      delete(ymdhm,1,i);
      mi:=strtoint(copy(ymdhm,1,length(s)));
    end;
    result:=encodedate(yy,mm,dd)+encodetime(hh,mi,0,0);
  end;
  function date2str(dt:tdatetime):string;
  var
    year,mon,day, ho,mi,sc,msc:word;
  begin
      decodedate(dt,year,mon,day);
      decodetime(dt,ho,mi,sc,msc);
      result:=inttostr(year)+'-'+inttostr(mon)+'-'+inttostr(day)+' '+inttostr(ho)+':'+inttostr(mi);
  end;
begin                  //begin of do_chuli
  if pos('#',s)<>1 then  begin
     result:=s;
  exit;
  end;
  ss:='';
  if pos('#0001',s)=1 then begin
    ss:=date2str(round(date));
    mon:=pos('0:0',ss);
    result:=midstr(ss,1,mon-2);
    //showmessage(midstr(result,1,10));
  exit; end;
  if pos('#0000',s)=1 then begin
    ss:=date2str(round(form1.DateTimePicker1.date));
    mon:=pos('0:0',ss);
    result:=midstr(ss,1,mon-2);
    //showmessage(midstr(result,1,10));
  exit; end;
  //if pos('#0001',s)=1 then begin  result:=date2str(round(date));exit; end;
  if pos('#0006',s)=1 then begin  decodedate(form1.DateTimePicker1.date,year,mon,day);result:=inttostr(year)+'-'+inttostr(mon);exit; end;
  if pos('#0007',s)=1 then begin  decodedate(form1.DateTimePicker1.date,year,mon,day);result:=inttostr(GetDaysOfMon(year,mon));exit; end;
  clh:=strtoint(copy(s,2,4));

  if ((clh>=9000) and (clh<=9999)) then begin
    get_colrow(s,col,row);
    //showmessage(form1.ExcelWorksheet1.Cells.Item[row,col]);
    try
    form1.XLS.Sheets[0].Asstring[col,row]:=do_chuli(form1.XLS.Sheets[0].Asstring[col,row]);
   except

   end;
   ss:=form1.XLS.Sheets[0].Asstring[col,row];

    try
      dttm:=str2date(ss);
      decodedate(dttm,year,mon,day);
      decodetime(dttm,ho,mi,sc,msc);
    except
      dttm:=0;
      year:=0;
      result:=format('ERR! %d,%d',[col,row]);
      exit;
    end;
  end;

    if clh<>8  then begin
       if clh<>9  then begin
       if clh<>10  then begin
       if not((clh>=9000) and (clh<=9999)) then begin
    try day:=get_day(s);except end;
    try mon:=get_mon(s);except end;

    try ho:=get_hour(s);except end;
    try mi:=get_minute(s);except end;
    end;
    end;
    end;
  end;
  if day=0 then day:=bb_day;
  if mon=0 then
    mon:=bb_mon;
  if mon=13 then begin
    if bb_mon=1 then begin
      mon:=12;
      bb_year:=bb_year-1;
    end else
      mon:=bb_mon-1;

  end;
  if not check_date(bb_year,bb_mon,day) then begin
    ss:='';
    result:='';
    exit;
  end;
  case clh of
    0008:begin //月高峰总时间
      result:=get_minutes(year,mon,1);exit;
    end;
    0009:begin //月腰荷总时间
      result:=get_minutes(year,mon,2);exit;
    end;
    0010:begin //月低谷总时间
      result:=get_minutes(year,mon,3);exit;
    end;
  end;
  ch:=get_ch(s);
  xh:=get_xh(s);



case clh of
  2..5,11..99,200..499,9001,1001..1006:ch2gno(ch,xh,2);  //0 yctime,
   100..199,500..799:ch2gno(ch,xh,3);      //1 event_tables 2,4//yx
   800..899,9002:ch2gno(ch,xh,4); //dn
end;
  case clh of
    9001:if year>0 then begin
      try ss:=formatfloat('0.000',get_yc_Value(ch,xh,mon,day,ho,mi)) except ss:='ERROR9001';end;

    end;
    9002:if year>0 then begin
      try ss:=formatfloat('0.000',get_dn_Value(ch,xh,mon,day,ho,mi)) except ss:='ERROR9001';end;
    end;
   {0001:begin
      ss:=date2str(round(date));
    end; }
    0002:begin //遥测日最大时间
      ss:=numtodate(get_yc_DayMax2(ch,xh,mon,day)); {日最大时间}
    end;
    0003:begin //遥测日最小时间
      ss:=numtodate(get_yc_DayMin2(ch,xh,mon,day)); {日最小时间}
    end;
    0004:begin //遥测月最大时间
      ss:=numtodate(get_yc_MonTJ(ch,xh,mon,5));
    end;
    0005:begin //遥测月最小时间
      ss:=numtodate(get_yc_MonTJ(ch,xh,mon,6));
    end;
    0011:begin //日高峰最大时间
       ss:=numtodate(get_yc_daytj(ch,xh,mon,day,7));
    end;
    0012:begin //日腰荷最大时间
      ss:=numtodate(get_yc_daytj(ch,xh,mon,day,8));
    end;
    0013:begin //日低谷最大时间
      ss:=numtodate(get_yc_daytj(ch,xh,mon,day,10));
    end;
    0014:begin //月高峰最大时间
      ss:=numtodate(get_yc_MonTJ(ch,xh,mon,16));
    end;
    0015:begin //月腰荷最大时间
      ss:=numtodate(get_yc_MonTJ(ch,xh,mon,17));
    end;
    0016:begin //月低谷最大时间
      ss:=numtodate(get_yc_MonTJ(ch,xh,mon,18));
    end;
    0017:begin //日高峰最小时间
      ss:=numtodate(get_yc_daytj(ch,xh,mon,day,11));
    end;
    0018:begin //日腰荷最小时间
      ss:=numtodate(get_yc_daytj(ch,xh,mon,day,13));
    end;
    0019:begin //日低谷最小时间
      ss:=numtodate(get_yc_daytj(ch,xh,mon,day,14));
    end;
    0020:begin //月高峰最小时间
      ss:=numtodate(get_yc_MonTJ(ch,xh,mon,19));
    end;
    0021:begin //月腰荷最小时间
      ss:=numtodate(get_yc_MonTJ(ch,xh,mon,20));
    end;
    0022:begin //月低谷最小时间
      ss:=numtodate(get_yc_MonTJ(ch,xh,mon,21));
    end;

    
    0101:begin
      ss:=inttostr(get_kgfhcs(ch,xh,mon,day,1));
    end;
    0102:begin
      ss:=inttostr(get_kgfhcs(ch,xh,mon,day,2));
    end;
    0103:begin
      ss:=inttostr(get_kgfhcs(ch,xh,mon,day,3));
    end;
    0104:begin
      ss:=inttostr(get_kgfhcs(ch,xh,mon,day,4));
    end;
    0105:begin
      ss:=inttostr(get_kgfhcs_mon(ch,xh,mon,1));
    end;
    0106:begin
      ss:=inttostr(get_kgfhcs_mon(ch,xh,mon,3));
    end;
    0107:begin
      ss:=inttostr(get_kgfhcs_mon(ch,xh,mon,4));
    end;
    0108:begin
      ss:=inttostr(get_kgfhcs_mon(ch,xh,mon,2));
    end;
    0201:begin {1:高峰超上时间}
      ss:=inttostr(get_allyxsj(ch,xh,mon,day,1));
    end;  {1:高峰超上时间，2:低谷超上时间, 3:高峰超上时间,
          {4:日超上时间，  5:日超上限率，6:日超下限率, 7:日合格率
          {8:日超下限时间 }
    0202:begin {4:日超上时间}
      ss:=inttostr(get_allyxsj(ch,xh,mon,day,4));
    end;
    0203:begin {8:日超下限时间}
      ss:=inttostr(get_allyxsj(ch,xh,mon,day,8));
    end;
    0204:begin // 1:高峰月超上时间
      ss:=inttostr(get_monyxsj(ch,xh,mon,1));
    end;
    0205:begin // 4:月超上时间
      ss:=inttostr(get_monyxsj(ch,xh,mon,4));
    end;
    0206:begin // 8:月超下时间
      ss:=inttostr(get_monyxsj(ch,xh,mon,8));
    end;
    0207:begin
      ss:=inttostr(get_allyxsj(ch,xh,mon,day,3));
    end;
    0208:begin
      ss:=inttostr(get_allyxsj(ch,xh,mon,day,2));
    end;
    0209:begin //月腰荷超上时间
      ss:=inttostr(get_monyxsj(ch,xh,mon,3));
    end;
    0210:begin //月低谷超上时间
      ss:=inttostr(get_monyxsj(ch,xh,mon,2));
    end;
    0211:begin {1:高峰超下时间}
      ss:=inttostr(get_allyxsj(ch,xh,mon,day,5));
    end;
    0212:begin {4:日腰荷超下时间}
      ss:=inttostr(get_allyxsj(ch,xh,mon,day,6));
    end;
    0213:begin {8:日低谷超下限时间}
      ss:=inttostr(get_allyxsj(ch,xh,mon,day,7));
    end;
    0214:begin // 1:高峰月超下时间
      ss:=inttostr(get_monyxsj(ch,xh,mon,5));
    end;
    0215:begin // 4:月腰荷超下时间
      ss:=inttostr(get_monyxsj(ch,xh,mon,6));
    end;
    0216:begin // 8:月低谷超下时间
      ss:=inttostr(get_monyxsj(ch,xh,mon,7));
    end;
    0301:begin
      ss:=formatfloat('0.000',get_yc_Value(ch,xh,mon,day,ho,mi));
    end;
    0302:begin
      ss:=formatfloat('0.000',get_yc_HourMax(ch,xh,mon,day,ho));
    end;
    0303:begin
      ss:=formatfloat('0.000',get_yc_DayMax(ch,xh,mon,day));
    end;
    0306:begin
      ss:=formatfloat('0.000',get_yc_mon_cz(ch,xh,mon,day)); {月差值}
    end;
    0304:begin
      ss:=formatfloat('0.000',get_yc_DayAvg(ch,xh,mon,day)); {日平均}
    end;
    0305:begin
      ss:=formatfloat('0.000',get_yc_DayMin(ch,xh,mon,day)); {日最小v}
    end;
    0307:begin //月最大
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,1));
//1:月最大  2:月最小  3:月平均  4:月负荷率
//5:月最大时间  6:月最小时间
    end;
    0308:begin //月最小
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,2));
    end;
    0309:begin //月平均
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,3));
    end;
    0310:begin //月平均负荷率
      ss:=formatfloat('0.000',get_yc_MonFhl(ch,xh,mon));
    end;
    0311:begin //月高峰最大
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,7));
    end;
    0312:begin //月高峰最小
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,8));
    end;
    0313:begin //月高峰平均
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,9));
    end;
    0314:begin //月腰荷最大
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,10));
    end;
    0315:begin //月腰荷最小
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,11));
    end;
    0316:begin //月腰荷平均
      ss:=floattostr(get_yc_MonTJ2(ch,xh,mon,12));
    end;
    0317:begin //月低谷最大
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,13));
    end;
    0318:begin //月低谷最小
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,14));
    end;
    0319:begin //月低谷平均
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,15));
    end;
    0320:begin
      ss:=formatfloat('0.000',get_yc_MonTJ2(ch,xh,mon,23));
    end;
    0321:begin
      ss:=formatfloat('0.000',get_yc_dayfhl(ch,xh,mon,day));
    end;

     401:begin {日超上限率}
      ss:=formatfloat('0.000',get_allyxsj2(ch,xh,mon,day,5));
    end;
    402:begin {日超下限率}
      ss:=formatfloat('0.000',get_allyxsj2(ch,xh,mon,day,6));
    end;
    403:begin {日合格率}
      fn:=get_allyxsj2(ch,xh,mon,day,7);
      if fn<>-1 then begin
      ss:=formatfloat('00.00%',fn*100);
      end  else ss:='';
    end;
    407:begin {日超上限率}
      ss:=formatfloat('0.000',get_allyxsj2(ch,xh,mon,day,1));
    end;
    408:begin {日超下限率}
      ss:=formatfloat('0.000',get_allyxsj2(ch,xh,mon,day,8));
    end;
    409:begin
      ss:=formatfloat('0.000',get_allyxsj2(ch,xh,mon,day,2));
    end;
    410:begin {日超上限率}
      ss:=formatfloat('0.000',get_allyxsj2(ch,xh,mon,day,9));
    end;
    411:begin {日超下限率}
      ss:=formatfloat('0.000',get_allyxsj2(ch,xh,mon,day,3));
    end;
    412:begin {日合格率}
      ss:=formatfloat('0.000',get_allyxsj2(ch,xh,mon,day,10));
    end;

    404:begin //月超上限率
      ss:=formatfloat('0.000',get_monyxsj2(ch,xh,mon,5));
    end;
    405:begin //月超下限率
      ss:=formatfloat('0.000',get_monyxsj2(ch,xh,mon,6));
    end;
    406:begin //月合格率
      fn:=get_monyxsj2(ch,xh,mon,7);
      if fn<>-1 then begin
      ss:=formatfloat('00.00%',fn*100);
      end  else ss:='';
    end;
    413:begin //月超下限率
      ss:=formatfloat('0.000',get_monyxsj2(ch,xh,mon,1));
    end;
    414:begin //月合格率
      ss:=formatfloat('0.000',get_monyxsj2(ch,xh,mon,8));
    end;
     415:begin //月超上限率
      ss:=formatfloat('0.000',get_monyxsj2(ch,xh,mon,2));
    end;
    416:begin //月超下限率
      ss:=formatfloat('0.000',get_monyxsj2(ch,xh,mon,9));
    end;
    417:begin //月合格率
      ss:=formatfloat('0.000',get_monyxsj2(ch,xh,mon,3));
    end;
    418:begin //月超上限率
      ss:=formatfloat('0.000',get_monyxsj2(ch,xh,mon,10));
    end;
     419:begin //glys 日合格率
      ss:=formatfloat('00%',get_cos(ch,xh,mon,day,1)*100);
      //ss:='$$$';

    end;
     420:begin //glys 月合格率
      ss:=formatfloat('00%',get_cos(ch,xh,mon,day,2)*100);
    end;
   501:begin
      ss:=inttostr(get_eventcs(ch,xh,mon,day,1));
    end;
    502:begin
      ss:=inttostr(get_eventcs(ch,xh,mon,day,2));
    end;

    506:begin
      ss:=inttostr(get_eventcs_mon(ch,xh,mon,3));
    end;
    507:begin
      ss:=inttostr(get_eventcs_mon(ch,xh,mon,2));
    end;
     0601:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,5));
    end;
    0602:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,1));
    end;
    0603:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,5)-get_ykcs(ch,xh,mon,day,1));
    end;
    0604:begin
     ss:=floattostr(round(get_yk_moncgl(ch,xh,mon,day,1)*10000+0.5)/100)+'%';
    end;
    0605:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,6));
    end;
    0606:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,2));
    end;
    0607:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,6)-get_ykcs(ch,xh,mon,day,2));
    end;
    0608:begin
    ss:=floattostr(round(get_yk_moncgl(ch,xh,mon,day,2)*10000+0.5)/100)+'%';
    end;

    0609:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,7));
    end;
    0610:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,3));
    end;
    0611:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,7)-get_ykcs(ch,xh,mon,day,3));
    end;
    0612:begin
    ss:=floattostr(round(get_yk_moncgl(ch,xh,mon,day,3)*10000+0.5)/100)+'%';
    end;
     0613:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,8));
    end;
    0614:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,4));
    end;
    0615:begin
      ss:=inttostr(get_ykcs(ch,xh,mon,day,8)-get_ykcs(ch,xh,mon,day,4));
    end;
    0616:begin
     ss:=floattostr(round(get_yk_moncgl(ch,xh,mon,day,4)*10000+0.5)/100)+'%';
    end;

    0661:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,5));
    end;
    0662:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,1));
    end;
    0663:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,5)-get_ykcs_mon(ch,xh,mon,1));
    end;
    0664:begin
     ss:=floattostr(round(get_yk_moncgl(ch,xh,mon,day,5)*10000+0.5)/100)+'%';
    end;
    0665:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,6));
    end;
    0666:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,2));
    end;
    0667:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,6)-get_ykcs_mon(ch,xh,mon,2));
    end;
    0668:begin
     ss:=floattostr(round(get_yk_moncgl(ch,xh,mon,day,6)*10000+0.5)/100)+'%';
    end;

    0669:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,7));
    end;
    0670:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,3));
    end;
    0671:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,7)-get_ykcs_mon(ch,xh,mon,3));
    end;
    0672:begin
     ss:=floattostr(round(get_yk_moncgl(ch,xh,mon,day,7)*10000+0.5)/100)+'%';
    end;
     0673:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,8));
    end;
    0674:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,4));
    end;
    0675:begin
      ss:=inttostr(get_ykcs_mon(ch,xh,mon,8)-get_ykcs_mon(ch,xh,mon,4));
    end;
    0676:begin
    ss:=floattostr(round(get_yk_moncgl(ch,xh,mon,day,7)*10000+0.5)/100)+'%';
    end;
    0701:begin
      ss:=inttostr(get_eventcs(ch,xh,mon,day,3));
    end;
    0702:begin
      ss:=inttostr(get_eventcs_mon(ch,xh,mon,1));
    end;
    0801:begin
      ss:=formatfloat('0.000',get_dayDN(ch,xh,mon,day,4));
    end;
    0802:begin
      ss:=formatfloat('0.000',get_dayDN(ch,xh,mon,day,1));
    end;
    0803:begin
      ss:=formatfloat('0.000',get_dayDN(ch,xh,mon,day,2));
    end;
    0804:begin
      ss:=formatfloat('0.000',get_dayDN(ch,xh,mon,day,3));
    end;
    0805:begin
      ss:=formatfloat('0.000',get_monDN(ch,xh,mon,day,4));
    end;
    0806:begin
      ss:=formatfloat('0.000',get_monDN(ch,xh,mon,day,1));
    end;
    0807:begin
      ss:=formatfloat('0.000',get_monDN(ch,xh,mon,day,2));
    end;
    0808:begin
      ss:=formatfloat('0.000',get_monDN(ch,xh,mon,day,3));
    end;
    0809:begin
      ss:=formatfloat('0.000',get_dn_Value(ch,xh,mon,day,ho,mi));
    end;
     1001:begin
      ss:=formatfloat('0',get_cos_day(ch,xh,mon,day,1));
    end;
     1002:begin
      ss:=formatfloat('0',get_cos_day(ch,xh,mon,day,2));
    end;
     1003:begin
      ss:=formatfloat('0.000',get_cos_day2(ch,xh,mon,day));
    end;
     1004:begin
      ss:=formatfloat('0',get_cos_mon(ch,xh,mon,1));
    end;
     1005:begin
      ss:=formatfloat('0',get_cos_mon(ch,xh,mon,2));
    end;
     1006:begin
      ss:=formatfloat('0.000',get_cos_mon2(ch,xh,mon));
    end;
  end;
  result:=ss;
end;

procedure cre_view(irow,jcol:smallint);
 var
   i,j,clh:smallint;
  s:shortstring;

  r:real;
  vm,vd,vem,ved,vdm,vdd:string;
  range1:ExcelRange ;
  str1,str2,str3,sqls:string;
  ymd1,ymd2:integer;
 begin                        ///////////////////////////////////
  for i:=1 to jcol do begin
    for j:=1 to irow do begin
      try
      s:=form1.XLS.Sheets[0].Asstring[i,j];
      except
      continue;
      end;
      if (s[1]='#')and(pos('#9001',s)<>1)and(pos('#9002',s)<>1)and(length(s)>5) then begin
              clh:=strtoint(copy(s,2,4));
               case clh of
                 4,5,14,15,16,20..22,204..206,209,210,
                301,306..313,320,404,405,406,420,1004..1006: begin
                  if pos(copy(s,25,2),vm)=0 then
                    if length(vm)=0 then
                      vm:=copy(s,25,2)
                    else
                      vm:=vm+','+copy(s,25,2);
                end;
                105..108,506,507,661..676,702:begin
                  if pos(copy(s,25,2),vem)=0 then
                    if length(vem)=0 then
                      vem:=copy(s,25,2)
                    else
                      vem:=vem+','+copy(s,25,2);
                end;
                101..104,501,502,601..616,701: begin
                  if pos(copy(s,16,2),ved)=0 then
                   if length(ved)=0 then
                      ved:=copy(s,16,2)
                    else
                      ved:=ved+','+copy(s,16,2);
                end;
                801..804,809:begin
                  if pos(copy(s,16,2),vdd)=0 then
                    if length(vdd)=0 then
                      vdd:=copy(s,16,2)
                    else
                      vdd:=vdd+','+copy(s,16,2);
                end;
                805..808:begin
                  if pos(copy(s,25,2),vdm)=0 then
                    if length(vdm)=0 then
                      vdm:=copy(s,25,2)
                    else
                      vdm:=vdm+','+copy(s,25,2);
                end;
                else begin
                  if pos(copy(s,16,2),vd)=0 then
                    if length(vd)=0 then
                      vd:=copy(s,16,2)
                    else
                      vd:=vd+','+copy(s,16,2);
                end;
                end;
                if (clh>800) and (clh<900) then
                bendin(dnsd);
              end;
            end;

     end;
  //showmessage(vm+'_'+vd+'_'+vem+'_'+ved);
  if vm<>'' then begin

     with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add('truncate table yc_table');
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

          end; //with   **  do begin
    sqls:='insert into yc_table  select * from '+ycname+' where ';
    if length(vm)=2 then  begin
      if vm='00' then vm:=inttostr(bb_mon);
      str1:=vm+'010000';
      str2:=inttostr(getdaysofmon(bb_year,strtoint(vm))*10000)+'+'+str1;
      str3:='savetime>='+str1+' and savetime<'+str2;
      sqls:=sqls+str3;
    end else begin
      i:=pos(',',vm);
      j:=1;
      while(i<>0) do begin
        str3:=copy(vm,1,2);
        if str3='00' then str3:=inttostr(bb_mon);
       str1:=str3+'010000';
      str2:=inttostr(getdaysofmon(bb_year,strtoint(str3))*10000)+'+'+str1;
        str3:='(savetime>='+str1+' and savetime<'+str2+')';
        if j=1 then
          sqls:=sqls+str3
        else
          sqls:=sqls+' or '+str3;
        i:=pos(',',vm);
        vm:=copy(vm,i+1,length(vm)-i);
        j:=j+1;
      end;
    end;
    //form1.edit1.Text:=sqls;
    with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add(sqls);
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

          end; //with   **  do begin
  end;

 if vdm<>'' then begin
     with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add('truncate table dn_table');
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

          end; //with   **  do begin
    sqls:='insert into  dn_table select * from '+dnname+' where ';
    if length(vdm)=2 then  begin
      if vdm='00' then vdm:=inttostr(bb_mon);
      if vdm='13' then  begin
        if bb_mon=1 then begin
          vdm:=inttostr(12);
          bb_year:=bb_year-1;
        end else
          vdm:=inttostr(bb_mon-1);
      end;
       str1:=gettimes(strtoint(vdm),bb_day,0);
        str2:=gettimed(strtoint(vdm),bb_day,0);
      //str1:='to_date('''+inttostr(bb_year)+'-'+vdm+'-1'',''yyyy-mm-dd'')';
      //str2:=inttostr(getdaysofmon(bb_year,strtoint(vdm)))+'+'+str1;
      str3:='savetime>='+str1+' and savetime<='+str2;
      sqls:=sqls+str3;
    end else begin
      i:=pos(',',vdm);
      j:=1;
      while(i<>0) do begin
        str3:=copy(vdm,1,2);
        if str3='00' then str3:=inttostr(bb_mon);
        if vdm='13' then  begin
          if bb_mon=1 then begin
            vdm:=inttostr(12);
            bb_year:=bb_year-1;
          end else
            vdm:=inttostr(bb_mon-1);
        end;
         str1:=vm+'010000';
      str2:=inttostr(getdaysofmon(bb_year,strtoint(vm))*10000)+'+'+str1;
      str3:='savetime>='+str1+' and savetime<'+str2;
         if j=1 then
          sqls:=sqls+str3
        else
          sqls:=sqls+' or '+str3;
        i:=pos(',',vdm);
        vdm:=copy(vdm,i+1,length(vdm)-i);
        j:=j+1;
      end;
    end;
    //edit5.Text:=sqls;
   with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add(sqls);
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

          end; //with   **  do begin
  end;


  if (vdm='')and(vdd<>'') then begin
  with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add('truncate table dn_table');
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

  end; //with   **  do begin
  sqls:='insert into  dn_table select * from '+dnname+' where ';
    if length(vdd)=2 then  begin
      if vdd='00' then begin
          vdd:=inttostr(bb_day);
          if bb_day<10 then
             vdd:='0'+inttostr(bb_day);
        end;

       str1:=inttostr(bb_mon)+vdd+'0000';
        
      str2:=inttostr(bb_mon)+inttostr(strtoint(vdd)+1)+'0000';
        if   strtoint(vdd)<9 then  str2:=inttostr(bb_mon)+'0'+inttostr(strtoint(vdd)+1)+'0000';
      str3:='savetime>='+str1+' and savetime<'+str2;






      sqls:=sqls+str3;
    end else begin
      i:=pos(',',vdd);
      j:=1;
      while(i<>0) do begin
        str3:=copy(vdd,1,2);
       if str3='00' then begin
          str3:=inttostr(bb_day);
          if bb_day<10 then
             str3:='0'+inttostr(bb_day);
        end;
        str1:=inttostr(bb_mon)+str3+'0000';
      str2:=inttostr(bb_mon)+inttostr(strtoint(str3)+1)+'0000';
       if   strtoint(str3)<9 then  str2:=inttostr(bb_mon)+'0'+inttostr(strtoint(str3)+1)+'0000';
        str3:='(savetime>='+str1+' and savetime<'+str2+')';
        if j=1 then
          sqls:=sqls+str3
        else
          sqls:=sqls+' or '+str3;
        i:=pos(',',vdd);
        vdd:=copy(vdd,i+1,length(vdd)-i);
        j:=j+1;
      end;
    end;
    with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add(sqls);
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

          end; //with   **  do begin
  end;

  if vem<>'' then begin
   with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add('truncate table event_table');
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

          end; //with   **  do begin

    sqls:='insert into  event_table  select * from '+evename+' where ';
    if length(vem)=2 then  begin
      if vem='00' then
        vem:=inttostr(bb_mon);
      ymd1:=bb_year*10000+strtoint(vem)*100+1;
      ymd2:=getdaysofmon(bb_year,strtoint(vem))+ymd1;
      {str1:='to_date('''+inttostr(bb_year)+'-'+vem+'-1'',''yyyy-mm-dd'')';
      str2:=inttostr(getdaysofmon(bb_year,strtoint(vem)))+'+'+str1; }
      str3:='ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2);
      sqls:=sqls+str3;
    end else begin
      i:=pos(',',vem);
      j:=1;
      while(i<>0) do begin
        str3:=copy(vem,1,2);
        if str3='00' then str3:=inttostr(bb_mon);
        ymd1:=bb_year*10000+strtoint(str3)*100+1;
        ymd2:=getdaysofmon(bb_year,strtoint(str3))+ymd1;
        str3:='(ymd>='+inttostr(ymd1)+' and ymd<'+inttostr(ymd2)+')';
        if j=1 then
          sqls:=sqls+str3
        else
          sqls:=sqls+' or '+str3;
        i:=pos(',',vem);
        vem:=copy(vem,i+1,length(vem)-i);
        j:=j+1;
      end;
    end;
    //edit5.Text:=sqls;
    with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add(sqls);
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

    end; //with   **  do begin
  end;

  if (vm='')and(vd<>'') then begin
  with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add('truncate table yc_table');
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

          end; //with   **  do begin
  sqls:='insert into yc_table  select * from '+ycname+' where ';
    if length(vd)=2 then  begin
     if vd='00' then begin
          vd:=inttostr(bb_day);
          if bb_day<10 then
             vd:='0'+inttostr(bb_day);
        end;
      str1:=inttostr(bb_mon)+vd+'0000';
      str2:=inttostr(bb_mon)+inttostr(strtoint(vd)+1)+'0000';
      str3:='savetime>='+str1+' and savetime<'+str2;
      sqls:=sqls+str3;
    end else begin
      i:=pos(',',vd);
      j:=1;
      while(i<>0) do begin
        str3:=copy(vd,1,2);
        if str3='00' then begin
          str3:=inttostr(bb_day);
          if bb_day<10 then
             str3:='0'+inttostr(bb_day);
        end;
        str1:=inttostr(bb_mon)+str3+'0000';

      str2:=inttostr(bb_mon)+inttostr(strtoint(str3)+1)+'0000';
      if   strtoint(str3)<9 then  str2:=inttostr(bb_mon)+'0'+inttostr(strtoint(str3)+1)+'0000';
        str3:='(savetime>='+str1+' and savetime<'+str2+')';
        if j=1 then
          sqls:=sqls+str3
        else
          sqls:=sqls+' or '+str3;
        i:=pos(',',vd);
        vd:=copy(vd,i+1,length(vd)-i);
        j:=j+1;
      end;
    end;
    with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add(sqls);
            try
              execsql;
            except
            end;

          end; //with   **  do begin
  end;
  if (vem='') and (ved<>'') then begin
  with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add('truncate table event_table');
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

          end; //with   **  do begin
    sqls:='insert into event_table  select * from '+evename+' where ';
    if length(ved)=2 then  begin
      if ved='00' then
        ved:=inttostr(bb_day);
      ymd1:=bb_year*10000+bb_mon*100+strtoint(ved);
      //ymd2:=1+ymd1;
      {str1:='to_date('''+inttostr(bb_year)+'-'+vem+'-1'',''yyyy-mm-dd'')';
      str2:=inttostr(getdaysofmon(bb_year,strtoint(vem)))+'+'+str1; }
      str3:='ymd='+inttostr(ymd1);
      sqls:=sqls+str3;
    end else begin
      i:=pos(',',ved);
      j:=1;
      while(i<>0) do begin
        str3:=copy(ved,1,2);
        if str3='00' then str3:=inttostr(bb_day);
        ymd1:=bb_year*10000+bb_mon*100+strtoint(str3);
        //ymd2:=1+ymd1;
        str3:='ymd='+inttostr(ymd1);
        if j=1 then
          sqls:=sqls+str3
        else
          sqls:=sqls+' or '+str3;
        i:=pos(',',ved);
        ved:=copy(ved,i+1,length(ved)-i);
        j:=j+1;
      end;
    end;

    with form1.adoquery1 do begin
            if active then close;
            sql.Clear;
           sql.Add(sqls);
           // edit1.Text:=sqls;
            try
              execsql;
            except
            end;

          end; //with   **  do begin
  end;
////////////////////////////////
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  try
    adodataset1.Close;
    adodataset3.Close;
    adodataset5.Close;
    adoquery1.Close;
    adoconnection1.Close;
  except
  end;
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
var
   s:string;
begin
    s:=combobox2.Text;
    if MessageDlg('删除 "'+s+'" ，确认乎?',mtInformation, [mbYes, mbNo], 0) = mrYes then
    begin

      with form1.adoquery1 do begin

        if active then close;
        sql.Clear;
        sql.add('delete rptlist where  valid=0 and  filename like ''%'+s+'%''');
        try
          Execsql;
        except
        end;
      end;
    end;
end;

procedure TForm1.BitBtn4Click(Sender: TObject);
begin
  ShellExecute(handle,'open','D:\YJDUNIX\tsbb.exe','-s','',SW_SHOWNORMAL);
end;




procedure TForm1.DBGrid1CellClick(Column: TColumn);
var
  ss:shortstring;
  i:integer;
begin
   ss:=adodataset3.fieldbyname('name').AsString;
   i:=adodataset3.fieldbyname('rtuno').AsInteger;

   with adodataset5 do begin
         if active then close;
           if i<>63 then
             commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where  filename like ''%'+ss+'%.XLS%'' order by filename desc'
           else
             commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid  from rptlist where   filename like ''%.XLS%'' order by filename ';
         try
           open;
         except
           showmessage('ERR:'+adodataset5.CommandText);
         end;
       end;
    form1.DBGrid5CellClick(nil);
end;

procedure TForm1.DBGrid5CellClick(Column: TColumn);
begin
  bb_valid:= adodataset5.fieldbyname('valid').AsInteger;
  combobox2.Text:=copy(adodataset5.fieldbyname('filename').AsString,1,length(adodataset5.fieldbyname('filename').AsString)-5);
  if pos('电量月',combobox2.Text)<>0 then combobox1.ItemIndex:=3
    else if  pos('能耗月',combobox2.Text)<>0 then combobox1.ItemIndex:=3
   else if  pos('能耗日',combobox2.Text)<>0 then combobox1.ItemIndex:=4
    else if  pos('巡视',combobox2.Text)<>0 then combobox1.ItemIndex:=1
  else if  pos('日报表',combobox2.Text)<>0 then combobox1.ItemIndex:=1
  else if  pos('时报表',combobox2.Text)<>0 then combobox1.ItemIndex:=2
  else if pos('月报表',combobox2.text)<>0 then combobox1.ItemIndex:=0

  else  combobox1.ItemIndex:=5;
  memo1.lines.Add(inttostr(DBGrid5.DataSource.DataSet.RecNo));
end;


{procedure TForm1.ExcelWorksheet1Change(ASender: TObject;
  const Target: ExcelRange);
begin
  if done then

    showmessage('ddd')
  else
    exit;
end; }

procedure TForm1.N1Click(Sender: TObject);
begin
  if MessageDlg('清除所有定义，确认乎?',
    mtInformation, [mbYes, mbNo], 0) = mrYes then
  begin
  end;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  td,sqls: string;
  tdt :tdate;
  sdt :tdate;
  len :integer;
  AFormat,AFormat2: TFormatSettings;
begin
             tdt:=trunc(date-1);
             AFormat.shortDateFormat:='yyyy/mm/dd';
              AFormat2.shortDateFormat:='yyyymmdd';
             AFormat.DateSeparator := '/';
             with adodataset1 do begin
                  if active then close;
                  commandText:='select max(did) time from xiebo_hgl ';

                 memo1.lines.Add(commandText);
                  open;
                  if  adodataset1.eof then  begin
                     td:='20171101';
                  end else begin

                    first;
                    while(not adodataset1.eof) do begin
                      td:=adodataset1.fieldbyname('time').asstring;
                      next;
                    end;
                  end;
                end;
                memo1.lines.Add('谐波统计最后日期是:'+td);
                if length(td)<8 then td:='20171101'; 
                td := copy(td,1,4)+'/'+copy(td,5,2)+'/'+copy(td,7,2);
                sdt:= StrToDate( td,AFormat);
                if sdt<tdt then begin
                     len := daysbetween(sdt,tdt);
                     sdt := sdt+1;
                     td := DateToStr(sdt,AFormat2);
                     sqls := 'call xb_hgl_range('''+td+''','+inttostr(len)+')';
                     memo1.lines.Add('进行谐波统计:'+td+','+inttostr(len));
                     with form1.adoquery1 do begin
                      if active then close;
                      sql.Clear;
                      sql.Add(sqls);
                      try
                        execsql;
                      except
                      end;
                       memo1.lines.Add('谐波统计 done');
                     end;
                end;
end;

procedure TForm1.UpDown2Click(Sender: TObject; Button: TUDBtnType);
var
  i:integer;
  ss:shortstring;

begin
   //if combobox1.ItemIndex<>2 then exit;
   i:=strtoint(qs.text);
   if Button = btnext then   begin
     i:=i+1;
     if i>23 then i:=0;
     if i<10 then ss:=inttostr(i)
     else ss:=inttostr(i);
     qs.Text:=ss;
   end;
   if Button = btprev then  begin
     i:=i-1;
     if i<0 then
     //showmessage('')
     i:=23;
     if i<10 then ss:=inttostr(i)
     else ss:=inttostr(i);
     qs.Text:=ss;
   end;
end;

procedure TForm1.UpDown3Click(Sender: TObject; Button: TUDBtnType);
var
  i:integer;
  ss:shortstring;

begin
   //if combobox1.ItemIndex<>2 then exit;
   i:=strtoint(qf.text);
   if Button = btnext then   begin
     i:=i+5;
     if i>55 then i:=0;
     if i<10 then ss:='0'+inttostr(i)
     else ss:=inttostr(i);
     qf.Text:=ss;
   end;
   if Button = btprev then  begin
     i:=i-5;
     if i<0 then
     //showmessage('')
     i:=55;
     if i<10 then ss:='0'+inttostr(i)
     else ss:=inttostr(i);
     qf.Text:=ss;
   end;
end;
procedure TForm1.DBGrid5KeyPress(Sender: TObject; var Key: Char);
begin
  if key='d' then bitbtn3click(sender);
end;

procedure TForm1.CheckBox2Click(Sender: TObject);
begin
  if checkbox2.Checked then memo1.Visible:=true;
  if not checkbox2.Checked then memo1.Visible:=false;

end;

procedure TForm1.Button2Click(Sender: TObject);
var
  tdt,sdt:tdate;
   AFormat,AFormat2: TFormatSettings;
   len :integer;
   tdts,sdts,sqls,td:string;
begin


             AFormat.shortDateFormat:='yyyy/mm/dd';
             AFormat2.shortDateFormat:='yyyymmdd';
             AFormat.DateSeparator := '/';
            tdt:=datetimepicker2.Date;
              try
              len:=strtoint(edit2.Text);
            except
              len:=1;
            end;
              sdt:=tdt+len;
             if checkbox3.Checked then begin
                with adodataset1 do begin
                  if active then close;
                  commandText:='select max(did) time from h_inc where did is not null ';

                 memo1.lines.Add(commandText);
                  open;
                  if  adodataset1.eof then  begin
                     td:='20190101';
                  end else begin

                    first;
                    while(not adodataset1.eof) do begin
                      td:=adodataset1.fieldbyname('time').asstring;
                      next;
                    end;
                  end;
                end;
                memo1.lines.Add('电量统计最后日期是:'+td);
                if length(td)<8 then td:='20190101';
                 td := copy(td,1,4)+'/'+copy(td,5,2)+'/'+copy(td,7,2);

                 tdt:=strtodate(td,AFormat);
                 sdt:=date;
                 if tdt=sdt then tdt:=tdt-1;
                 len:=daysbetween(tdt,sdt);

            end;


            tdts:=DateToStr(tdt,AFormat2);
            sdts:=DateToStr(sdt,AFormat2);
                     sqls := 'call dz_gen_range(str_to_date('''+tdts+'2000'',''%Y%m%d%H%i''),str_to_date('''+sdts+'2000'',''%Y%m%d%H%i''))';
                     memo1.lines.Add('电量增量计算:'+sqls);
                     with form1.adoquery1 do begin
                      if active then close;
                      sql.Clear;
                      sql.Add(sqls);
                      try
                        execsql;
                      except
                      end;
                       memo1.lines.Add('增量计算 done');
                     end;

                    sqls := 'call calc_ff_range(str_to_date('''+tdts+''',''%Y%m%d''),'+inttostr(len)+')';
                     memo1.lines.Add('峰谷电量计算:'+tdts+','+inttostr(len));
                     with form1.adoquery1 do begin
                      if active then close;
                      sql.Clear;
                      sql.Add(sqls);
                      try
                        execsql;
                      except
                      end;
                       memo1.lines.Add('峰谷电量计算 done');
                     end;




end;

procedure TForm1.Timer1Timer(Sender: TObject);
var
 ToTal,totall:longint;
begin
 totall:=24*60*60*1000;
 timer1.Interval:=totall;
 timer1.Enabled:=false;
  timer1.Enabled:=true;
  checkbox3.Checked:=true;
  datetimepicker1.Date:=date;
 datetimepicker2.Date:=date;
 with adoconnection1 do begin
     if connected then close; 
     open;
 end;
  Button1Click(nil);
  Button2Click(nil);
  Button3Click(nil);
  Button5Click(nil);

end;

procedure TForm1.Button3Click(Sender: TObject);
var
  bookmark: TBookmark;
  fnm :string;
  can:integer;

  qzw:Variant;
  s1,s2:word;
  i,j,gno,vno,blan,blans:integer;
  k:byte;
  ss,str1,sqls,str2,clh,lb,str3,str4,xls:string;
  fn:real;
  pmax,pmin,pavg,fmax,fmin,favg:real;

begin
 with adoconnection1 do begin
     if not connected then open;
 end;
  decodedate(DateTimePicker1.Date,bb_year,bb_mon,bb_day);
   ycname:='hyc'+inttostr(bb_year mod 10);
  evename:='eve_v'+inttostr(bb_year mod 10);
  dnname:='hdn'+inttostr(bb_year mod 10);
   Button8Click(nil);
 Button7Click(nil);
   Button9Click(nil);
    Button10Click(nil);
   Button11Click(nil);

end;

procedure TForm1.Button5Click(Sender: TObject);
var
  td,sqls: string;
  tdt :tdate;
  sdt :tdate;
  ldt :tdate;
  len :integer;
  AFormat,AFormat2: TFormatSettings;
  yy,mm,dd :word;
begin
             tdt:=trunc(date-1);
             AFormat.shortDateFormat:='yyyy/mm/dd';
              AFormat2.shortDateFormat:='yyyymmdd';
             AFormat.DateSeparator := '/';
             with adodataset1 do begin
                  if active then close;
                  commandText:='select max(did) time from dl_max ';

                 memo1.lines.Add(commandText);
                  open;
                  if  adodataset1.eof then  begin
                     td:='20171101';
                  end else begin

                    first;
                    while(not adodataset1.eof) do begin
                      td:=adodataset1.fieldbyname('time').asstring;
                      next;
                    end;
                  end;
                end;
                memo1.lines.Add('电流统计最后日期是:'+td);
                if length(td)<8 then td:='20171101'; 
                td := copy(td,1,4)+'/'+copy(td,5,2)+'/'+copy(td,7,2);
                sdt:= StrToDate( td,AFormat);
                while sdt<tdt do begin


                     td := DateToStr(sdt,AFormat2);
                     sqls := 'call calc_dl_max_range('''+td+''','+inttostr(1)+')';
                     memo1.lines.Add('进行电流统计:'+td+','+inttostr(1));
                     with form1.adoquery1 do begin
                      if active then close;
                      sql.Clear;
                      sql.Add(sqls);
                      try
                        execsql;
                      except
                      end;
                       memo1.lines.Add('电流统计 done');
                     end;
                      decodedate(sdt,yy,mm,dd);
                     len:=getdaysofmon(yy,mm);
                     sdt := sdt+len;
                end;
end;

procedure TForm1.Button6Click(Sender: TObject);
var
  ss:string;
begin
try
   try
       XLS.Filename := rptdir+'\12\一二期电费明细能耗月报表.xlsx';
       xls.Read;
    except
     memo1.lines.Add('请确认已下载Excel文件:'+rptdir+'\12\一二期电费明细能耗月报表.xlsx');
     exit;
    end;

 {
  irow:= XLS[0].LastRow;
  jcol:=XLS[0].LastCol;

 XLS.Sheets[0].AsFloat[2,4]:= 3;
 XLS.Sheets[0].AsFloat[2,5]:=44;
 XLS.Sheets[0].AsFloat[2,6]:=22;
 XLS.Sheets[0].AsFloat[2,7]:=13;
 XLS.Sheets[0].AsFloat[2,8]:=14;
  }
  xls.Calculate;
 ss:= xls.Sheets[0].AsSimpleTags[2,2];
  memo1.lines.Add(ss) ;
   ss:= xls.Sheets[0].AsFmtString[2,2];
  memo1.lines.Add(ss) ;
   ss:= xls.Sheets[0].AsHyperlink[2,2];
  memo1.lines.Add(ss) ;
  xls.SaveToFile('xx.XLSX');

 except  on e:Exception do
ShowMessage(ExtractFileDir(ParamStr(0))+'\'+'report\1号变抄表日报表.XLS'+e.Message);
 end;

end;



procedure TForm1.Button7Click(Sender: TObject);
var
  bookmark: TBookmark;
  fnm :string;
  can:integer;

  qzw:Variant;
  s1,s2:word;
  i,j,gno,vno,blan,blans:integer;
  k:byte;
  ss,str1,sqls,str2,clh,lb,str3,str4,xls:string;
  fn:real;
  pmax,pmin,pavg,fmax,fmin,favg:real;


begin with adoconnection1 do begin
     if  connected then begin
      close;

     end;
     open;
 end;
  decodedate(DateTimePicker1.Date,bb_year,bb_mon,bb_day);
   ycname:='hyc'+inttostr(bb_year mod 10);
  evename:='eve_v'+inttostr(bb_year mod 10);
  dnname:='hdn'+inttostr(bb_year mod 10);
  load_data:=true;
   with form1.adoquery1 do begin
            if active then close;
              sql.Clear;
                sql.add('truncate table yc_table') ;
            try
              execsql;
            except
              //showmessage('ddfd');
            end;
            {
             if active then close;
              sql.Clear;

                sql.add('truncate table dn_table');
            try
              execsql;
            except
              //showmessage('ddfd');
            end;
            }
             str1:=gettimes(bb_mon,bb_day,1);
        str2:=gettimed(bb_mon,bb_day,1);
        sqls:='insert into yc_table  select * from '+ycname+' where savetime>=';
        sqls:=sqls+str1+' and savetime<'+str2+' and chgtime is not null';
        str4:=' and savetime >= ' +str1+' and savetime<'+str2;
         if active then close;
              sql.Clear;

                sql.add(sqls);
            try
              execsql;
            except
              //showmessage('ddfd');
            end;
            { sqls:='insert into dn_table  select * from '+dnname+' where savetime>=';
        sqls:=sqls+str1+' and savetime<'+str2;
        str4:=' and savetime >= ' +str1+' and savetime<'+str2;
         if active then close;
              sql.Clear;

                sql.add(sqls);
            try
              execsql;
            except
              //showmessage('ddfd');
            end; }
     end;

                  i:=1;
                 // 
                   with adodataset5 do begin
                     if active then close;

                         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where   filename like ''%日报表%''  order by valid ';
                    memo1.lines.Add(commandtext);
                     try
                       open;
                     except
                       showmessage('ERR:'+adodataset5.CommandText);
                     end;
                   end;
                   can := adodataset5.RecordCount;

                    while(i<=can) do begin
                       DBGrid5.DataSource.DataSet.RecNo:=i;
                      fnm:=adodataset5.fieldbyname('filename').asstring;
                     
                      if fileexists(rptdir+fnm) then begin

                          memo1.lines.Add(rptdir+fnm);
                          //fnm:= copy(fnm,1,length(fnm)-4);
                          ///////////////////////////////////////////////////


                          DBGrid5CellClick(nil)  ;
                          memo1.lines.Add(combobox2.Text);
                          BitBtn2Click(nil);
                          // DBGrid1CellClick(nil);
  ///////////////////////////////////////////////////

                        end;   // fileexists(rptdir+fnm
                        i:=i+1;

                    end;
                    load_data:=false;

end;

procedure TForm1.Button8Click(Sender: TObject);
var
  bookmark: TBookmark;
  fnm :string;
  can:integer;

  qzw:Variant;
  s1,s2:word;
  i,j,gno,vno,blan,blans:integer;
  k:byte;
  ss,str1,sqls,str2,clh,lb,str3,str4,xls:string;
  fn:real;
  pmax,pmin,pavg,fmax,fmin,favg:real;

begin with adoconnection1 do begin
     if  connected then begin
      close;

     end;
     open;
 end;

  decodedate(DateTimePicker1.Date,bb_year,bb_mon,bb_day);
   ycname:='hyc'+inttostr(bb_year mod 10);
  evename:='eve_v'+inttostr(bb_year mod 10);
  dnname:='hdn'+inttostr(bb_year mod 10);
  load_data:=true;
   with form1.adoquery1 do begin
            if active then close;
              sql.Clear;
                sql.add('truncate table yc_table') ;
            try
             // execsql;
            except
              //showmessage('ddfd');
            end;
            {
             if active then close;
              sql.Clear;

                sql.add('truncate table dn_table');
            try
              execsql;
            except
              //showmessage('ddfd');
            end;
            }
                str1:=gettimes(bb_mon,bb_day,1);
        str2:=gettimed(bb_mon,bb_day,1);
        sqls:='insert into yc_table  select * from '+ycname+' where savetime>=';
        sqls:=sqls+str1+' and savetime<'+str2+' and chgtime is not null';
        str4:=' and savetime >= ' +str1+' and savetime<'+str2+' and chgtime is not null';
         if active then close;
              sql.Clear;

                sql.add(sqls);
            try
            //  execsql;

            except
              //showmessage('ddfd');
            end;
         end;
          memo1.Lines.Add(sqls);
                  i:=1;
                 // 
                   with adodataset5 do begin
                     if active then close;

                         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where   filename like ''%月报表%'' and filename not like ''%电量%'' and filename not like ''%能耗%'' and valid<>1  order by valid ';

                     try
                       open;
                        memo1.Lines.Add(commandtext);
                     except
                       showmessage('ERR:'+adodataset5.CommandText);
                     end;
                   end;
                   can := adodataset5.RecordCount;
                    memo1.Lines.Add(inttostr(can));
                    while(i<=can) do begin
                       DBGrid5.DataSource.DataSet.RecNo:=i;
                      fnm:=adodataset5.fieldbyname('filename').asstring;
                       memo1.Lines.Add(rptdir+fnm);
                      if fileexists(rptdir+fnm) then begin

                          memo1.Lines.Add(rptdir+fnm);
                          //fnm:= copy(fnm,1,length(fnm)-4);
                          ///////////////////////////////////////////////////


                          DBGrid5CellClick(nil)  ;
                          memo1.Lines.Add(combobox2.Text);
                          BitBtn2Click(nil);
                          // DBGrid1CellClick(nil);
  ///////////////////////////////////////////////////

                        end;   // fileexists(rptdir+fnm
                        i:=i+1;

                    end;
                    load_data:=false;

end;

procedure TForm1.Button9Click(Sender: TObject);
var
  bookmark: TBookmark;
  fnm :string;
  can:integer;

  qzw:Variant;
  s1,s2:word;
  i,j,gno,vno,blan,blans:integer;
  k:byte;
  ss,str1,sqls,str2,clh,lb,str3,str4,xls:string;
  fn:real;
  pmax,pmin,pavg,fmax,fmin,favg:real;


begin with adoconnection1 do begin
     if  connected then begin
      close;

     end;
     open;
 end;
  decodedate(DateTimePicker1.Date,bb_year,bb_mon,bb_day);
   ycname:='hyc'+inttostr(bb_year mod 10);
  evename:='eve_v'+inttostr(bb_year mod 10);
  dnname:='hdn'+inttostr(bb_year mod 10);


                  i:=1;
                 // 
                   with adodataset5 do begin
                     if active then close;

                         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where   ( filename  like ''%电量月%'' or filename  like ''%能耗月%'') and valid<>1  order by valid ';
                   // memo1.lines.Add(commandtext);
                     try
                       open;
                     except
                       showmessage('ERR:'+adodataset5.CommandText);
                     end;
                   end;
                   can := adodataset5.RecordCount;

                    while(i<=can) do begin
                       DBGrid5.DataSource.DataSet.RecNo:=i;
                      fnm:=adodataset5.fieldbyname('filename').asstring;
                     
                      if fileexists(rptdir+fnm) then begin

                         // memo1.lines.Add(rptdir+fnm);
                          //fnm:= copy(fnm,1,length(fnm)-4);
                          ///////////////////////////////////////////////////


                          DBGrid5CellClick(nil)  ;
                          memo1.lines.Add(combobox2.Text);
                          BitBtn2Click(nil);
                          // DBGrid1CellClick(nil);
  ///////////////////////////////////////////////////

                        end;   // fileexists(rptdir+fnm
                        i:=i+1;

                    end;
                    load_data:=false;

end;

procedure TForm1.Button10Click(Sender: TObject);
var
  bookmark: TBookmark;
  fnm :string;
  can:integer;

  qzw:Variant;
  s1,s2:word;
  i,j,gno,vno,blan,blans:integer;
  k:byte;
  ss,str1,sqls,str2,clh,lb,str3,str4,xls:string;
  fn:real;
  pmax,pmin,pavg,fmax,fmin,favg:real;


begin with adoconnection1 do begin
     if  connected then begin
      close;

     end;
     open;
 end;
  decodedate(DateTimePicker1.Date,bb_year,bb_mon,bb_day);
   ycname:='hyc'+inttostr(bb_year mod 10);
  evename:='eve_v'+inttostr(bb_year mod 10);
  dnname:='hdn'+inttostr(bb_year mod 10);


                  i:=1;
                 // 
                   with adodataset5 do begin
                     if active then close;

                         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where     valid=1  order by valid ';
                    memo1.lines.Add(commandtext);
                     try
                       open;
                     except
                       showmessage('ERR:'+adodataset5.CommandText);
                     end;
                   end;
                   can := adodataset5.RecordCount;

                    while(i<=can) do begin
                       DBGrid5.DataSource.DataSet.RecNo:=i;
                      fnm:=adodataset5.fieldbyname('filename').asstring;
                     
                      if fileexists(rptdir+fnm) then begin

                          memo1.lines.Add(rptdir+fnm);
                          //fnm:= copy(fnm,1,length(fnm)-4);
                          ///////////////////////////////////////////////////


                          DBGrid5CellClick(nil)  ;
                          memo1.lines.Add(combobox2.Text);
                          BitBtn2Click(nil);
                          // DBGrid1CellClick(nil);
  ///////////////////////////////////////////////////

                        end;   // fileexists(rptdir+fnm
                        i:=i+1;

                    end;
                    load_data:=false;
                    end;

procedure TForm1.Button11Click(Sender: TObject);
var
  bookmark: TBookmark;
  fnm :string;
  can:integer;

  qzw:Variant;
  s1,s2:word;
  i,j,gno,vno,blan,blans:integer;
  k:byte;
  ss,str1,sqls,str2,clh,lb,str3,str4,xls:string;
  fn:real;
  pmax,pmin,pavg,fmax,fmin,favg:real;


begin with adoconnection1 do begin
     if  connected then begin
      close;

     end;
     open;
 end;
  decodedate(DateTimePicker1.DateTime,bb_year,bb_mon,bb_day);
   ycname:='hyc'+inttostr(bb_year mod 10);
  evename:='eve_v'+inttostr(bb_year mod 10);
  dnname:='hdn'+inttostr(bb_year mod 10);
  load_data:=true;
   with form1.adoquery1 do begin
            if active then close;
              sql.Clear;
                sql.add('truncate table dn_table') ;
            try
              execsql;
            except
              //showmessage('ddfd');
            end;
            {
             if active then close;
              sql.Clear;

                sql.add('truncate table dn_table');
            try
              execsql;
            except
              //showmessage('ddfd');
            end;
            }
        if not ((bb_mon=1) and (bb_day=1)) then begin
            str1:=gettimed_(bb_mon,bb_day,1);
            str2:=gettimes_(bb_mon,bb_day,1);
            sqls:='insert into dn_table  select * from dn_inc where savetime>=';
            sqls:=sqls+str1+' and savetime<'+str2;
            sqls:=sqls+' and floor(chgtime/10000)='+inttostr(bb_year);
            str4:=' and savetime >= ' +str1+' and savetime<'+str2;
        end else begin
            str1:=gettimed_(bb_mon,bb_day,1);
            str2:=gettimes_(bb_mon,bb_day,1);
            sqls:='insert into dn_table  select * from dn_inc where savetime>=';
            sqls:=sqls+str1+' or savetime<'+str2;
            sqls:=sqls+' and floor(chgtime/10000)='+inttostr(bb_year);
            str4:=' and savetime >= ' +str1+' and savetime<'+str2;
        end;
         memo1.Lines.Add(sqls);
         if active then close;
              sql.Clear;

                sql.add(sqls);
                 //memo1.lines.Add(sqls);
            try
              execsql;
            except
              //showmessage('ddfd');
            end;
            { sqls:='insert into dn_table  select * from '+dnname+' where savetime>=';
        sqls:=sqls+str1+' and savetime<'+str2;
        str4:=' and savetime >= ' +str1+' and savetime<'+str2;
         if active then close;
              sql.Clear;

                sql.add(sqls);
            try
              execsql;
            except
              //showmessage('ddfd');
            end; }
     end;

                  i:=1;
                 // 
                   with adodataset5 do begin
                     if active then close;

                         commandtext:='select substr(trim(filename),1,length(trim(filename))-5) filename,valid from rptlist where   filename like ''%能耗日%''  order by valid ';
                    memo1.lines.Add(commandtext);
                     try
                       open;
                     except
                       showmessage('ERR:'+adodataset5.CommandText);
                     end;
                   end;
                   can := adodataset5.RecordCount;

                    while(i<=can) do begin
                       DBGrid5.DataSource.DataSet.RecNo:=i;
                      fnm:=adodataset5.fieldbyname('filename').asstring;
                     
                      if fileexists(rptdir+fnm) then begin

                          memo1.lines.Add(rptdir+fnm);
                          //fnm:= copy(fnm,1,length(fnm)-4);
                          ///////////////////////////////////////////////////


                          DBGrid5CellClick(nil)  ;
                          memo1.Lines.Add(combobox2.Text);
                          BitBtn2Click(nil);
                          // DBGrid1CellClick(nil);
  ///////////////////////////////////////////////////

                        end;   // fileexists(rptdir+fnm
                        i:=i+1;

                    end;
                    load_data:=false;

end;

procedure TForm1.Timer2Timer(Sender: TObject);
var
  bookmark: TBookmark;
  fnm :string;
  can:integer;

  qzw:Variant;
  s1,s2:word;
  i,j,gno,vno,blan,blans:integer;
  k:byte;
  ss,str1,sqls,str2,clh,lb,str3,str4,xls:string;
  fn:real;
  pmax,pmin,pavg,fmax,fmin,favg:real;
begin
   with form1.adoquery1 do begin
            if active then close;
              sql.Clear;
                sql.add('select * from prtu') ;
            try
              adoquery1.Open;
            except
              //showmessage('ddfd');
            end;
            {
             if active then close;
              sql.Clear;

                sql.add('truncate table dn_table');
            try
              execsql;
            except
              //showmessage('ddfd');
            end;
            }

         end;
end;

end.
