<%
Dim Conn,ConnStr
Dim Connk3,ConnStrk3
Dim Connzxpt,ConnStrzxpt
Dim Connkq,ConnStrkq
dim AllOPENROWSET
AllOPENROWSET="OPENROWSET('SQLOLEDB ', '192.168.0.9'; 'sa'; 'lovemaster',"
On error resume next
Set Conn=Server.CreateObject("Adodb.Connection")
ConnStr="driver={SQL Server};server=192.168.0.9;UID=sa;PWD=lovemaster;Database=LDERP"
Conn.open ConnStr
Set Connk3=Server.CreateObject("Adodb.Connection")
ConnStrk3="driver={SQL Server};server=192.168.0.9;UID=sa;PWD=lovemaster;Database=AIS20081217153921"
Connk3.open ConnStrk3
Set Connzxpt=Server.CreateObject("Adodb.Connection")
'ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(SiteDataPath)
ConnStrzxpt="driver={SQL Server};server=192.168.0.9;UID=sa;PWD=lovemaster;Database=zxpt"
'ConnStrzxpt="driver={SQL Server};server=192.168.0.29;UID=sa;PWD=landao;Database=zxpt"
Connzxpt.open ConnStrzxpt
Set Connkq=Server.CreateObject("Adodb.Connection")
ConnStrkq=session("KQSQLSTR")
Connkq.open ConnStrkq
if err then
   err.clear
   Set Conn = Nothing
   Set Connk3 = Nothing
   Set Connzxpt = Nothing
   Set Connkq = Nothing
   Response.Write "系统错误：数据库连接出错，请检查'系统管理>>站点常量设置',请联系管理员!"
   Response.End
end if
%>
<!--#include file="Function.asp" -->
