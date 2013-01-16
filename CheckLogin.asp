<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012 - zbh-STUDIO" />
<META NAME="Author" CONTENT="---zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>检查用户登录</TITLE>
</HEAD>
<BODY>
<%
Dim Conn,ConnStr
Dim Connzxpt,ConnStrzxpt
On error resume next
Set Conn=Server.CreateObject("Adodb.Connection")
ConnStr="driver={SQL Server};server=192.168.0.9;UID=sa;PWD=lovemaster;Database=LDERP"
Conn.open ConnStr
Set Connzxpt=Server.CreateObject("Adodb.Connection")
ConnStrzxpt="driver={SQL Server};server=192.168.0.9;UID=sa;PWD=lovemaster;Database=zxpt"
Connzxpt.open ConnStrzxpt

dim LoginName,LoginPassword,AdminName,Password,AdminPurview,UserName,rs,sql,AdminPurviewFLW,Depart,DepartName

LoginName=UCase(trim(request.form("LoginName")))
LoginPassword=request.form("LoginPassword")
set rs = server.createobject("adodb.recordset")
sql="select a.员工代号,a.姓名,a.部门别 as Depart,a.密码,a.权限,a.权限_FLW,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where 员工代号='"&LoginName&"' and 是否离职=0 and a.部门别=b.部门代号"
rs.open sql,conn,1,3

if rs.eof then
   response.write "<script language=javascript> alert('用户名称不正确，请重新输入。');location.replace('u_Login.asp');</script>"
   response.end
else
   UserName=rs("员工代号")
   AdminName=rs("姓名")
   Password=rs("密码")
   AdminPurview=rs("权限")
   AdminPurviewFLW=rs("权限_FLW")
   Depart=rs("Depart")
   DepartName=rs("部门名称")
end if
if LoginPassword<>Password then
'   response.write LoginPassword
   response.write "<script language=javascript> alert('用户密码不正确，请重新输入。');location.replace('u_Login.asp');</script>"
   response.end
end if 

if LoginName=UserName and LoginPassword=Password then
'   rs("frighttype")=9
'   rs.update
   rs.close
   set rs=nothing 
   
	set rs = server.createobject("adodb.recordset")
	sql="select a.Permissions,a.PermissionsFLW from smmsys_PermissionGroup a,smmsys_PermissionGroupDetails b where a.ForbidFlag=0 and a.SerialNum=b.GroupSnum and b.UserID='"&UserName&"'"
	rs.open sql,connzxpt,1,1
	while (not rs.eof)
	  AdminPurview=rs("Permissions")&AdminPurview
	  AdminPurviewFLW=rs("PermissionsFLW")&AdminPurviewFLW
		rs.movenext
	wend
   session("UserName")=UserName
   session("AdminName")=AdminName
   session("AdminPurview")=AdminPurview
   session("AdminPurviewFLW")=AdminPurviewFLW
   session("LoginSystem")="Succeed"
   session("Depart")=Depart
	 session("DepartName")=DepartName
	 dim Connkq,ConnStrkq
Set Connkq=Server.CreateObject("Adodb.Connection")
ConnStrkq="driver={SQL Server};server=122.228.158.226;UID=sa;PWD=lovemaster;Database=att2000"
Connkq.open ConnStrkq
set rs = server.createobject("adodb.recordset")
	sql="select 1 from USERINFO where ssn='"&UserName&"'"
	rs.open sql,connkq,1,1
	response.Write(rs.eof)
	 if rs.eof then
	 	session("KQSQLSTR")="driver={SQL Server};server=192.168.0.184;UID=sa;PWD=ldrz;Database=KQ2011"
	 else
	 	session("KQSQLSTR")="driver={SQL Server};server=122.228.158.226;UID=sa;PWD=lovemaster;Database=att2000"
	 end if
   session.timeout=1000
   '==================================
   dim LoginIP,LoginTime,LoginSoft
   LoginIP=Request.ServerVariables("Remote_Addr")
   LoginSoft=Request.ServerVariables("Http_USER_AGENT")
   LoginTime=now()
   '====================================
	 sql="select * from smmsys_Online where UserName='"&UserName&"'"
	 set rs = server.createobject("adodb.recordset")
	 rs.open sql,connzxpt,1,1
	 if not rs.eof then
	 	sql="update smmsys_Online set o_ip='"&LoginIP&"',o_lasttime='"&LoginTime&"',LoginSoft='"&LoginSoft&"',AdminName='"&AdminName&"',o_state=1 where UserName='"&UserName&"'"
		connzxpt.Execute (sql)
	 else
     sql = "insert into smmsys_Online (o_ip,UserName, o_lasttime,LoginSoft,AdminName,o_state) values ('"&LoginIP&"','" & UserName & "', '" & LoginTime & "','"&LoginSoft&"','" & AdminName & "',1)"
     connzxpt.Execute (sql)
	 end if
   rs.close
   set rs=nothing 
   '========================================
   response.redirect "main.asp"
   response.end
end if
%>
</BODY>
</HTML>