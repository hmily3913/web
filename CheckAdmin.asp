<%
Response.Charset="utf-8"
'判断是否登录或登录超时===============================================================
On error resume next
Dim Connonline,ConnStronline
Set Connonline=Server.CreateObject("Adodb.Connection")
ConnStronline="driver={SQL Server};server=192.168.0.9;UID=sa;PWD=lovemaster;Database=zxpt"
Connonline.open ConnStronline
if err then
   err.clear
   Set Connonline = Nothing
   Response.End
end if

dim AdminAction,sqlonline,rsonline
AdminAction=request.QueryString("AdminAction")
select case AdminAction
  case "Out"
    call OutLogin()
  case else
    call Login()
end select
sub Login()
  if session("UserName")="" or session("LoginSystem")<>"Succeed" then
     response.write "您还没有登录或登录已超时，请<a href='/u_Login.asp' target='_parent'><font color='red'>返回登录</font></a>!"
     response.end
	else
		 dim LoginIP,LoginTime,LoginSoft
		 LoginIP=Request.ServerVariables("Remote_Addr")
		 LoginSoft=Request.ServerVariables("Http_USER_AGENT")
		 LoginTime=now()
		 sqlonline="select * from smmsys_Online where UserName='"&session("UserName")&"'"
		 set rsonline = server.createobject("adodb.recordset")
		 rsonline.open sqlonline,Connonline,1,1
		 if not rsonline.eof then
			sqlonline="update smmsys_Online set o_ip='"&LoginIP&"',o_lasttime='"&LoginTime&"',LoginSoft='"&LoginSoft&"',AdminName='"&session("AdminName")&"',o_state=1 where UserName='"&session("UserName")&"'"
			Connonline.Execute (sqlonline)
		 else
			 sqlonline = "insert into smmsys_Online (o_ip,UserName, o_lasttime,LoginSoft,AdminName,o_state) values ('"&LoginIP&"','" & session("UserName") & "', '" & LoginTime & "','"&LoginSoft&"','" & session("AdminName") & "',1)"
			 Connonline.Execute (sqlonline)
		 end if
		Connonline.Execute ("delete from smmsys_Online where DateDiff(d,o_lasttime,getdate())>7")
		Connonline.Execute ("update smmsys_Online set o_state=0 where DateDiff(n,o_lasttime,getdate())>10")
  end if
end sub
'========
sub OutLogin()
	Connonline.Execute ("update smmsys_Online set o_state=0 where UserName = '" &session("UserName")&"'")
  session.contents.remove "UserName"
  session.contents.remove "LoginSystem"
  session.contents.remove "AdminPurview"
  session.contents.remove "AdminPurviewFLW"
  session.contents.remove "VerifyCode"
  response.write "<script language=javascript>top.location.replace('u_Login.asp');</script>"
end sub



'浏览器、操作系统版本侦测
function browser(text)
    if Instr(text,"MSIE 5.5")>0 then
        browser="IE 5.5"
    elseif Instr(text,"MSIE 8.0")>0 then
        browser="IE 8.0"
    elseif Instr(text,"MSIE 7.0")>0 then
        browser="IE 7.0"
    elseif Instr(text,"MSIE 6.0")>0 then
        browser="IE 6.0"
    elseif Instr(text,"MSIE 5.01")>0 then
        browser="IE 5.01"
    elseif Instr(text,"MSIE 5.0")>0 then
        browser="IE 5.00"
    elseif Instr(text,"MSIE 4.0")>0 then
        browser="IE 4.01"
        else
        browser="未知"
    end if
end function
function system(text)
    if Instr(text,"NT 5.1")>0 then
        system="Windows XP"
    elseif Instr(text,"NT 6.1")>0 then
        system="Windows  7"
    elseif Instr(text,"NT 6.0")>0 then
        system="Windows Vista/Server 2008"
    elseif Instr(text,"NT 5.2")>0 then
        system="Windows Server 2003"
    elseif Instr(text,"NT 5.1")>0 then
        system="Windows XP"
    elseif Instr(text,"NT 5")>0 then
        system="Windows 2000"
    elseif Instr(text,"NT 4")>0 then
        system="Windows NT4"
    elseif Instr(text,"4.9")>0 then
        system="Windows ME"
    elseif Instr(text,"98")>0 then
        system="Windows 98"
    elseif Instr(text,"95")>0 then
        system="Windows 95"
    elseif Instr(text,"Mac")>0 then
        system="Mac"
    elseif Instr(text,"Unix")>0 then
        system="Unix"
    elseif Instr(text,"Linux")>0 then
        system="Linux"
    elseif Instr(text,"SunOS")>0 then
        system="SunOS"
        else
        system="未知"
    end if
end function

%>