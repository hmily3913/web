<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
dim FItemid,key
dim rsajax,sqlajax
dim arryword(5)
key=request.QueryString("key")
FItemid=request.QueryString("FItemid")
if key = "empname" then
	'获取员工信息
	sqlajax="select * from [N-基本资料单头] where 员工代号 like '%"&FItemid&"%'"
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,conn,0,1
	if rsajax.bof and rsajax.eof then
		response.Write(sqlajax&"!@#$")
	else
		arryword(0)=rsajax("员工代号")
		arryword(1)=rsajax("姓名")
		arryword(2)=rsajax("密码")
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2))
	end if
	rsajax.close
	set rsajax=nothing
end if
%>