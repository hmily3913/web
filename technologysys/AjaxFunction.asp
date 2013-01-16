<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
dim FItemid,key,UserName,AdminName,FEntryID
dim rsajax,sqlajax
dim arryword(5)
dim ReplyText
key=request.QueryString("key")
FItemid=request.QueryString("FItemid")
FEntryID=request.QueryString("FEntryID")
UserName=session("UserName")
AdminName=session("AdminName")
'=======================模具维修及时率========================
if key = "MRreply" then
	sqlajax="select * from technologysys_Moldrepair where SerialNum = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		response.Write("#########1")
	else
		arryword(0)=rsajax("Replyer")
		arryword(1)=rsajax("replyDate")
		arryword(2)=rsajax("ReplyText")
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="updateMRreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from technologysys_Moldrepair where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update technologysys_Moldrepair set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================模具制作及时率========================
elseif key = "MMreply" then
	sqlajax="select * from technologysys_Moldmake where SerialNum = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		response.Write("#########1")
	else
		arryword(0)=rsajax("Replyer")
		arryword(1)=rsajax("replyDate")
		arryword(2)=rsajax("ReplyText")
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="updateMMreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from technologysys_Moldmake where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update technologysys_Moldmake set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================设备维修及时率========================
elseif key = "DRreply" then
	sqlajax="select * from technologysys_Devicerepair where SerialNum = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		response.Write("#########1")
	else
		arryword(0)=rsajax("Replyer")
		arryword(1)=rsajax("replyDate")
		arryword(2)=rsajax("ReplyText")
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="updateDRreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from technologysys_Devicerepair where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update technologysys_Devicerepair set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
end if
%>