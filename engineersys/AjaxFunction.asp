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
'=======================打样按期完成率========================
if key = "PFRreply" then
	sqlajax="select * from engineersys_ProofingFinishRate where SerialNum = '"&FItemid&"' "
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
elseif key ="updatePFRreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from engineersys_ProofingFinishRate where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update engineersys_ProofingFinishRate set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================BOM表按期完成率========================
elseif key = "BFreply" then
	sqlajax="select * from engineersys_BomFinish where SerialNum = '"&FItemid&"' "
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
elseif key ="updateBFreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from engineersys_BomFinish where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update engineersys_BomFinish set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================新产品开发件数========================
elseif key = "NPDreply" then
	sqlajax="select * from engineersys_NewProductDev where SerialNum = '"&FItemid&"' "
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
elseif key ="updateNPDreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from engineersys_NewProductDev where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update engineersys_NewProductDev set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================新产品接单额========================
elseif key = "NPOreply" then
	sqlajax="select * from engineersys_NewProductOrder where SerialNum = '"&FItemid&"' "
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
elseif key ="updateNPOreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from engineersys_NewProductOrder where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update engineersys_NewProductOrder set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================产前试做合格率========================
elseif key = "TQreply" then
	sqlajax="select * from engineersys_TryQualified where SerialNum = '"&FItemid&"' "
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
elseif key ="updateCCreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from engineersys_TryQualified where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update engineersys_TryQualified set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
end if
%>