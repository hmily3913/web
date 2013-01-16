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
'=======================收料及时率========================
if key = "RPreply" then
	sqlajax="select * from stocksys_ReceiveProm where SerialNum = '"&FItemid&"' "
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
elseif key ="updateRPreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from stocksys_ReceiveProm where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update stocksys_ReceiveProm set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================发料达成率========================
elseif key = "SPreply" then
	sqlajax="select * from stocksys_SendmReach where SerialNum = '"&FItemid&"' "
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
elseif key ="updateSPreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from stocksys_SendmReach where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update stocksys_SendmReach set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================成品出货及时率========================
elseif key = "FDRreply" then
	sqlajax="select * from stocksys_FinishDeliveryRate where SerialNum = '"&FItemid&"' "
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
elseif key ="updateFDRreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from stocksys_FinishDeliveryRate where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update stocksys_FinishDeliveryRate set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================帐卡物准确率========================
elseif key = "CARreply" then
	sqlajax="select * from stocksys_CardAccuracyRate where SerialNum = '"&FItemid&"' "
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
elseif key ="updateCARreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from stocksys_CardAccuracyRate where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update stocksys_CardAccuracyRate set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================提货及时率========================
elseif key = "PPreply" then
	sqlajax="select * from stocksys_PickupProm where SerialNum = '"&FItemid&"' "
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
elseif key ="updatePPreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from stocksys_PickupProm where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update stocksys_PickupProm set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================派车及时率========================
elseif key = "SCPreply" then
	sqlajax="select * from stocksys_SendCarProm where SerialNum = '"&FItemid&"' "
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
elseif key ="updateSCPreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from stocksys_SendCarProm where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update stocksys_SendCarProm set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
end if
%>