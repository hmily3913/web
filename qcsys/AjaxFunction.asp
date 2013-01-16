<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
dim FItemid,key,UserName,AdminName,FEntryID
dim ReplyType
dim rsajax,sqlajax
dim arryword(5)
dim ReplyText
key=request.QueryString("key")
FItemid=request.QueryString("FItemid")
FEntryID=request.QueryString("FEntryID")
UserName=session("UserName")
AdminName=session("AdminName")
'=======================成品合格率========================
if key = "FQreply" then
	sqlajax="select * from hrsys_PersonnelLoss where qcsys_FinishQualified = '"&FItemid&"' "
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
elseif key ="updateFQreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from qcsys_FinishQualified where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update qcsys_FinishQualified set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================不合格来料处理率========================
elseif key = "UQMDreply" then
	sqlajax="select * from qcsys_UnQualifiMtrDeal where SerialNum = '"&FItemid&"' "
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
elseif key ="updateUQMDreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from qcsys_UnQualifiMtrDeal where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update qcsys_UnQualifiMtrDeal set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================进料检验的及时率========================
elseif key = "CCPreply" then
	sqlajax="select * from qcsys_ComeCheckProm where SerialNum = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		response.Write("#########1")
	else
		arryword(0)=rsajax("Replyer")
		arryword(1)=rsajax("replyDate")
		arryword(2)=rsajax("ReplyText")
		arryword(3)=rsajax("ReplyType")
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1###"&arryword(3))
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="updateCCPreply" then
	ReplyText=request.QueryString("ReplyText")
	ReplyType=request.QueryString("ReplyType")
	sqlajax="select * from qcsys_ComeCheckProm where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update qcsys_ComeCheckProm set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',ReplyType='"&ReplyType&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###"&ReplyType)
	end if
	rsajax.close
	set rsajax=nothing
'=======================进料检验准确率========================
elseif key = "CCAreply" then
	sqlajax="select * from qcsys_ComeCheckAccur where SerialNum = '"&FItemid&"' "
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
elseif key ="updateCCAreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from qcsys_ComeCheckAccur where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update qcsys_ComeCheckAccur set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================客户投诉========================
elseif key = "CCreply" then
	sqlajax="select * from qcsys_CustomComplain where SerialNum = '"&FItemid&"' "
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
	sqlajax="select * from qcsys_CustomComplain where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update qcsys_CustomComplain set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================检测及时率========================
elseif key = "CPreply" then
	sqlajax="select * from qcsys_CheckProm where SerialNum = '"&FItemid&"' "
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
elseif key ="updateCPreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from qcsys_CheckProm where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update qcsys_CheckProm set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
'=======================首检错误========================
elseif key = "FCAreply" then
	sqlajax="select * from qcsys_FirstCheckAccur where SerialNum = '"&FItemid&"' "
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
elseif key ="updateFCAreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from qcsys_FirstCheckAccur where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update qcsys_FirstCheckAccur set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
end if
%>