<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
dim FItemid,key,UserName,AdminName,FEntryID,Depart,OtherEle
dim rsajax,sqlajax
dim arryword(5)
dim ReplyText
key=request.QueryString("key")
FItemid=request.QueryString("FItemid")
FEntryID=request.QueryString("FEntryID")
UserName=session("UserName")
AdminName=session("AdminName")
Depart=session("Depart")
'=======================生产工作流========================
if key = "PCreply" then
	'生管回复
	sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|101.1,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("PCReplyer")
		arryword(1)=rsajax("PCreplyDate")
		arryword(2)=rsajax("PCReplyText")
	  if Instr(session("AdminPurviewFLW"),"|101.1,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="MCreply" then
	'生管回复
	sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|101.2,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("MCReplyer")
		arryword(1)=rsajax("MCreplyDate")
		arryword(2)=rsajax("MCReplyText")
	  if Instr(session("AdminPurviewFLW"),"|101.2,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="STreply" then
	'仓库回复
	sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|101.3,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("STReplyer")
		arryword(1)=rsajax("STreplyDate")
		arryword(2)=rsajax("STReplyText")
	  if Instr(session("AdminPurviewFLW"),"|101.3,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="SEreply" then
	'仓库回复
	sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|101.4,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("SEReplyer")
		arryword(1)=rsajax("SEreplyDate")
		arryword(2)=rsajax("SEReplyText")
	  if Instr(session("AdminPurviewFLW"),"|101.4,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="SUreply" then
	'采购回复
	sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|101.5,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("SUReplyer")
		arryword(1)=rsajax("SUreplyDate")
		arryword(2)=rsajax("SUReplyText")
	  if Instr(session("AdminPurviewFLW"),"|101.5,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="QCreply" then
	'品保回复
	sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|101.6,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("QCReplyer")
		arryword(1)=rsajax("QCreplyDate")
		arryword(2)=rsajax("QCReplyText")
	  if Instr(session("AdminPurviewFLW"),"|101.6,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="ENreply" then
	'工程回复
	sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|101.7,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("ENReplyer")
		arryword(1)=rsajax("ENreplyDate")
		arryword(2)=rsajax("ENReplyText")
	  if Instr(session("AdminPurviewFLW"),"|101.7,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="updatePCreply" then
	'仓库回复
	if Instr(session("AdminPurviewFLW"),"|101.1,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_Icmo (FInterID,PCReplyer,PCreplyDate,PCReplyText,PCreplyFlag) values ("&FItemid&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_Icmo set PCReplyer='"&AdminName&"',PCreplyDate='"&Now()&"',PCReplyText='"&ReplyText&"',PCreplyFlag=PCreplyFlag+1 where FInterID="&FItemid
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateMCreply" then
	'仓库回复
	if Instr(session("AdminPurviewFLW"),"|101.2,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_Icmo (FInterID,MCReplyer,MCreplyDate,MCReplyText,MCreplyFlag) values ("&FItemid&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_Icmo set MCReplyer='"&AdminName&"',MCreplyDate='"&Now()&"',MCReplyText='"&ReplyText&"',MCreplyFlag=MCreplyFlag+1 where FInterID="&FItemid
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateSTreply" then
	'仓库回复
	if Instr(session("AdminPurviewFLW"),"|101.3,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_Icmo (FInterID,STReplyer,STreplyDate,STReplyText,STreplyFlag) values ("&FItemid&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_Icmo set STReplyer='"&AdminName&"',STreplyDate='"&Now()&"',STReplyText='"&ReplyText&"',STreplyFlag=STreplyFlag+1 where FInterID="&FItemid
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateSEreply" then
	'仓库回复
	if Instr(session("AdminPurviewFLW"),"|101.4,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_Icmo (FInterID,SEReplyer,SEreplyDate,SEReplyText,SEreplyFlag) values ("&FItemid&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_Icmo set SEReplyer='"&AdminName&"',SEreplyDate='"&Now()&"',SEReplyText='"&ReplyText&"',SEreplyFlag=SEreplyFlag+1 where FInterID="&FItemid
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateSUreply" then
	'采购回复
	if Instr(session("AdminPurviewFLW"),"|101.5,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_Icmo (FInterID,SUReplyer,SUreplyDate,SUReplyText,SUreplyFlag) values ("&FItemid&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_Icmo set SUReplyer='"&AdminName&"',SUreplyDate='"&Now()&"',SUReplyText='"&ReplyText&"',SUreplyFlag=SUreplyFlag+1 where FInterID="&FItemid
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateQCreply" then
	'品保回复
	if Instr(session("AdminPurviewFLW"),"|101.6,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_Icmo (FInterID,QCReplyer,QCreplyDate,QCReplyText,QCreplyFlag) values ("&FItemid&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_Icmo set QCReplyer='"&AdminName&"',QCreplyDate='"&Now()&"',QCReplyText='"&ReplyText&"',QCreplyFlag=QCreplyFlag+1 where FInterID="&FItemid
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateENreply" then
	'工程回复
	if Instr(session("AdminPurviewFLW"),"|101.7,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_Icmo where FInterID = '"&FItemid&"' "
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_Icmo (FInterID,ENReplyer,ENreplyDate,ENReplyText,ENreplyFlag) values ("&FItemid&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_Icmo set ENReplyer='"&AdminName&"',ENreplyDate='"&Now()&"',ENReplyText='"&ReplyText&"',ENreplyFlag=ENreplyFlag+1 where FInterID="&FItemid
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
'==========================1.生产工作流结束========================
'==========================2.采购单工作流开始========================
elseif key = "MtrPCreply" then
	'生管回复
	sqlajax="select * from Flw_vwICBill_26 where FInterID = '"&FItemid&"' and FEntryID = '"&FEntryID&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|102.1,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("PCReplyer")
		arryword(1)=rsajax("PCreplyDate")
		arryword(2)=rsajax("PCReplyText")
	  if Instr(session("AdminPurviewFLW"),"|102.1,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="MtrMCreply" then
	'生管回复
	sqlajax="select * from Flw_vwICBill_26 where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|102.2,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("MCReplyer")
		arryword(1)=rsajax("MCreplyDate")
		arryword(2)=rsajax("MCReplyText")
	  if Instr(session("AdminPurviewFLW"),"|102.2,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="MtrSTreply" then
	'仓库回复
	sqlajax="select * from Flw_vwICBill_26 where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|102.3,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("STReplyer")
		arryword(1)=rsajax("STreplyDate")
		arryword(2)=rsajax("STReplyText")
	  if Instr(session("AdminPurviewFLW"),"|102.3,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="MtrSEreply" then
	'仓库回复
	sqlajax="select * from Flw_vwICBill_26 where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|102.4,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("SEReplyer")
		arryword(1)=rsajax("SEreplyDate")
		arryword(2)=rsajax("SEReplyText")
	  if Instr(session("AdminPurviewFLW"),"|102.4,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="updateMtrPCreply" then
	'仓库回复
	if Instr(session("AdminPurviewFLW"),"|102.1,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_vwICBill_26 where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_vwICBill_26 (FInterID,FEntryID,PCReplyer,PCreplyDate,PCReplyText,PCreplyFlag) values ("&FItemid&","&FEntryID&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_vwICBill_26 set PCReplyer='"&AdminName&"',PCreplyDate='"&Now()&"',PCReplyText='"&ReplyText&"',PCreplyFlag=PCreplyFlag+1 where FInterID="&FItemid&" and FEntryID="&FEntryID
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateMtrMCreply" then
	'仓库回复
	if Instr(session("AdminPurviewFLW"),"|102.2,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_vwICBill_26 where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_vwICBill_26 (FInterID,FEntryID,MCReplyer,MCreplyDate,MCReplyText,MCreplyFlag) values ("&FItemid&","&FEntryID&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_vwICBill_26 set MCReplyer='"&AdminName&"',MCreplyDate='"&Now()&"',MCReplyText='"&ReplyText&"',MCreplyFlag=MCreplyFlag+1 where FInterID="&FItemid&" and FEntryID="&FEntryID
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateMtrSTreply" then
	'仓库回复
	if Instr(session("AdminPurviewFLW"),"|102.3,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_vwICBill_26 where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_vwICBill_26 (FInterID,FEntryID,STReplyer,STreplyDate,STReplyText,STreplyFlag) values ("&FItemid&","&FEntryID&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_vwICBill_26 set STReplyer='"&AdminName&"',STreplyDate='"&Now()&"',STReplyText='"&ReplyText&"',STreplyFlag=STreplyFlag+1 where FInterID="&FItemid&" and FEntryID="&FEntryID
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateMtrSEreply" then
	'仓库回复
	if Instr(session("AdminPurviewFLW"),"|102.4,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_vwICBill_26 where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_vwICBill_26 (FInterID,FEntryID,SEReplyer,SEreplyDate,SEReplyText,SEreplyFlag) values ("&FItemid&","&FEntryID&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',1)"
			connzxpt.Execute(sqlajax)
		else
			sqlajax="update Flw_vwICBill_26 set SEReplyer='"&AdminName&"',SEreplyDate='"&Now()&"',SEReplyText='"&ReplyText&"',SEreplyFlag=SEreplyFlag+1 where FInterID="&FItemid&" and FEntryID="&FEntryID
			connzxpt.Execute(sqlajax)
		end if
		Response.Write(ReplyText&"###")
		rsajax.close
		set rsajax=nothing
	end if
'==========================2.采购单工作流结束========================
'==========================3.6S工作流开始========================
elseif key ="T7reply" then
	'改善对策
	sqlajax="select * from t_5sEntry where FEntryID = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connk3,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|103.1,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("FText11")
		arryword(1)=rsajax("FDate2")
		arryword(2)=rsajax("FText7")
	  if Instr(session("AdminPurviewFLW"),"|103.1,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="T8reply" then
	'改善结果确认
	sqlajax="select * from t_5sEntry where FEntryID = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connk3,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|103.2,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("FText11")
		arryword(1)=rsajax("FDate2")
		arryword(2)=rsajax("FText7")
	  if Instr(session("AdminPurviewFLW"),"|103.2,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="T9reply" then
	'6S分录结案
	sqlajax="select * from t_5sEntry where FEntryID = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connk3,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|103.3,")>0 then
		response.Write("#########1")
	  else
	    response.Write("#########0")
	  end if
	else
		arryword(0)=rsajax("FText11")
		arryword(1)=rsajax("FDate2")
		arryword(2)=rsajax("FText7")
	  if Instr(session("AdminPurviewFLW"),"|103.3,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="updateT7reply" then
	'改善对策
	if Instr(session("AdminPurviewFLW"),"|103.1,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from t_5sEntry where FEntryID = "&FItemid
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connk3,0,1
		if rsajax.bof and rsajax.eof then
			Response.Write("@@@")
		else
		  dim temflag
		  temflag=0
		  if Depart="KD01.0001.0001" and rsajax("FComboBox1")="人资部" then
		    temflag=1
		  elseif Depart="KD01.0001.0002" and rsajax("FComboBox1")="工程部" then
		    temflag=1
		  elseif Depart="KD01.0001.0003" and rsajax("FComboBox1")="采购部" then
		    temflag=1
		  elseif Depart="KD01.0005.0004" and rsajax("FComboBox1")="营销部" then
		    temflag=1
		  elseif Depart="KD01.0001.0005" and rsajax("FComboBox1")="生技部" then
		    temflag=1
		  elseif Depart="KD01.0001.0006" and rsajax("FComboBox1")="仓储科" then
		    temflag=1
		  elseif Depart="KD01.0001.0007" and rsajax("FComboBox1")="二分厂" then
		    temflag=1
		  elseif Depart="KD01.0001.0008" and rsajax("FComboBox1")="三分厂" then
		    temflag=1
		  elseif Depart="KD01.0001.0009" and rsajax("FComboBox1")="财务部" then
		    temflag=1
		  elseif Depart="KD01.0001.0010" and rsajax("FComboBox1")="一分厂" then
		    temflag=1
		  elseif Depart="KD01.0001.0011" and rsajax("FComboBox1")="品保部" then
		    temflag=1
		  elseif Depart="KD01.0001.0017" and rsajax("FComboBox1")="生管部" then
		    temflag=1
		  elseif (Depart="KD01.0004.0001" or Depart="KD01.0004.0002") and rsajax("FComboBox1")="娄桥办" then
		    temflag=1
		  elseif Depart="KD01.0001.0012" and rsajax("FComboBox1")="总经办" then
		    temflag=1
		  end if
		  if temflag=1 then
			sqlajax="update t_5sEntry set FText11='"&AdminName&"',FDate2='"&Now()&"',FText7='"&ReplyText&"' where FEntryID="&FItemid
			connk3.Execute(sqlajax)
		  end if
			Response.Write(ReplyText&"###")
		end if
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateT8reply" then
	'改善结果确认
	if Instr(session("AdminPurviewFLW"),"|103.2,")=0 then
	  Response.Write("@@@")
	else
		sqlajax="select * from t_5sEntry where FEntryID = "&FItemid
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connk3,0,1
		if rsajax.bof and rsajax.eof then
			Response.Write("@@@")
		else
			sqlajax="update t_5sEntry set FText8='Y' where FEntryID="&FItemid
			connk3.Execute(sqlajax)
			Response.Write("Y###")
		end if
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateT9reply" then
	'6s分录结案
	if Instr(session("AdminPurviewFLW"),"|103.3,")=0 then
	  Response.Write("@@@")
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from t_5sEntry where FEntryID = "&FItemid
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connk3,0,1
		if rsajax.bof and rsajax.eof then
			Response.Write("@@@")
		else
			sqlajax="update t_5sEntry set FText9='Y' where FEntryID="&FItemid
			connk3.Execute(sqlajax)
			Response.Write("Y###")
		end if
		rsajax.close
		set rsajax=nothing
	end if
'==========================3.6S工作流结束========================
'==========================4.出货样工作流开始========================
elseif key ="SHDSPLreply" then
	'供货人确认
	sqlajax="select * from Flw_SamplesHandover where FInterID = '"&FItemid&"' and FEntryID = '"&FEntryID&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|104.1,")>0 then
		response.Write("############1")
	  else
	    response.Write("############0")
	  end if
	else
		arryword(0)=rsajax("SampleReplyer")
		arryword(1)=rsajax("SamplereplyDate")
		arryword(2)=rsajax("SampleReplyText")
		arryword(3)=rsajax("SampleNum")
	  if Instr(session("AdminPurviewFLW"),"|103.1,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###"&arryword(3)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###"&arryword(3)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="SHDQCreply" then
	'品保确认
	sqlajax="select * from Flw_SamplesHandover where FInterID = '"&FItemid&"' and FEntryID = '"&FEntryID&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|104.2,")>0 then
		response.Write("############1")
	  else
	    response.Write("############0")
	  end if
	else
		arryword(0)=rsajax("QCReplyer")
		arryword(1)=rsajax("QCreplyDate")
		arryword(2)=rsajax("QCReplyText")
		arryword(3)=rsajax("QCCheckPoint")
	  if Instr(session("AdminPurviewFLW"),"|104.2,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###"&arryword(3)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###"&arryword(3)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="SHDMCreply" then
	'生管确认
	sqlajax="select * from Flw_SamplesHandover where FInterID = '"&FItemid&"' and FEntryID = '"&FEntryID&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|104.3,")>0 then
		response.Write("############1")
	  else
	    response.Write("############0")
	  end if
	else
		arryword(0)=rsajax("MCReplyer")
		arryword(1)=rsajax("MCreplyDate")
		arryword(2)=rsajax("MCReplyText")
		arryword(3)=""
	  if Instr(session("AdminPurviewFLW"),"|104.3,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###"&arryword(3)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###"&arryword(3)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="SHDSEreply" then
	'业务确认
	sqlajax="select * from Flw_SamplesHandover where FInterID = '"&FItemid&"' and FEntryID = '"&FEntryID&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
	  if Instr(session("AdminPurviewFLW"),"|104.4,")>0 then
		response.Write("############1")
	  else
	    response.Write("############0")
	  end if
	else
		arryword(0)=rsajax("SEReplyer")
		arryword(1)=rsajax("SEreplyDate")
		arryword(2)=rsajax("SEReplyText")
		arryword(3)=""
	  if Instr(session("AdminPurviewFLW"),"|104.4,")>0 then
	  '编辑状态
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###"&arryword(3)&"###1")
	  else
	  '查看状态
	    response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###"&arryword(3)&"###0")
	  end if
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="updateSHDSPLreply" then
	'仓库回复
	if Instr(session("AdminPurviewFLW"),"|104.1,")=0 then
	  Response.Write("您没有权限进行此操作！")
	  Response.End()
	else
		ReplyText=request.QueryString("ReplyText")
		OtherEle=request.QueryString("OtherEle")
		sqlajax="select * from Flw_SamplesHandover where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
			sqlajax="insert into Flw_SamplesHandover (FInterID,FEntryID,SampleReplyer,SamplereplyDate,SampleReplyText,SampleNum,replyFlag) values ("&FItemid&","&FEntryID&",'"&AdminName&"','"&Now()&"','"&ReplyText&"',"&OtherEle&",1)"
			connzxpt.Execute(sqlajax)
		else
		  if rsajax("replyFlag")>1 then
		    Response.Write("当前状态不允许保存，请联系管理员！")
			Response.End()
		  else
			sqlajax="update Flw_SamplesHandover set SampleReplyer='"&AdminName&"',SamplereplyDate='"&Now()&"',SampleReplyText='"&ReplyText&"',SampleNum="&OtherEle&",replyFlag=1 where FInterID="&FItemid&" and FEntryID="&FEntryID
			connzxpt.Execute(sqlajax)
		  end if
		end if
		Response.Write(OtherEle&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateSHDQCreply" then
	'品保确认
	if Instr(session("AdminPurviewFLW"),"|104.2,")=0 then
	  Response.Write("您没有权限进行此操作！")
	  Response.End()
	else
		ReplyText=request.QueryString("ReplyText")
		OtherEle=request.QueryString("OtherEle")
		sqlajax="select * from Flw_SamplesHandover where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
		    Response.Write("尚未提供出货样，无需确认！")
			Response.End()
		else
		  if rsajax("replyFlag")>2 then
		    Response.Write("当前状态不允许保存，请联系管理员！")
			Response.End()
		  else
			sqlajax="update Flw_SamplesHandover set QCReplyer='"&AdminName&"',QCreplyDate='"&Now()&"',QCReplyText='"&ReplyText&"',QCCheckPoint='"&OtherEle&"',replyFlag=2 where FInterID="&FItemid&" and FEntryID="&FEntryID
			connzxpt.Execute(sqlajax)
		  end if
		end if
		Response.Write(AdminName&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateSHDMCreply" then
	'生管确认
	if Instr(session("AdminPurviewFLW"),"|104.3,")=0 then
	  Response.Write("您没有权限进行此操作！")
	  Response.End()
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_SamplesHandover where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
		    Response.Write("尚未提供出货样，无需确认！")
			Response.End()
		else
		  if rsajax("replyFlag")>3 then
		    Response.Write("当前状态不允许保存，请联系管理员！")
			Response.End()
		  else
			sqlajax="update Flw_SamplesHandover set MCReplyer='"&AdminName&"',MCreplyDate='"&Now()&"',MCReplyText='"&ReplyText&"',replyFlag=3 where FInterID="&FItemid&" and FEntryID="&FEntryID
			connzxpt.Execute(sqlajax)
		  end if
		end if
		Response.Write(AdminName&"###")
		rsajax.close
		set rsajax=nothing
	end if
elseif key ="updateSHDSEreply" then
	'业务确认
	if Instr(session("AdminPurviewFLW"),"|104.4,")=0 then
	  Response.Write("您没有权限进行此操作！")
	  Response.End()
	else
		ReplyText=request.QueryString("ReplyText")
		sqlajax="select * from Flw_SamplesHandover where FInterID = '"&FItemid&"'  and FEntryID = '"&FEntryID&"'"
		set rsajax=server.createobject("adodb.recordset")
		rsajax.open sqlajax,connzxpt,0,1
		if rsajax.bof and rsajax.eof then
		    Response.Write("尚未提供出货样，无需确认！")
			Response.End()
		else
		  if rsajax("replyFlag")>4 then
		    Response.Write("当前状态不允许保存，请联系管理员！")
			Response.End()
		  else
			sqlajax="update Flw_SamplesHandover set SEReplyer='"&AdminName&"',SEreplyDate='"&Now()&"',SEReplyText='"&ReplyText&"',replyFlag=4 where FInterID="&FItemid&" and FEntryID="&FEntryID
			connzxpt.Execute(sqlajax)
		  end if
		end if
		Response.Write(AdminName&"###")
		rsajax.close
		set rsajax=nothing
	end if
end if
%>