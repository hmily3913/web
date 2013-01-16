<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
dim FItemid,key
dim rsajax,sqlajax
dim arryword(5)
key=request.QueryString("key")
FItemid=request.QueryString("FItemid")

if key = "emp" then
	'获取员工信息
	sqlajax="select * from t_Emp where Fnumber = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connk3,0,1
	if rsajax.bof and rsajax.eof then
		response.Write("!@#$")
	else
		arryword(0)=rsajax("FItemid")
		arryword(1)=rsajax("Fnumber")
		arryword(2)=rsajax("Fname")
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2))
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="department" then
	'获取部门信息
	dim wordforout
	wordforout=""
	sqlajax="select * from t_item where (Fnumber like '%"&FItemid&"%' or  fname like '%"&FItemid&"%') and fitemclassid=2"
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connk3,0,1
	while(not rsajax.eof)
		arryword(0)=rsajax("FItemid")
		arryword(1)=rsajax("Fname")
		wordforout=wordforout&arryword(0)&"###"&arryword(1)&"@@@"
	  rsajax.movenext
    wend
	response.Write(wordforout)
	rsajax.close
	set rsajax=nothing
elseif key ="DPdeleteDetail" then
	'删除子表明细
	sqlajax="update t_dormperson set sumperson=sumperson-1 where fid in (select fid from t_dormpersonEntry where FEntryID="&FItemid&")"
	connk3.Execute(sqlajax)
	sqlajax="delete from t_dormpersonEntry where FEntryID="&FItemid
	connk3.Execute(sqlajax)
	response.Write("###")
elseif key ="deleteDP" then
	'删除宿舍员工信息
	sqlajax="delete from t_dormperson where Fid="&FItemid
	connk3.Execute(sqlajax)
	sqlajax="delete from t_dormpersonEntry where Fid="&FItemid
	connk3.Execute(sqlajax)
	response.Write("###")
elseif key ="checkDorm" then
	'检查宿舍号
'	sqlajax="select * from t_dormperson where ftext='"&FItemid&"'"
    set rsajax = connk3.execute("select * from t_dormperson where ftext='"&FItemid&"'")
    if not (rsajax.bof and rsajax.eof) then '判断此产品编号是否存在
	response.Write("###")
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="deleteEW" then
	'删除宿舍员工信息
	dim tempFid
	sqlajax="select fid from t_electwaterEntry where Fentryid="&FItemid
	set rsajax=connk3.execute(sqlajax)
	tempFid=rsajax("fid")
	sqlajax="delete from t_electwaterEntry where Fentryid="&FItemid
	connk3.Execute(sqlajax)
	sqlajax="select * from t_electwaterEntry where fid="&tempFid
    set rsajax = connk3.execute(sqlajax)
    if rsajax.bof and rsajax.eof then '判断此产品编号是否存在
	sqlajax="delete from t_electwater where Fid="&tempFid
	connk3.Execute(sqlajax)
	end if
	rsajax.close
	set rsajax=nothing
	response.Write("###")
elseif key = "EW" then
	'获取员工信息
	sqlajax="select * from t_dormperson where ftext = '"&FItemid&"' "
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connk3,0,1
	if rsajax.bof and rsajax.eof then
		response.Write("!@#$")
	else
		arryword(0)=rsajax("ftext")
		arryword(1)=rsajax("waternum")
		arryword(2)=rsajax("electnum")
		arryword(3)=rsajax("Hotwater")
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2)&"###"&arryword(3))
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="EWdeleteDetail" then
	'删除子表明细
	sqlajax="delete from t_electwaterEntry where FEntryID="&FItemid
	connk3.Execute(sqlajax)
	response.Write("###")
elseif key = "empname" then
	'获取员工信息
	sqlajax="select * from t_Emp where Fnumber = '"&FItemid&"' or Fname = '"&FItemid&"'"
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connk3,0,1
	if rsajax.bof and rsajax.eof then
		response.Write("!@#$")
	else
		arryword(0)=rsajax("Fnumber")
		arryword(1)=rsajax("Fname")
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2))
	end if
	rsajax.close
	set rsajax=nothing
elseif key ="deleteSC" then
	'删除子表明细
	sqlajax="delete from z_SendCar where SerialNum="&FItemid
	connk3.Execute(sqlajax)
	response.Write("###")
elseif key ="GOPdeleteDetail" then
	'删除子表明细
	sqlajax="select * from z_GoodsCarryOutMain a,z_GoodsCarryOutDetails b where a.SerialNum=b.SerialNum and a.OutCheckFlag=0 and b.FEntryID="&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connk3,0,1
	if rsajax.bof and rsajax.eof then
		response.Write("!@#$")
	else
	
	sqlajax="delete from z_GoodsCarryOutDetails where FEntryID="&FItemid
	connk3.Execute(sqlajax)
	response.Write("###")
	end if
elseif key ="getDriver" then
	'获取司机信息
	sqlajax="select b.fnumber,b.fname,a.mileageNum from z_Car a,t_emp b where a.Driver=b.fname and a.CarID='"&FItemid&"'"
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connk3,0,1
	if rsajax.bof and rsajax.eof then
		response.Write("!@#$")
	else
		arryword(0)=rsajax("Fnumber")
		arryword(1)=rsajax("Fname")
		arryword(2)=rsajax("mileageNum")
		response.Write(arryword(0)&"###"&arryword(1)&"###"&arryword(2))
	end if
'=======================网络维修及时率========================
elseif key = "NRreply" then
	sqlajax="select * from managesys_Networkrepair where SerialNum = '"&FItemid&"' "
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
elseif key ="updateNRreply" then
	ReplyText=request.QueryString("ReplyText")
	sqlajax="select * from managesys_Networkrepair where SerialNum = "&FItemid
	set rsajax=server.createobject("adodb.recordset")
	rsajax.open sqlajax,connzxpt,0,1
	if rsajax.bof and rsajax.eof then
		Response.Write("@@@")
	else
		sqlajax="update managesys_Networkrepair set Replyer='"&AdminName&"',replyDate='"&Now()&"',ReplyText='"&ReplyText&"',replyFlag=replyFlag+1 where SerialNum="&FItemid
		connzxpt.Execute(sqlajax)
		Response.Write(ReplyText&"###")
	end if
	rsajax.close
	set rsajax=nothing
end if
%>