<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurview"),"|409,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")

if showType="DetailsList" then 
  dim datawhere'数据条件
	datawhere=" WHERE a.Finterid=b.Finterid and b.FItemID=e.FItemID and a.FEmpID=c.Fitemid and a.FBase=d.Fitemid and FCheckerID>0 and (FCheckBox3=1 or FCheckBox11=1 or FCheckBox1=1) "
	if request("AllQuery")<>"" then
		datawhere=datawhere&" and "&request("AllQuery")
	end if

  dim i:i=1
  dim rs,sql,idCount,page,temstr,pages
	page=clng(request("page"))
	pages=request("rp")
	temstr="###"
  sql="select count(1) as idCount  from t_DDBGTZD a,t_DDBGTZDEntry b,t_emp c,t_Organization d,t_icitem e "&datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  idCount=rs("idCount")

%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
    sql="select a.Fbillno,a.FAlterDate,b.fbillno_src,d.fname as provider,e.fnumber,e.fname as product,e.Fmodel,b.FAuxOrgQty,a.LossMoney,a.GatherMoney,c.Fname,a.FAlterReason,a.FNote,a.FNOTE1,a.FInterID from t_DDBGTZD a,t_DDBGTZDEntry b,t_emp c,t_Organization d,t_icitem e "&datawhere&_
" order by a.Finterid desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
		rs.pagesize = pages
		rs.absolutepage = page  
		for i=1 to rs.pagesize
		if rs.eof then exit for  
		if(i=1)then
		else
		  Response.Write ","
		end if
%>		
		{"id":"<%=i%>",
		"cell":["<%=rs("Fbillno")%>","<%=rs("FAlterDate")%>","<%=JsonStr(rs("provider"))%>","<%=rs("fbillno_src")%>","<%=JsonStr(rs("product"))%>","<%=JsonStr(rs("Fmodel"))%>","<%=rs("FAuxOrgQty")%>","<%=rs("Fname")%>","<%=JsonStr(rs("FAlterReason"))%>","<%=JsonStr(rs("FNote"))%>","<%=JsonStr(rs("FNOTE1"))%>","<%=rs("LossMoney")%>","<%=rs("GatherMoney")%>","<%=rs("SAFlag")%>","<%=rs("CWFlag")%>","<%=rs("FInterID")%>"]}
<%		
	    rs.movenext
    next
  rs.close
  set rs=nothing
	response.Write"]}"
elseif showType="FInterID" then
	InfoID=request("InfoID")
	sql="select a.FInterID,a.FInterID,a.Fbillno,a.FAlterDate,d.fname as provider,c.Fname,a.FAlterReason,a.FNote,a.FNOTE1,a.LossMoney,a.GatherMoney,case when FCheckBox3=1 then '订单取消' when FCheckBox1=1 then '数量变更' when FCheckBox11=1 then '订单暂停' end as typechange  from t_DDBGTZD a,t_emp c,t_Organization d WHERE a.FEmpID=c.Fitemid and a.FBase=d.Fitemid and FCheckerID>0 and (FCheckBox3=1 or FCheckBox11=1 or FCheckBox1=1) and a.FInterid="&InfoID
'	response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
	if rs.bof and rs.eof then
		response.write ("对应单据不存在，请检查！")
		response.end
	else
		response.write "{""Info"":""###"",""fieldValue"":[{"
		for i=0 to rs.fields.count-1
		if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
		else
		response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
		end if
		next
		response.write ("""Entrys"":[")
		sql="select a.*,b.FNumber,b.FName,b.FModel from t_DDBGTZDEntry a,t_icitem b where a.FItemID=b.FItemID and a.Finterid="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		do until rs.eof
			response.write "{"
			for i=0 to rs.fields.count-2
			if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
			end if
			next
			if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}")
			end if
			rs.movenext
		If Not rs.eof Then
			Response.Write ","
		End If
		loop
		response.write "]}]}"
	end if
	rs.close
	set rs=nothing 
elseif showType="DataProcess" then
	set rs=server.createobject("adodb.recordset")
	sql="select * from t_DDBGTZD where FInterid="&request("FInterid")
	rs.open sql,connk3,1,3
	if Instr(session("AdminPurview"),"|105.1,")>0 then 
		rs("LossMoney")=request("LossMoney")
		rs("MLoss")=request("MLoss")
		rs("PLoss")=request("PLoss")
		rs("PMer")=AdminName
		rs("PMID")=UserName
		rs("PMDate")=now()
		rs("PMFlag")=1
	end if
	if Instr(session("AdminPurview"),"|105.2,")>0 then 
		rs("FNote1")=request("FNote1")
		rs("SAer")=AdminName
		rs("SAID")=UserName
		rs("SADate")=now()
		rs("SAFlag")=1
	end if
	if Instr(session("AdminPurview"),"|105.3,")>0 then 
		rs("GatherMoney")=request("GatherMoney")
		rs("CWer")=AdminName
		rs("CWID")=UserName
		rs("CWDate")=now()
		rs("CWFlag")=1
	end if
	rs.update
	rs.close
	set rs=nothing
	response.Write("保存成功！")
end if
 %>
