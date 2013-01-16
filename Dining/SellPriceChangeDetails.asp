<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|602,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
Depart=session("Depart")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="DetailsList" then 
  dim page'页码
      page=clng(request("page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=request("rp")
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Dining_SellPriceChange a,Dining_Material b "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where a.SellId=b.SerialNum "
	Dim searchterm,searchcols
	
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = Request.Form("qtype")
	if isnumeric(searchterm) then
	datawhere = datawhere&" and " & searchcols & " = " & searchterm & " "
	else
	datawhere = datawhere&" and " & searchcols & " LIKE '%" & searchterm & "%' "
	end if
	End if
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "a.SerialNum" 
	Else
	sortname = Request.Form("sortname")
	End If
	Dim sortorder
	if Request.Form("sortorder") = "" then
	sortorder = "desc"
	Else
	sortorder = Request.Form("sortorder")
	End If
      taxis=" order by "&sortname&" "&sortorder
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
  idCount=rs("idCount")
  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
  end if
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select a.SerialNum from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("SerialNum")
	  else
	    sqlid=sqlid &","&rs("SerialNum")
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
'-----------------------------------------------------------
'-----------------------------------------------------------
%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    sql="select a.*,b.MaterialName,b.Unit from "& datafrom &" where a.SellId=b.SerialNum and a.SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格'
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("ChangeDate")%>","<%=rs("SellId")%>","<%=rs("MaterialName")%>","<%=rs("Unit")%>","<%=rs("OldPrice")%>","<%=rs("NewPrice")%>","<%=rs("CheckFlag")%>","<%=rs("Remark")%>","<%=rs("Biller")%>"]}
<%		
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
  end if
  rs.close
  set rs=nothing
response.Write"]}"
'-----------------------------------------------------------'
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
  	if  Instr(session("AdminPurviewFLW"),"|602.1,")>0 then
			for   i=2   to   Request.form("SerialNum").count
				if Request.form("DeleteFlag")(i)="0" and Request.form("SellId")(i)<>"" then
					connzxpt.Execute("insert into Dining_SellPriceChange (ChangeDate,SellId,OldPrice,NewPrice,Remark,BillerID,Biller,BillDate) values ('"&Request.form("ChangeDate")(i)&"','"&Request.form("SellId")(i)&"','"&Request.form("OldPrice")(i)&"','"&Request.form("NewPrice")(i)&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"')")
				end if
			next
			response.write "###"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
		if Instr(session("AdminPurviewFLW"),"|602.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("Delete from Dining_SellPriceChange where SerialNum="&Request.Form("SerialNum")(i))
			elseif Request.Form("SerialNum")(i)<>"" then
				connzxpt.Execute("update Dining_SellPriceChange set ChangeDate='"&Request.form("ChangeDate")(i)&"',SellId='"&Request.form("SellId")(i)&"',OldPrice='"&Request.form("OldPrice")(i)&"',NewPrice='"&Request.form("NewPrice")(i)&"',Remark='"&Request.form("Remark")(i)&"',BillerID='"&UserName&"',Biller='"&AdminName&"',BillDate='"&now()&"' where SerialNum="&Request.Form("SerialNum")(i))
			else
				connzxpt.Execute("insert into Dining_SellPriceChange (ChangeDate,SellId,OldPrice,NewPrice,Remark,BillerID,Biller,BillDate) values ('"&Request.form("ChangeDate")(i)&"','"&Request.form("SellId")(i)&"','"&Request.form("OldPrice")(i)&"','"&Request.form("NewPrice")(i)&"','"&Request.form("Remark")(i)&"','"&UserName&"','"&AdminName&"','"&now()&"')")
			end if
		next
		response.write "###"
  elseif detailType="Delete" then
		if Instr(session("AdminPurviewFLW"),"|602.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
    SerialNum=request("SerialNum")
		sql="select * from Dining_SellPriceChange where CheckFlag=1 and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from Dining_SellPriceChange where SerialNum in ("&SerialNum&")")
			response.write "###"
		else
			response.Write("已审核不允许删除！")
			response.End()
		end if
  elseif detailType="Check"  then
		SerialNum=request("SerialNum")
		if Instr(session("AdminPurviewFLW"),"|602.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		sql="select * from Dining_SellPriceChange where SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if rs("CheckFlag")=0 then
				rs("Checker")=AdminName
				rs("CheckerID")=UserName
				rs("CheckDate")=now()
				rs("CheckFlag")=1
				connzxpt.Execute("update Dining_Material set Price="&rs("NewPrice")&",UseFlag=1 where SerialNum ="&rs("SellId"))
			end if		
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="unCheck"  then
		if Instr(session("AdminPurviewFLW"),"|602.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		SerialNum=request("SerialNum")
		connzxpt.Execute("update Dining_SellPriceChange set CheckFlag=0 where SerialNum in ("&SerialNum&")")
		response.write "###"
  end if
elseif showType="getInfo" then 
	sql="select SerialNum,MaterialName,Unit,Price from Dining_Material where Type='菜品' and ForbidFlag=0 "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	response.write "["
	do until rs.eof
	Response.Write("{")
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
	Response.Write("]")
	rs.close
	set rs=nothing 
end if
 %>
