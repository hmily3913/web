<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|404,")=0 then 
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
      datafrom=" parametersys_PMProductCycle a, "&AllOPENROWSET&" AIS20081217153921.dbo.t_item) as b "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where b.FItemClassId=4 and b.FDetail=0 and a.ProductTypeId=b.FItemid "
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
		searchterm = Request.Form("query")
		searchcols = Request.Form("qtype")
		datawhere = datawhere&" and "&searchcols&" like '%"&searchterm&"%' "
	End if
	
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "SerialNum" 
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
    sql="select SerialNum from "& datafrom &" " & datawhere & taxis
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
    sql="select a.*,b.FNumber,b.FName,b.FFullName from "& datafrom &" "& datawhere &"and a.SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("ProductTypeId")%>","<%=rs("FNumber")%>","<%=rs("FName")%>","<%=rs("FFullName")%>","<%=rs("ProductCycle")%>","<%=rs("MonthCapacity")%>","<%=rs("OrderCycle")%>"]}
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
'-----------------------------------------------------------
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
  	if  Instr(session("AdminPurviewFLW"),"|404.1,")>0 then
		set rs = server.createobject("adodb.recordset")
		sql="select * from parametersys_PMProductCycle"
		rs.open sql,connzxpt,1,3
		rs.addnew
		rs("ProductTypeId")=Request.Form("ProductTypeId")
		rs("ProductCycle")=Request.Form("ProductCycle")
		rs("MonthCapacity")=Request.Form("MonthCapacity")
		rs("OrderCycle")=Request.Form("OrderCycle")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
			SerialNum=request("SerialNum")
		sql="select * from parametersys_PMProductCycle where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|404.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		rs("ProductTypeId")=Request.Form("ProductTypeId")
		rs("ProductCycle")=Request.Form("ProductCycle")
		rs("MonthCapacity")=Request.Form("MonthCapacity")
		rs("OrderCycle")=Request.Form("OrderCycle")
		response.write "###"
		rs.update
		rs.close
		set rs=nothing 
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from parametersys_PMProductCycle where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|404.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from parametersys_PMProductCycle where SerialNum="&SerialNum)
	response.write "###"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="SearchData" then
    InfoID=request("InfoID")
	sql="select FItemId,FNumber,FName,FFullName from t_item where FItemClassId=4 and FDetail=0 order by FNUmber asc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
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
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
	sql="select a.*,b.FNumber,b.FName,b.FFullName from parametersys_PMProductCycle a, "&AllOPENROWSET&" AIS20081217153921.dbo.t_item) as b where b.FItemClassId=4 and b.FDetail=0 and a.ProductTypeId=b.FItemid and a.SerialNum="&InfoID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
    if rs.bof and rs.eof then
        response.write ("对应单据不存在，请检查！")
        response.end
	else
	  response.write "{""Info"":""###"",""fieldValue"":[{"
	  for i=0 to rs.fields.count-2
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
		else
		response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
		end if
	  next
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}]}")
		else
		response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}]}")
		end if
	end if
	rs.close
	set rs=nothing 
  end if
end if
 %>
