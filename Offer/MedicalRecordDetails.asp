<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurviewFLW"),"|305,")=0 then 
  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
  response.end
end if
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
      datafrom=" Offer_MedicalRecord a,"&AllOPENROWSET&"LDERP.dbo.[N-基本资料单头]) b "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where a.EmpId=b.员工代号 "
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
    sql="select a.SerialNum,a.EmpId,姓名,职等,部门别,工作岗位,性别,婚姻状况,出生日期,datediff(yy,出生日期,getdate()) as 年龄,到职日,MedicalDate,MedicalType,身份证号,户籍地址 from "& datafrom &" where a.EmpId=b.员工代号 and a.SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格'
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":[
<%		
	  for i=0 to rs.fields.count-2
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).value&""",")
		else
		response.write (""""&JsonStr(rs.fields(i).value)&""",")
		end if
	  next
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).value&"""]}")
		else
		response.write (""""&JsonStr(rs.fields(i).value)&"""]}")
		end if

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
  	if  Instr(session("AdminPurviewFLW"),"|305.1,")>0 then
		set rs = server.createobject("adodb.recordset")
		sql="select * from Offer_MedicalRecord"
		rs.open sql,connzxpt,1,3
		rs.addnew
		for i=0 to rs.fields.count-1
		  rs.fields(i).value=Request(rs.fields(i).name)
		next

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
	sql="select * from Offer_MedicalRecord where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|305.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
    end if
		for i=0 to rs.fields.count-1
			rs.fields(i).value=Request(rs.fields(i).name)
		next
		response.write "###"
    rs.update
	rs.close
	set rs=nothing 
  elseif detailType="Delete" then
		if Instr(session("AdminPurviewFLW"),"|305.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
    SerialNum=request("SerialNum")
		connzxpt.Execute("Delete from Offer_MedicalRecord where SerialNum in ("&SerialNum&")")
		response.write "###"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="EmpId" or detailType="n1" then
    InfoID=request("InfoID")
	sql="select 姓名 n1,职等 n3,部门别 n2,工作岗位 as n4,性别 as n5,婚姻状况 as n6,出生日期 as n7,datediff(yy,出生日期,getdate()) as n8,到职日 as n9,身份证号 as n10 from [N-基本资料单头] where 员工代号 like '%"&InfoID&"%' or 姓名 like '%"&InfoID&"%' "
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
    if rs.bof and rs.eof then
        response.write ("对应员工不存在！")
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
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
	sql="select a.SerialNum,a.EmpId,姓名 n1,职等 n3,部门别 n2,工作岗位 as n4,性别 as n5,婚姻状况 as n6,出生日期 as n7,datediff(yy,出生日期,getdate()) as n8,到职日 as n9,MedicalDate,MedicalType,身份证号 as n10 from Offer_MedicalRecord a,"&AllOPENROWSET&"LDERP.dbo.[N-基本资料单头]) b where a.EmpId=b.员工代号 and a.SerialNum ="&InfoID
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
elseif showType="xls2sql" then
	if Instr(session("AdminPurviewFLW"),"|305.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
	end if
	Server.ScriptTimeout = 999999
	set rs=server.createobject("adodb.recordset")
	sql="select * from Offer_MedicalRecord"
	rs.open sql,connzxpt,3,3
	InfoID=request("InfoID")
	Set xlApp=Server.CreateObject("Excel.Application")          '/******** VBA方法 连接Excel *********/
	Set xlbook=xlApp.Workbooks.Open(Server.mappath(InfoID))  
	Set xlsheet=xlbook.Worksheets(1)  
	i=2
	While cstr(xlsheet.cells(i,1))<>""           '/********** 使用第3列 帐号为空时判断为结束标志  **********/
	
	rs.Addnew()  
	rs("EmpID")=cstr(xlsheet.cells(i,1))
	rs("MedicalDate")=cdate(xlsheet.cells(i,2))
	rs("MedicalType")=cstr(xlsheet.cells(i,3))
	rs.Update  
	
	i=i+1  
	Wend  
	xlsheet.close
	Set xlsheet=nothing  
	xlbook.Close  
	Set xlbook=Nothing  
	xlApp.DisplayAlerts=false
	xlApp.Quit  
	
	rs.Close
	Set rs=Nothing
	response.Write("共计"&i-1&"个体检记录导入成功!")
end if
 %>
