<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|308,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
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
      datafrom=" Offer_Moju "
  dim datawhere'数据条件
	datawhere=" where 1=1 "
	Dim searchterm,searchcols
	if Request.Form("PositionID") <> "" then
		datawhere=datawhere&" and PositionID="&Request.Form("PositionID")
	End if
	if Request.Form("KeyWork") <> "" then datawhere=datawhere&" and Moju like '%"&Request.Form("KeyWork")&"%' "

  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "PositionID,OrderId" 
	Else
	sortname = Request.Form("sortname")
	End If
	Dim sortorder
	if Request.Form("sortorder") = "" then
	sortorder = "asc"
	Else
	sortorder = Request.Form("sortorder")
	End If
      taxis=" order by "&sortname&" "&sortorder
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
'	response.Write(sql)
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
    sql="select * from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("PositionID")%>","<%=rs("OrderId")%>","<%=rs("Moju")%>","<%=rs("Shuliang")%>","<%=Jsonstr(rs("Remark"))%>","<%=rs("Biller")%>","<%=rs("BillDate")%>"]}
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
  	if  Instr(session("AdminPurviewFLW"),"|308.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
		sql="select * from Offer_ModeLayer"
		rs.open sql,connzxpt,1,3
		rs.addnew
		rs("PositionName")=Request("PositionName")
		rs("PSNum")=Request("TypePsnum")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
		set rs = server.createobject("adodb.recordset")
		sql="select top 1 * from Offer_ModeLayer order by serialnum desc"
		rs.open sql,connzxpt,1,1
		response.write("[{""SerialNum"": """&rs("SerialNum")&""", ""name"": """&rs("PositionName")&""", ""PSNum"": """&rs("PSNum")&"""}]")
		rs.close
		set rs=nothing 
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from Offer_ModeLayer where SerialNum="&Request("PositionID")
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|308.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
			end if
'		if rs("CheckFlag")>0 then
'			response.write ("当前状态不允许编辑，请检查！")
'			response.end
'		end if
		rs("PositionName")=Request("PositionName")
		rs("PSNum")=Request("TypePsnum")
		response.write "###"
    rs.update
		rs.close
		set rs=nothing 
  elseif detailType="Delete" then
		if Instr(session("AdminPurviewFLW"),"|308.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from Offer_ModeLayer where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		rs.close
		set rs=nothing 
		connzxpt.Execute("Delete from Offer_ModeLayer where SerialNum="&SerialNum)
		response.write "###"
  elseif detailType="AddFile" then
  	if  Instr(session("AdminPurviewFLW"),"|308.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
		sql="select * from Offer_Moju"
		rs.open sql,connzxpt,1,3
		rs.addnew
		rs("Moju")=Request("Moju")
		rs("OrderId")=Request("OrderId")
		rs("PositionID")=Request("PositionIDDetail")
		rs("Remark")=Request("Remark")
		rs("Shuliang")=Request("Shuliang")
		rs("Biller")=AdminName
		rs("BillerID")=UserName
		rs("BillDate")=now()
		rs.update
		rs.close
		set rs=nothing 
		response.write "保存成功！"
  elseif detailType="EditFile" then
  	if  Instr(session("AdminPurviewFLW"),"|308.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
		sql="select * from Offer_Moju where SerialNum="&request("DocumentID")
		rs.open sql,connzxpt,1,3
		rs("Moju")=Request("Moju")
		rs("OrderId")=Request("OrderId")
		rs("PositionID")=Request("PositionIDDetail")
		rs("Remark")=Request("Remark")
		rs("Shuliang")=Request("Shuliang")
		rs("Biller")=AdminName
		rs("BillerID")=UserName
		rs("BillDate")=now()
		rs.update
		rs.close
		set rs=nothing 
		response.write "保存成功！"
  elseif detailType="DeleteFile" then
  	if  Instr(session("AdminPurviewFLW"),"|308.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		connzxpt.Execute("Delete from Offer_Moju where SerialNum="&request("SerialNum"))
		response.write "删除成功！"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from Offer_Moju where SerialNum="&InfoID
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
	elseif detailType="DocumentType" then
    dim tempstr1
		if request("SerialNum")="" then
			tempstr1=" psnum=0"
		else
			tempstr1=" psnum="&request("SerialNum")
		end if
    set rs = server.createobject("adodb.recordset")
    sql="select * from Offer_ModeLayer order by SerialNum"
    rs.open sql,connzxpt,1,1
		response.Write("[")
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("SerialNum")%>", "name": "<%=rs("PositionName")%>", "PSNum": "<%=rs("PSNum")%>"
	<%
		set rs2=server.createobject("adodb.recordset")
		sql2="select count(1) as idcount from Offer_ModeLayer where PSNum="&rs("SerialNum")
		rs2.open sql2,connzxpt,1,1
		if rs2("idcount")>0 then
	%>
	, isParent:true
	<%
		end if
		rs2.close
		set rs2=nothing 
		response.Write("}")
			rs.movenext
			If Not rs.eof Then
			  Response.Write ","
			End If
		loop
		Response.Write "]"
		rs.close
		set rs=nothing
	elseif detailType="getModel" then
		sql="select Moju from Offer_Moju "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.write "["
		do until rs.eof
		Response.Write("{""Moju"":"""&JsonStr(rs("Moju"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
  end if
end if
 %>
