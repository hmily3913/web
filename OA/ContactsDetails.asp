<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|306,")=0 then 
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
      datafrom=" t_Base_Emp,LDERP.dbo.[N-基本资料单头] "
  dim datawhere'数据条件
	datawhere=" where FNumber=员工代号 and 离职否='在职' and (isnull(Femail,'') <> '' or isnull(FPhone,'')<>'' or isnull(shortMobile,'')<>'') "
	Dim searchterm,searchcols
	if Request.Form("FDepartmentID") <> "" then
		datawhere=datawhere&" and 二级部门='"&Request.Form("FDepartmentID")&"' "
	End if
	if Request.Form("KeyWork") <> "" then datawhere=datawhere&" and FNumber like '%"&Request.Form("KeyWork")&"%' "

  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "二级部门,FNumber" 
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
  rs.open sql,connk3,0,1
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
    sql="select FitemID from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("FitemID")
	  else
	    sqlid=sqlid &","&rs("FitemID")
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
    sql="select * from "& datafrom &" where FNumber=员工代号 and 离职否='在职' and FitemID in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    do until rs.eof'填充数据到表格
%>		
		{"id":"<%=rs("FitemID")%>",
		"cell":["<%=rs("FitemID")%>","<%=rs("二级部门")%>","<%=rs("FNumber")%>","<%=rs("FName")%>","<%=rs("FPhone")%>","<%=rs("FMobilePhone")%>","<%=rs("shortMobile")%>","<%=rs("Femail")%>"]}
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
  if detailType="AddFile" then
  	if  Instr(session("AdminPurviewFLW"),"|507.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
		sql="select * from t_Base_Emp where FItemid="&Request("FItemid")
		rs.open sql,connk3,1,3
		rs("FPhone")=Request("FPhone")
		rs("shortMobile")=Request("shortMobile")
		rs("FMobilePhone")=Request("FMobilePhone")
		rs("Femail")=Request("Femail")
		rs.update
		rs.close
		set rs=nothing 
		response.write "保存成功！"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="FNumber" then
		sql="select FNumber,FName,FItemid,FPhone,FMobilePhone,shortMobile,Femail from t_Base_Emp where FNumber='"&request("InfoID")&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
			if rs.bof and rs.eof then
					response.write ("对应员工不存在，请检查！")
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
    set rs = server.createobject("adodb.recordset")
    sql="select distinct c.部门代号 as pk,c.部门名称 as name,'0' as pspk from [N-基本资料单头] a,[G-部门资料表] c  where a.离职否='在职' and a.部门别=c.部门代号 union all select distinct a.二级部门 as pk,c.部门名称 as name,a.部门别 as pspk from [N-基本资料单头] a,[G-部门资料表] c  where a.离职否='在职' and a.二级部门=c.部门代号 and a.二级部门<>a.部门别 order by pk"
    rs.open sql,conn,1,1
		response.Write("[")
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("pk")%>", "name": "<%=rs("name")%>", "PSNum": "<%=rs("pspk")%>"}
	<%
			rs.movenext
			If Not rs.eof Then
			  Response.Write ","
			End If
		loop
		Response.Write "]"
		rs.close
		set rs=nothing
	elseif detailType="getModel" then
		sql="select 员工代号,姓名 from [N-基本资料单头] where 离职否='在职' "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,0,1
		response.write "["
		do until rs.eof
		Response.Write("{""FNumber"":"""&rs("员工代号")&""",""FName"":"""&JsonStr(rs("姓名"))&"""}")
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
