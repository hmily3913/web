<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|214,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName,Depart
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
      datafrom=" qcsys_CustomInspect "
  dim datawhere'数据条件
    datawhere=" where 1=1 "
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
		if Request.Form("qtype")="CtrlDate" then
			datawhere = datawhere&" and datediff(d,"&searchcols&",'"&searchterm&"')=0 "
		else
			datawhere = datawhere&" and "&searchcols&" like '%"&searchterm&"%' "
		end if
	End if
	datawhere = datawhere&Session("AllMessage44")
	session.contents.remove "AllMessage44"
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
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(distinct SerialNum) as idCount from "& datafrom &" " & datawhere
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
    sql="select distinct SerialNum from "& datafrom &" " & datawhere & taxis
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
	dim ys:ys="#f7f7f7"
	if rs("CheckFlag")>0 then ys="#ffff66"
	dim CheckState:CheckState=""
	if rs("CheckFlag")="1" then
	  CheckState="品保确认"
	else
	  CheckState="待审核"
	end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("BillDate")%>","<%=rs("OrderID")%>","<%=rs("Product")%>","<%=rs("Employer")%>","<%=rs("InspectDate")%>","<%=rs("Count")%>","<%=JsonStr(rs("Result"))%>","<%=JsonStr(rs("Problem"))%>","<%=rs("Biller")%>","<%=CheckState%>","<%=JsonStr(rs("Judge"))%>"]}
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
  	if  Instr(session("AdminPurview"),"|411.1,")>0 then
			set rs = server.createobject("adodb.recordset")
			sql="select * from qcsys_CustomInspect"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
'			rs("Register")=Request("Register")
'			rs("RegisterID")=Request("RegisterID")
			rs("InspectDate")=Request("InspectDate")
			rs("OrderID")=Request("OrderID")
			rs("Employer")=Request("Employer")
			rs("Product")=Request("Product")
			rs("Remark")=Request("Remark")
			set rs2=connzxpt.Execute("select count(1) as ct from qcsys_CustomInspect where OrderID='"&Request("OrderID")&"' and Product='"&Request("Product")&"'")
			if not rs2.eof then
			rs("Count")=rs2("ct")+1
			else
			rs("Count")=1
			end if
			rs.update
			rs.close
			set rs=nothing 
			response.write "###"
		SendMail "2213@loverdoor.cn","客户验货通知",Request("OrderID"),Request("Product")&"("&AdminName&")",""
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
		response.Write("该单据不允许修改，只能删除或由品保确认！")
		response.end
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from qcsys_CustomInspect where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
'		if Instr(session("AdminPurviewFLW"),"|214.1,")=0 then
'			response.write ("你没有权限进行此操作！")
'			response.end
'			end if
		if rs("Biller")<>UserName and rs("Register")<>UserName then
			response.write ("只能编辑自己添加的数据！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			response.write ("当前状态不允许编辑，请检查！")
			response.end
		end if
		rs("Biller")=AdminName
		rs("BillerID")=UserName
		rs("BillDate")=now()
'		rs("Register")=Request("Register")
'		rs("RegisterID")=Request("RegisterID")
		rs("InspectDate")=Request("InspectDate")
		rs("OrderID")=Request("OrderID")
		rs("Employer")=Request("Employer")
		rs("Product")=Request("Product")
		rs("Remark")=Request("Remark")
    rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="Delete" then
    SerialNum=request("SerialNum")
		sql="select * from qcsys_CustomInspect where SerialNum="&SerialNum
		set rs = server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("BillerID")<>UserName and rs("RegisterID")<>UserName then
			response.write ("只能删除本人自己添加的数据！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			response.write ("已经审核不允许删除！")
			response.end
		end if
		rs.close
		set rs=nothing 
		connzxpt.Execute("Delete from qcsys_CustomInspect where SerialNum="&SerialNum)
		response.write "###"
  elseif detailType="QCCheck" then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from qcsys_CustomInspect where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")>1 then
			response.write ("此单据已经在确认，不需要审核！")
			response.end
		end if
		if Instr(session("AdminPurview"),"|411.2,")>0 then
			rs("Judge")=Request("Judge")
			rs("Result")=Request("Result")
			rs("Problem")=Request("Problem")
			rs("Treatment")=Request("Treatment")
			rs("CheckerID")=UserName
			rs("Checker")=AdminName
			rs("CheckDate")=now()
			rs("CheckFlag")=1
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="RegisterID" then
    InfoID=request("InfoID")
		sql="select 员工代号,姓名 from [N-基本资料单头] where 员工代号 like '%"&InfoID&"%' or 姓名 like '%"&InfoID&"%'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
				response.write ("员工编号不存在！")
				response.end
		else
			response.write(rs("员工代号")&"###"&rs("姓名"))
		end if
		rs.close
		set rs=nothing 
  elseif detailType="OrderID" then
    InfoID=request("InfoID")
	sql="select a.fbillno,c.fname,e.fname as name2 from seorder a inner join  "&_
" seorderentry b on a.finterid=b.finterid left join  "&_
" t_emp c on a.fempid=c.fitemid left join "&_
" t_ICItem e on b.fitemid=e.fitemid "&_
" where fbillno='"&InfoID&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
				response.write ("定单号不存在！")
				response.end
		else
			response.write(rs("fbillno")&"###"&rs("fname")&"###"&"<select id=""Product"" name=""Product"">")
			while (not rs.eof)
				response.write("<option value='"&rs("name2")&"'>"&rs("name2")&"</option>")
				rs.movenext
			wend
			response.write("</select>")
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from qcsys_CustomInspect where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
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
		rs.close
		set rs=nothing 
  end if
end if
 %>
