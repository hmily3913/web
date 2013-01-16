<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|601,")=0 then 
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
      datafrom=" purchasesys_SupplierEvaluat "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where 1=1 "
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
	datawhere=datawhere&Session("AllMessage52")&Session("AllMessage53")&Session("AllMessage54")&Session("AllMessage55")
	session.contents.remove "AllMessage52"
	session.contents.remove "AllMessage53"
	session.contents.remove "AllMessage54"
	session.contents.remove "AllMessage55"
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
    sql="select * from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格'
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("Supplier")%>","<%=rs("Evaluatdate")%>","<%=rs("Telephone")%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("pall")%>","<%=rs("EnFlag")%>","<%=rs("QcFlag")%>","<%=rs("PoFlag")%>","<%=rs("CheckFlag")%>","<%=JsonStr(rs("Remark"))%>"]}
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
		if Instr(session("AdminPurview"),"|207.1,")=0 then
			response.Write("你没有权限进行此操作")
			response.End()
		end if
		set rs = server.createobject("adodb.recordset")
		sql="select * from purchasesys_SupplierEvaluat"
		rs.open sql,connzxpt,1,3
		rs.addnew
		rs("BillerID")=UserName
		rs("Biller")=AdminName
		rs("BillDate")=now()
		rs("Supplier")=Request("Supplier")
		rs("Evaluatdate")=Request("Evaluatdate")
		rs("Telephone")=Request("Telephone")
		rs("OldOrNew")=Request("OldOrNew")
		rs("Remark")=Request("Remark")
		if Instr(session("AdminPurview"),"|207.2,")>0 then
			rs("p1")=Request("p1")
			rs("p2")=Request("p2")
			rs("p3")=Request("p3")
			rs("p4")=Request("p4")
			rs("p5")=Request("p5")
			rs("EnerID")=UserName
			rs("Ener")=AdminName
			rs("EnDate")=now()
			rs("EnFlag")=1
			rs("Enopinion")=Request("Enopinion")
		end if
		if Instr(session("AdminPurview"),"|207.3,")>0 then
			rs("p6")=Request("p6")
			rs("p7")=Request("p7")
			rs("p8")=Request("p8")
			rs("p9")=Request("p9")
			rs("p10")=Request("p10")
			rs("QcerID")=UserName
			rs("Qcer")=AdminName
			rs("QcDate")=now()
			rs("QcFlag")=1
			rs("Qcopinion")=Request("Qcopinion")
		end if
		if Instr(session("AdminPurview"),"|207.4,")>0 then
			rs("p11")=Request("p11")
			rs("p12")=Request("p12")
			rs("p13")=Request("p13")
			rs("p14")=Request("p14")
			rs("p15")=Request("p15")
			rs("p16")=Request("p16")
			rs("PoerID")=UserName
			rs("Poer")=AdminName
			rs("PoDate")=now()
			rs("PoFlag")=1
			rs("Poopinion")=Request("Poopinion")
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
		sql="select * from purchasesys_SupplierEvaluat where SerialNum="&request("SerialNum")
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurview"),"|207.1,")>0 then
			if rs("CheckFlag")>0 then
				response.write ("已审核不允许编辑，请检查！")
				response.end
			end if
			rs("BillerID")=UserName
			rs("Biller")=AdminName
			rs("BillDate")=now()
			rs("Supplier")=Request("Supplier")
			rs("Evaluatdate")=Request("Evaluatdate")
			rs("Telephone")=Request("Telephone")
			rs("Remark")=Request("Remark")
			rs("OldOrNew")=Request("OldOrNew")
		end if
		if Instr(session("AdminPurview"),"|207.2,")>0 then
			if rs("CheckFlag")>0 then
				response.write ("已审核不允许编辑，请检查！")
				response.end
			end if
			rs("p1")=Request("p1")
			rs("p2")=Request("p2")
			rs("p3")=Request("p3")
			rs("p4")=Request("p4")
			rs("p5")=Request("p5")
			rs("pall")=Request("pall")
			rs("EnerID")=UserName
			rs("Ener")=AdminName
			rs("EnDate")=now()
			rs("EnFlag")=1
			rs("Enopinion")=Request("Enopinion")
		end if
		if Instr(session("AdminPurview"),"|207.3,")>0 then
			if rs("CheckFlag")>0 then
				response.write ("已审核不允许编辑，请检查！")
				response.end
			end if
			rs("p6")=Request("p6")
			rs("p7")=Request("p7")
			rs("p8")=Request("p8")
			rs("p9")=Request("p9")
			rs("p10")=Request("p10")
			rs("pall")=Request("pall")
			rs("QcerID")=UserName
			rs("Qcer")=AdminName
			rs("QcDate")=now()
			rs("QcFlag")=1
			rs("Qcopinion")=Request("Qcopinion")
		end if
		if Instr(session("AdminPurview"),"|207.4,")>0 then
			if rs("CheckFlag")>0 then
				response.write ("已审核不允许编辑，请检查！")
				response.end
			end if
			rs("p11")=Request("p11")
			rs("p12")=Request("p12")
			rs("p13")=Request("p13")
			rs("p14")=Request("p14")
			rs("p15")=Request("p15")
			rs("p16")=Request("p16")
			rs("pall")=Request("pall")
			rs("PoerID")=UserName
			rs("Poer")=AdminName
			rs("PoDate")=now()
			rs("PoFlag")=1
			rs("Poopinion")=Request("Poopinion")
		end if
		if Instr(session("AdminPurview"),"|207.5,")>0 then
			rs("CheckerID")=UserName
			rs("Checker")=AdminName
			rs("CheckDate")=now()
			rs("Remark")=Request("Remark")
			rs("CheckFlag")=1
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="Delete" then
		if Instr(session("AdminPurview"),"|207.1,")=0 then
			response.Write("你没有权限进行此操作")
			response.End()
		end if
    SerialNum=request("SerialNum")
		sql="select * from purchasesys_SupplierEvaluat where CheckFlag>0 and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from purchasesys_SupplierEvaluat where SerialNum in ("&SerialNum&")")
			response.write "###"
		else
			response.Write("已审核不允许删除！")
			response.End()
		end if
		rs.close
		set rs=nothing
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="RegisterID" then
    InfoID=request("InfoID")
		sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号 like '%"&InfoID&"%' "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
				response.write ("员工编号不存在！")
				response.end
		else
			response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###"&rs("性别")&"###"&rs("工作岗位")&"###"&rs("到职日")&"###"&rs("职等"))
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from purchasesys_SupplierEvaluat where SerialNum="&InfoID
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
