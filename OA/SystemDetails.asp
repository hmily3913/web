<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|210,")=0 then 
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
      datafrom=" Bill_SoftOperateRule "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where OldFlag=0 "
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = Request.Form("qtype")
	datawhere = " WHERE " & searchcols & " LIKE '%" & searchterm & "%' "
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
    sql="select * from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格'
	dim oldstr,checkstr
	oldstr="正常"
	checkstr="待审"
	if rs("OldFlag")="1" then oldstr="过期"
	if rs("CheckFlag")="1" then
	checkstr="审核"
	elseif rs("CheckFlag")="2" then
	checkstr="审批"
	if datediff("d",rs("EffectDate"),now())>=0 then checkstr="生效"
	end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("RuleName")%>","<%=rs("EffectDate")%>","<%=rs("VersionNum")%>","<%=rs("Platform")%>","<%=rs("Module")%>","<%=rs("Register")%>","<%=rs("RegisterName")%>","<%=rs("ReceivDepartment")%>","<%=rs("Signman")%>","<%=rs("OldVersion")%>","<%=oldstr%>","<%=checkstr%>","<%=rs("Biller")%>"]}
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
  	if  Instr(session("AdminPurviewFLW"),"|210.1,")>0 then
	  if Request("OldVersion")<>"" then
	    set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_SoftOperateRule where VersionNum='"&Request("OldVersion")&"' and RuleName='"&Request("RuleName")&"'"
		rs.open sql,connzxpt,1,1
		if rs("CheckFlag")=0 then
		response.write "版本为："&Request("OldVersion")&"的 "&Request("RuleName")&" 可以直接编辑，不需要变更"
		response.end
		end if
	    set rs = server.createobject("adodb.recordset")
		sql="select count(1) as idcount from Bill_SoftOperateRule where VersionNum='"&Request("VersionNum")&"' and RuleName='"&Request("RuleName")&"'"
		rs.open sql,connzxpt,1,1
		if rs("idcount")>0 then
		response.write "版本为："&Request("VersionNum")&"的 "&Request("RuleName")&" 已存在，修改新版本号再保存！"
		response.end
		end if
	  end if
		set rs = server.createobject("adodb.recordset")
		sql="select * from Bill_SoftOperateRule"
		rs.open sql,connzxpt,1,3
		rs.addnew
		rs("RuleName")=Request("RuleName")
		rs("VersionNum")=Request("VersionNum")
		rs("Platform")=Request("Platform")
		rs("Module")=Request("Module")
		rs("Register")=Request("Register")
		rs("RegisterName")=Request("RegisterName")
		rs("ReceivDepartment")=Request("RD1")&Request("RD2")&Request("RD3")&Request("RD4")&Request("RD5")&Request("RD6")&Request("RD7")&Request("RD8")&Request("RD9")&Request("RD10")&Request("RD11")&Request("RD12")&Request("RD13")&Request("RD14")&Request("RD15")&Request("RD16")
		rs("RuleDescrib")=Request("RuleDescrib")
		rs("OldVersion")=Request("OldVersion")
		rs("Biller")=UserName
		rs("BillDate")=now()
		if Request("EffectDate")<>"" then rs("EffectDate")=Request("EffectDate")
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
	sql="select * from Bill_SoftOperateRule where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|210.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
    end if
	if rs("CheckFlag")>0 then
		response.write ("当前单据不允许编辑，保存失败！")
		response.end
    end if
		rs("RuleName")=Request("RuleName")
		rs("VersionNum")=Request("VersionNum")
		rs("Platform")=Request("Platform")
		rs("Module")=Request("Module")
		rs("Register")=Request("Register")
		rs("RegisterName")=Request("RegisterName")
		rs("ReceivDepartment")=Request("RD1")&Request("RD2")&Request("RD3")&Request("RD4")&Request("RD5")&Request("RD6")&Request("RD7")&Request("RD8")&Request("RD9")&Request("RD10")&Request("RD11")&Request("RD12")&Request("RD13")&Request("RD14")&Request("RD15")&Request("RD16")
		rs("RuleDescrib")=Request("RuleDescrib")
		rs("OldVersion")=Request("OldVersion")
		rs("Biller")=UserName
		rs("BillDate")=now()
		if Request("EffectDate")<>"" then rs("EffectDate")=Request("EffectDate")
	response.write "###"
    rs.update
	rs.close
	set rs=nothing 
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftOperateRule where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|210.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
	end if
	if rs("CheckFlag")>0 then
		response.write ("当前状态不允许删除！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_SoftOperateRule where SerialNum="&SerialNum)
	response.write "###"
  elseif detailType="Check" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftOperateRule where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if rs("CheckFlag")=2 then
		response.write ("此单据已经审批，不需要审核！")
		response.end
	end if
	if rs("CheckFlag")=0 and Instr(session("AdminPurviewFLW"),"|210.2,")>0 then
    rs("Checker")=AdminName
    rs("CheckDate")=now()
    rs("CheckFlag")=1
	'如果是变更单，审核时同时更新旧版本为不生效状态
	if rs("OldVersion")<>"" then
	  connzxpt.Execute("update Bill_SoftOperateRule set OldFlag=1 where VersionNum='"&rs("OldVersion")&"' and RuleName='"&rs("RuleName")&"'")
	end if
	elseif rs("CheckFlag")=1 and Instr(session("AdminPurviewFLW"),"|210.3,")>0 then
    rs("Approvaler")=AdminName
    rs("ApprovalDate")=now()
    rs("CheckFlag")=2
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write("###")
  elseif detailType="Sign" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftOperateRule where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|210.4,")=0 then
		response.write ("你没有权限进行当前操作！")
		response.end
	end if
	if Instr(rs("Signman"),AdminName)=0 then
    rs("Signman")=rs("Signman")&session("AdminName")&","
	rs.update
	end if
	rs.close
	set rs=nothing 
	response.write("###")
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="Register" or detailType="RegisterName" then
    InfoID=request("InfoID")
	if InfoID="" then InfoID=UserName
	sql="select 员工代号,姓名 from [N-基本资料单头] where 员工代号='"&InfoID&"' or 姓名='"&InfoID&"'"
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
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
	sql="select * from Bill_SoftOperateRule where SerialNum="&InfoID
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
		response.write (""""&rs.fields(i).name & """:"""&replace(replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r"),chr(34),"\""")&""",")
		end if
	  next
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}]}")
		else
		response.write (""""&rs.fields(i).name & """:"""&replace(replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r"),chr(34),"\""")&"""}]}")
		end if
	end if
	rs.close
	set rs=nothing 
  end if
end if
 %>
