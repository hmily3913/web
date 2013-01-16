<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurview"),"|808,")=0 then 
  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！<\/font>""]}]}")
  response.end
end if
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
      datafrom=" Financesys_SupplyBill "
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
	if Instr(session("AdminPurview"),"|808.1,")=0 and Instr(session("AdminPurview"),"|808.2,")=0 and Instr(session("AdminPurview"),"|808.3,")=0 then datawhere = datawhere&" and Employer = '" & AdminName & "' "
	datawhere=datawhere&Session("AllMessage70")
	session.contents.remove "AllMessage70"
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
    sql="select SerialNum,CountDate,No,Paymentattr,Employer,Custom,Supplier,InvoiceCode,Invoice,Product,Quality,Unit,Price,Money,Tariff,Tax,Total,Remark,Payment,Advance,Contract,PayDate,PayMoney,CheckFlag,AccountFlag,ApprovFlag from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
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
		if rs.fields(i).value="1900-1-1" then
		response.write (""""",")
		else
		response.write (""""&JsonStr(rs.fields(i).value)&""",")
		end if
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
  	if  Instr(session("AdminPurview"),"|808.1,")>0 then
		dim billymd,billid
		billymd=(YEAR(date())*10000+MONTH(date())*100+DAY(date())) mod 1000000
		set rs2=connzxpt.Execute("select max(No) as ids from Financesys_SupplyBill where No like '"&billymd&"%'")
		if cdbl(rs2("ids"))=0 then
			billid=billymd*1000+1
		else
		billid=cdbl(rs2("ids"))+1
		end if
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="0" then
				connzxpt.Execute("insert into Financesys_SupplyBill (CountDate,Paymentattr,Employer,Custom,Supplier,InvoiceCode,Invoice,Product,Quality,Unit,Price,Money,Remark,Advance,Tariff,Tax,Total,Contract,PayDate,PayMoney,Biller,BillerID,BillDate,Payment,No) values ('"&Request.form("CountDate")&"','"&Request.form("Paymentattr")&"','"&Request.form("Employer")&"','"&Request.form("Custom")&"','"&Request.form("Supplier")&"','"&Request.form("InvoiceCode")&"','"&Request.form("Invoice")&"','"&Request.form("Product")(i)&"',"&Request.form("Quality")(i)&",'"&Request.form("Unit")(i)&"',"&Request.form("Price")(i)&","&Request.form("Money")(i)&",'"&Request.form("Remark")(i)&"','"&Request.form("Advance")(i)&"','"&Request.form("Tariff")(i)&"','"&Request.form("Tax")(i)&"','"&Request.form("Total")(i)&"','"&Request.form("Contract")(i)&"','"&Request.form("PayDate")(i)&"',"&Request.form("PayMoney")(i)&",'"&AdminName&"','"&UserName&"','"&now()&"','"&Request.form("Payment")&"','"&billid&"')")
			end if
		next
		response.write "###"
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
  elseif detailType="Edit"  then
		if Instr(session("AdminPurview"),"|808.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNum")(i)<>"" then
				set rs=connzxpt.Execute("select CheckFlag from Financesys_SupplyBill where SerialNum="&Request.Form("SerialNum")(i))
				if rs("CheckFlag") then
					response.Write("已审核不允许删除！")
					response.End()
				end if
				connzxpt.Execute("Delete from Financesys_SupplyBill where SerialNum="&Request.Form("SerialNum")(i))
			elseif Request.Form("SerialNum")(i)<>"" then
				set rs=connzxpt.Execute("select CheckFlag from Financesys_SupplyBill where SerialNum="&Request.Form("SerialNum")(i))
				if rs("CheckFlag") then
					response.Write("已审核不允许修改！")
					response.End()
				end if
				connzxpt.Execute("update Financesys_SupplyBill set CountDate='"&Request.form("CountDate")&"',No='"&Request.form("No")&"',Paymentattr='"&Request.form("Paymentattr")&"',Employer='"&Request.form("Employer")&"',Custom='"&Request.form("Custom")&"',Supplier='"&Request.form("Supplier")&"',InvoiceCode='"&Request.form("InvoiceCode")&"',Invoice='"&Request.form("Invoice")&"',Product='"&Request.form("Product")(i)&"',Quality="&Request.form("Quality")(i)&",Unit='"&Request.form("Unit")(i)&"',Price="&Request.form("Price")(i)&",Money="&Request.form("Money")(i)&",Remark='"&Request.form("Remark")(i)&"',Tariff='"&Request.form("Tariff")(i)&"',Tax='"&Request.form("Tax")(i)&"',Total='"&Request.form("Total")(i)&"',Advance='"&Request.form("Advance")(i)&"',Contract='"&Request.form("Contract")(i)&"',PayDate='"&Request.form("PayDate")(i)&"',PayMoney="&Request.form("PayMoney")(i)&",Payment='"&Request.form("Payment")&"' where SerialNum="&Request.Form("SerialNum")(i))
			else
				connzxpt.Execute("insert into Financesys_SupplyBill (CountDate,Paymentattr,Employer,Custom,Supplier,InvoiceCode,Invoice,Product,Quality,Unit,Price,Money,Remark,Advance,Tariff,Tax,Total,Contract,PayDate,PayMoney,Biller,BillerID,BillDate,Payment,No) values ('"&Request.form("CountDate")&"','"&Request.form("Paymentattr")&"','"&Request.form("Employer")&"','"&Request.form("Custom")&"','"&Request.form("Supplier")&"','"&Request.form("InvoiceCode")&"','"&Request.form("Invoice")&"','"&Request.form("Product")(i)&"',"&Request.form("Quality")(i)&",'"&Request.form("Unit")(i)&"',"&Request.form("Price")(i)&","&Request.form("Money")(i)&",'"&Request.form("Remark")(i)&"','"&Request.form("Advance")(i)&"','"&Request.form("Tariff")(i)&"','"&Request.form("Tax")(i)&"','"&Request.form("Total")(i)&"','"&Request.form("Contract")(i)&"','"&Request.form("PayDate")(i)&"',"&Request.form("PayMoney")(i)&",'"&AdminName&"','"&UserName&"','"&now()&"','"&Request.form("Payment")&"','"&Request.form("No")&"')")
			end if
		next
		response.write "###"
  elseif detailType="Check" then
		if Instr(session("AdminPurview"),"|808.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		connzxpt.Execute("update Financesys_SupplyBill set CheckFlag="&request("operattext")&",Checker='"&AdminName&"',CheckerID='"&UserName&"',CheckDate='"&now()&"' where SerialNum in ("&request("SerialNum")&")")
		response.Write("审核成功！")
  elseif detailType="Delete" then
		if Instr(session("AdminPurview"),"|808.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
    SerialNum=request("SerialNum")
		set rs=connzxpt.Execute("select 1 from Financesys_SupplyBill where CheckFlag=1 where SerialNum in ("&SerialNum&")")
		if not rs.eof then
			response.Write("选择单据中存在已审核单据，删除失败！")
			response.End()
		end if
		connzxpt.Execute("Delete from Financesys_SupplyBill where SerialNum in ("&SerialNum&")")
		response.write "###"
  elseif detailType="Account" then
		if Instr(session("AdminPurview"),"|808.3,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		connzxpt.Execute("update Financesys_SupplyBill set AccountFlag="&request("operattext")&",Accounter='"&AdminName&"',AccounterID='"&UserName&"',AccountDate='"&now()&"' where SerialNum in ("&request("SerialNum")&")")
		response.Write("记账成功！")
  elseif detailType="Approv" then
		if Instr(session("AdminPurview"),"|808.3,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		connzxpt.Execute("update Financesys_SupplyBill set ApprovFlag="&request("operattext")&",Approver='"&AdminName&"',ApproverID='"&UserName&"',ApprovDate='"&now()&"' where SerialNum in ("&request("SerialNum")&")")
		response.Write("认证成功！")
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="getEmp" then
		sql="select 员工代号,姓名 from [N-基本资料单头] where 是否离职=0 and 工作岗位 like '%采购%'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		response.write "["
		do until rs.eof
		Response.Write("{""ClientName"":"""&JsonStr(rs("姓名"))&""",""FNumber"":"""&JsonStr(rs("员工代号"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
	elseif detailType="getCtmsLD" then
		sql="select top 10 Fshortname,fnumber,F_105 from t_Supplier where Fshortname like '%"&request("q")&"%' or fname like '%"&request("q")&"%' or fnumber like '%"&request("q")&"%'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		response.write "["
		do until rs.eof
		Response.Write("{""ClientName"":"""&JsonStr(rs("Fshortname"))&""",""FNumber"":"""&JsonStr(rs("fnumber"))&""",""Emp"":"""&JsonStr(rs("F_105"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
	elseif detailType="Supplier" then
		sql="select count(1),Supplier,InvoiceCode from Financesys_SupplyBill where Custom like '%"&request("q")&"%' group by Supplier,InvoiceCode order by count(1) desc"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.write "["
		do until rs.eof
		Response.Write("{""ClientName"":"""&JsonStr(rs("Supplier"))&""",""FNumber"":"""&JsonStr(rs("InvoiceCode"))&"""}")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
  end if
elseif showType="Excel" then 
	sql="select * from Financesys_SupplyBill where 1=1"
	if Instr(session("AdminPurview"),"|808.1,")=0 and Instr(session("AdminPurview"),"|808.2,")=0 and Instr(session("AdminPurview"),"|808.3,")=0 then datawhere = datawhere&" and Employer = '" & AdminName & "' "
	if request("SDate")<>"" then sql=sql&" and datediff(d,'"&request("SDate")&"',CountDate)>=0 "
	if request("EDate")<>"" then sql=sql&" and datediff(d,'"&request("EDate")&"',CountDate)<=0 "
	if request("Paymentattr")<>"" then sql=sql&" and Paymentattr='"&request("Paymentattr")&"' "
	sql=sql&" order by SerialNum "
%>
<div id="listtable" style="width:100%; height:420; overflow:scroll">
<table width="1500px" border="0" id="editDetails" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" style=" overflow:auto">
<tr>    <td height="20" width="100%" class="tablemenu" colspan="22"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="$('#listtable').hide().css('z-index','550');" >&nbsp;<strong>页面查看明细</strong></font></td>
</tr>
		  <tr>
			<td width="40" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>id</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>日期</strong></font></td>
			<td width="40" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>序号</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>货款归属单位</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>采购员</strong></font></td>
			<td width="50" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>供方名称</strong></font></td>
			<td width="80" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>发票开具单位</strong></font></td>
			<td width="100" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>发票代码</strong></font></td>
			<td width="100" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>发票号码</strong></font></td>
			<td width="120" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>货物名称</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>数量</strong></font></td>
			<td width="40" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单位</strong></font></td>
			<td width="50" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单价（不含税）</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>金额（不含税）</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>税率</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>税额</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>价税合计</strong></font></td>
			<td width="100" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>备注</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>付款方式</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>预付款</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>购销合同</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>付款日期</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>付款金额</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>审核</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>登记人</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工号</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>登记时间</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>审核人</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工号</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>审核时间</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>记账</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>记账人</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工号</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>记账时间</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>认证</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>认证人</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工号</strong></font></td>
			<td width="60" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>认证时间</strong></font></td>
		  </tr>
<%
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
	while (not rs.eof)
		response.Write("<tr bgColor='#EBF2F9'>")
	  for i=0 to rs.fields.count-1
	    if IsNull(rs.fields(i).value) then
		response.write ("<td>"&rs.fields(i).value&"</td>")
		else
		if rs.fields(i).value="1900-1-1" then
		response.write ("<td></td>")
		else
		response.write ("<td>"&JsonStr(rs.fields(i).value)&"</td>")
		end if
		end if
	  next
		rs.movenext
	wend
	rs.close
	set rs=nothing 
%>
</table>
</div>
<%
end if
 %>
