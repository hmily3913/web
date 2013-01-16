<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="报表系统" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="../Images/CssAdmin.css">

</HEAD>
<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<BODY>
<%
'if Instr(session("AdminPurviewFLW"),"|207,")=0 then 
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
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>选择</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>采购订单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物料名称</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>规格</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单位</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>品质要求</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>紧急数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>紧急说明</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>品保</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>仓库</strong></font></td>
  </tr>
 <%
  dim Depart,wherestr,seachword,flag4search
  wherestr=""
  seachword=request("seachword")
  if seachword<>"" then
  start_date=dateadd("d",4,split(getDateRange(seachword,2012),"###")(0))
  end_date=dateadd("d",4,split(getDateRange(seachword,2012),"###")(1))
  wherestr=" and datediff(d,RegDate,'"&start_date&"')<=0 and datediff(d,RegDate,'"&end_date&"')>=0 "
  end if
  flag4search=request("flag4search")
  if flag4search="1" then 
    if Instr(session("AdminPurviewFLW"),"|207.2,")>0 then
	 wherestr=wherestr&" and QCFlag=1 "
    elseif Instr(session("AdminPurviewFLW"),"|207.3,")>0 then
	 wherestr=wherestr&" and STFlag=1 "
    else
	 wherestr=wherestr&" and (QCFlag=1 or STFlag=1) "
	end if
  elseif flag4search="2" then 
    if Instr(session("AdminPurviewFLW"),"|207.2,")>0 then
	 wherestr=wherestr&" and QCFlag=2 "
    elseif Instr(session("AdminPurviewFLW"),"|207.3,")>0 then
	 wherestr=wherestr&" and STFlag=2 "
    else
	 wherestr=wherestr&" and (QCFlag=2 or STFlag=2) "
	end if
  elseif flag4search="0" then 
    wherestr=wherestr&" and QCFlag=0 and STFlag=0 "
  end if
  dim page'页码
      page=clng(request("Page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Bill_PMMTRboard "
  dim datawhere'数据条件
		 datawhere="where 1=1"&wherestr&Session("AllMessage13")&Session("AllMessage14")
		 Session("AllMessage13")=""
		 Session("AllMessage14")=""
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by SerialNum desc"
  dim i'用于循环的整数
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
    rs.pagesize = 20 '每页显示记录数
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
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    dim sql2,rs2,hejikoufen
	hejikoufen=0
	dim formdata(3),bgcolors
    sql="select *,left(NeedNote,10) as a2,left(Quality,10) as a1 from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
'		if Len(rs("ReplyText"))>0 then
'		  bgcolors="#ff99ff"'粉色
'		end if
	dim QCstr,STstr
	if rs("QCFlag")=1 then
	QCstr="验前接收"
	if Instr(session("AdminPurviewFLW"),"|207.2,")>0 then bgcolors="#ff99ff"'
	elseif rs("QCFlag")=2 then
	QCstr="验后确认"
	if Instr(session("AdminPurviewFLW"),"|207.2,")>0 then bgcolors="#7CFC00"'
	else
	QCstr="未接收"
	end if
	if rs("STFlag")=1 then
	STstr="验前接收"
	if Instr(session("AdminPurviewFLW"),"|207.3,")>0 then bgcolors="#ff99ff"'
	elseif rs("STFlag")=2 then
	STstr="验后确认"
	if Instr(session("AdminPurviewFLW"),"|207.3,")>0 then bgcolors="#7CFC00"'
	else
	STstr="未接收"
	end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap><input type='radio' name='Snum' ForCheck='"&rs("QCFlag")&"_"&rs("STFlag")&"' value='"&rs("SerialNum")&"'></td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegisterName")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("OrderID")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ProductName")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Model")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Quantity")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Unit")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a1")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("NeedQuantity")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a2")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&QCstr&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&STstr&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  else
    response.write "<tr><td height='50' align='center' colspan='14' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
	response.Write("###"&pagec&"###"&idcount&"###")
elseif showType="AddEditShow" then 
  dim detailType
  detailType=request("detailType")
'数据处理
  dim SerialNum,Register,RegisterName,RegDate,Department,Departmentname,OrderID,ProductId,ProductName,Model
  dim Unit,Quantity,Quality,NeedQuantity,NeedNote
  dim style1,style2,style3,QCReplyer,QCReplyText,QCReplyDate,QCFlag
  dim STReplyer,STReplyText,STReplyDate,STFlag
  if detailType="AddNew" then
	sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&UserName&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
	Register=UserName
	RegisterName=AdminName
	RegDate=date()
	Department=rs("部门别")
	Departmentname=rs("部门名称")
	style1="block;"
	style2="none;"
  elseif detailType="Edit" or detailType="Check" then
    SerialNum=request("SerialNum")
	sql="select * from Bill_PMMTRboard where SerialNum="&SerialNum
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	Register=rs("Register")
	RegisterName=rs("RegisterName")
	RegDate=rs("RegDate")
	Department=rs("Department")
	Departmentname=rs("Departmentname")
	OrderID=rs("OrderID")
	ProductId=rs("ProductId")
	ProductName=rs("ProductName")
	Model=rs("Model")
	Unit=rs("Unit")
	Quantity=rs("Quantity")
	Quality=rs("Quality")
	NeedQuantity=rs("NeedQuantity")
	NeedNote=rs("NeedNote")
	QCReplyer=rs("QCReplyer")
	QCReplyText=rs("QCReplyText")
	QCReplyDate=rs("QCReplyDate")
	QCFlag=rs("QCFlag")
	STReplyText=rs("STReplyText")
	STReplyer=rs("STReplyer")
	STReplyDate=rs("STReplyDate")
	STReplyDate=rs("STReplyDate")
	STFlag=rs("STFlag")
	if detailType="Edit" then
	style1="block;"
	style2="block;"
	else
	style1="none;"
	style2="block;"
	if Instr(session("AdminPurviewFLW"),"|207.2,")>0 then
	  style3="block;"
	else
	  style3="none;"
	end if
	end if	
  end if
  %>
 <div id="AddandEditdiv" style="width:790px; height:460px; overflow:auto; ">
 <form id="AddForm" name="AddForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>生管紧急物料看板</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews width="100%">
      <tr>
        <td height="20" align="left">单据号：</td>
        <td>
		<input name="SerialNum" type="text" class="textfield" id="SerialNum" style="WIDTH: 140;" value="<%= SerialNum %>" maxlength="100" readonly="true"></td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left">提交人工号：</td>
        <td><input name="Register" type="text" class="textfield" id="Register" style="WIDTH: 140;" value="<%= Register %>" maxlength="100" onBlur="getInfo('Register')"></td>
        <td height="20" align="left">提交人姓名：</td>
        <td><input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 140;" value="<%= RegisterName %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">提交日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 140;" value="<%= RegDate %>" maxlength="100"  readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">提交部门编号：</td>
        <td>
		<input name="Department" type="text" class="textfield" id="Department" style="WIDTH: 140;" value="<%= Department %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">提交部门名称：</td>
        <td>
		<input name="Departmentname" type="text" class="textfield" id="Departmentname" style="WIDTH: 140;" value="<%= Departmentname %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">采购定单号：</td>
        <td><input name="OrderID" type="text" class="textfield" id="OrderID" style="WIDTH: 140;" value="<%= OrderID %>" maxlength="100" onChange="getInfo('OrderID')"></td>
      </tr>
      <tr>
        <td height="20" align="left">物品名称：</td>
        <td id="Product_td">
		<input name="ProductName" type="text" class="textfield" id="ProductName" style="WIDTH: 140;" value="<%= ProductName %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">物料编号：</td>
        <td><input name="ProductId" type="text" class="textfield" id="ProductId" style="WIDTH: 140;" value="<%= ProductId %>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">规格：</td>
        <td>
		<input name="Model" type="text" class="textfield" id="Model" style="WIDTH: 140;" value="<%= Model %>" maxlength="100"  readonly="true">
		</td>
      </tr>
      <tr>
        <td height="20" align="left">单位：</td>
        <td><input name="Unit" type="text" class="textfield" id="Unit" style="WIDTH: 140;" value="<%= Unit %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">数量：</td>
        <td><input name="Quantity" type="text" class="textfield" id="Quantity" style="WIDTH: 140;" value="<%= Quantity %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">品质要求：</td>
        <td>
		<input name="Quality" type="text" class="textfield" id="Quality" style="WIDTH: 140;" value="<%= Quality %>" maxlength="100"  readonly="true">
		</td>
      </tr>
      <tr>
        <td height="20" align="left">急需数量：</td>
        <td><input name="NeedQuantity" type="text" class="textfield" id="NeedQuantity" style="WIDTH: 140;" value="<%= NeedQuantity %>" maxlength="100" onBlur="return checkNum(this)"></td>
        <td height="20" align="left">急需说明：</td>
        <td colspan="3">
		<input name="NeedNote" type="text" class="textfield" id="NeedNote" style="WIDTH: 340;" value="<%= NeedNote %>" maxlength="500" >
		</td>
      </tr>
      <tr>
        <td height="20" align="left">品保接收状态：</td>
        <td><%
		 if QCFlag = 1 then
		 response.Write("验前接收")
		 elseif QCFlag = 2 then
		 response.Write("验后确认")
		 else
		 response.Write("未接收")
		 end if
		 %></td>
		 <td > 品保接收人： </td>
		 <td >
		<%= QCReplyer %></td>
		 <td > 品保接收日期： </td>
		 <td >
		 <%= QCReplyDate %></td>
      </tr>
      <tr>
        <td height="20" align="left">仓库接收状态：</td>
        <td><%
		 if STFlag = 1 then
		 response.Write("验前接收")
		 elseif STFlag = 2 then
		 response.Write("验后确认")
		 else
		 response.Write("未接收")
		 end if
		 %></td>
		 <td > 仓库接收人： </td>
		 <td >
		<%= STReplyer %></td>
		 <td > 仓库接收日期： </td>
		 <td >
		 <%= STReplyDate %></td>
      </tr>
	  </table>
<div id="Buttondiv" style="display:<%= style1 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="1">  </td>
  </tr>
	<tr>
	  <td align="center">
	  <input type="hidden" name="CheckFlag" id="CheckFlag" value="<%= CheckFlag %>">
	  <input type="hidden" name="detailType" id="detailType" value="<%= detailType %>">
			<input name="submitSaveAdd" type="button" class="button"  id="submitSaveAdd" value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">&nbsp;
			<input name="submitDelete" type="button" class="button"  id="submitDelete" value="删除" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('Delete');toSubmit(this);">
	  </td>
	</tr>
  <tr>  <td height="1">  </td>
  </table>
</div>
<div id="Buttondiv" style="display:<%= style2 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="1" colspan="4">  </td></tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 品保意见 </td>
<td colspan="3">
  <textarea name="QCReplyText" id="QCReplyText" style="width:500px; height:100px; "><%= QCReplyText %></textarea>
</td>
</tr> 
  <tr>  <td height="1" colspan="4">  </td>
  </tr>
	<tr>
	  <td align="center" colspan="4">
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="品保意见" style="WIDTH: 80; display:<%= style3 %>;"  onClick="javascript:$('#detailType').val('Check');toSubmit(this);">&nbsp;
	  </td>
	</tr>
  <tr>  <td height="5">  </td>
  </table>
</div>
	</td>
  </tr>
</table>
</form>
</div>
  <%
    rs.close
    set rs=nothing
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" and Instr(session("AdminPurviewFLW"),"|207.1,")>0 then
	set rs = server.createobject("adodb.recordset")
	sql="select * from Bill_PMMTRboard"
	rs.open sql,connzxpt,1,3
	rs.addnew
	rs("Biller")=UserName
	rs("BillDate")=now()
	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("OrderID")=Request("OrderID")
	rs("ProductId")=Request("ProductId")
	rs("ProductName")=Request("ProductName")
	rs("Model")=Request("Model")
	rs("Unit")=Request("Unit")
	rs("Quantity")=Request("Quantity")
	rs("Quality")=Request("Quality")
	rs("NeedQuantity")=Request("NeedQuantity")
	rs("NeedNote")=Request("NeedNote")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Edit" and Instr(session("AdminPurviewFLW"),"|207.1,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_PMMTRboard where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs("Biller")=UserName
	rs("BillDate")=now()
	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("OrderID")=Request("OrderID")
	rs("ProductId")=Request("ProductId")
	rs("ProductName")=Request("ProductName")
	rs("Model")=Request("Model")
	rs("Unit")=Request("Unit")
	rs("Quantity")=Request("Quantity")
	rs("Quality")=Request("Quality")
	rs("NeedQuantity")=Request("NeedQuantity")
	rs("NeedNote")=Request("NeedNote")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Delete" and Instr(session("AdminPurviewFLW"),"|207.1,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_PMMTRboard where SerialNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_PMMTRboard where SerialNum="&SerialNum)
	response.write "###"
  elseif detailType="Check" and Instr(session("AdminPurviewFLW"),"|207.2,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_PMMTRboard where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
    rs("QCReplyer")=session("AdminName")
    rs("QCReplyDate")=now()
    rs("QCReplyText")=Request("QCReplyText")
	rs("QCFlag")=2
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&Request("QCReplyText")&"###")
  elseif detailType="QC" then
   if Instr(session("AdminPurviewFLW"),"|207.2,")=0 then
		response.write ("@@@你没有权限执行此操作！@@@")
		response.end
   end if
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_PMMTRboard where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("@@@数据库读取记录出错！@@@")
		response.end
	end if
	if rs("QCFlag")<2 then
    rs("QCReplyer")=session("AdminName")
    rs("QCReplyDate")=now()
	rs("QCFlag")=1
	rs.update
	rs.close
	else
		response.write ("@@@当前状态不允许此操作！@@@")
		response.end
	end if
	set rs=nothing 
	response.write("###")
  elseif detailType="ST" then
   if Instr(session("AdminPurviewFLW"),"|207.3,")=0 then
		response.write ("@@@你没有权限执行此操作！@@@")
		response.end
   end if
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_PMMTRboard where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("@@@数据库读取记录出错！@@@")
		response.end
	end if
    rs("STReplyer")=session("AdminName")
    rs("STReplyDate")=now()
	if rs("STFlag")=1 then
	rs("STFlag")=2
	elseif rs("STFlag")=0 then
	rs("STFlag")=1
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write("###")
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="OrderID" then
    InfoID=request("InfoID")
	sql="select c.Fname as a1,c.fnumber,c.Fmodel,d.Fname as a2,b.Fauxqty,b.FentrySelfP0247 from poorder a,POOrderEntry b,t_ICitem c,t_measureUnit d "&_
		"where a.finterid=b.finterid and a.fcheckerid>0 and b.fitemid=c.fitemid  "&_
		"and b.Funitid=d.fitemid and a.fbillno='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
    if rs.bof and rs.eof then
        response.write ("采购定单号不存在！")
        response.end
	else
	  response.write(rs("a1")&"###"&rs("fnumber")&"###"&rs("Fmodel")&"###"&rs("a2")&"###"&rs("Fauxqty")&"###"&rs("FentrySelfP0247")&"###<select id='ProductName' name='ProductName' onChange='changePro()' style='width:140'>")
    end if
	while (not rs.eof)
	  response.write("<option value='"&rs("a1")&"'>"&rs("a1")&"@"&rs("fnumber")&"@"&rs("Fmodel")&"@"&rs("a2")&"@"&rs("Fauxqty")&"@"&rs("FentrySelfP0247")&"</option>")
	  rs.movenext
	wend
	response.write("</select>###")
	rs.close
	set rs=nothing 
  elseif detailType="Register" then
    InfoID=request("InfoID")
	sql="select a.姓名,a.部门别,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write("###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###")
	end if
	rs.close
	set rs=nothing 
  end if
end if
 %>
</body>
</html>
