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
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp" -->
<BODY>
<%
'if Instr(session("AdminPurviewFLW"),"|203,")=0 then 
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
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>接受部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>是否损失</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品型号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客户排行</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客户等级</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单状态</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>经办人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>下单日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>交货日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生管回复交期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>交货天数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>异常类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>异常描述</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>改善回复</strong></font></td>
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
    wherestr=wherestr&" and len(ReplyText)=0 and checkFlag=0 "
  elseif flag4search="2" then 
    wherestr=wherestr&" and len(ReplyText)>0 and checkFlag=0 "
  elseif flag4search="3" then 
    wherestr=wherestr&" and checkFlag>0 "
	else
	  wherestr=wherestr&" and checkFlag=0 "
  end if
  dim page'页码
      page=clng(request("Page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Bill_OrderAbnormal "
  dim datawhere'数据条件
		 datawhere="where 1=1"&wherestr&Session("AllMessage6")
		 Session("AllMessage6")=""
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
  if print_tag=1 then
    rs.pagesize = idcount '每页显示记录数
	rs.absolutepage = 1  
  else
    rs.pagesize = 20 '每页显示记录数
	rs.absolutepage = page  
  end if
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
    sql="select *,left(AbnormalNote,10) as a1,left(ReplyText,10) as a2,datediff(d,CustomDate,MCReplyDate) as ddate from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
		if Len(rs("ReplyText"))>0 then
		  bgcolors="#ff99ff"'粉色
		end if
		if rs("CheckFlag")>0 then
		  bgcolors="#66ff66"'粉色
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegisterName")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Departmentname")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ReceivDepartment")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("IsLoss")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("LossAmount")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("OrderID")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Product")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ProductType")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("CustomRanke")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("CustomLevel")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("OrderState")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("OrderQuantity")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Agenter")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("OrderDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("CustomDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("MCReplyDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ddate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("AbnormalType")&"</td>" & vbCrLf
	  if print_tag=1 then
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("AbnormalNote")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Check',"&rs("SerialNum")&")"">"&rs("ReplyText")&"</td>" & vbCrLf
	  else
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Check',"&rs("SerialNum")&")"">"&rs("a2")&"</td>" & vbCrLf
	  end if
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  else
    response.write "<tr><td height='50' align='center' colspan='21' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
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
  dim SerialNum,Register,RegisterName,RegDate,Department,Departmentname,OrderID,Product,ReceivDepartment
  dim ProductType,CustomRanke,CustomLevel,OrderState,OrderQuantity,Agenter,AbnormalType,AbnormalNote
  dim CustomID,CheckFlag,style1,style2,style3,Replyer,ReplyText,ReplyDate
  dim IsLoss,LossAmount,OrderDate,CustomDate,MCReplyDate
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
	sql="select * from Bill_OrderAbnormal where SerialNum="&SerialNum
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	Register=rs("Register")
	RegisterName=rs("RegisterName")
	RegDate=rs("RegDate")
	Department=rs("Department")
	Departmentname=rs("Departmentname")
	ReceivDepartment=rs("ReceivDepartment")
	OrderID=rs("OrderID")
	Product=rs("Product")
	ProductType=rs("ProductType")
	CustomRanke=rs("CustomRanke")
	CustomLevel=rs("CustomLevel")
	OrderState=rs("OrderState")
	OrderQuantity=rs("OrderQuantity")
	Agenter=rs("Agenter")
	AbnormalType=rs("AbnormalType")
	AbnormalNote=rs("AbnormalNote")
	CustomID=rs("CustomID")
	IsLoss=rs("IsLoss")
	LossAmount=rs("LossAmount")
	OrderDate=rs("OrderDate")
	CustomDate=rs("CustomDate")
	MCReplyDate=rs("MCReplyDate")
	Replyer=rs("Replyer")
	ReplyText=rs("ReplyText")
	ReplyDate=rs("ReplyDate")
	if detailType="Edit" then
	style1="block;"
	style2="block;"
	else
	style1="none;"
	style2="block;"
	if Instr(session("AdminPurviewFLW"),"|203.2,")>0 or Instr(session("AdminPurviewFLW"),"|203.3,")>0 then
	  style3="visible;"
	else
	  style3="hidden;"
	end if
	end if	
  end if
  %>
 <div id="AddandEditdiv" style="width:790px; height:460px; overflow:auto; ">
 <form id="AddForm" name="AddForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap id="formove"><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>订单异常反馈处理</strong></font></td>
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
        <td height="20" align="left">反馈人工号：</td>
        <td><input name="Register" type="text" class="textfield" id="Register" style="WIDTH: 140;" value="<%= Register %>" maxlength="100" onBlur="getInfo('Register')"></td>
        <td height="20" align="left">反馈人姓名：</td>
        <td><input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 140;" value="<%= RegisterName %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">反馈日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 140;" value="<%= RegDate %>" maxlength="100"  readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">反馈部门编号：</td>
        <td>
		<input name="Department" type="text" class="textfield" id="Department" style="WIDTH: 140;" value="<%= Department %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">反馈部门名称：</td>
        <td>
		<input name="Departmentname" type="text" class="textfield" id="Departmentname" style="WIDTH: 140;" value="<%= Departmentname %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">接收部门：</td>
        <td>
		<select id="ReceivDepartment" name="ReceivDepartment">
		  <option value="一分厂" <% If ReceivDepartment="一分厂" Then Response.Write("selected")%>>一分厂</option>
		  <option value="二分厂" <% If ReceivDepartment="二分厂" Then Response.Write("selected")%>>二分厂</option>
		  <option value="三分厂" <% If ReceivDepartment="三分厂" Then Response.Write("selected")%>>三分厂</option>
		  <option value="营销部" <% If ReceivDepartment="营销部" Then Response.Write("selected")%>>营销部</option>
		  <option value="采购部" <% If ReceivDepartment="采购部" Then Response.Write("selected")%>>采购部</option>
		  <option value="工程部" <% If ReceivDepartment="工程部" Then Response.Write("selected")%>>工程部</option>
		  <option value="生管部" <% If ReceivDepartment="生管部" Then Response.Write("selected")%>>生管部</option>
		  <option value="仓储科" <% If ReceivDepartment="仓储科" Then Response.Write("selected")%>>仓储科</option>
		  <option value="品保部" <% If ReceivDepartment="品保部" Then Response.Write("selected")%>>品保部</option>
		  <option value="生技部" <% If ReceivDepartment="生技部" Then Response.Write("selected")%>>生技部</option>
		  <option value="财务部" <% If ReceivDepartment="财务部" Then Response.Write("selected")%>>财务部</option>
		  <option value="人资部" <% If ReceivDepartment="人资部" Then Response.Write("selected")%>>人资部</option>
		  <option value="总经办" <% If ReceivDepartment="总经办" Then Response.Write("selected")%>>总经办</option>
		  <option value="娄桥" <% If ReceivDepartment="娄桥" Then Response.Write("selected")%>>娄桥</option>
		</select>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">定单号：</td>
        <td><input name="OrderID" type="text" class="textfield" id="OrderID" style="WIDTH: 140;" value="<%= OrderID %>" maxlength="100" onChange="getInfo('OrderID')"></td>
        <td height="20" align="left">产品型号：</td>
        <td id="Product_td">
		<input name="Product" type="text" class="textfield" id="Product" style="WIDTH: 140;" value="<%= Product %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">产品类别：</td>
        <td>
		<input name="ProductType" type="text" class="textfield" id="ProductType" style="WIDTH: 140;" value="<%= ProductType %>" maxlength="100"  readonly="true">
		</td>
      </tr>
      <tr>
        <td height="20" align="left">客户代号：</td>
        <td><input name="CustomID" type="text" class="textfield" id="CustomID" style="WIDTH: 140;" value="<%= CustomID %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">客户排行：</td>
        <td><input name="CustomRanke" type="text" class="textfield" id="CustomRanke" style="WIDTH: 140;" value="<%= CustomRanke %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">客户等级：</td>
        <td>
		<input name="CustomLevel" type="text" class="textfield" id="CustomLevel" style="WIDTH: 140;" value="<%= CustomLevel %>" maxlength="100"  readonly="true">
		</td>
      </tr>
      <tr>
        <td height="20" align="left">订单状态：</td>
        <td>
		<select id="OrderState" name="OrderState">
		  <option value="未生产" <% If OrderState="未生产" Then Response.Write("selected")%>>未生产</option>
		  <option value="已生产" <% If OrderState="已生产" Then Response.Write("selected")%>>已生产</option>
		  <option value="等待物料" <% If OrderState="等待物料" Then Response.Write("selected")%>>等待物料</option>
		  <option value="等待包装" <% If OrderState="等待包装" Then Response.Write("selected")%>>等待包装</option>
		  <option value="退单" <% If OrderState="退单" Then Response.Write("selected")%>>退单</option>
		</select>
		</td>
        <td height="20" align="left">订单数量：</td>
        <td><input name="OrderQuantity" type="text" class="textfield" id="OrderQuantity" style="WIDTH: 140;" value="<%= OrderQuantity %>" maxlength="100" onBlur="return checkNum(this)" readonly="true"></td>
        <td height="20" align="left">经办人：</td>
        <td>
		<input name="Agenter" type="text" class="textfield" id="Agenter" style="WIDTH: 140;" value="<%= Agenter %>" maxlength="100"  readonly="true">
		</td>
      </tr>
      <tr>
        <td height="20" align="left">下单日期：</td>
        <td><input name="OrderDate" type="text" class="textfield" id="OrderDate" style="WIDTH: 140;" value="<%= OrderDate %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">客户交期：</td>
        <td><input name="CustomDate" type="text" class="textfield" id="CustomDate" style="WIDTH: 140;" value="<%= CustomDate %>" maxlength="100" onChange="checkDate(this)" readonly="true"></td>
        <td height="20" align="left">生管回复交期：</td>
        <td>
		<input name="MCReplyDate" type="text" class="textfield" id="MCReplyDate" style="WIDTH: 140;" value="<%= MCReplyDate %>" maxlength="100"  readonly="true">
		</td>
      </tr>
      <tr>
        <td height="20" align="left">异常类别：</td>
        <td>
		<select id="AbnormalType" name="AbnormalType">
		  <option value="交期太长" <% If AbnormalType="交期太长" Then Response.Write("selected")%>>交期太长</option>
		  <option value="交期延后" <% If AbnormalType="交期延后" Then Response.Write("selected")%>>交期延后</option>
		  <option value="产前样未确认" <% If AbnormalType="产前样未确认" Then Response.Write("selected")%>>产前样未确认</option>
		  <option value="急单" <% If AbnormalType="急单" Then Response.Write("selected")%>>急单</option>
		  <option value="插单" <% If AbnormalType="插单" Then Response.Write("selected")%>>插单</option>
		  <option value="订单下错" <% If AbnormalType="数量异常" Then Response.Write("selected")%>>订单下错</option>
		  <option value="订单暂停" <% If AbnormalType="订单暂停" Then Response.Write("selected")%>>订单暂停</option>
		  <option value="分批交货" <% If AbnormalType="分批交货" Then Response.Write("selected")%>>分批交货</option>
		  <option value="订单异常未回复" <% If AbnormalType="订单异常未回复" Then Response.Write("selected")%>>订单异常未回复</option>
		  <option value="订单重新确认" <% If AbnormalType="订单重新确认" Then Response.Write("selected")%>>订单重新确认</option>
		  <option value="客户更改" <% If AbnormalType="客户更改" Then Response.Write("selected")%>>客户更改</option>
		  <option value="客供品未提供" <% If AbnormalType="客供品未提供" Then Response.Write("selected")%>>客供品未提供</option>
		  <option value="客人样品未确认" <% If AbnormalType="客人样品未确认" Then Response.Write("selected")%>>客人样品未确认</option>
		</select>
		</td>
        <td height="20" align="left">是否损失：</td>
        <td>
		<select id="IsLoss" name="IsLoss">
		  <option value="否" <% If IsLoss="否" Then Response.Write("selected")%>>否</option>
		  <option value="是" <% If IsLoss="是" Then Response.Write("selected")%>>是</option>
		</select>
		</td>
        <td height="20" align="left">损失金额：</td>
        <td>
		<input name="LossAmount" type="text" class="textfield" id="LossAmount" style="WIDTH: 140;" value="<%= LossAmount %>" maxlength="100"  onBlur="return checkNum(this)">
		</td>
      </tr>
      <tr>
        <td height="20" align="left">异常描述：</td>
        <td colspan="5">
  <textarea name="AbnormalNote" id="AbnormalNote" style="width:'540px'; height:'80px'; "><%= AbnormalNote %></textarea>
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
  <tr>  <td height="1" colspan="4">  </td>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td width="60"> 改善人 </td>
 <td width="60">
 <input name="Replyer" type="text" id="Replyer" value="<%= Replyer %>"></td>
 <td width="60"> 改善日期 </td>
 <td width="60">
 <input name="ReplyDate" type="text" id="ReplyDate" value="<%= ReplyDate %>"></td>
</tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 改善意见 </td>
<td colspan="3">
  <textarea name="ReplyText" id="ReplyText" style="width:500px; height:100px; "><%= ReplyText %></textarea>
</td>
</tr> 
  <tr>  <td height="1" colspan="4">  </td>
  </tr>
	<tr>
	  <td align="center" colspan="4">
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="改善" style="WIDTH: 80; visibility:<%= style3 %>;"  onClick="javascript:$('#detailType').val('Check');toSubmit(this);">&nbsp;
			<input name="submitend" type="button" class="button"  id="submitend" value="结案" style="WIDTH: 80; visibility:<%= style3 %>;"  onClick="javascript:$('#detailType').val('End');toSubmit(this);">&nbsp;
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
  if detailType="AddNew" and Instr(session("AdminPurviewFLW"),"|203.1,")>0 then
	set rs = server.createobject("adodb.recordset")
	sql="select * from Bill_OrderAbnormal"
	rs.open sql,connzxpt,1,3
	rs.addnew
	rs("Biller")=UserName
	rs("BillDate")=now()
	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("ReceivDepartment")=Request("ReceivDepartment")
	rs("OrderID")=Request("OrderID")
	rs("Product")=Request("Product")
	rs("ProductType")=Request("ProductType")
	rs("CustomRanke")=Request("CustomRanke")
	rs("CustomLevel")=Request("CustomLevel")
	rs("OrderState")=Request("OrderState")
	rs("OrderQuantity")=Request("OrderQuantity")
	rs("Agenter")=Request("Agenter")
	rs("AbnormalType")=Request("AbnormalType")
	rs("AbnormalNote")=Request("AbnormalNote")
	rs("CustomID")=Request("CustomID")
	rs("IsLoss")=Request("IsLoss")
	rs("LossAmount")=Request("LossAmount")
	rs("OrderDate")=Request("OrderDate")
	rs("CustomDate")=Request("CustomDate")
	rs("MCReplyDate")=Request("MCReplyDate")
	SendMail "sg@loverdoor.cn","订单异常反馈单-新增",Request("OrderID"),Request("AbnormalNote"),""
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Edit" and Instr(session("AdminPurviewFLW"),"|203.1,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_OrderAbnormal where SerialNum="&SerialNum
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
	rs("ReceivDepartment")=Request("ReceivDepartment")
	rs("OrderID")=Request("OrderID")
	rs("Product")=Request("Product")
	rs("ProductType")=Request("ProductType")
	rs("CustomRanke")=Request("CustomRanke")
	rs("CustomLevel")=Request("CustomLevel")
	rs("OrderState")=Request("OrderState")
	rs("OrderQuantity")=Request("OrderQuantity")
	rs("Agenter")=Request("Agenter")
	rs("AbnormalType")=Request("AbnormalType")
	rs("AbnormalNote")=Request("AbnormalNote")
	rs("CustomID")=Request("CustomID")
	rs("IsLoss")=Request("IsLoss")
	rs("LossAmount")=Request("LossAmount")
	rs("OrderDate")=Request("OrderDate")
	rs("CustomDate")=Request("CustomDate")
	rs("MCReplyDate")=Request("MCReplyDate")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Delete" and Instr(session("AdminPurviewFLW"),"|203.1,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_OrderAbnormal where SerialNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_OrderAbnormal where SerialNum="&SerialNum)
	response.write "###"
  elseif detailType="Check" and Instr(session("AdminPurviewFLW"),"|203.2,")>0 then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		dim strMail:strMail=""
	sql="select * from Bill_OrderAbnormal where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
		wherestr="FName='"&rs("Agenter")&"' or FNumber='"&rs("Register")&"' or FNumber='"&rs("Biller")&"'"
    rs("Replyer")=session("AdminName")
    rs("Replydate")=now()
    rs("ReplyText")=Request("ReplyText")
	rs.update
		set rs = server.createobject("adodb.recordset")
		sql="select distinct FEmail from t_Base_Emp where "&wherestr
		rs.open sql,connk3,1,1
		do until rs.eof
			strMail=strMail&rs("FEmail")
	    rs.movenext
		If Not rs.eof Then
		  strMail=strMail&";"
		End If
    loop
		SendMail strMail,"订单异常反馈单-回复",Request("OrderID"),Request("ReplyText")&"("&session("AdminName")&")",""
		rs.close
		set rs=nothing 
		response.write("###"&Request("ReplyText")&"###")
  elseif detailType="End" and Instr(session("AdminPurviewFLW"),"|203.3,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_OrderAbnormal where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
    rs("CheckFlag")=1
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
	sql="select a.fheadselfs0143,c.fname,d.fnumber,d.f_105,a.fheadselfs0153,b.fdate,e.fname as name2,f.fname as name3,b.fqty,g.fdate11 from seorder a inner join  "&_
" seorderentry b on a.finterid=b.finterid left join  "&_
" t_emp c on a.fempid=c.fitemid left join "&_
" t_Organization d on a.fcustid=d.fitemid left join "&_
" t_ICItem e on b.fitemid=e.fitemid left join "&_
" t_Item f on LEFT(e.FNumber, 7) = f.FNumber and f.FItemClassID = 4 inner join "&_
" t_dhtzdentry g on g.FEntryID=b.FSourceEntryID "&_
" where fbillno='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
    if rs.bof and rs.eof then
        response.write ("定单号不存在！")
        response.end
	else
	  response.write(rs("fheadselfs0143")&"###"&rs("fnumber")&"###"&rs("fname")&"###"&rs("f_105")&"###"&rs("fheadselfs0143")&"###"&rs("fheadselfs0153")&"###"&rs("fdate")&"###"&rs("name3")&"###"&rs("fqty")&"###"&rs("Fdate11")&"###<select id='Product' name='Product' onChange='changePro()' style='width:140'>")
    end if
	while (not rs.eof)
	  response.write("<option value='"&rs("name2")&"'>"&rs("name2")&"@"&rs("fdate")&"@"&rs("name3")&"@"&rs("fqty")&"@"&rs("fdate11")&"</option>")
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
