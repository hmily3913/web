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
if Instr(session("AdminPurviewFLW"),"|202,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
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
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>打样单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品型号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客户排行</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客户等级</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单状态</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>经办人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>异常类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>改善回复</strong></font></td>
  </tr>
 <%
  dim Depart,wherestr,seachword
  wherestr=""
  seachword=request("seachword")
  if seachword<>"" then
  wherestr=" and SerialNum="&seachword
  end if
  dim page'页码
      page=clng(request("Page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Bill_ProofingAbnormal "
  dim datawhere'数据条件
		 datawhere="where 1=1"&wherestr&Session("AllMessage5")
		 Session("AllMessage5")=""
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
    sql="select *,left(ReplyText,10) as a2 from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
		if Len(rs("ReplyText"))>0 then
		  bgcolors="#ff99ff"'粉色
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegisterName")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Departmentname")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ProofingID")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Product")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ProductType")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("CustomRanke")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("CustomLevel")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("OrderState")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("OrderQuantity")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Agenter")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("AbnormalType")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Check',"&rs("SerialNum")&")"">"&rs("a2")&"</td>" & vbCrLf
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
  dim SerialNum,Register,RegisterName,RegDate,Department,Departmentname,ProofingID,Product
  dim ProductType,CustomRanke,CustomLevel,OrderState,OrderQuantity,Agenter,AbnormalType,AbnormalNote
  dim CustomID,CheckFlag,style1,style2,style3,Replyer,ReplyText,ReplyDate,Pic
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
	sql="select * from Bill_ProofingAbnormal where SerialNum="&SerialNum
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	Register=rs("Register")
	RegisterName=rs("RegisterName")
	RegDate=rs("RegDate")
	Department=rs("Department")
	Departmentname=rs("Departmentname")
	ProofingID=rs("ProofingID")
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
	Replyer=rs("Replyer")
	ReplyText=rs("ReplyText")
	ReplyDate=rs("ReplyDate")
	Pic=rs("Pic")
	if detailType="Edit" then
	style1="block;"
	style2="none;"
	else
	style1="none;"
	style2="block;"
	if Instr(session("AdminPurviewFLW"),"|202.2,")>0 then
	  style3="inline;"
	else
	  style3="none;"
	end if
	end if	
  end if
  %>
 <div id="AddandEditdiv" style="width:100%; height:100%; overflow:auto; ">
 <form id="AddForm" name="AddForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap id="formove"><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>打样异常反馈处理</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews>
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
        <td height="20" align="left">打样单号：</td>
        <td><input name="ProofingID" type="text" class="textfield" id="ProofingID" style="WIDTH: 140;" value="<%= ProofingID %>" maxlength="100" onChange="getInfo('ProofingID')"></td>
      </tr>
      <tr>
        <td height="20" align="left">产品型号：</td>
        <td id="Product_td">
		<input name="Product" type="text" class="textfield" id="Product" style="WIDTH: 140;" value="<%= Product %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">产品类别：</td>
        <td>
		<input name="ProductType" type="text" class="textfield" id="ProductType" style="WIDTH: 140;" value="<%= ProductType %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">客户代号：</td>
        <td><input name="CustomID" type="text" class="textfield" id="CustomID" style="WIDTH: 140;" value="<%= CustomID %>" maxlength="100"  readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">客户排行：</td>
        <td><input name="CustomRanke" type="text" class="textfield" id="CustomRanke" style="WIDTH: 140;" value="<%= CustomRanke %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">客户等级：</td>
        <td>
		<input name="CustomLevel" type="text" class="textfield" id="CustomLevel" style="WIDTH: 140;" value="<%= CustomLevel %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">订单状态：</td>
        <td>
		<select id="OrderState" name="OrderState">
		  <option value="已接订单" <% If OrderState="已接订单" Then Response.Write("selected")%>>已接订单</option>
		  <option value="已谈订单" <% If OrderState="已谈订单" Then Response.Write("selected")%>>已谈订单</option>
		  <option value="未知订单" <% If OrderState="未知订单" Then Response.Write("selected")%>>未知订单</option>
		  <option value="没有订单" <% If OrderState="没有订单" Then Response.Write("selected")%>>没有订单</option>
		</select>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">订单数量：</td>
        <td><input name="OrderQuantity" type="text" class="textfield" id="OrderQuantity" style="WIDTH: 140;" value="<%= OrderQuantity %>" maxlength="100" onBlur="return checkNum(this)"></td>
        <td height="20" align="left">经办人：</td>
        <td>
		<input name="Agenter" type="text" class="textfield" id="Agenter" style="WIDTH: 140;" value="<%= Agenter %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">异常类别：</td>
        <td>
		<select id="AbnormalType" name="AbnormalType">
		  <option value="操作不规范" <% If AbnormalType="操作不规范" Then Response.Write("selected")%>>操作不规范</option>
		  <option value="审核不及时" <% If AbnormalType="审核不及时" Then Response.Write("selected")%>>审核不及时</option>
		  <option value="交期异常" <% If AbnormalType="交期异常" Then Response.Write("selected")%>>交期异常</option>
		  <option value="样品未确认" <% If AbnormalType="样品未确认" Then Response.Write("selected")%>>样品未确认</option>
		  <option value="描述不规范" <% If AbnormalType="描述不规范" Then Response.Write("selected")%>>描述不规范</option>
		  <option value="数量异常" <% If AbnormalType="数量异常" Then Response.Write("selected")%>>数量异常</option>
		  <option value="图稿未提供" <% If AbnormalType="图稿未提供" Then Response.Write("selected")%>>图稿未提供</option>
		  <option value="型号下错" <% If AbnormalType="型号下错" Then Response.Write("selected")%>>型号下错</option>
		  <option value="业务员错误" <% If AbnormalType="业务员错误" Then Response.Write("selected")%>>业务员错误</option>
		  <option value="客人赶空运" <% If AbnormalType="客人赶空运" Then Response.Write("selected")%>>客人赶空运</option>
		  <option value="采购物料延期" <% If AbnormalType="采购物料延期" Then Response.Write("selected")%>>采购物料延期</option>
		  <option value="客户提前交货" <% If AbnormalType="客户提前交货" Then Response.Write("selected")%>>客户提前交货</option>
		</select>
		</td>
      </tr>
      <tr>
        <td height="80" align="left">异常描述：</td>
        <td colspan="5">
        <textarea name="AbnormalNote" id="AbnormalNote" class="xheditor {skin:'vista',width:'100%',height:'200px',upImgUrl:'../Include/UpLoadAjax.asp?immediate=1',upLinkUrl:'../Include/UpLoadAjax.asp?immediate=1',upLinkExt:'doc,xls'}"><%= AbnormalNote %></textarea>
</td>
      </tr>
      <tr>
        <td height="20" align="left">图片：</td>
        <td colspan="5">
		直接在上面输入框中选择图片按钮进行上传图片！
		</td>
      </tr>
	  </table>
<div id="Buttondiv" style="display:<%= style1 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="5">  </td>
  </tr>
	<tr>
	  <td align="center">
	  <input type="hidden" name="CheckFlag" id="CheckFlag" value="<%= CheckFlag %>">
	  <input type="hidden" name="detailType" id="detailType" value="<%= detailType %>">
			<input name="submitSaveAdd" type="button" class="button"  id="submitSaveAdd" value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">&nbsp;
			<input name="submitDelete" type="button" class="button"  id="submitDelete" value="删除" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('Delete');toSubmit(this);">
	  </td>
	</tr>
  </table>
</div>
<div id="Buttondiv" style="display:<%= style2 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="5" colspan="4">  </td>
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
  <tr>  <td height="5" colspan="4">  </td>
  </tr>
	<tr>
	  <td align="center" colspan="4">
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="改善" style="WIDTH: 80; display:<%= style3 %>;"  onClick="toSubmit(this)">&nbsp;
	  </td>
	</tr>
  </table>
</div>
	</td>
  </tr>
</table>
</form>
<!--
  <form action="../Include/UpFileSave.asp" method="post" enctype="multipart/form-data" name="formUpload" id="formUpload">
	<TABLE BORDER=0 style="display:<%= style1 %>;">
  <tr>
    <td bgcolor="#EBF2F9" id="callbackHtml">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="60" height="30" nowrap>选择文件：</td>
        <td><input name="FromFile" type="file" id="FromFile" size="41" class="multi" accept="gif|jpg|bmp|jpeg" maxlength="2"/></td>
      </tr>
      <tr>
        <td height="36" colspan="2" align="center" valign="bottom">
	<input type="hidden" name="SaveToPath" id="SaveToPath" value="../Upload/">
	<input type="hidden" name="detailType" value="test">
          &nbsp;<input name="Submit" type="button" class="button" value=" 上传 " onClick="uploadPic()">
		  </td>
        </tr>
    </table>
	</td>
  </tr>
	</TABLE>
</FORM>
-->
</div>
  <%
    rs.close
    set rs=nothing
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" and Instr(session("AdminPurviewFLW"),"|202.1,")>0 then
	set rs = server.createobject("adodb.recordset")
	sql="select * from Bill_ProofingAbnormal"
	rs.open sql,connzxpt,1,3
	rs.addnew
	rs("Biller")=UserName
	rs("BillDate")=now()
	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("ProofingID")=Request("ProofingID")
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
'	rs("Pic")=Request("Pic")
	
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="test" then
  	Dim Uploader, File
	Set Uploader = New FileUploader
	' This starts the upload process
	Uploader.Upload()
	If Uploader.Files.Count > 0 Then
		For Each File In Uploader.Files.Items
			File.SaveToDisk "E:\"
		Next
	End If
	response.write "###"
  elseif detailType="Edit" and Instr(session("AdminPurviewFLW"),"|202.1,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_ProofingAbnormal where SerialNum="&SerialNum
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
	rs("ProofingID")=Request("ProofingID")
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
'	rs("Pic")=Request("Pic")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Delete" and Instr(session("AdminPurviewFLW"),"|202.1,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_ProofingAbnormal where SerialNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_ProofingAbnormal where SerialNum="&SerialNum)
	response.write "###"
  elseif detailType="Check" and Instr(session("AdminPurviewFLW"),"|102.2,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_ProofingAbnormal where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
    rs("Replyer")=session("AdminName")
    rs("Replydate")=now()
    rs("ReplyText")=Request("ReplyText")
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&Request("ReplyText")&"###")
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="ProofingID" then
    InfoID=request("InfoID")
	sql="select * from [K-打样单查询] where 打样单号='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
    if rs.bof and rs.eof then
        response.write ("打样单号不存在！")
        response.end
	else
	  response.write(rs("品质等级")&"###"&rs("打样单号")&"###"&rs("客户代号")&"###"&rs("打样排行")&"###"&rs("品质等级")&"###"&rs("业务员姓名")&"###"&rs("名称")&"###<select id='Product' name='Product' onChange='changePro()'>")
    end if
	while (not rs.eof)
	  response.write("<option value='"&rs("品号")&"'>"&rs("品号")&"/"&rs("名称")&"</option>")
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
