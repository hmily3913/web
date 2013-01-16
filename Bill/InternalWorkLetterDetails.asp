<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!--#include file="../CheckAdmin.asp" -->
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
<BODY>
<%
if Instr(session("AdminPurviewFLW"),"|206,")=0 then 
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
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>事项名称</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>发出日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>发出人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>发出部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>接收部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>审批进度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>抄送</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>制单人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>审核人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>审批人</strong></font></td>
  </tr>
 <%
  dim Depart,wherestr,seachword,flag4search,type4search
  wherestr=""
  seachword=request("seachword")
  if seachword<>"" then
  wherestr=" and (ProjectDescrib like '%"&seachword&"%' or ProjectName like '%"&seachword&"%')"
  end if
  flag4search=request("flag4search")
  if flag4search<>"" then wherestr=wherestr&" and CheckFlag="&flag4search

  dim page'页码
      page=clng(request("Page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Bill_InternalWorkLetter "
  dim datawhere'数据条件
		 datawhere="where 1=1"&wherestr&Session("AllMessage10")&Session("AllMessage11")&Session("AllMessage12")
		 Session("AllMessage10")=""
		 Session("AllMessage11")=""
		 Session("AllMessage12")=""
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
    sql="select *,left(ProjectDescrib,10) as a1,left(ProjectName,10) as a2,left(Signman,10) as a3 from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
		if rs("CheckFlag")="1" then
		  formdata(0)="已审核"
		  bgcolors="#ffff66"'黄色
		elseif rs("CheckFlag")="2" then
		  formdata(0)="已批准"
		  bgcolors="#66ff66"'粉色
		else
		  formdata(0)="未审核"
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a2")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegisterName")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Departmentname")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ReceivDepartment")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&formdata(0)&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Ccman")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Biller")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Checker")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Approvaler")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  else
    response.write "<tr><td height='50' align='center' colspan='11' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
	response.Write("###"&pagec&"###"&idcount&"###"&datawhere)
elseif showType="AddEditShow" then 
  dim detailType
  detailType=request("detailType")
'数据处理
  dim SerialNum,Register,RegisterName,RegDate,Department,Departmentname,ReceivDepartment,Ccman
  dim ProjectName,ProjectDescrib,Signman
  dim style1,style2,style3,buttonval
  dim CheckFlag,Biller,BillDate,Checker,CheckDate,Approvaler,ApprovalDate
  if detailType="AddNew" then
	sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&UserName&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
	Register=UserName
	RegisterName=AdminName
	RegDate=date()
	Department=rs("部门别")
	Departmentname=rs("部门名称")
	CheckFlag=0
	style1="block;"
	style2="none;"
  elseif detailType="Edit" or detailType="Check" then
    SerialNum=request("SerialNum")
	sql="select * from Bill_InternalWorkLetter where SerialNum="&SerialNum
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	Register=rs("Register")
	RegisterName=rs("RegisterName")
	RegDate=rs("RegDate")
	Department=rs("Department")
	Departmentname=rs("Departmentname")
	ReceivDepartment=rs("ReceivDepartment")
	ProjectName=rs("ProjectName")
	ProjectDescrib=rs("ProjectDescrib")
	Ccman=rs("Ccman")
	Signman=rs("Signman")

	Checker=rs("Checker")
	CheckDate=rs("CheckDate")
	Approvaler=rs("Approvaler")
	ApprovalDate=rs("ApprovalDate")
	Biller=rs("Biller")
	BillDate=rs("BillDate")
	CheckFlag=rs("CheckFlag")
	style1="none;"
	style2="none;"
	style3="none;"
	if CheckFlag=0 then
	style1="block;"
	elseif CheckFlag=1 and Instr(session("AdminPurviewFLW"),"|206.3,")>0 then
	style2="block;"
	buttonval="审批"
	elseif CheckFlag=2 and Instr(session("AdminPurviewFLW"),"|206.4,")>0 then
	style2="block;"
	buttonval="会签"
	elseif Register=UserName or rs("Biller")=UserName then
	style1="block;"
	end if
  end if
  %>
 <div id="AddandEditdiv" style="width:100%; height:460px; overflow:auto; ">
 <form id="AddForm" name="AddForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>内部联络函进度处理</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews >
      <tr>
        <td width="120" height="20" align="left">单据号：</td>
        <td>
		<input name="SerialNum" type="text" class="textfield" id="SerialNum" style="WIDTH: 140;" value="<%= SerialNum %>" maxlength="100" readonly="true"></td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left">发出人工号：</td>
        <td><input name="Register" type="text" class="textfield" id="Register" style="WIDTH: 140;" value="<%= Register %>" maxlength="100" onBlur="getInfo('Register')"></td>
        <td height="20" align="left">发出人姓名：</td>
        <td><input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 140;" value="<%= RegisterName %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">发出日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 140;" value="<%= RegDate %>" maxlength="100" onBlur="checkDate(this)"></td>
      </tr>
      <tr>
        <td height="20" align="left">发出部门编号：</td>
        <td>
		<input name="Department" type="text" class="textfield" id="Department" style="WIDTH: 140;" value="<%= Department %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">发出部门名称：</td>
        <td>
		<input name="Departmentname" type="text" class="textfield" id="Departmentname" style="WIDTH: 140;" value="<%= Departmentname %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left"></td>
        <td>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">接收部门：</td>
        <td colspan="5">
		  <input type="checkbox" name="RD1" id="ReceivDepartment1" value="财务部," <% If Instr(ReceivDepartment,"财务部")>0 Then Response.Write("checked")%>><label for="ReceivDepartment1">财务部</label>
		  <input type="checkbox" name="RD2" id="ReceivDepartment2" value="总经办," <% If Instr(ReceivDepartment,"总经办")>0 Then Response.Write("checked")%>><label for="ReceivDepartment2">总经办</label>
		  <input type="checkbox" name="RD3" id="ReceivDepartment3" value="一分厂," <% If Instr(ReceivDepartment,"一分厂")>0 Then Response.Write("checked")%>><label for="ReceivDepartment3">一分厂</label>
		  <input type="checkbox" name="RD4" id="ReceivDepartment4" value="二分厂," <% If Instr(ReceivDepartment,"二分厂")>0 Then Response.Write("checked")%>><label for="ReceivDepartment4">二分厂</label>
		  <input type="checkbox" name="RD5" id="ReceivDepartment5" value="三分厂," <% If Instr(ReceivDepartment,"三分厂")>0 Then Response.Write("checked")%>><label for="ReceivDepartment5">三分厂</label>
		  <input type="checkbox" name="RD6" id="ReceivDepartment6" value="营销部," <% If Instr(ReceivDepartment,"营销部")>0 Then Response.Write("checked")%>><label for="ReceivDepartment6">营销部</label>
		  <input type="checkbox" name="RD7" id="ReceivDepartment7" value="采购部," <% If Instr(ReceivDepartment,"采购部")>0 Then Response.Write("checked")%>><label for="ReceivDepartment7">采购部</label>
		  <input type="checkbox" name="RD8" id="ReceivDepartment8" value="工程部," <% If Instr(ReceivDepartment,"工程部")>0 Then Response.Write("checked")%>><label for="ReceivDepartment8">工程部</label>
		  <input type="checkbox" name="RD9" id="ReceivDepartment9" value="生管部," <% If Instr(ReceivDepartment,"生管部")>0 Then Response.Write("checked")%>><label for="ReceivDepartment9">生管部</label>
		  <input type="checkbox" name="RD10" id="ReceivDepartment10" value="仓储科," <% If Instr(ReceivDepartment,"仓储科")>0 Then Response.Write("checked")%>><label for="ReceivDepartment10">仓储科</label>
		  <input type="checkbox" name="RD11" id="ReceivDepartment11" value="品保部," <% If Instr(ReceivDepartment,"品保部")>0 Then Response.Write("checked")%>><label for="ReceivDepartment11">品保部</label>
		  <input type="checkbox" name="RD12" id="ReceivDepartment12" value="生技部," <% If Instr(ReceivDepartment,"生技部")>0 Then Response.Write("checked")%>><label for="ReceivDepartment12">生技部</label>
		  <input type="checkbox" name="RD13" id="ReceivDepartment13" value="人资部," <% If Instr(ReceivDepartment,"人资部")>0 Then Response.Write("checked")%>><label for="ReceivDepartment13">人资部</label>
		  <input type="checkbox" name="RD14" id="ReceivDepartment14" value="娄桥," <% If Instr(ReceivDepartment,"娄桥")>0 Then Response.Write("checked")%>><label for="ReceivDepartment14">娄桥</label>
		  <input type="checkbox" name="RD15" id="ReceivDepartment15" value="蓝驰," <% If Instr(ReceivDepartment,"蓝驰")>0 Then Response.Write("checked")%>><label for="ReceivDepartment15">蓝驰</label>
		  <input type="checkbox" name="RD16" id="ReceivDepartment16" value="蓝劲," <% If Instr(ReceivDepartment,"蓝劲")>0 Then Response.Write("checked")%>><label for="ReceivDepartment16">蓝劲</label>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">抄送：</td>
        <td colspan="2">
		<input name="Ccman" type="text" class="textfield" id="Ccman" style="WIDTH: 200;" value="<%= Ccman %>" maxlength="100"  readonly="true" <% If CheckFlag<1 Then Response.Write("onFocus='ShowCcmanDiv()'")%>>
<div id="CcmanDiv" style="width:'202px';height:'200px';position:absolute;display:none;text-align:center;width:0px;height:0px;overflow:visible; z-index:9999">
<table id="ReplyTable" border="0" width="200px" cellspacing="0" cellpadding="1" align="center" bgcolor="black" height="100%">
<tbody id="TbDetails">
<tr style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%;border-bottom: 1px solid #000000;">
<td  bgcolor="#8DB5E9" height="24"  width="70">工号</td>
<td  bgcolor="#8DB5E9" height="24"  width="90">姓名</td>
<td  bgcolor="#8DB5E9" width="40">操作</td>
</tr>
<tr height="24" id="CloneNodeTr" style='background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%;border-bottom: 1px solid #000000;cursor:hand; display:none'>
 <td >
 <input name="PeronD" type="text" class="textfield" style="WIDTH: 65;" id="PeronD" onChange="getEmpName(this)"></td>
 <td >
		<input name="PeronDName" type="text" class="textfield" id="PeronDName" style="WIDTH: 88;" value="" maxlength="20" onChange="getEmpName(this)"></td>
 <td  align="right" onClick="deleted(this)"><img src="../images/close.jpg"/></td>
</tr>
</tbody>
<tr style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%;border-bottom: 1px solid #000000;">
<td valign="bottom" colspan="3" align="center">
&nbsp;<input style='HEIGHT: 18px;WIDTH: 40px;' name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="关闭"  onClick="closead()">
&nbsp;<input style='HEIGHT: 18px;WIDTH: 40px;' name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="增加"   onClick="AddRow()">
&nbsp;<input style='HEIGHT: 18px;WIDTH: 40px;' name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="确定"  onClick="SaveRow()">
</td>
</tr>
</tbody>
</table>
</div>
		
		</td>
        <td height="20" align="left">事项名称：</td>
        <td colspan="2">
		<input name="ProjectName" type="text" class="textfield" id="ProjectName" style="WIDTH: 200;" value="<%= ProjectName %>" maxlength="100">
		</td>
      </tr>
      <tr>
		<td> 事项内容：</td>
		<td colspan="5">
		  <textarea name="ProjectDescrib" id="ProjectDescrib" class="xheditor {skin:'vista',width:'100%',height:'250',upImgUrl:'../Include/UpLoadAjax.asp?immediate=1',upLinkUrl:'../Include/UpLoadAjax.asp?immediate=1',upLinkExt:'doc,xls'}"><%= ProjectDescrib %></textarea>
		</td>
		</tr> 
      <tr>
        <td height="20" align="left">制单：</td>
        <td >
		<%= Biller %>|<%= BillDate %>
		</td>
        <td height="20" align="left">审核：</td>
        <td >
		<%= Checker %>|<%= CheckDate %>
		</td>
        <td height="20" align="left">批准：</td>
        <td >
		<%= Approvaler %>|<%= ApprovalDate %>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">会签：</td>
        <td colspan="5">
		<%= Signman %>
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
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="审核" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('Check');toSubmit(this);">
			<input type="button" class="button"  value="反审核" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('反审核');toSubmit(this);">&nbsp;
			<input type="button" class="button" value="关闭" style="WIDTH: 80;"  onClick='$("#addDiv").hide()'>&nbsp;
	  </td>
	</tr>
  <tr>  <td height="5">  </td></tr>
  </table>
</div>
<div id="Buttondiv" style="display:<%= style2 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="1" colspan="4">  </td></tr>
	<tr>
	  <td align="center" colspan="4">
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="<%= buttonval %>" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val(this.value);toSubmit(this);">&nbsp;
			<input type="button" class="button"  value="反审核" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('反审核');toSubmit(this);">&nbsp;
			<input type="button" class="button" value="关闭" style="WIDTH: 80;"  onClick='$("#addDiv").hide()'>&nbsp;
	  </td>
	</tr>
  <tr>  <td height="1" colspan="4">  </td></tr>
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
  if detailType="AddNew" then
	set rs = server.createobject("adodb.recordset")
	sql="select * from Bill_InternalWorkLetter"
	rs.open sql,connzxpt,1,3
	rs.addnew
	rs("Biller")=AdminName
	rs("BillDate")=now()
	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("ReceivDepartment")=Request("RD1")&Request("RD2")&Request("RD3")&Request("RD4")&Request("RD5")&Request("RD6")&Request("RD7")&Request("RD8")&Request("RD9")&Request("RD10")&Request("RD11")&Request("RD12")&Request("RD13")&Request("RD14")&Request("RD15")&Request("RD16")
	rs("Ccman")=Request("Ccman")
	rs("ProjectName")=Request("ProjectName")
	rs("ProjectDescrib")=Request("ProjectDescrib")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Edit" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_InternalWorkLetter where (Biller='"&AdminName&"' or Register='"&UserName&"') and CheckFlag=0 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许编辑，只能编辑本人的单据！")
		response.end
	end if
	rs("Biller")=AdminName
	rs("BillDate")=now()
	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("ReceivDepartment")=Request("RD1")&Request("RD2")&Request("RD3")&Request("RD4")&Request("RD5")&Request("RD6")&Request("RD7")&Request("RD8")&Request("RD9")&Request("RD10")&Request("RD11")&Request("RD12")&Request("RD13")
	rs("Ccman")=Request("Ccman")
	rs("ProjectName")=Request("ProjectName")
	rs("ProjectDescrib")=Request("ProjectDescrib")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_InternalWorkLetter where (Biller='"&AdminName&"' or Register='"&UserName&"') and CheckFlag=0 and SerialNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("当前状态不允许删除，只能删除本人的单据！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_InternalWorkLetter where SerialNum="&SerialNum)
	response.write "###"
  elseif detailType="Check" and Instr(session("AdminPurviewFLW"),"|206.2,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_InternalWorkLetter where CheckFlag=0 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许审核！")
		response.end
	end if
	rs("CheckFlag")="1"
    rs("Checker")=AdminName
    rs("Checkdate")=now()
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&session("AdminName")&"###")
  elseif detailType="反审核" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_InternalWorkLetter where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许审核！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|206.2,")>0 and rs("CheckFlag")=1 then
		rs("CheckFlag")=0
    rs("Checker")=AdminName
    rs("Checkdate")=now()
	elseif Instr(session("AdminPurviewFLW"),"|206.3,")>0 and rs("CheckFlag")=2 then
		rs("CheckFlag")=1
    rs("Approvaler")=AdminName
    rs("ApprovalDate")=now()
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&session("AdminName")&"###")
  elseif detailType="审批" and Instr(session("AdminPurviewFLW"),"|206.3,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_InternalWorkLetter where CheckFlag=1 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许此操作！")
		response.end
	end if
	rs("CheckFlag")="2"
    rs("Approvaler")=AdminName
    rs("ApprovalDate")=now()
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&session("AdminName")&"###")
  elseif detailType="会签" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_InternalWorkLetter where CheckFlag=2 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许此操作，只能反馈本人单据！")
		response.end
	end if
	if Instr(rs("Signman"),AdminName)=0 then
    rs("Signman")=rs("Signman")&session("AdminName")&","
	rs.update
	end if
	rs.close
	set rs=nothing 
	response.write("###"&session("AdminName")&"###")
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="Register" then
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
  elseif detailType="getEmpName" then
    InfoID=request("InfoID")
	sql="select a.姓名,a.员工代号 from [N-基本资料单头] a where a.员工代号 like '%"&InfoID&"%' or a.姓名 like '%"&InfoID&"%'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write("###"&rs("员工代号")&"###"&rs("姓名")&"###")
	end if
	rs.close
	set rs=nothing 
  end if
end if
 %>
</body>
</html>
