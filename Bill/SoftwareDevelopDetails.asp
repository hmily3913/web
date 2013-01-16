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
if Instr(session("AdminPurviewFLW"),"|205,")=0 then 
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
  <tr class="TitleRow">
    <td nowrap bgcolor="#8DB5E9" class="TitleCol"><font color="#FFFFFF"><strong>单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" class="TitleCol"><font color="#FFFFFF"><strong>项目名称</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>针对部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>接收日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>接收人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>需求描述</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>需求分析</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>存在问题</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>开发状态</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>计划完成日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>实际完成日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>实施进度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>实施人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>实施日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>实施进度说明</strong></font></td>
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
      datafrom=" Bill_SoftwareDevelop "
  dim datawhere'数据条件
		 datawhere="where 1=1"&wherestr
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
    sql="select *,left(ProjectDescrib,10) as a1,left(DevelopAnaly,10) as a2,left(Schedule,10) as a3,left(Problem,10) as a4 from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
		if rs("CheckFlag")="1" then
		  formdata(0)="已审核"
		  bgcolors="#ffff66"'黄色
		elseif rs("CheckFlag")="2" then
		  formdata(0)="开发中"
		  bgcolors="#ff99ff"'粉色
		elseif rs("CheckFlag")="3" then
		  formdata(0)="已完成"
		  bgcolors="#66ff66"'绿色
		else
		  formdata(0)="未审核"
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td class='DataCol' nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td class='DataCol' nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ProjectName")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegisterName")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Departmentname")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("According")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ReceivDate")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Receiver")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a1")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a2")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a4")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&formdata(0)&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("PlanFinishDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ActualFinishDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ScheduleProgress")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Scheduler")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ScheduleDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a3")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  else
    response.write "<tr><td height='50' align='center' colspan='18' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
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
  dim SerialNum,Register,RegisterName,RegDate,Department,Departmentname,ReceivDate,Receiver,ReceivDepartment
  dim According,ProjectName,ProjectDescrib,DevelopAnaly,Problem,PlanDevelopDate
  dim style1,style2,style3,buttonval
  dim PlanFinishDate,ActualDevelopDate,ActualFinishDate,NextProjectDescrib,Schedule,ScheduleProgress
  dim CheckFlag,Scheduler,ScheduleDate
  if detailType="AddNew" then
	sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&UserName&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
	Register=UserName
	RegisterName=AdminName
	RegDate=now()
	Department=rs("部门别")
	Departmentname=rs("部门名称")
	style1="block;"
	style2="none;"
  elseif detailType="Edit" or detailType="Check" then
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftwareDevelop where SerialNum="&SerialNum
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	Register=rs("Register")
	RegisterName=rs("RegisterName")
	RegDate=rs("RegDate")
	Department=rs("Department")
	Departmentname=rs("Departmentname")
	ReceivDepartment=rs("ReceivDepartment")
	ReceivDate=rs("ReceivDate")
	Receiver=rs("Receiver")
	According=rs("According")
	ProjectName=rs("ProjectName")
	ProjectDescrib=rs("ProjectDescrib")
	DevelopAnaly=rs("DevelopAnaly")
	Problem=rs("Problem")
	PlanDevelopDate=rs("PlanDevelopDate")
	PlanFinishDate=rs("PlanFinishDate")
	ActualDevelopDate=rs("ActualDevelopDate")
	ActualFinishDate=rs("ActualFinishDate")
	NextProjectDescrib=rs("NextProjectDescrib")
	Schedule=rs("Schedule")
	ScheduleProgress=rs("ScheduleProgress")
	Scheduler=rs("Scheduler")
	ScheduleDate=rs("ScheduleDate")
	CheckFlag=rs("CheckFlag")
	style1="none;"
	style2="none;"
	style3="none;"
	if CheckFlag=1 and Instr(session("AdminPurviewFLW"),"|205.3,")>0 then
	PlanDevelopDate=now()
	PlanFinishDate=now()
	style2="block;"
	buttonval="接收"
	style3="block;"
	elseif CheckFlag=2 then
	style2="block;"
	if Instr(session("AdminPurviewFLW"),"|205.3,")>0 then 
	style3="block;"
	buttonval="进度更新"
	end if
	elseif CheckFlag=3 then
	style2="block;"
	if Register=UserName or rs("Biller")=UserName then 
	style3="block;"
	end if
	elseif Register=UserName or rs("Biller")=UserName then
	style1="block;"
	end if
  end if
  %>
 <div id="AddandEditdiv" style="width:790px; height:460px; overflow:auto; ">
 <form id="AddForm" name="AddForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>程序开发处理进度处理</strong></font></td>
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
        <td height="20" align="left">申请人工号：</td>
        <td><input name="Register" type="text" class="textfield" id="Register" style="WIDTH: 140;" value="<%= Register %>" maxlength="100" onBlur="getInfo('Register')"></td>
        <td height="20" align="left">申请人姓名：</td>
        <td><input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 140;" value="<%= RegisterName %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">申请日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 140;" value="<%= RegDate %>" maxlength="100"  readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">申请部门编号：</td>
        <td>
		<input name="Department" type="text" class="textfield" id="Department" style="WIDTH: 140;" value="<%= Department %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">申请部门名称：</td>
        <td>
		<input name="Departmentname" type="text" class="textfield" id="Departmentname" style="WIDTH: 140;" value="<%= Departmentname %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">接收部门：</td>
        <td>
		<select id="ReceivDepartment" name="ReceivDepartment">
		  <option value="财务部" <% If ReceivDepartment="财务部" Then Response.Write("selected")%>>财务部</option>
		  <option value="总经办" <% If ReceivDepartment="总经办" Then Response.Write("selected")%>>总经办</option>
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
		  <option value="人资部" <% If ReceivDepartment="人资部" Then Response.Write("selected")%>>人资部</option>
		</select>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">项目名称：</td>
        <td ><input name="ProjectName" type="text" class="textfield" id="ProjectName" style="WIDTH: 140;" value="<%= ProjectName %>" maxlength="100"></td>
        <td height="20" align="left">针对部门：</td>
        <td ><input name="According" type="text" class="textfield" id="According" style="WIDTH: 140;" value="<%= According %>" maxlength="100"></td>
        <td height="20" align="left"></td>
        <td ></td>
      </tr>
      <tr>
        <td height="20" align="left">审核状态：</td>
        <td >
		<% If CheckFlag =1 Then Response.Write("已审核")%>
		<% If CheckFlag =2 Then Response.Write("开发中")%>
		<% If CheckFlag =3 Then Response.Write("已完成")%>
		</td>
        <td height="20" align="left">接收人：</td>
        <td >		<%= Receiver %>		</td>
        <td height="20" align="left">接收日期：</td>
        <td >		<%= ReceivDate %>		</td>
      </tr>
      <tr>
		<td> 需求描述：</td>
		<td colspan="5">
		  <textarea name="ProjectDescrib" id="ProjectDescrib" style="width:'500px'; height:'50px'; "><%= ProjectDescrib %></textarea>
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
			<input name="submitDelete" type="button" class="button"  id="submitDelete" value="审核" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('Check');toSubmit(this);">
	  </td>
	</tr>
  <tr>  <td height="5">  </td></tr>
  </table>
</div>
<div id="Buttondiv" style="display:<%= style2 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="1" colspan="4">  </td></tr>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td height="20" align="left">计划完成日期：</td>
 <td>
<input name="PlanFinishDate" type="text" class="textfield" style="width:140px " id="PlanFinishDate" value="<%= PlanFinishDate %>" onBlur="return checkFullTime(this)">
 </td>
<td height="20" align="left">开发进度：</td>
 <td>
 <select id="DevelopProcess" name="DevelopProcess">
 <option value="2" <% If CheckFlag="2" Then Response.Write("selected")%>>开发中</option>
 <option value="3" <% If CheckFlag="3" Then Response.Write("selected")%>>开发完毕</option>
 </select>
 </td>
</tr>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td width="100"> 接收人 </td>
 <td width="60">
 <input name="Receiver" class="textfield" type="text" id="Receiver" value="<%= Receiver %>"></td>
 <td width="60"> 接收日期 </td>
 <td width="60">
 <input name="ReceivDate" class="textfield" type="text" id="ReceivDate" value="<%= ReceivDate %>"></td>
</tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="100"> 需求分析 </td>
<td colspan="3">
  <textarea name="DevelopAnaly" id="DevelopAnaly" style="width:'500px'; height:'50px'; "><%= DevelopAnaly %></textarea>
</td>
</tr> 
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 可能问题 </td>
<td colspan="3">
 <input name="Problem" class="textfield" style="width:500px; " type="text" id="Problem" value="<%= Problem %>"></td>
</tr> 
  <tr>  <td height="1" colspan="4">  </td></tr>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td height="20" align="left">实施日期：</td>
 <td>
<input name="ScheduleDate" type="text" class="textfield" style="width:140px " id="ScheduleDate" value="<%= ScheduleDate %>" onBlur="return checkDate(this)">
 </td>
<td height="20" align="left">实施进度：</td>
 <td>
 <select id="ScheduleProgress" name="ScheduleProgress">
 <option value="未实施" <% If ScheduleProgress="未实施" Then Response.Write("selected")%>>未实施</option>
 <option value="实施中" <% If ScheduleProgress="实施中" Then Response.Write("selected")%>>实施中</option>
 <option value="实施完毕" <% If ScheduleProgress="实施完毕" Then Response.Write("selected")%>>实施完毕</option>
 </select>
 </td>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="100"> 实施进度说明 </td>
<td colspan="3">
 <input name="Schedule" class="textfield" style="width:500px; " type="text" id="Schedule" value="<%= Schedule %>"></td>
</tr> 
  </table>
</div>
<div id="Buttondiv" style="display:<%= style2 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="1" colspan="4">  </td></tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 后续需求 </td>
<td colspan="3">
  <textarea name="NextProjectDescrib" id="NextProjectDescrib" style="width:'500px'; height:'50px'; "><%= NextProjectDescrib %></textarea>
</td>
</tr> 
  <tr>  <td height="5" colspan="4">  </td></tr>
	<tr>
	  <td align="center" colspan="4">
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="<%= buttonval %>" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val(this.value);toSubmit(this);">&nbsp;
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="后续" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val(this.value);toSubmit(this);">&nbsp;
	  </td>
	</tr>
  <tr>  <td height="5">  </td></tr>
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
	sql="select * from Bill_SoftwareDevelop"
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
	rs("According")=Request("According")
	rs("ProjectName")=Request("ProjectName")
	rs("ProjectDescrib")=Request("ProjectDescrib")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Edit" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftwareDevelop where (Biller='"&UserName&"' or Register='"&UserName&"') and CheckFlag=0 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许编辑，只能编辑本人的单据！")
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
	rs("According")=Request("According")
	rs("ProjectName")=Request("ProjectName")
	rs("ProjectDescrib")=Request("ProjectDescrib")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftwareDevelop where (Biller='"&UserName&"' or Register='"&UserName&"') and CheckFlag=0 and SerialNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("当前状态不允许删除，只能删除本人的单据！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_SoftwareDevelop where SerialNum="&SerialNum)
	response.write "###"
  elseif detailType="Check" and Instr(session("AdminPurviewFLW"),"|205.2,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftwareDevelop where CheckFlag=0 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许审核！")
		response.end
	end if
	rs("CheckFlag")="1"
'    rs("Checker")=session("AdminName")
'    rs("Checkdate")=now()
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&session("AdminName")&"###")
  elseif detailType="接收" and Instr(session("AdminPurviewFLW"),"|205.3,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftwareDevelop where CheckFlag=1 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许此操作！")
		response.end
	end if
	rs("CheckFlag")="2"
    rs("Receiver")=session("AdminName")
    rs("ReceivDate")=now()
    rs("DevelopAnaly")=Request("DevelopAnaly")
	rs("Problem")=Request("Problem")
	rs("PlanFinishDate")=Request("PlanFinishDate")
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&session("AdminName")&"###")
  elseif detailType="进度更新" and Instr(session("AdminPurviewFLW"),"|205.3,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftwareDevelop where CheckFlag=2 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许此操作！")
		response.end
	end if
	if Request("DevelopProcess")="3" and rs("CheckFlag")=2 then rs("ActualFinishDate")=now()
	rs("CheckFlag")=Request("DevelopProcess")
    rs("DevelopAnaly")=Request("DevelopAnaly")
    rs("Problem")=Request("Problem")
	if Request("ScheduleDate")<>"" and Request("Schedule")<>"未实施" then rs("ScheduleDate")=Request("ScheduleDate")
    rs("ScheduleProgress")=Request("ScheduleProgress")
    rs("Schedule")=Request("Schedule")
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&Request("RReplyText")&"###")
  elseif detailType="后续" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_SoftwareDevelop where CheckFlag>1 and (Biller='"&UserName&"' or Register='"&UserName&"') and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许此操作，只能反馈本人单据！")
		response.end
	end if
	rs("CheckFlag")="2"
    rs("NextProjectDescrib")=Request("NextProjectDescrib")
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&Request("NextProjectDescrib")&"###")
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="Register" then
    InfoID=request("InfoID")
	sql="select a.员工代号,a.姓名,a.部门别,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and (a.员工代号 like '%"&InfoID&"%' or a.姓名 like '%"&InfoID&"%')"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write("###"&rs("员工代号")&"###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###")
	end if
	rs.close
	set rs=nothing 
  end if
end if
 %>
</body>
</html>
