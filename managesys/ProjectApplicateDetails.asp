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
if Instr(session("AdminPurview"),"|1007,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
dim KeyWord
 KeyWord=request("KeyWord")
' response.Write(KeyWord&"####")
 if KeyWord="ShowList" then
%>
 <div id="listtable" style="width:100%; height:410px; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>计划申报项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>项目类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>经办部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>经办人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>计划申报日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>应准备资料及内容</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申报情况说明</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>应配合部门/人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>计划完成报批日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进度跟踪情况</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>请报部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>联系人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>联系电话</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>注意事项</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>实际批复情况</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>备注</strong></font></td>
  </tr>
 <%
  dim page'页码
      page=clng(request("Page"))
  dim datafrom'数据表名
      datafrom=" z_ProjectApplicate "
  dim datawhere'数据条件
'		 datawhere="where a.fid=b.fid and a.fuser>0 and ftext9!='Y'"+wherestr
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by SerialNum desc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select SerialNum from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
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
    dim sql2,rs2
	dim formdata(3),bgcolors
    sql="select a.*,b.fname as name1,c.fname as name2,e.fname as name3,f.fname as name4,left(AppliCase,10) as a1 from z_ProjectApplicate a left join t_item b on b.fitemid=a.RegistDepartment left join t_item c on c.fitemid=a.AgencyDepartment "&_
	" left join t_emp e on e.fnumber=a.Register left join t_emp f on f.fnumber=a.Agencyer where a.SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
		if rs("StatusFlag")=1 then
		  bgcolors="#ffff66"'黄色
		end if
		if rs("StatusFlag")=2 then
		  bgcolors="#ff99ff"'粉色
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand' onDblClick='ShowEdit("&rs("Serialnum")&")'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("name1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("name3")&"</td>"
      Response.Write "<td nowrap>"&rs("ProjectName")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ProjectClass")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("name2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("name4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("PlanAppliDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("MaterialContent")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("AppliCase")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Coordinater")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("PlanFinishDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ProgessTrack")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("QuoteDepart")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Contacter")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ContactPhone")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Note")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ActualReply")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Remark")&"</td>" & vbCrLf
'      Response.Write "<td nowrap onDblClick=""return S6ClickTd(this,'T9reply',"&rs("fentryid")&")"">"&rs("FText9")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  else
    response.write "<tr><td height='50' align='center' colspan='19' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
elseif KeyWord = "ShowAdd" then
 %>
<div id="AddDiv" style="width:'780px';height:'200px';background-color:#888888;overflow-y: hidden; overflow-x: hidden;">
<form name="AddForm" id="AddForm" action="test1.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews idth="100%">
      <tr>
        <td height="20" align="left">项目名称：</td>
        <td>
		<input name="ProjectName" type="text" class="textfield" id="ProjectName" style="WIDTH: 140;" value="" maxlength="100" ></td>
        <td height="20" align="left">申请人：</td>
        <td>
		<input name="Register" type="hidden" id="Register" value="<%= session("UserName") %>">
		<input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 140;" value="<%= session("AdminName") %>" maxlength="100" onBlur="getEmpName(this)" >
		</td>
        <td height="20" align="left">申请日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 140;" value="<%= date() %>" maxlength="100" onBlur="return checkDate(this)" ></td>
      </tr>
      <tr>
        <td height="20" align="left">申请部门：</td>
        <td><input name="RegistDepartment" type="hidden" id="RegistDepartment" value="">
		<input name="RegistDepartmentname" type="text" class="textfield" id="RegistDepartmentname" style="WIDTH: 140;" value="" maxlength="100" onBlur="return getDepartment(this)" ></td>
        <td height="20" align="left">经办人：</td>
        <td>
		<input name="Agencyer" type="hidden" id="Agencyer" value="">
		<input name="AgencyerName" type="text" class="textfield" id="AgencyerName" style="WIDTH: 140;" value="" maxlength="100" onBlur="getEmpName(this)" >
		</td>
        <td height="20" align="left">经办部门：</td>
        <td><input name="AgencyDepartment" type="hidden" id="AgencyDepartment" value="">
		<input name="AgencyDepartmentname" type="text" class="textfield" id="AgencyDepartmentname" style="WIDTH: 140;" value="" maxlength="100" onBlur="return getDepartment(this)" ></td>
      </tr>
      <tr>
        <td height="20" align="left">项目类别：</td>
        <td>
		<select name="ProjectClass" id="ProjectClass"  >
		<option value="类别一" >类别一</option>
		</select></td>
        <td height="20" align="left">计划申报日期：</td>
        <td><input name="PlanAppliDate" type="text" class="textfield" id="PlanAppliDate" style="WIDTH: 140;" value="<%= date() %>" maxlength="100" onBlur="return checkDate(this)" ></td>
        <td height="20" align="left">应配合部门/人：</td>
        <td ><input name="Coordinater" type="text" class="textfield" id="Coordinater" style="WIDTH: 140;" maxlength="100" ></td>
      </tr>
      <tr>
        <td height="20" align="left" width="100">资料及内容：</td>
        <td width="120"><input name="MaterialContent" type="text" class="textfield" id="MaterialContent" style="WIDTH: 140;" maxlength="100"></td>
        <td height="20" align="left" width="100">情况说明：</td>
        <td colspan="3"><input name="AppliCase" type="text" class="textfield" id="AppliCase" style="WIDTH: 300;" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left">计划完成报批日期：</td>
        <td><input name="PlanFinishDate" type="text" class="textfield" id="PlanFinishDate" style="WIDTH: 140;" value="<%= date() %>" maxlength="100" onBlur="return checkDate(this)" ></td>
        <td height="20" align="left" width="100">进度跟踪情况：</td>
        <td colspan="3"><input name="ProgessTrack" type="text" class="textfield" id="ProgessTrack" style="WIDTH: 300;" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left">请报部门：</td>
        <td><input name="QuoteDepart" type="text" class="textfield" id="QuoteDepart" style="WIDTH: 140;" value="" maxlength="100" ></td>
        <td height="20" align="left">联系人：</td>
        <td><input name="Contacter" type="text" class="textfield" id="Contacter" style="WIDTH: 140;" value="" maxlength="100" ></td>
        <td height="20" align="left">联系电话：</td>
        <td><input name="ContactPhone" type="text" class="textfield" id="ContactPhone" style="WIDTH: 140;" value="" maxlength="100" ></td>
      </tr>
      <tr>
        <td height="20" align="left">注意事项：</td>
        <td><input name="Note" type="text" class="textfield" id="Note" style="WIDTH: 140;" value="" maxlength="100" ></td>
        <td height="20" align="left" width="100">实际批复情况：</td>
        <td colspan="3"><input name="ActualReply" type="text" class="textfield" id="ActualReply" style="WIDTH: 300;" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left" width="100">备注：</td>
        <td colspan="5"><input name="Remark" type="text" class="textfield" id="Remark" value="" style="WIDTH: 450;" maxlength="450" ></td>
      </tr>
      <tr>
        <td valign="bottom" colspan="6" align="center">
		<input type="hidden" name="Keyword" id="Keyword" value="SaveAdd">
		&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">
		&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="关闭" style="WIDTH: 80;"  onClick="closead1(this)">
		</td>
      </tr>
   </table>
	</td>
  </tr>
</table>
</form>
</div>
<% 
elseif KeyWord = "SaveAdd" and Instr(session("AdminPurview"),"|1007.1,")>0 then
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_ProjectApplicate"
		  rs.open sql,connk3,1,3
		  rs.addnew
		  rs("RegDate")=Request("RegDate")
		  rs("Register")=trim(Request("Register"))
		  rs("RegistDepartment")=Request("RegistDepartment")
		  rs("ProjectName")=Request("ProjectName")
		  rs("ProjectClass")=Request("ProjectClass")
		  rs("Agencyer")=Request("Agencyer")
		  rs("AgencyDepartment")=Request("AgencyDepartment")
		  if Request("PlanAppliDate") <> "" then rs("PlanAppliDate")=Request("PlanAppliDate")
		  rs("MaterialContent")=Request("MaterialContent")
		  rs("AppliCase")=Request("AppliCase")
		  rs("Coordinater")=Request("Coordinater")
		  if Request("PlanFinishDate") <> "" then rs("PlanFinishDate")=Request("PlanFinishDate")
		  rs("ProgessTrack")=Request("ProgessTrack")
		  rs("QuoteDepart")=Request("QuoteDepart")
		  rs("Contacter")=Request("Contacter")
		  rs("ContactPhone")=Request("ContactPhone")
		  rs("Note")=Request("Note")
		  rs("ActualReply")=Request("ActualReply")
		  rs("Remark")=Request("Remark")
		  rs.update
		  rs.close
		  set rs=nothing 
elseif KeyWord = "ShowEdit" then
dim SerialNum
  SerialNum=request("SerialNum")
  sql="select * from z_ProjectApplicate where SerialNum="& SerialNum
  set rs = server.createobject("adodb.recordset")
  rs.open sql,connk3,1,1
 %>
<div id="EditDiv" style="width:'780px';height:'200px';background-color:#888888;overflow-y: hidden; overflow-x: hidden;">
<form name="EditForm" id="EditForm" action="test1.asp">
<%
      if rs.bof and rs.eof then
        response.write ("数据库读取记录出错！")
        response.end
	  else
 %>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews idth="100%">
      <tr>
        <td height="20" align="left">项目名称：</td>
        <td>
		<input name="ProjectName" type="text" class="textfield" id="ProjectName" style="WIDTH: 140;" value="<%= rs("ProjectName") %>" maxlength="100" ></td>
        <td height="20" align="left">申请人：</td>
        <td>
		<input name="Register" type="hidden" id="Register" value="<%= rs("Register") %>">
		<input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 140;" value="<%= getUser(rs("Register")) %>" maxlength="100" onBlur="getEmpName(this)" >
		</td>
        <td height="20" align="left">申请日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 140;" value="<%= rs("RegDate") %>" maxlength="100" onBlur="return checkDate(this)" ></td>
      </tr>
      <tr>
        <td height="20" align="left">申请部门：</td>
        <td><input name="RegistDepartment" type="hidden" id="RegistDepartment" value="<%= rs("RegistDepartment") %>">
		<input name="RegistDepartmentname" type="text" class="textfield" id="RegistDepartmentname" style="WIDTH: 140;" value="<%= getDepartment(rs("RegistDepartment")) %>" maxlength="100" onBlur="return getDepartment(this)" ></td>
        <td height="20" align="left">经办人：</td>
        <td>
		<input name="Agencyer" type="hidden" id="Agencyer" value="<%= rs("Agencyer") %>">
		<input name="AgencyerName" type="text" class="textfield" id="AgencyerName" style="WIDTH: 140;" value="<%= getUser(rs("Agencyer")) %>" maxlength="100" onBlur="getEmpName(this)" >
		</td>
        <td height="20" align="left">经办部门：</td>
        <td><input name="AgencyDepartment" type="hidden" id="AgencyDepartment" value="<%= rs("AgencyDepartment") %>">
		<input name="AgencyDepartmentname" type="text" class="textfield" id="AgencyDepartmentname" style="WIDTH: 140;" value="<%= getDepartment(rs("AgencyDepartment")) %>" maxlength="100" onBlur="return getDepartment(this)" ></td>
      </tr>
      <tr>
        <td height="20" align="left">项目类别：</td>
        <td>
		<select name="ProjectClass" id="ProjectClass"  >
		<option value="类别一" <%if rs("ProjectClass")="类别一" then response.write ("selected")%>>类别一</option>
		</select></td>
        <td height="20" align="left">计划申报日期：</td>
        <td><input name="PlanAppliDate" type="text" class="textfield" id="PlanAppliDate" style="WIDTH: 140;" value="<%= rs("PlanAppliDate") %>" maxlength="100" onBlur="return checkDate(this)" ></td>
        <td height="20" align="left">应配合部门/人：</td>
        <td ><input name="Coordinater" type="text" class="textfield" id="Coordinater" style="WIDTH: 140;"  value="<%= rs("Coordinater") %>" maxlength="100" ></td>
      </tr>
      <tr>
        <td height="20" align="left" width="100">资料及内容：</td>
        <td width="120"><input name="MaterialContent" type="text" class="textfield" id="MaterialContent" value="<%= rs("MaterialContent") %>" style="WIDTH: 140;" maxlength="100"></td>
        <td height="20" align="left" width="100">情况说明：</td>
        <td colspan="3"><input name="AppliCase" type="text" class="textfield" id="AppliCase" value="<%= rs("AppliCase") %>" style="WIDTH: 300;" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left">计划完成报批日期：</td>
        <td><input name="PlanFinishDate" type="text" class="textfield" id="PlanFinishDate" style="WIDTH: 140;" value="<%= rs("PlanFinishDate") %>" maxlength="100" onBlur="return checkDate(this)" ></td>
        <td height="20" align="left" width="100">进度跟踪情况：</td>
        <td colspan="3"><input name="ProgessTrack" type="text" class="textfield" id="ProgessTrack" value="<%= rs("ProgessTrack") %>" style="WIDTH: 300;" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left">请报部门：</td>
        <td><input name="QuoteDepart" type="text" class="textfield" id="QuoteDepart" style="WIDTH: 140;" value="<%= rs("QuoteDepart") %>" maxlength="100" ></td>
        <td height="20" align="left">联系人：</td>
        <td><input name="Contacter" type="text" class="textfield" id="Contacter" style="WIDTH: 140;" value="<%= rs("Contacter") %>" maxlength="100" ></td>
        <td height="20" align="left">联系电话：</td>
        <td><input name="ContactPhone" type="text" class="textfield" id="ContactPhone" style="WIDTH: 140;" value="<%= rs("ContactPhone") %>" maxlength="100" ></td>
      </tr>
      <tr>
        <td height="20" align="left">注意事项：</td>
        <td><input name="Note" type="text" class="textfield" id="Note" style="WIDTH: 140;" value="<%= rs("Note") %>" maxlength="100" ></td>
        <td height="20" align="left" width="100">实际批复情况：</td>
        <td colspan="3"><input name="ActualReply" type="text" class="textfield" id="ActualReply" value="<%= rs("ActualReply") %>" style="WIDTH: 300;" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left" width="100">备注：</td>
        <td colspan="5"><input name="Remark" type="text" class="textfield" id="Remark" value="<%= rs("Remark") %>" style="WIDTH: 450;" maxlength="450" ></td>
      </tr>
      <tr>
        <td valign="bottom" colspan="6" align="center">
		<input type="hidden" name="Keyword" id="Keyword" value="SaveEdit">
		<input type="hidden" name="SerialNum" id="SerialNum" value="<%=  SerialNum%>">
		&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;"  onClick="toSubmitEdit(this)">
		&nbsp;<input name="delete" type="button" class="button"  id="delete" value="删除" style="WIDTH: 80;"  onClick="toDelete(this)">
		&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="关闭" style="WIDTH: 80;"  onClick="closead1(this)">
		</td>
      </tr>
   </table>
	</td>
  </tr>
</table>
</form>
<% 
	end if
	%>
</div>
<% 
	  rs.close
      set rs=nothing 
elseif KeyWord = "SaveEdit" and Instr(session("AdminPurview"),"|1007.1,")>0 then
  		SerialNum=request("SerialNum")
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_ProjectApplicate where SerialNum="& SerialNum
		  rs.open sql,connk3,1,3
		  if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		  end if
		  rs("RegDate")=Request("RegDate")
		  rs("Register")=trim(Request("Register"))
		  rs("RegistDepartment")=Request("RegistDepartment")
		  rs("ProjectName")=Request("ProjectName")
		  rs("ProjectClass")=Request("ProjectClass")
		  rs("Agencyer")=Request("Agencyer")
		  rs("AgencyDepartment")=Request("AgencyDepartment")
		  if Request("PlanAppliDate") <> "" then rs("PlanAppliDate")=Request("PlanAppliDate")
		  rs("MaterialContent")=Request("MaterialContent")
		  rs("AppliCase")=Request("AppliCase")
		  rs("Coordinater")=Request("Coordinater")
		  if Request("PlanFinishDate") <> "" then rs("PlanFinishDate")=Request("PlanFinishDate")
		  rs("ProgessTrack")=Request("ProgessTrack")
		  rs("QuoteDepart")=Request("QuoteDepart")
		  rs("Contacter")=Request("Contacter")
		  rs("ContactPhone")=Request("ContactPhone")
		  rs("Note")=Request("Note")
		  rs("ActualReply")=Request("ActualReply")
		  rs("Remark")=Request("Remark")
		  rs.update
		  rs.close
		  set rs=nothing 
elseif KeyWord = "Delete" and Instr(session("AdminPurview"),"|1007.1,")>0 then
  		SerialNum=request("SerialNum")
		  sql="Delete from z_ProjectApplicate where SerialNum="& SerialNum
		  connk3.Execute(sql)
else
		response.write ("你没有权限进行此操作！@@@")
		response.end
end if
Function getUser(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From t_emp where fnumber='"&ID&"'"
  rs.open sql,connk3,1,1
  if rs.bof and rs.eof then
  getUser=""
  else
  getUser=rs("Fname")
  end if
  rs.close
  set rs=nothing
End Function    
Function getDepartment(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From t_item where fitemclassid=2 and Fitemid="&ID
  rs.open sql,connk3,1,1
  getDepartment=rs("Fname")
  rs.close
  set rs=nothing
End Function    
 %>
</body>
</html>
