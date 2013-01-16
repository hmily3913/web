<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>编辑</TITLE>
<link rel="stylesheet" href="../Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/CustomAjax.js"></script>
<script language="javascript" >
function toSubmit(obj){
  var submitname=document.getElementById("Keyword");
  var theform=document.getElementById("editForm");
//  var checkflag=document.getElementById("CheckFlag").value;
  var snumber=document.getElementById("SerialNum").value;
  var subflag=true;
  switch (obj.value){
    case "保存" :
	  submitname.value="SaveEdit";
	  break;
    case "删除" :
	  submitname.value="Delete";
	  break;
	default : 
	  break; 
  }
  if(subflag){
  obj.disabled = true; 
  theform.submit();
  }else{
    alert("该单据当前状态不允许此操作！");
  }
}
</script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|1004,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result,Action,SerialNum,AdminName,UserName
Result=request.QueryString("Result")
Action=request.QueryString("Action")
SerialNum=request("SerialNum")
UserName=session("UserName")
AdminName=session("AdminName")
dim i,j '用于循环的整数
i=0
'定义保安单主表变量
dim Classes,OccurTime,OccurAddr,MainPerson,MainPersonName,Details,ReasonAnaly,Measure
dim PunishResult,LossAmount,FBase1,FBase1Name,Remark,FBiller,FBillerName,FDate
call ProcessFun()
if Result="Security" then

%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>保安单查看：添加，修改，删除宿舍人员信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="SecurityEdit.asp?Result=Security&Action=Add" onClick='changeAdminFlag("添加宿舍人员信息")'>添加保安单信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="SecurityMana.asp?Result=Person" onClick='changeAdminFlag("保安单列表")'>查看所有保安单信息</a></td>
  </tr>
</table>
<br>
  <form name="editForm" id="editForm" method="post" action="SecurityEdit.asp?Result=Security&Action=<%=Action%>&SerialNum=<%=SerialNum%>">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews idth="100%">

      <tr>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left">单据号：</td>
        <td>
		<input name="SerialNum" type="text" class="textfield" id="SerialNum" style="WIDTH: 140;" value="<%=SerialNum%>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">类别：</td>
        <td>
		<select name="Classes" id="Classes" >
		<option value="偷盗" <%if Classes="偷盗" then response.write ("selected")%>>偷盗</option>
		<option value="打架" <%if Classes="打架" then response.write ("selected")%>>打架</option>
		<option value="赌博" <%if Classes="赌博" then response.write ("selected")%>>赌博</option>
		</select></td>
        <td height="20" align="left">发生时间：</td>
        <td><input name="OccurTime" type="text" class="textfield" id="OccurTime" style="WIDTH: 140;" value="<%=OccurTime%>" maxlength="100" onBlur="return checkFullTime(this)"></td>
      </tr>
      <tr>
        <td height="20" align="left">发生地点：</td>
        <td colspan="3"><input name="OccurAddr" type="text" class="textfield" id="OccurAddr" style="WIDTH: 350;" value="<%=OccurAddr%>" maxlength="200"></td>
        <td height="20" align="left">损失金额：</td>
        <td><input name="LossAmount" type="text" class="textfield" id="LossAmount" style="WIDTH: 140;" value="<%=LossAmount%>" maxlength="100" onChange="return checkNum(this)"></td>
      </tr>
      <tr>
        <td height="20" align="left">主要人员：</td>
        <td>
		<input name="MainPerson" type="hidden" id="MainPerson" value="<%=MainPerson%>">
		<input name="MainPersonName" type="text" class="textfield" id="MainPersonName" style="WIDTH: 140;" value="<%=MainPersonName%>" maxlength="100" onBlur="getEmpName(this)">
		</td>
        <td height="20" align="left">部门：</td>
        <td><input name="FBase1" type="hidden" id="FBase1" value="<%=FBase1%>">
		<input name="FBase1name" type="text" class="textfield" id="FBase1name" style="WIDTH: 140;" value="<%=FBase1name%>" maxlength="100" onBlur="return getDepartment(this)"></td>
        <td height="20" align="left"></td>
        <td></td>
      </tr>
      <tr>
        <td height="20" align="left">详细情况：</td>
        <td colspan="5"><input name="Details" type="text" class="textfield" id="Details" style="WIDTH: 550;" value="<%=Details%>" maxlength="500"></td>
      </tr>
      <tr>
        <td height="20" align="left">原因分析：</td>
        <td colspan="5"><input name="ReasonAnaly" type="text" class="textfield" id="ReasonAnaly" style="WIDTH: 550;" value="<%=ReasonAnaly%>" maxlength="500"></td>
      </tr>
      <tr>
        <td height="20" align="left">改正及预防措施：</td>
        <td colspan="5"><input name="Measure" type="text" class="textfield" id="Measure" style="WIDTH: 550;" value="<%=Measure%>" maxlength="200"></td>
      </tr>
      <tr>
        <td height="20" align="left">处罚结果：</td>
        <td colspan="5"><input name="PunishResult" type="text" class="textfield" id="PunishResult" style="WIDTH: 550;" value="<%=PunishResult%>" maxlength="200"></td>
      </tr>
      <tr>
        <td height="20" align="left">备注：</td>
        <td colspan="5"><input name="Remark" type="text" class="textfield" id="Remark" style="WIDTH: 550;" value="<%=Remark%>" maxlength="500"></td>
      </tr>
      <tr>
        <td height="20" align="left">制单人：</td>
        <td><input type="hidden" name="FBiller" id="FBiller" value="<%=FBiller%>">
		<input name="FBillerName" type="text" class="textfield" id="FBillerName" style="WIDTH: 140;" value="<%=FBillerName%>" maxlength="100" readonly></td>
        <td height="20" align="left">制单日期：</td>
        <td><input name="FDate" type="text" class="textfield" id="FDate" style="WIDTH: 140;" value="<%=FDate%>" maxlength="100" readonly></td>
        <td height="20" align="left"></td>
        <td></td>
      </tr>
 
      <tr>
        <td height="20" align="left" colspan="3">&nbsp;</td>
        <td valign="bottom" colspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td valign="bottom" colspan="6" align="center">
		<input type="hidden" name="Keyword" id="Keyword" value="">
		&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">
		&nbsp;<input name="delete" type="button" class="button"  id="delete" value="删除" style="WIDTH: 80;"  onClick="toSubmit(this)">
		</td>
      </tr>
      <tr>
        <td height="20" align="left" colspan="3">&nbsp;</td>
        <td valign="bottom" colspan="3">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
  </form>
<%
end if
%>
</BODY>
</HTML>

<%
sub ProcessFun()
  dim Keyword,rsRepeat,rs,sql
  Keyword=request("Keyword")
  if Keyword="SaveEdit" then '保存事务处理
	  if Action="Add" then '增加记录
		  '主表信息添加
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_SecurityMana"
		  rs.open sql,connk3,1,3
		  rs.addnew
		  rs("Classes")=Request.Form("Classes")
		  rs("OccurTime")=Request.Form("OccurTime")
		  rs("Department")=Request.Form("FBase1")
		  rs("OccurAddr")=Request.Form("OccurAddr")
		  rs("MainPerson")=Request.Form("MainPerson")
		  rs("Details")=Request.Form("Details")
		  rs("ReasonAnaly")=Request.Form("ReasonAnaly")
		  rs("FBiller")=Request.Form("FBiller")
		  rs("FDate")=Request.Form("FDate")
		  rs("LossAmount")=Request.Form("LossAmount")
		  rs("Measure")=Request.Form("Measure")
		  rs("PunishResult")=Request.Form("PunishResult")
		  rs("Remark")=Request.Form("Remark")
		  rs.update
		  rs.close
		  set rs=nothing 
		response.write "<script language=javascript> alert('成功增加保安信息！');changeAdminFlag('保安管理信息');location.replace('SecurityMana.asp');</script>"
	  end if
	  if Action="Modify" then '修改记录
	  	'保存主表信息编辑
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_SecurityMana where SerialNum="& SerialNum
		  rs.open sql,connk3,1,3
		  if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		  end if
		  rs("Classes")=Request.Form("Classes")
		  rs("OccurTime")=trim(Request.Form("OccurTime"))
		  rs("Department")=Request.Form("FBase1")
		  rs("OccurAddr")=Request.Form("OccurAddr")
		  rs("MainPerson")=Request.Form("MainPerson")
		  rs("Details")=Request.Form("Details")
		  rs("ReasonAnaly")=Request.Form("ReasonAnaly")
		  rs("LossAmount")=Request.Form("LossAmount")
		  rs("Measure")=Request.Form("Measure")
		  rs("PunishResult")=Request.Form("PunishResult")
		  rs("Remark")=Request.Form("Remark")
		  rs.update
		  rs.close
		  set rs=nothing 
		response.write "<script language=javascript> alert('成功编辑保安信息！');changeAdminFlag('保安管理信息');location.replace('SecurityMana.asp');</script>"
	  end if
  elseif Keyword="Delete" then
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_SecurityMana where SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  else
		  sql="delete from z_SecurityMana where SerialNum="& SerialNum
		  connk3.execute(sql)
	  end if
	  rs.close
	  set rs=nothing 
	response.write "<script language=javascript> alert('保安信息删除成功！');changeAdminFlag('保安管理信息');location.replace('SecurityMana.asp');</script>"
  else
  	if Action="Modify" then'提出编辑信息
	  '提取主表信息
	  set rs = server.createobject("adodb.recordset")
      sql="select * from z_SecurityMana where SerialNum="& SerialNum
      rs.open sql,connk3,1,1
      if rs.bof and rs.eof then
        response.write ("数据库读取记录出错！")
        response.end
      end if
	  Classes=rs("Classes")
	  MainPerson=rs("MainPerson")
	  MainPersonName=getUser(rs("MainPerson"))
	  FBase1=rs("Department")
	  FBase1Name=getDepartment(rs("Department"))
	  OccurTime=rs("OccurTime")
	  OccurAddr=rs("OccurAddr")
	  Details=rs("Details")
      ReasonAnaly=rs("ReasonAnaly")
	  Remark=rs("Remark")
      FBiller=rs("FBiller")
	  FBillerName=getUser(rs("FBiller"))
      FDate=rs("FDate")
	  Measure=rs("Measure")
	  PunishResult=rs("PunishResult")
	  LossAmount=rs("LossAmount")
	  rs.close
      set rs=nothing 
	else'提取增加时所需信息,制单人，制单日期，单号
	  OccurTime=now()
      FBiller=UserName
	  FBillerName=AdminName
	  LossAmount="0"
	  Remark=" "
      FDate=date()
	  
	end if
  end if
end sub

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
  sql="Select * From t_Department where Fitemid="&ID
  rs.open sql,connk3,1,1
  getDepartment=rs("Fname")
  rs.close
  set rs=nothing
End Function    
%>