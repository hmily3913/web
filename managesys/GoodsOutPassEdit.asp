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
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" >
function toSubmit(obj){
  if($('#RegDate').val()==""){alert("申请日期不能为空，请检查！");return false;}
  if($('#Register').val()==""){alert("申请人不能为空，请检查！");return false;}
  if($('#RegistDepartment').val()==""){alert("申请部门不能为空，请检查！");return false;}
  var submitname=document.getElementById("Keyword");
  var theform=document.getElementById("editForm");
  var Outcheckflag=document.getElementById("OutCheckFlag").value;
  var Incheckflag=document.getElementById("InCheckFlag").value;
  var snumber=document.getElementById("SerialNum").value;
  var GetOutDate=document.getElementById("GetOutDate").value;
  var GetInDate=document.getElementById("GetInDate").value;
  var subflag=false;
  switch (obj.value){
    case "保存" :
	  if(Outcheckflag!=2)subflag=true;
	  submitname.value="SaveEdit";
	  break;
    case "删除" :
	  if(Outcheckflag==0 && snumber!='')subflag=true;
	  submitname.value="Delete";
	  break;
	case "主管审核-出" :
	  if(Outcheckflag==0 && snumber!='')subflag=true;
	  submitname.value="check1";
	  break;
	case "门卫审核-出" :
	 if(GetOutDate!=''){
	  if(Outcheckflag==1 && snumber!='')subflag=true;
	  submitname.value="check2";
	  break;
	 }else{
	   alert("请输入携出日期");
	   break;
	 }
	case "门卫审核-入" :
	 if(GetInDate!=''){
	  if(Incheckflag==0 && snumber!='')subflag=true;
	  submitname.value="check3";
	  break;
	 }else{
	   alert("请输入取回日期");
	   break;
	 }
	case "主管审核-入" :
	  if(Incheckflag==1 && snumber!='')subflag=true;
	  submitname.value="check4";
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
'if Instr(session("AdminPurview"),"|1006.1,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
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
'定义宿舍水电主表变量
dim RegDate,Register,RegisterName,GetOutDate,GetInDate,FBiller,FBillerName,FDate,OutCheckFlag,RegistDepartment,RegistDepartmentName
dim OutChecker1,OutCheckDate1,OutChecker2,OutCheckDate2,InCheckFlag,InChecker1,InCheckDate1,InChecker2,InCheckDate2
dim strInCheckFlag,strOutCheckFlag
'定义宿舍水电子表变量
dim FEntryID(),Goods(),FQty(),UseState(),ReturnFlag(),FNumber()
call ProcessFun()
if Result="GoodsOut" then

%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>物品携出放行条查看：添加，修改，删除，审核信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" align="center" nowrap  bgcolor="#EBF2F9"><a href="GoodsOutPassEdit.asp?Result=GoodsOut&Action=Add" onClick='changeAdminFlag("添加物品携出信息")'>添加物品携出信息</a><font color="#0000FF">&nbsp;|&nbsp;</font><a href="GoodsOutPassMana.asp" onClick='changeAdminFlag("宿舍水电列表")'>查看所有物品携出信息</a></td>
  </tr>
</table>
<br>
  <form name="editForm" id="editForm" method="post" action="GoodsOutPassEdit.asp?Result=GoodsOut&Action=<%=Action%>&SerialNum=<%=SerialNum%>">
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
        <td><input name="SerialNum" type="text" class="textfield" id="SerialNum" style="WIDTH: 140;" value="<%=SerialNum%>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">申请人：</td>
        <td>
		<input name="Register" type="hidden" id="Register" value="<%=Register%>">
		<input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 140;" value="<%=RegisterName%>" maxlength="100" onBlur="getEmpName(this)" <%if OutCheckFlag=1 then response.Write("readonly") end if%>>
		</td>
        <td height="20" align="left">申请部门：</td>
        <td><input name="RegistDepartment" type="hidden" id="RegistDepartment" value="<%=RegistDepartment%>">
		<input name="RegistDepartmentName" type="text" class="textfield" id="RegistDepartmentName" style="WIDTH: 140;" value="<%=RegistDepartmentName%>" maxlength="100" onBlur="return getDepartment(this)" <%if OutCheckFlag=1 then response.Write("readonly") end if%>></td>
      </tr>
      <tr>
        <td height="20" align="left">申请日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 140;" value="<%=RegDate%>" maxlength="100" onBlur="return checkDate(this)" <%if OutCheckFlag=1 then response.Write("readonly") end if%>></td>
        <td height="20" align="left">放行日期：</td>
        <td><input name="GetOutDate" type="text" class="textfield" id="GetOutDate" style="WIDTH: 140;" value="<%=GetOutDate%>" onBlur="return checkDate(this)" maxlength="100" <%if OutCheckFlag=2 then response.Write("readonly") end if%>></td>
        <td height="20" align="left">取回日期：</td>
        <td><input name="GetInDate" type="text" class="textfield" id="GetInDate" style="WIDTH: 140;" value="<%=GetInDate%>" onBlur="return checkDate(this)" maxlength="100"></td>
      </tr>
      <tr>
        <td height="20" colspan="6">
		<table width="100%" border="0" id="editDetails" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
		  <tbody id="TbDetails">
		  <tr>
			<td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物品名称</strong></font></td>
			<td width="80" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>件数</strong></font></td>
			<td width="80" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>数量</strong></font></td>
			<td width="250" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>携出用途说明</strong></font></td>
			<td width="80" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>是否回厂</strong></font></td>
			<td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong></td>
		  </tr>
		  <tr bgcolor='#EBF2F9' id="CloneNodeTr" onMouseOver = "this.style.backgroundColor = '#FFFFFF'" onMouseOut = "this.style.backgroundColor = ''" style='cursor:hand; display:none'>
		  <td nowrap><input type="hidden" name="FEntryID" id="FEntryID" value="">
		  <input type="text" name="Goods" id="Goods" value="" style="font-size:12px; width:140;"></td>
		  <td nowrap><input type="text" name="FNumber" id="FNumber" value="" style="font-size:12px; size:50; width:60;" onChange="return checkNum(this)"></td>
		  <td nowrap><input type="text" name="FQty" id="FQty" value="" style="font-size:12px; width:80;" onChange="return checkNum(this)"></td>
		  <td nowrap><input type="text" name="UseState" id="UseState" value="" style="font-size:12px; size:250;width:250" ></td>
		  <td nowrap>
		<select name="ReturnFlag" id="ReturnFlag" >
		<option value="0">不回厂</option>
		<option value="1">回厂</option>
		</select></td>
		  </td>
		  <td nowrap onClick="DeleteRow(this,'GOP')">删除</td>
		  </tr>
		  <%
		  for j=0 to i-1 
		  %>
		  <tr bgcolor='#EBF2F9' onMouseOver = "this.style.backgroundColor = '#FFFFFF'" onMouseOut = "this.style.backgroundColor = ''" style='cursor:hand'>
		  <td nowrap><input type="hidden" name="FEntryID" id="FEntryID" value="<%=FEntryID(j)%>">
		  <input type="text" name="Goods" id="Goods" value="<%=Goods(j)%>" style="font-size:12px; width:140;"></td>
		  <td nowrap><input type="text" name="FNumber" id="FNumber" value="<%=FNumber(j)%>" style="font-size:12px; width:60;" onChange="return checkNum(this)"></td>
		  <td nowrap><input type="text" name="FQty" id="FQty" value="<%=FQty(j)%>" style="font-size:12px; width:80;" onChange="return checkNum(this)"></td>
		  <td nowrap><input type="text" name="UseState" id="UseState" value="<%=UseState(j)%>" style="font-size:12px;width:250" ></td>
		  <td nowrap>
		<select name="ReturnFlag" id="ReturnFlag"  <%if OutCheckFlag=1 then response.Write("readonly") end if%>>
		<option value="0" <%if ReturnFlag(j)="0" then response.write ("selected")%>>不回厂</option>
		<option value="1" <%if ReturnFlag(j)="1" then response.write ("selected")%>>回厂</option>
		</select></td>
		  </td>
		  <td nowrap onClick="DeleteRow(this,'GOP')">删除</td>
		  </tr>
		  <%
		  next
		  %>
		  </tbody>
		</table>
		</td>
      </tr>
      <tr>
        <td height="20" align="left" colspan="5">&nbsp;</td>
        <td valign="bottom" colspan="1">&nbsp;<%if OutCheckFlag=0 then%><input name="addrow" type="button" class="button"  id="addrow" value="增加一行" style="WIDTH: 80;" onClick="AddRow()"><%end if%></td>
      </tr>
 
      <tr>
        <td height="20" align="left">制单人：</td>
        <td><input type="hidden" name="FBiller" id="FBiller" value="<%=FBiller%>"><input name="FBillerName" type="text" class="textfield" id="FBillerName" style="WIDTH: 140;" value="<%=FBillerName%>" maxlength="100" readonly></td>
        <td height="20" align="left">制单日期：</td>
        <td><input name="FDate" type="text" class="textfield" id="FDate" style="WIDTH: 140;" value="<%=FDate%>" maxlength="100" readonly></td>
        <td height="20" align="left"></td>
        <td></td>
      </tr>
      <tr>
        <td height="20" align="left">放行状态：</td>
        <td><input name="OutCheckFlag" type="hidden" id="OutCheckFlag"  value="<%=OutCheckFlag%>">
		<input name="strOutCheckFlag" type="text" class="textfield" id="strOutCheckFlag" style="WIDTH: 140;" value="<%=strOutCheckFlag%>" maxlength="100" readonly="true"></td>
        <td><input name="OutChecker1" type="text" class="textfield" id="OutChecker1" style="WIDTH: 140;" value="<%=OutChecker1%>" maxlength="100" readonly="true"></td>
        <td><input name="OutCheckDate1" type="text" class="textfield" id="OutCheckDate1" style="WIDTH: 140;" value="<%=OutCheckDate1%>" maxlength="100" readonly="true"></td>
        <td><input name="OutChecker2" type="text" class="textfield" id="OutChecker2" style="WIDTH: 140;" value="<%=OutChecker2%>" maxlength="100" readonly="true"></td>
        <td><input name="OutCheckDate2" type="text" class="textfield" id="OutCheckDate2" style="WIDTH: 140;" value="<%=OutCheckDate2%>" maxlength="100" readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">取回状态：</td>
        <td><input name="InCheckFlag" type="hidden" id="InCheckFlag"  value="<%=InCheckFlag%>">
		<input name="strInCheckFlag" type="text" class="textfield" id="strInCheckFlag" style="WIDTH: 140;" value="<%=strInCheckFlag%>" maxlength="100" readonly="true"></td>
        <td><input name="InChecker1" type="text" class="textfield" id="InChecker1" style="WIDTH: 140;" value="<%=InChecker1%>" maxlength="100" readonly="true"></td>
        <td><input name="InCheckDate1" type="text" class="textfield" id="InCheckDate1" style="WIDTH: 140;" value="<%=InCheckDate1%>" maxlength="100" readonly="true"></td>
        <td><input name="InChecker2" type="text" class="textfield" id="InChecker2" style="WIDTH: 140;" value="<%=InChecker2%>" maxlength="100" readonly="true"></td>
        <td><input name="InCheckDate2" type="text" class="textfield" id="InCheckDate2" style="WIDTH: 140;" value="<%=InCheckDate2%>" maxlength="100" readonly="true"></td>
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
        <td height="20" align="center" colspan="3">放行管制</td>
        <td align="center" colspan="3">取回管制</td>
      </tr>
      <tr>
        <td valign="bottom" colspan="3" align="center">
		&nbsp;<input name="check1" type="button" class="button"  id="check1" value="主管审核-出" style="WIDTH: 80;" onClick="toSubmit(this)" >
		&nbsp;<input name="check2" type="button" class="button"  id="check2" value="门卫审核-出" style="WIDTH: 80;"  onClick="toSubmit(this)">
		</td>
        <td valign="bottom" colspan="3" align="center">
		&nbsp;<input name="check3" type="button" class="button"  id="check3" value="门卫审核-入" style="WIDTH: 80;" onClick="toSubmit(this)" >
		&nbsp;<input name="check4" type="button" class="button"  id="check4" value="主管审核-入" style="WIDTH: 80;"  onClick="toSubmit(this)">
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
	  	  '子表信息添加
		  dim formdata(4)
		  formdata(0)=Split(Replace(Request("Goods")," ",""),",")
		  formdata(1)=Split(Replace(Request("FQty")," ",""),",")
		  formdata(2)=Split(Replace(Request("UseState")," ",""),",")
		  formdata(3)=Split(Replace(Request("ReturnFlag")," ",""),",")
		  formdata(4)=Split(Replace(Request("FNumber")," ",""),",")
		  if Ubound(formdata(0)) >0 then
		  '主表信息添加
		  if Request.Form("RegistDepartment")="" then
		  response.write "<script language=javascript> alert('部门编号不能为空！');history.back();</script>"
		  response.End()
		  end if
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_GoodsCarryOutMain"
		  rs.open sql,connk3,1,3
		  rs.addnew
		  rs("RegDate")=Request.Form("RegDate")
		  rs("Register")=Request.Form("Register")
		  rs("FDate")=Request.Form("FDate")
		  rs("FBiller")=Request.Form("FBiller")
		  rs("RegistDepartment")=Request.Form("RegistDepartment")
		  rs.update
		  SerialNum=rs("SerialNum")
		  rs.close
		  set rs=nothing 
		  For i=1 To Ubound(formdata(0)) 
			if formdata(0)(i)<>"" then
			  sql="insert into z_GoodsCarryOutDetails values ("&SerialNum&","&i&",'"&formdata(0)(i)&"',"&formdata(1)(i)&",'"&formdata(2)(i)&"',"&formdata(3)(i)&","&formdata(4)(i)&")"
			  connk3.Execute(sql)
			end if
		  Next
		response.write "<script language=javascript> alert('成功编辑物品携出信息！');changeAdminFlag('物品携出');location.replace('GoodsOutPassMana.asp');</script>"
		else
		  response.write "<script language=javascript> alert('无明细信息，登记失败！');history.back();</script>"
		  response.End()
		end if
	  end if
	  if Action="Modify" then '修改记录
	  	'保存主表信息编辑
		'保存子表信息编辑
		  dim formdata2(5)
		  formdata2(0)=Split(Replace(Request("Goods")," ",""),",")
		  formdata2(1)=Split(Replace(Request("FQty")," ",""),",")
		  formdata2(2)=Split(Replace(Request("UseState")," ",""),",")
		  formdata2(3)=Split(Replace(Request("ReturnFlag")," ",""),",")
		  formdata2(4)=Split(Replace(Request("FEntryID")," ",""),",")
		  formdata2(5)=Split(Replace(Request("FNumber")," ",""),",")
		  if Ubound(formdata2(0)) >0 then
		  set rs = server.createobject("adodb.recordset")
		  sql="select * from z_GoodsCarryOutMain where SerialNum="& SerialNum
		  rs.open sql,connk3,1,3
		  if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		  end if
		  rs("RegDate")=trim(Request.Form("RegDate"))
		  rs("Register")=Request.Form("Register")
		  rs("RegistDepartment")=Request.Form("RegistDepartment")
		  rs.update
		  rs.close
		  set rs=nothing 

		  For i=1 To Ubound(formdata2(0)) 
			    if formdata2(0)(i)<>"" then
				  if formdata2(4)(i)="" then
					sql="insert into z_GoodsCarryOutDetails values ("&SerialNum&","&i&",'"&formdata2(0)(i)&"',"&formdata2(1)(i)&",'"&formdata2(2)(i)&"',"&formdata2(3)(i)&","&formdata2(5)(i)&")"
					connk3.Execute(sql)
				  else
				    sql="update z_GoodsCarryOutDetails set SerialNum="&SerialNum&",Findex="&i&",Goods='"&formdata2(0)(i)&"',FQty="&formdata2(1)(i)&",UseState='"&formdata2(2)(i)&"',ReturnFlag='"&formdata2(3)(i)&"',FNumber='"&formdata2(5)(i)&"' where FEntryID="&formdata2(4)(i)
				  	connk3.Execute(sql)
				  end if
				end if
		  Next
		response.write "<script language=javascript> alert('成功编辑物品携出信息！');changeAdminFlag('物品携出');location.replace('GoodsOutPassMana.asp');</script>"
		  else
		  response.write "<script language=javascript> alert('无明细信息，登记失败！');history.back();</script>"
		  end if
	  end if
  elseif Keyword="Delete" then
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_GoodsCarryOutMain where Outcheckflag =0 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  else
	    if UserName=rs("FBiller") then
		  sql="delete from z_GoodsCarryOutMain where SerialNum="& SerialNum
		  connk3.execute(sql)
		  sql="delete from z_GoodsCarryOutDetails where SerialNum="& SerialNum
		  connk3.execute(sql)
		else
		response.write ("只能删除自己建立的单据！")
		response.end
		end if		  
	  end if
	  rs.close
	  set rs=nothing 
		response.write "<script language=javascript> alert('成功编辑物品携出信息！');changeAdminFlag('物品携出');location.replace('GoodsOutPassMana.asp');</script>"
  elseif Keyword="check1" and Instr(session("AdminPurview"),"|1006.2,")>0 then
  '申请部门主管审核
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_GoodsCarryOutMain where OutCheckFlag =0 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	    rs("OutCheckFlag")=1
		rs("OutChecker1")=UserName
		rs("OutCheckDate1")=now()
	  rs.update
	  rs.close
	  set rs=nothing 
		response.write "<script language=javascript> alert('成功编辑物品携出信息！');changeAdminFlag('物品携出');location.replace('GoodsOutPassMana.asp');</script>"
  elseif Keyword="check2" and Instr(session("AdminPurview"),"|1006.3,")>0 then
  '申请部门主管审核
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_GoodsCarryOutMain where Outcheckflag =1 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	    rs("OutCheckFlag")=2
		rs("OutChecker2")=UserName
		rs("GetOutDate")=Request.Form("GetOutDate")
		rs("OutCheckDate2")=now()
	  rs.update
	  rs.close
	  set rs=nothing 
		response.write "<script language=javascript> alert('成功编辑物品携出信息！');changeAdminFlag('物品携出');location.replace('GoodsOutPassMana.asp');</script>"
  elseif Keyword="check3" and Instr(session("AdminPurview"),"|1006.3,")>0 then
  '申请部门主管审核
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_GoodsCarryOutMain where Incheckflag =0 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	    rs("InCheckFlag")=1
		rs("InChecker1")=UserName
		rs("GetInDate")=Request.Form("GetInDate")
		rs("InCheckDate1")=now()
	  rs.update
	  rs.close
	  set rs=nothing 
		response.write "<script language=javascript> alert('成功编辑物品携出信息！');changeAdminFlag('物品携出');location.replace('GoodsOutPassMana.asp');</script>"
  elseif Keyword="check4" and Instr(session("AdminPurview"),"|1006.2,")>0 then
  '申请部门主管审核
	  set rs = server.createobject("adodb.recordset")
	  sql="select * from z_GoodsCarryOutMain where Incheckflag =1 and SerialNum="& SerialNum
	  rs.open sql,connk3,1,3
	  if rs.bof and rs.eof then
		response.write ("数据库读取记录出错 或者 该单据当前状态不允许此操作！")
		response.end
	  end if
	    rs("InCheckFlag")=2
		rs("InChecker2")=UserName
		rs("InCheckDate2")=now()
	  rs.update
	  rs.close
	  set rs=nothing 
		response.write "<script language=javascript> alert('成功编辑物品携出信息！');changeAdminFlag('物品携出');location.replace('GoodsOutPassMana.asp');</script>"
  else
  	if Action="Modify" then'提出编辑信息
	  '提取主表信息
	  set rs = server.createobject("adodb.recordset")
      sql="select * from z_GoodsCarryOutMain where SerialNum="& SerialNum
      rs.open sql,connk3,1,1
      if rs.bof and rs.eof then
        response.write ("数据库读取记录出错！")
        response.end
      end if
	  FBillerName=getUser(rs("FBiller"))
      FDate=rs("FDate")
	  FBiller=rs("FBiller")
      RegDate=rs("RegDate")
	  RegisterName=getUser(rs("Register"))
      Register=rs("Register")
	  GetOutDate=rs("GetOutDate")
      GetInDate=rs("GetInDate")
	  RegistDepartment=rs("RegistDepartment")
	  RegistDepartmentName=getDepartment(rs("RegistDepartment"))
	  OutCheckFlag=rs("OutCheckFlag")
	  OutCheckDate1=rs("OutCheckDate1")
	  OutCheckDate2=rs("OutCheckDate2")
	  OutChecker1=getUser(rs("OutChecker1"))
	  OutChecker2=getUser(rs("OutChecker2"))
	  InCheckFlag=rs("InCheckFlag")
	  InCheckDate1=rs("InCheckDate1")
	  InCheckDate2=rs("InCheckDate2")
	  InChecker1=getUser(rs("InChecker1"))
	  InChecker2=getUser(rs("InChecker2"))
	  if OutCheckFlag=1 then
	  GetOutDate=now()
	  end if
	  if OutCheckFlag=2 and InCheckFlag=0 then
	  GetInDate=now()
	  end if
	  if rs("OutCheckFlag")=1 then
	  strOutCheckFlag="主管审核"
	  elseif rs("OutCheckFlag")=2 then
	  strOutCheckFlag="门卫审核"
	  else
	  strOutCheckFlag="未审核"
	  end if
	  if rs("InCheckFlag")=1 then
	  strInCheckFlag="主管审核"
	  elseif rs("InCheckFlag")=2 then
	  strInCheckFlag="门卫审核"
	  else
	  strInCheckFlag="未审核"
	  end if
	  '提取子表信息
	  set rs = server.createobject("adodb.recordset")
      sql="select count(Fentryid) as idCount from z_GoodsCarryOutDetails where SerialNum="& SerialNum
      rs.open sql,connk3,0,1
	  dim idCount
	  idCount=rs("idCount")
	  ReDim Preserve Goods(idCount)
	  ReDim Preserve FQty(idCount)
	  ReDim Preserve UseState(idCount)
	  ReDim Preserve ReturnFlag(idCount)
	  ReDim Preserve FEntryID(idCount)
	  ReDim Preserve FNumber(idCount)
	  set rs = server.createobject("adodb.recordset")
      sql="select * from z_GoodsCarryOutDetails where SerialNum="& SerialNum&" order by findex asc "
      rs.open sql,connk3,1,1
	  while(not rs.eof)
	    FEntryID(i)=rs("FEntryID")
	    Goods(i)=rs("Goods")
		FQty(i)=rs("FQty")
		UseState(i)=rs("UseState")
		ReturnFlag(i)=rs("ReturnFlag")
		FNumber(i)=rs("FNumber")
		i=i+1
		rs.movenext
	  wend
	  rs.close
      set rs=nothing 
	else'提取增加时所需信息,制单人，制单日期，单号
	  RegDate=date()
      FBiller=UserName
	  FBillerName=AdminName
      Register=UserName
	  RegisterName=AdminName
      FDate=date()
	end if
  end if
end sub

Function getDepartment(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From t_item where fitemclassid=2 and Fitemid="&ID
  rs.open sql,connk3,1,1
  getDepartment=rs("Fname")
  rs.close
  set rs=nothing
End Function    

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
Function getBillNo(ID,id2)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select max("&id2&") as maxno From "&ID
  rs.open sql,connk3,1,1
  getBillNo=rs("maxno")
  rs.close
  set rs=nothing
End Function  

%>