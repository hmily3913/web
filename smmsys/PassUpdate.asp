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
  var theform=document.getElementById("editForm");
  theform.submit();
}
</script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<BODY>
<% 
dim AdminName,UserName
UserName=session("UserName")
AdminName=session("AdminName")

%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>修改用户密码信息</strong></font></td>
  </tr>
</table>
<br>
  <form name="editForm" id="editForm" method="post" action="PassUpdate.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews idth="100%">

      <tr>
        <td width="60" height="20" align="right">&nbsp;</td>
        <td align="left">&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="right">当前用户名/姓名：</td>
        <td  align="left"><%=UserName%>/<%=AdminName%>
		</td>
      </tr>
      <tr>
        <td height="20" align="right">旧密码：</td>
        <td  align="left"><input name="CurrPass" type="password" class="textfield" id="CurrPass" style="WIDTH: 140;"  maxlength="50"></td>
      </tr>
      <tr>
        <td height="20" align="right">新密码：</td>
        <td  align="left"><input name="NewPass" type="password" class="textfield" id="NewPass" style="WIDTH: 140;"  maxlength="50"></td>
      </tr>
      <tr>
        <td height="20" align="right">新密码确认：</td>
        <td  align="left"><input name="ReNewPass" type="password" class="textfield" id="ReNewPass" style="WIDTH: 140;"  maxlength="50"></td>
      </tr>
 
      <tr>
        <td height="20" align="left" colspan="3">&nbsp;</td>
        <td valign="bottom" colspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td valign="bottom" colspan="6" align="center">
		<input type="hidden" name="Keyword" id="Keyword" value="Update">
		&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">
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
</BODY>
</HTML>

<%
dim Keyword,rsRepeat,rs,sql
Keyword=request("Keyword")
if Keyword="Update" then '保存事务处理
  dim CurrPass,NewPass,ReNewPass
  set rs = server.createobject("adodb.recordset")
  Keyword=request("Keyword")
  CurrPass=request("CurrPass")
  NewPass=request("NewPass")
  ReNewPass=request("ReNewPass")
  sql="select * from [N-基本资料单头] where 员工代号='"&UserName&"' and 密码='"&CurrPass&"'"
  rs.open sql,conn,1,3
  if rs.bof and rs.eof then
	response.write ("旧密码不正确，密码修改失败！")
	response.end
  end if
  if NewPass<>ReNewPass then
    response.write ("新密码两次输入不相同，密码修改失败！")
	response.end
  else
    rs("密码")=Request.Form("NewPass")
    rs.update
  end if
  rs.close
  set rs=nothing 
response.write "<script language=javascript> alert('用户密码修改成功！');changeAdminFlag('密码修改');location.replace('PassUpdate.asp');</script>"
end if
%>