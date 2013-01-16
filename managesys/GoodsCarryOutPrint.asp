<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="application/vnd.ms-excel; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="../Images/CssAdmin.css">
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style>
<style type="text/css">
td{
 border:1px solid;
 bgcolor:'#ffffff';
}
</style>
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|1006,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim i,j '用于循环的整数
i=0
j=0
'定义派车单主表变量
dim Result,Action,SerialNum,AdminName,UserName,SerialNumOne
SerialNum=request("SerialNum")
dim RegDate,Register,RegisterName,GetOutDate,GetInDate,FBiller,FBillerName,FDate,OutCheckFlag,RegistDepartment,RegistDepartmentName
dim OutChecker1,OutCheckDate1,OutChecker2,OutCheckDate2,InCheckFlag,InChecker1,InCheckDate1,InChecker2,InCheckDate2
'定义宿舍水电子表变量
dim FEntryID(),Goods(),FQty(),UseState(),ReturnFlag()

%>
<table class="Noprint">
 <tr>
 <td><div><OBJECT id="WebBrowser" classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0></OBJECT><input type="button" value="打印" onclick="javascript:window.print()">&nbsp;<input type="button" value="页面设置" onclick=document.all.WebBrowser.ExecWB(8,1)>&nbsp;<input type="button" value="打印预览" onclick=document.all.WebBrowser.ExecWB(7,1)></div></td>
 </tr>
</table>
<%
  dim Keyword,rsRepeat,rs,sql,names
  names = Split(SerialNum, ",")
  set rs = server.createobject("adodb.recordset")
  connk3.Execute("update z_GoodsCarryOutMain set PrintFlag=1 where SerialNum in ("& SerialNum&")")
  sql="select * from z_GoodsCarryOutMain where SerialNum in ("& SerialNum&")"
  rs.open sql,connk3,0,1
  while(i<=UBound(names))
  	SerialNumOne=rs("SerialNum")
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
	  if rs("OutCheckFlag")=1 then
	  OutCheckFlag="主管审核"
	  elseif rs("OutCheckFlag")=2 then
	  OutCheckFlag="门卫审核"
	  else
	  OutCheckFlag="未审核"
	  end if
	  if rs("InCheckFlag")=1 then
	  InCheckFlag="主管审核"
	  elseif rs("InCheckFlag")=2 then
	  InCheckFlag="门卫审核"
	  else
	  InCheckFlag="未审核"
	  end if
	  
%>
<table width="100%" cellpadding="3" cellspacing="1" style="border: 1px solid; ">
      <tr style="border: 1px solid; ">
        <td bgcolor='#ffffff' colspan="6" align="center"><b><strong>蓝道物品携出放行条</strong></b></td>
      </tr>
      <tr>
        <td height="20" align="left">单据号：</td>
        <td><%=SerialNum%>&nbsp;</td>
        <td height="20" align="left">申请人：</td>
        <td>
<%=RegisterName%>&nbsp;
		</td>
        <td height="20" align="left">申请部门：</td>
        <td><%=RegistDepartmentName%>&nbsp;<td>
      </tr>
      <tr>
        <td height="20" align="left">申请日期：</td>
        <td><%=RegDate%>&nbsp;</td>
        <td height="20" align="left">放行日期：</td>
        <td><%=GetOutDate%>&nbsp;</td>
        <td height="20" align="left">取回日期：</td>
        <td><%=GetInDate%>&nbsp;</td>
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
		  </tr>
		  <%
	  set rsRepeat = server.createobject("adodb.recordset")
      sql="select * from z_GoodsCarryOutDetails where SerialNum="& SerialNumOne&" order by findex asc "
      rsRepeat.open sql,connk3,1,1
	  while(not rsRepeat.eof)
		  %>
		  <tr bgcolor='#EBF2F9' onMouseOver = "this.style.backgroundColor = '#FFFFFF'" onMouseOut = "this.style.backgroundColor = ''" style='cursor:hand'>
		  <td nowrap><%=rsRepeat("Goods")%>&nbsp;</td>
		  <td nowrap><%=rsRepeat("FNumber")%>&nbsp;</td>
		  <td nowrap><%=rsRepeat("FQty")%>&nbsp;</td>
		  <td nowrap><%=rsRepeat("UseState")%>&nbsp;</td>
		  <td nowrap>
		 <%if rsRepeat("ReturnFlag")="0" then response.write ("不回厂")%>
		<%if rsRepeat("ReturnFlag")="1" then response.write ("回厂")%>&nbsp;
		</td>
		  </tr>
		  <%
		rsRepeat.movenext
	  wend
		  %>
		  </tbody>
		</table>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">制单人：</td>
        <td><%=FBillerName%></td>
        <td height="20" align="left">制单日期：</td>
        <td><%=FDate%></td>
        <td height="20" align="left"></td>
        <td></td>
      </tr>
      <tr>
        <td height="20" align="left">放行状态：</td>
        <td><%=OutCheckFlag%>&nbsp;</td>
        <td><%=OutChecker1%>&nbsp;</td>
        <td><%=OutCheckDate1%>&nbsp;</td>
        <td><%=OutChecker2%>&nbsp;</td>
        <td><%=OutCheckDate2%>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left">取回状态：</td>
        <td>
<%=InCheckFlag%>&nbsp;</td>
        <td><%=InChecker1%>&nbsp;</td>
        <td><%=InCheckDate1%>&nbsp;</td>
        <td><%=InChecker2%>&nbsp;</td>
        <td><%=InCheckDate2%>&nbsp;</td>
      </tr>
</table>
<% if i<>UBound(names) then %>
<div class="PageNext"></div>
<% 
   end if
	  i=i+1
	  rs.movenext
    wend
    rs.close
    set rs=nothing 
 %>
</BODY>
</HTML>


<%
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