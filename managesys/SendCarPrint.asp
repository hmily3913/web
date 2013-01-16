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
if Instr(session("AdminPurview"),"|1003,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<% 
dim Result,Action,SerialNum,AdminName,UserName
SerialNum=request("SerialNum")
dim i,j '用于循环的整数
i=0
'定义派车单主表变量
dim SerialNumOne,RegDate,Register,RegisterName,FBase1,FBase1Name,SendReason,GoodsName,DeliveryAddr,mileage,Remark,TeamSugg
dim FBiller,FBillerName,FDate,PlanStarteDate,PlanEndDate,StarteDate,EndDate,CheckFlag,Checker1,Checker2,CheckDate1,CheckDate2
dim Checker3,CheckDate3,Importance,CarNumber,Driver,DriverName,Fee,FeeDepartment,FeeDepartmentName,DPhone,Startemil,Endmil,Planmileage 
dim totalTime,CarryGoods,CarryMans,UseCarFlag,packages,checkflagname,OutPeron,OutPeronName

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
  connk3.Execute("update z_SendCar set PrintFlag=1 where SerialNum in ("& SerialNum&")")
  sql="select * from z_SendCar where SerialNum in ("& SerialNum&")"
  rs.open sql,connk3,0,1
  while(i<=UBound(names))
  	SerialNumOne=rs("SerialNum")
	  packages=rs("packages")
	  RegDate=rs("RegDate")
	  Register=rs("Register")
	  RegisterName=getUser(rs("Register"))
	  FBase1=rs("RegistDepartment")
	  FBase1Name=getDepartment(rs("RegistDepartment"))
	  SendReason=rs("SendReason")
	  GoodsName=rs("GoodsName")
	  DeliveryAddr=rs("DeliveryAddr")
      mileage=rs("mileage")
	  Remark=rs("Remark")
      FBiller=rs("FBiller")
	  FBillerName=getUser(rs("FBiller"))
      FDate=rs("FDate")
	  PlanStarteDate=rs("PlanStarteDate")
	  PlanEndDate=rs("PlanEndDate")
	  StarteDate=rs("StarteDate")
	  EndDate=rs("EndDate")
	  CheckFlag=rs("CheckFlag")
	  Checker1=getUser(rs("Checker1"))
	  Checker2=getUser(rs("Checker2"))
	  CheckDate1=rs("CheckDate1")
	  CheckDate2=rs("CheckDate2")
	  Checker3=getUser(rs("Checker3"))
	  CheckDate3=rs("CheckDate3")
	  Importance=rs("Importance")
	  CarNumber=rs("CarNumber")
	  Driver=rs("Driver")
	  DriverName=getUser(rs("Driver"))
	  Fee=rs("Fee")
	  TeamSugg=rs("TeamSugg")
	  FeeDepartment=rs("FeeDepartment")
	  DPhone=rs("DPhone")
	  Startemil=rs("Startemil")
	  Endmil=rs("Endmil")
	  Planmileage=rs("Planmileage")
	  CarryGoods=rs("CarryGoods")
	  CarryMans=rs("CarryMans")
	  totalTime=rs("totalTime")
	  UseCarFlag=rs("UseCarFlag")
	  if rs("CheckFlag")=1 then
	  checkflagname="主管审核"
	  elseif rs("CheckFlag")=2 then
	  checkflagname="车队一审"
	  elseif rs("CheckFlag")=3 then
	  checkflagname="门卫一审"
	  elseif rs("CheckFlag")=4 then
	  checkflagname="门卫二审"
	  elseif rs("CheckFlag")=5 then
	  checkflagname="车队二审"
	  else
	  checkflagname="未审核"
	  end if
	  OutPeron=rs("OutPeron")
	  if OutPeron<>"" then
	  dim iii
	  iii=0
	  while (iii<=UBound(split(OutPeron,",")))
	  if iii<>UBound(split(OutPeron,",")) then
	  OutPeronName=OutPeronName+getUser(split(OutPeron,",")(iii))+","
	  else
	  OutPeronName=OutPeronName+getUser(split(OutPeron,",")(iii))
	  end if
	  iii=iii+1
	  wend
	  end if
	  if rs("FeeDepartment")<>"" then FeeDepartmentName=getDepartment(rs("FeeDepartment"))
%>
<table width="100%" cellpadding="3" cellspacing="1" bgcolor="#000000" style="border: 1px solid; ">
      <tr style="border: 1px solid; ">
        <td bgcolor='#ffffff' colspan="6" align="center"><b><strong>蓝道外出单</strong></b></td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">单据号：</td>
        <td bgcolor='#ffffff'>
		<%=SerialNumOne%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">申请人：</td>
        <td bgcolor='#ffffff'>
		<%=RegisterName%>&nbsp;
		</td>
        <td bgcolor='#ffffff' height="20" align="left">申请日期：</td>
        <td bgcolor='#ffffff'><%=RegDate%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">申请部门：</td>
        <td bgcolor='#ffffff'>
		<%=FBase1name%></td>
        <td bgcolor='#ffffff' height="20" align="left">计划出发时间：</td>
        <td bgcolor='#ffffff'><%=PlanStarteDate%></td>
        <td bgcolor='#ffffff' height="20" align="left">计划回来时间：</td>
        <td bgcolor='#ffffff'><%=PlanEndDate%></td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">派车分类：</td>
        <td bgcolor='#ffffff'><%=SendReason%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">重要性：</td>
        <td bgcolor='#ffffff'><%=Importance%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">事由及内容：</td>
        <td bgcolor='#ffffff' ><%=GoodsName%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">目的地：</td>
        <td bgcolor='#ffffff' colspan="3"><%=DeliveryAddr%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">联系方式：</td>
        <td bgcolor='#ffffff' ><%=DPhone%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">外出人员：</td>
        <td bgcolor='#ffffff'><%=OutPeronName%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">外出人数/个：</td>
        <td bgcolor='#ffffff'><%=CarryMans%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">&nbsp;</td>
        <td bgcolor='#ffffff'>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">本次载货量/立方：</td>
        <td bgcolor='#ffffff'><%=CarryGoods%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">货物件数：</td>
        <td bgcolor='#ffffff'><%=packages%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">&nbsp;</td>
        <td bgcolor='#ffffff'>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">车牌号：</td>
        <td bgcolor='#ffffff'><%=CarNumber%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">驾驶员：</td>
        <td bgcolor='#ffffff'><%=DriverName%>&nbsp;
		</td>
        <td bgcolor='#ffffff' height="20" align="left">预计里程数：</td>
        <td bgcolor='#ffffff'><%=Planmileage%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">出发里程表数：</td>
        <td bgcolor='#ffffff'><%=Startemil%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">回来里程表数：</td>
        <td bgcolor='#ffffff'><%=Endmil%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">实际里程数：</td>
        <td bgcolor='#ffffff'><%=mileage%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">实际出发时间：</td>
        <td bgcolor='#ffffff'><%=StarteDate%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">实际回来时间：</td>
        <td bgcolor='#ffffff'><%=EndDate%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">累计使用时间/分：</td>
        <td bgcolor='#ffffff'><%=totalTime%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">费用：</td>
        <td bgcolor='#ffffff'><%=Fee%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">费用归属单位：</td>
        <td bgcolor='#ffffff'>
		<%=FeeDepartmentName%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">车队意见：</td>
        <td bgcolor='#ffffff'><%=TeamSugg%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">备注：</td>
        <td bgcolor='#ffffff' colspan="5"><%=Remark%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">制单人：</td>
        <td bgcolor='#ffffff'><%=FBillerName%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">制单日期：</td>
        <td bgcolor='#ffffff'><%=FDate%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left"></td>
        <td bgcolor='#ffffff'></td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left">审核状态：</td>
        <td bgcolor='#ffffff'><%=CheckFlag%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">一审人名：</td>
        <td bgcolor='#ffffff'><%=Checker1%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">一审时间：</td>
        <td bgcolor='#ffffff'><%=CheckDate1%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left"></td>
        <td bgcolor='#ffffff'></td>
        <td bgcolor='#ffffff' height="20" align="left">二审人名：</td>
        <td bgcolor='#ffffff'><%=Checker2%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">二审时间：</td>
        <td bgcolor='#ffffff'><%=CheckDate2%>&nbsp;</td>
      </tr>
      <tr>
        <td bgcolor='#ffffff' height="20" align="left"></td>
        <td bgcolor='#ffffff'></td>
        <td bgcolor='#ffffff' height="20" align="left">三审人名：</td>
        <td bgcolor='#ffffff'><%=Checker3%>&nbsp;</td>
        <td bgcolor='#ffffff' height="20" align="left">三审时间：</td>
        <td bgcolor='#ffffff'><%=CheckDate3%>&nbsp;</td>
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