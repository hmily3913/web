<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="zbh" />
<META NAME="Keywords" CONTENT="" />
<TITLE>产品列表</TITLE>
<link rel="stylesheet" href="../Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/CustomAjax.js"></script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
if Instr(session("AdminPurview"),"|1003,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<%
dim Result,StartDate,EndDate,Keyword,Reachsum,unReachsum,Reachper,sqlstr,checkst,checkdx
Result=request("Result")
StartDate=request("start_date")
if StartDate="" then StartDate=date()
EndDate=request("end_date")
if EndDate="" then EndDate=date()
Keyword=request("Keyword")
Reachsum=request("Reachsum")
checkst=request("checkst")
checkdx=request("checkdx")
sqlstr="z_SendCar,t_item a "
 
%>

  <% ProductsList() %>

</BODY>
</HTML>
<%
'-----------------------------------------------------------
function ProductsList()
 if Result="Search" then
 %>
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单号</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td width="60" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请部门</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>外出工号</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>外出人员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>派车分类</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>用车</strong></font></td>
    <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>事由及内容</strong></font></td>
    <td width="120" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">目的地</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">计划出发</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">计划回来</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">出发时间</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">回来时间</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">里程</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">状态</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">打印</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">重要性</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">车牌号</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">驾驶员</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">预计里程</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">使用时间</font></strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">备注</font></strong></td>
  </tr>
 <%
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数
  dim page'页码
      page=clng(request("Page"))
  dim pagenc'每页显示的分页页码数量=pagenc*2+1
      pagenc=2
  dim pagenmax'每页显示的分页的最大页码
  dim pagenmin'每页显示的分页的最小页码
  dim datafrom'数据表名
      datafrom=sqlstr
  dim datawhere'数据条件
      if Keyword="list" then
	     datawhere=" where DeleteFlag<1 and RegistDepartment=a.fitemid  and (SendReason like '%"&Reachsum&"%' or GoodsName like '%"&Reachsum&"%' or DeliveryAddr like '%"&Reachsum&"%' or Register like '%"&Reachsum&"%' or OutPeron like '%"&Reachsum&"%' or CarNumber like '%"&Reachsum&"%') and RegDate>='"&StartDate&"' and RegDate<='"&EndDate&"'"
		 if checkst<>"" then
		   datawhere=datawhere&" and checkflag"&checkdx&checkst
		 end if
	  else
		 datawhere=" where DeleteFlag<1 and RegistDepartment=a.fitemid  "
 	  end if
  if Instr(session("AdminPurview"),"|1003.2,")=0 and Instr(session("AdminPurview"),"|1003.3,")=0 then
	dim Depart:Depart=session("Depart")
	  if Depart="KD01.0001.0001"  then
		datawhere=datawhere&" and (left(a.fnumber,2)='06' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0002" then
		datawhere=datawhere&" and (left(a.fnumber,2)='03' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0003" then
		datawhere=datawhere&" and (left(a.fnumber,2)='05' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0005.0004" then
		datawhere=datawhere&" and (left(a.fnumber,2)='02' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0005" then
		datawhere=datawhere&" and (left(a.fnumber,2)='08' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0006" then
		datawhere=datawhere&" and (left(a.fnumber,2)='07' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0007" then
		datawhere=datawhere&" and (left(a.fnumber,2)='11' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0008" then
		datawhere=datawhere&" and (left(a.fnumber,2)='12' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0009" then
		datawhere=datawhere&" and (left(a.fnumber,2)='04' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0010" then
		datawhere=datawhere&" and (left(a.fnumber,2)='10' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0011" then
		datawhere=datawhere&" and (left(a.fnumber,2)='09' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0012" AND Instr(session("AdminPurview"),"|1003.5,")=0 then
		datawhere=datawhere&" and (left(a.fnumber,2)='01' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  end if
  end if
		 dim queryFlag:queryFlag=request("queryFlag")
		 if queryFlag<>"" then
		   if queryFlag="none" then
		     datawhere=datawhere&" and checkflag=0"
		   elseif queryFlag="chedui" then
		     datawhere=datawhere&" and (checkflag=1 or checkFlag=4) and UseCarFlag='是' "
		   elseif queryFlag="menwei" then
		     datawhere=datawhere&" and ((checkflag=1 and UseCarFlag='否') or checkFlag=2 or checkFlag=3) "
		   elseif queryFlag="kqgl" then
		     datawhere=datawhere&" and ISCQ=1 "
		   end if
		 end if
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by SerialNum desc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(SerialNum) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  idCount=rs("idCount")
  '获取记录总数
  if(idcount>0) then'如果记录总数=0,则不处理
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select SerialNum from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
	i=1
    while(not rs.eof)
	  if(i=1)then

	    sqlid=rs("SerialNum")
	  else
	    sqlid=sqlid &","&rs("SerialNum")
	  end if
	  i=i+1
	  rs.movenext
    wend
  '获取本页需要用到的id结束============================================
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  if(idcount>0 and sqlid<>"") then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
	dim CheckFlag,PrintFlag
    sql="select z_SendCar.*,t_item.fname as name1,t_emp.fname as name2 from z_SendCar left join t_item on t_item.fitemid=RegistDepartment left join t_emp on t_emp.fnumber=z_SendCar.Register where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while(not rs.eof)'填充数据到表格
	dim outpersonname,OutPeron
	outpersonname=""
	  OutPeron=rs("OutPeron")
	  if OutPeron<>"" then
	  dim ooo
	  ooo=0
	  while (ooo<=UBound(split(OutPeron,",")))
	  if ooo<>UBound(split(OutPeron,",")) then
	  outpersonname=outpersonname+getUser(split(OutPeron,",")(ooo))+","
	  else
	  outpersonname=outpersonname+getUser(split(OutPeron,",")(ooo))
	  end if
	  ooo=ooo+1
	  wend
	  end if
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("name2")&"</td>"
      Response.Write "<td nowrap>"&rs("name1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&OutPeron&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&outpersonname&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("UseCarFlag")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("SendReason")&"</td>" & vbCrLf
      Response.Write "<td nowrap width='150'>"&rs("GoodsName")&"</td>" & vbCrLf
      Response.Write "<td nowrap width='150' >"&rs("DeliveryAddr")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("PlanStarteDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("PlanEndDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("StarteDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("EndDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("mileage")&"</td>" & vbCrLf
	  if rs("CheckFlag")=1 then
	  CheckFlag="主管审核"
	  elseif rs("CheckFlag")=2 then
	  CheckFlag="车队一审"
	  elseif rs("CheckFlag")=3 then
	  CheckFlag="门卫一审"
	  elseif rs("CheckFlag")=4 then
	  CheckFlag="门卫二审"
	  elseif rs("CheckFlag")=5 then
	  CheckFlag="车队二审"
	  else
	  CheckFlag="未审核"
	  end if
	  if rs("PrintFlag")=1 then
	  PrintFlag="√"
	  else
	  PrintFlag="×"
	  end if
      Response.Write "<td nowrap>"&CheckFlag&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&PrintFlag&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Importance")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("CarNumber")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Driver")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Planmileage")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("totalTime")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Remark")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
  end if
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
<%
  end if
end function 
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

%>


