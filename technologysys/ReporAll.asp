<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>产品列表</TITLE>
<link rel="stylesheet" href="../Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|60,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<%
dim sdaynum,edaynum
dim stryear,strmonth,strdate,laststrmonth
stryear=Year(now())
if Month(now() <=9) then
strmonth="0"&Month(now())
else
strmonth=Month(now())
end if
if Month(now())= 1 then
laststrmonth=12&"#"&eval(stryear-1)
else
laststrmonth=eval(Month(now())-1)
end if
strdate=date()
sdaynum=1
select case Month(now())
	case 2
	  if ((stryear mod 4=0) and (stryear mod 100>0)) or (stryear mod 400=0) then
	    edaynum=29
	  else
	    edaynum=28
	  end if
    case 4
	  edaynum=30
    case 6
	  edaynum=30
    case 9
	  edaynum=30
    case 11
	  edaynum=30
	case 1
	  edaynum=31
	  stryear=Year(now())-1
	case else
	  edaynum=31
end select
'response.Write(stryear&"-"&strmonth&"-"&edaynum)

dim rs,sql,sqlstr,StartDate,EndDate
dim Reachsum,unReachsum,Reachper
dim num11,num12,num13
dim num21,num22,num23
dim num31,num32,num33
dim num41,num42,num43
dim num51,num52,num53
StartDate=stryear&"-"&strmonth&"-"&sdaynum
EndDate=stryear&"-"&strmonth&"-"&edaynum
sql="select top 1 * from technologysys order by SerialNum desc" 
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num11=rs("MoldrepairOne")	
	num21=rs("MoldmakeOne")
	num31=rs("DevicerepairOne")
	num12=rs("Moldrepair")	
	num22=rs("Moldmake")
	num32=rs("Devicerepair")
  rs.close
  set rs=nothing
sql="select * from technologysys where UPTdate='"&split(getDateRangebyMonth(laststrmonth),"###")(1)&"'" 
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num13=rs("Moldrepair")	
	num23=rs("Moldmake")
	num33=rs("Devicerepair")
  rs.close
  set rs=nothing
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap  class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>生技系统报表汇总信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><font style="left:inherit ">&nbsp;报表周期从<%response.Write(stryear&"-"&strmonth&"-"&sdaynum)%>至<%response.Write(stryear&"-"&strmonth&"-"&edaynum)%></font></td>
  </tr>  <tr>
    <td height="36"  align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="1">
      <tr bgcolor='#99BBE8'>
          <td nowrap width="100"> 名称
          </td>
          <td nowrap width="60"> 上月数据
          </td>
          <td nowrap width="60"> 昨日数据
          </td>
          <td nowrap width="120"> 本月累计平均数据
          </td>
          <td nowrap width="40"> 单位
          </td>
          <td nowrap width="60"> 针对部门
          </td>
          <td nowrap width="200"> 计算描述
          </td>
	  </tr>
	  <tr>
          <td nowrap> 模具维修及时率
          </td>
          <td nowrap> <%=num13%>
          </td>
          <td nowrap> <%=num11%>
          </td>
          <td nowrap> <%=num12%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>生技部</td>
          <td nowrap> (延误件数/总件数)*100%</td>
          
      </tr>
      <tr>
          <td nowrap> 模具制作及时率</td>
           <td nowrap> <%=num23%>
          </td>
           <td nowrap> <%=num21%>
          </td>
          <td nowrap> <%=num22%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>生技部</td>
          <td nowrap> (延误件数/总件数)*100%</td>
      </tr>
      <tr>
          <td nowrap> 设备维修及时率</td>
           <td nowrap> <%=num33%>
          </td>
           <td nowrap> <%=num31%>
          </td>
          <td nowrap> <%=num32%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>生技部</td>
          <td nowrap> (延误件数/总件数)*100%</td>
      </tr>
    </table>      </td>    
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td height="30"></td>
  </tr>
</table>

<%	
  rs.close
  set rs=nothing
%>
</BODY>
</HTML>
