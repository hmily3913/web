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
if Instr(session("AdminPurview"),"|80,")=0 then 
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
if Month(now()) <=10 then
laststrmonth="0"&eval(Month(now())-1)
else
laststrmonth=Month(now())
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
dim num11,num12
dim num21,num22
dim num31,num32
dim num41,num42
dim num51,num52
StartDate=stryear&"-"&strmonth&"-"&sdaynum
EndDate=stryear&"-"&strmonth&"-"&edaynum
sql="select * from financesys where YearNum='"&stryear&"' and period='"&laststrmonth&"'" 
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num11=rs("PayPrompt")	
	num21=rs("GatherPrompt")	
	num31=rs("PayCheckPrompt")	
	num41=rs("GatherCheckPrompt")	
sql="select * from financesys where YearNum='"&stryear&"' and period='"&strmonth&"'" 
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num12=rs("PayPrompt")	
	num22=rs("GatherPrompt")	
	num32=rs("PayCheckPrompt")	
	num42=rs("GatherCheckPrompt")	
  rs.close
  set rs=nothing
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap  class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>财务系统报表汇总信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><font style="left:inherit ">&nbsp;报表周期从<%response.Write(stryear&"-"&strmonth&"-"&sdaynum)%>至<%response.Write(stryear&"-"&strmonth&"-"&edaynum)%></font></td>
  </tr>  <tr>
    <td height="36"  align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="1">
      <tr bgcolor='#99BBE8'>
          <td nowrap width="162"> 名称
          </td>
          <td nowrap width="54"> 上月数据
          </td>
          <td nowrap width="54"> 本月数据
          </td>
          <td nowrap width="30"> 单位
          </td>
          <td nowrap width="433"> 计算描述
          </td>
	  </tr>
	  <tr>
          <td nowrap> 同种物料的采购频率
          </td>
          <td nowrap> 查看明细
          </td>
          <td nowrap>
          </td>
		  <td nowrap></td>
          <td nowrap> 同种物料的下单次数</td>
          
      </tr>
	  <tr>
          <td nowrap> 付款及时率
          </td>
          <td nowrap> 0
          </td>
          <td nowrap> 0
          </td>
		  <td nowrap>%</td>
          <td nowrap> (及时付款笔数/付款总笔数)*100%</td>
          
      </tr>
	  <tr>
          <td nowrap> 应收账款收款到账率
          </td>
          <td nowrap> <%=num21%>
          </td>
          <td nowrap> <%=num22%>
          </td>
		  <td nowrap>%</td>
          <td nowrap> (到期已收金额/到期应收金额)*100%</td>
          
      </tr>
	  <tr>
          <td nowrap> 供应商应付账款对账未及时率
          </td>
          <td nowrap> <%=num31%>
          </td>
          <td nowrap> <%=num32%>
          </td>
		  <td nowrap>%</td>
          <td nowrap> （及时核对笔数/应核对总笔数）*100%</td>
          
      </tr>
	  <tr>
          <td nowrap> 客户应收账款对账及时率
          </td>
          <td nowrap> <%=num41%>
          </td>
          <td nowrap> <%=num42%>
          </td>
		  <td nowrap>%</td>
          <td nowrap> (及时核对笔数/应核对总笔数)*100%</td>
          
      </tr>
    </table>      </td>    
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td height="30"></td>
  </tr>
</table>


</BODY>
</HTML>
