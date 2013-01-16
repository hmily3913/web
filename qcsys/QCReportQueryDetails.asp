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
<style>
A:link {
    color: #FF00FF;
	text-decoration: none;
}
</style>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp" -->
<BODY>
<%
'if Instr(session("AdminPurview"),"|40,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
dim Rtype,Rclass,start_date,end_date,zhouqi,Dtype,years
  Rtype=request("Rtype")
  Dtype=request("Dtype")
  Rclass=request("Rclass")
  zhouqi=request("zhouqi")
	years=request("years")
  dim rs,sql'sql语句
  dim i,d1,d2 '循环，百分率
  dim tcategories(),tvalues(),qctego(1),qcval(1),tcategories2(),tvalues2(),tcategories3(),tvalues3(),tcategories4(),tvalues4(),tcategories5(),tvalues5()'画图变量
 dim a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12'累计变量
 dim b1,b2,b3,b4,b5,b6'包材或者例外变量
 dim c1,c2,c3,c4
if Rtype="OneDay" and Rclass="QC" then
  start_date=zhouqi
  end_date=zhouqi
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、IQC进料批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>让步接收</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退货</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
  qctego(0)="让不接受"
  qctego(1)="退货"
    sql="select * from qcsys where uptdate='"&start_date&"'"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    if (not rs.eof) then
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&start_date&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("unComeCheckAOne")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(rs("unComeCheckAOne")-rs("unUnQualifiMtrOne"))&"</td>"
	  if rs("unComeCheckAOne") = 0 then
	  d1=0
	  else
	  d1=(rs("unComeCheckAOne")-rs("unUnQualifiMtrOne"))*100/rs("unComeCheckAOne")
	  end if
    sql="select count(1) as a From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1077 and datediff(d,a.fdate,'"&start_date&"')=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(0)=rs("a")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&start_date&"')""><font color='#FF00FF'>"&rs("a")&"</font></a></td>" & vbCrLf
    sql="select count(1) as b From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1036 and datediff(d,a.fdate,'"&start_date&"')=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(1)=rs("b")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&start_date&"')""><font color='#FF00FF'>"&rs("b")&"</font></a></td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
    end if
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <table style="display:none" id="datatb">
    <tr><td colspan="6"></td></tr>
    <tr>
    <td >日期</td>
    <td ></td>
    <td ></td>
    <td ></td>
    <td ></td>
    <td >批次达成率</td>
    </tr>
<% 
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
i=0
    sql="select * from qcsys where uptdate>=dateadd(d,-9,'"&start_date&"') and uptdate<='"&start_date&"'"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
	while (not rs.eof) 
	  tcategories(i)=day(rs("uptdate"))&"/"&month(rs("uptdate"))
	  if rs("unComeCheckAOne") = 0 then
	  tvalues(i)=0
	  else
	  tvalues(i)=formatnumber((rs("unComeCheckAOne")-rs("unUnQualifiMtrOne"))*100/rs("unComeCheckAOne"),2)
	  end if
		response.Write "<tr>"
		response.Write "<td>"&tcategories(i)&"</td>"
		response.Write "<td></td>"
		response.Write "<td></td>"
		response.Write "<td></td>"
		response.Write "<td></td>"
		response.Write "<td>"&tvalues(i)&"</td>"
		response.Write "</tr>"
		
	  i=i+1
	  rs.movenext
	wend
		response.Write "</table>"
	dim ssss,bbbbb

%>
    <table style="display:none" id="datatb1">
    <tr><td colspan="3"></td></tr>
    <tr><td><%=qctego(0)%></td><td><%=qcval(0)%></td><td><%=formatnumber((qcval(0)*100/(qcval(0)+qcval(1))),2)%>%</td></tr>
    <tr><td><%=qctego(1)%></td><td><%=qcval(1)%></td><td><%=formatnumber((qcval(1)*100/(qcval(0)+qcval(1))),2)%>%</td></tr>
    </table>
<div id="container" style="width: 600px; height: 300px; margin: 0 auto; display:inline;"></div>	
<div id="container1" style="width: 500px; height: 300px; margin: 0 auto; display:inline;"></div>

	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><strong>二、OQC出货批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>出货批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格原因</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格批次</strong></font></td>
  </tr>
 <%
    sql="select count(1) as a1 from t_DeliverCheck where datediff(d,Fdate,'"&start_date&"')=0 and fuser>0 "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then c1=rs("a1")
    sql="select count(1) as a1 from t_DeliverCheck where datediff(d,Fdate,'"&start_date&"')=0 and fuser>0 and FComboBox='合格'"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then c2=rs("a1")
	  if c1 = 0 then
	  d1=0
	  else
	  d1=c2*100/c1
	  end if
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zhouqi&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('OQC','"&start_date&"','"&start_date&"')""><font color='#FF00FF'>"&(c1-c2)&"</font></a></td>" & vbCrLf
     Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  
elseif Rtype="OneDay" and Rclass="MN1" then
  start_date=zhouqi
  end_date=zhouqi
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>工段</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
  qctego(0)="让不接受"
  qctego(1)="退货"
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6,fname,FBase1 "&_
"from ( "&_
"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
"from t_t_productqualitycountEntry b inner join  "&_
"t_productqualitycount a on a.FID=b.FID left join  "&_
"t_workcenter c on c.fitemid=b.FBase1 left join "&_
"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
"where left(d.fnumber,2)='10' and d.fnumber<>'10.04' and datediff(d,b.fdate1,'"&start_date&"')=0) aaa "&_
"group by fname,FBase1 "
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
i=0
	dim rs2,sql2
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  tcategories(i)=rs("fname")
	  tcategories2(i)=rs("fname")
	  tcategories3(i)=rs("fname")
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fname")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='首件' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='制程' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  rs.movenext
    wend
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
</table>
  </div>

<%  
elseif Rtype="OneDay" and Rclass="MN2" then
  start_date=zhouqi
  end_date=zhouqi
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>工段</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6,fname,FBase1 "&_
"from ( "&_
"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
"from t_t_productqualitycountEntry b inner join  "&_
"t_productqualitycount a on a.FID=b.FID left join  "&_
"t_workcenter c on c.fitemid=b.FBase1 left join "&_
"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
"where left(d.fnumber,2)='11' and datediff(d,b.fdate1,'"&start_date&"')=0) aaa "&_
"group by fname,FBase1 "
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
i=0
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  tcategories(i)=rs("fname")
	  tcategories2(i)=rs("fname")
	  tcategories3(i)=rs("fname")
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fname")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='首件' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+cdbl(rs2("a1"))
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='制程' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'> "&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  rs.movenext
    wend
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'> "&a8&" </font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>

	</td>
  </tr>
</table>
  </div>

<%  
elseif Rtype="OneDay" and Rclass="MN3" then
  start_date=zhouqi
  end_date=zhouqi
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>工段</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6,fname,FBase1 "&_
"from ( "&_
"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
"from t_t_productqualitycountEntry b inner join  "&_
"t_productqualitycount a on a.FID=b.FID left join  "&_
"t_workcenter c on c.fitemid=b.FBase1 left join "&_
"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
"where left(d.fnumber,2)='12' and datediff(d,b.fdate1,'"&start_date&"')=0) aaa "&_
"group by fname,FBase1 "
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
i=0
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  tcategories(i)=rs("fname")
	  tcategories2(i)=rs("fname")
	  tcategories3(i)=rs("fname")
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fname")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='首件' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='制程' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'> "&rs2("a1")&" </font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  rs.movenext
    wend
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'> "&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>

	</td>
  </tr>
</table>
  </div>

<%  
elseif Rtype="OneDay" and Rclass="MN4" then
  start_date=zhouqi
  end_date=zhouqi
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>工段</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
  qctego(0)="让不接受"
  qctego(1)="退货"
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6,fname,FBase1 "&_
"from ( "&_
"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
"from t_t_productqualitycountEntry b inner join  "&_
"t_productqualitycount a on a.FID=b.FID left join  "&_
"t_workcenter c on c.fitemid=b.FBase1 left join "&_
"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
"where d.fnumber='10.04' and datediff(d,b.fdate1,'"&start_date&"')=0) aaa "&_
"group by fname,FBase1 "
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
i=0
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  tcategories(i)=rs("fname")
	  tcategories2(i)=rs("fname")
	  tcategories3(i)=rs("fname")
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fname")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='首件' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='制程' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  rs.movenext
    wend
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
</table>
  </div>

<%  
elseif Rtype="OneDay" and Rclass="MN" then
  start_date=zhouqi
  end_date=zhouqi
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>工段</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
  qctego(0)="让不接受"
  qctego(1)="退货"
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6,fname,FBase1 "&_
"from ( "&_
"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
"from t_t_productqualitycountEntry b inner join  "&_
"t_productqualitycount a on a.FID=b.FID left join  "&_
"t_workcenter c on c.fitemid=b.FBase1 left join "&_
"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
"where datediff(d,b.fdate1,'"&start_date&"')=0) aaa "&_
"group by fname,FBase1 "
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
i=0
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  tcategories(i)=rs("fname")
	  tcategories2(i)=rs("fname")
	  tcategories3(i)=rs("fname")
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fname")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='首件' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified where FBase2="&rs("FBase1")&" and FText='制程' and datediff(d,FDate1,'"&start_date&"')=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  rs.movenext
    wend
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
</table>
  </div>

<%  
elseif Rtype="OneWeek" and Rclass="QC" then
	if years="2011" then
  start_date=dateadd("d",4,split(getDateRange(zhouqi,years),"###")(0))
  end_date=dateadd("d",4,split(getDateRange(zhouqi,years),"###")(1))
	else'if years="2012"
  start_date=split(getDateRange(zhouqi,years),"###")(0)
  end_date=split(getDateRange(zhouqi,years),"###")(1)
	end if
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、OQC出货批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>周别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>出货批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>返工</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>评审放行</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
    sql="select count(1) as a1 from t_DeliverCheck where fuser>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c1=rs("a1")
    end if
    sql="select count(1) as a2 from t_DeliverCheck where fuser>0 and FComboBox='合格' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c2=rs("a2")
    end if
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zhouqi&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c2&"</td>"
	  if c1= 0 then
	  d1=0
	  else
	  d1=c2*100/c1
	  end if
    sql="select count(1) as a from t_DeliverCheck where fuser>0 and FComboBox='不合格' and FComboBox1='返工' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(0)=rs("a")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('OQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'> "&rs("a")&"</font></a></td>" & vbCrLf
    sql="select count(1) as b from t_DeliverCheck where fuser>0 and FComboBox='不合格' and FComboBox1='出货' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(1)=rs("b")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('OQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'> "&rs("b")&"</font></a></td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF">&nbsp;<strong>二、IQC批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>周别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>让步接收</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退货</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率包材除外</strong></font></td>
  </tr>
 <%
    sql="select count(1) as a1 from t_supplycheck where fuser>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c1=rs("a1")
    end if
    sql="select count(1) as a2 from t_supplycheck where fuser>0 and FComboBox2='合格' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c2=rs("a2")
    end if
    sql="select count(1) as a1 from t_supplycheck where fuser>0 and FText3<>'包材' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c3=rs("a1")
    end if
    sql="select count(1) as a2 from t_supplycheck where fuser>0 and FComboBox2='合格' and FText3<>'包材' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c4=rs("a2")
    end if
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zhouqi&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c2&"</td>"
	  if c1 = 0 then
	  d1=0
	  else
	  d1=c2*100/c1
	  end if
	  if c3 = 0 then
	  d2=0
	  else
	  d2=c4*100/c3
	  end if
    sql="select count(1) as a From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1077 and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(0)=rs("a")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("a")&"</font></a></td>" & vbCrLf
    sql="select count(1) as b From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1036 and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(1)=rs("b")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("b")&"</font></a></td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF">&nbsp;<strong>三、IQC各类原材料进料情况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
 <tr style="display:none"><td colspan="7"></td></tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>特采批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退货批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格扣款额</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
i=0
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,a6 as b6 "&_
	" from (select 1 as a1, "&_
	" case when FComboBox2='不合格' then 1 else 0 end as a2, "&_
	" 0 as a3, "&_
	" 0 as a4, "&_
	" FDecimal2 as a5,FText3 as a6 "&_
	" from t_supplycheck where fuser>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0 "&_
	" union all "&_
	" select 0 as a1,0 as a2, "&_
	" case when b.FDefectHandlingID=1077 then 1 else 0 end as a3, "&_
	" case when b.FDefectHandlingID=1036 then 1 else 0 end as a4,0 as a5,FText3 as a6 "&_
	"  from qmreject a,qmrejectentry b,t_supplycheck c where a.FID_SRC=c.FID and a.fid=b.fid and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0 "&_
	" ) bbb "&_
	" group by a6 "&_
	" order by b6 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>" & vbCrLf
	  d1=(rs("b1")-rs("b2"))*100/rs("b1")
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  if rs("b6")="包材" then
		  b1=rs("b1")
		  b2=rs("b2")
		  b3=rs("b3")
		  b4=rs("b4")
		  b5=rs("b5")
	  end if
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  a5=a5+rs("b5")
	  tcategories(i)=rs("b6")
	  tvalues(i)=formatnumber(d1,2)
	  i=i+1
	  rs.movenext
    wend
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>" & vbCrLf
	  d1=(a1-a2)*100/a1
	  tcategories(i)="合计"
	  tvalues(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>包材除外</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a1-b1)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a2-b2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a3-b3)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a4-b4)&"</td>" & vbCrLf
	  d1=((a1-b1)-(a2-b2))*100/(a1-b1)
	  tcategories(i+1)="包材除外"
	  tvalues(i+1)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a5-b5)&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 700px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
</table>
</div>
<%
elseif Rtype="OneWeek" and Rclass="MN1" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>周别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
dim lastzhouqi
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0
dim b
for b=lastzhouqi to zhouqi
	if years="2011" then
  start_date=dateadd("d",4,split(getDateRange(b,years),"###")(0))
  end_date=dateadd("d",4,split(getDateRange(b,years),"###")(1))
	else'if years="2012"
  start_date=split(getDateRange(b,years),"###")(0)
  end_date=split(getDateRange(b,years),"###")(1)
	end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='10' and d.fnumber<>'10.04' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"周"
	  tcategories2(i)=b&"周"
	  tcategories3(i)=b&"周"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"周</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 部门代码='KD01.0001.0010' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      Response.Write "</tr>" & vbCrLf
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneWeek" and Rclass="MN2" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>周别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0

for b=lastzhouqi to zhouqi
	if years="2011" then
  start_date=dateadd("d",4,split(getDateRange(b,years),"###")(0))
  end_date=dateadd("d",4,split(getDateRange(b,years),"###")(1))
	else'if years="2012"
  start_date=split(getDateRange(b,years),"###")(0)
  end_date=split(getDateRange(b,years),"###")(1)
	end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='11' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"周"
	  tcategories2(i)=b&"周"
	  tcategories3(i)=b&"周"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"周</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'> "&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 部门代码='KD01.0001.0007' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      Response.Write "</tr>" & vbCrLf
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'> "&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneWeek" and Rclass="MN3" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>周别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0
for b=lastzhouqi to zhouqi
	if years="2011" then
  start_date=dateadd("d",4,split(getDateRange(b,years),"###")(0))
  end_date=dateadd("d",4,split(getDateRange(b,years),"###")(1))
	else'if years="2012"
  start_date=split(getDateRange(b,years),"###")(0)
  end_date=split(getDateRange(b,years),"###")(1)
	end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='12' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"周"
	  tcategories2(i)=b&"周"
	  tcategories3(i)=b&"周"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"周</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""></font>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 部门代码='KD01.0001.0008' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      Response.Write "</tr>" & vbCrLf
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneWeek" and Rclass="MN4" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>周别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0
for b=lastzhouqi to zhouqi
	if years="2011" then
  start_date=dateadd("d",4,split(getDateRange(b,years),"###")(0))
  end_date=dateadd("d",4,split(getDateRange(b,years),"###")(1))
	else'if years="2012"
  start_date=split(getDateRange(b,years),"###")(0)
  end_date=split(getDateRange(b,years),"###")(1)
	end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where d.fnumber='10.04' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"周"
	  tcategories2(i)=b&"周"

	  tcategories3(i)=b&"周"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"周</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 部门代码='KD01.0001.0010' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      Response.Write "</tr>" & vbCrLf
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneWeek" and Rclass="MN" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>周别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tcategories2(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tcategories3(9)
ReDim Preserve tvalues3(9)
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0
for b=lastzhouqi to zhouqi
	if years="2011" then
  start_date=dateadd("d",4,split(getDateRange(b,years),"###")(0))
  end_date=dateadd("d",4,split(getDateRange(b,years),"###")(1))
	else'if years="2012"
  start_date=split(getDateRange(b,years),"###")(0)
  end_date=split(getDateRange(b,years),"###")(1)
	end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"周"
	  tcategories2(i)=b&"周"

	  tcategories3(i)=b&"周"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"周</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 部门代码='KD01.0001.0010' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      Response.Write "</tr>" & vbCrLf
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneMonth" and Rclass="QC" then
  start_date=split(getDateRangebyMonth(zhouqi&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(zhouqi&"#"&years),"###")(1)
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、OQC出货批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
 <tr style="display:none"><td colspan="7"></td></tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>出货批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>返工</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>评审放行</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
    sql="select sum(unFinishQOne) as 出货批次,sum(FinishQOne) as 合格批次 from qcsys where uptdate>='"&start_date&"' and uptdate<='"&end_date&"'"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    if (not rs.eof) then
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zhouqi&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("出货批次")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("合格批次")&"</td>"
	  if rs("出货批次") = 0 then
	  d1=0
	  else
	  d1=rs("合格批次")*100/rs("出货批次")
	  end if
    sql="select count(1) as a from t_DeliverCheck where fuser>0 and FComboBox='不合格' and FComboBox1='返工' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(0)=rs("a")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('OQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("a")&"</font></a></td>" & vbCrLf
    sql="select count(1) as b from t_DeliverCheck where fuser>0 and FComboBox='不合格' and FComboBox1='出货' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(1)=rs("b")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('OQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("b")&"</font></a></td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
    end if
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF">&nbsp;<strong>二、IQC批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>让步接收</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退货</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率包材除外</strong></font></td>
  </tr>
 <%
  qctego(0)="让不接受"
  qctego(1)="退货"
    sql="select count(1) as a1 from t_supplycheck where fuser>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c1=rs("a1")
    end if
    sql="select count(1) as a2 from t_supplycheck where fuser>0 and FComboBox2='合格' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c2=rs("a2")
    end if
    sql="select count(1) as a1 from t_supplycheck where fuser>0 and FText3<>'包材' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c3=rs("a1")
    end if
    sql="select count(1) as a2 from t_supplycheck where fuser>0 and FComboBox2='合格' and FText3<>'包材' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c4=rs("a2")
    end if
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zhouqi&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c2&"</td>"
	  if c1 = 0 then
	  d1=0
	  else
	  d1=c2*100/c1
	  end if
	  if c3 = 0 then
	  d2=0
	  else
	  d2=c4*100/c3
	  end if
    sql="select count(1) as a From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1077 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(0)=rs("a")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("a")&"</font></a></td>" & vbCrLf
    sql="select count(1) as b From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1036 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(1)=rs("b")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("b")&"</font></a></td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF">&nbsp;<strong>三、IQC各类原材料进料情况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
 <tr style="display:none"><td colspan="7"></td></tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>特采批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退货批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格扣款额</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
i=0
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,a6 as b6 "&_
	" from (select 1 as a1, "&_
	" case when FComboBox2='不合格' then 1 else 0 end as a2, "&_
	" 0 as a3, "&_
	" 0 as a4, "&_
	" FDecimal2 as a5,FText3 as a6 "&_
	" from t_supplycheck where fuser>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0 "&_
	" union all "&_
	" select 0 as a1,0 as a2, "&_
	" case when b.FDefectHandlingID=1077 then 1 else 0 end as a3, "&_
	" case when b.FDefectHandlingID=1036 then 1 else 0 end as a4,0 as a5,FText3 as a6 "&_
	"  from qmreject a,qmrejectentry b,t_supplycheck c where a.FID_SRC=c.FID and a.fid=b.fid and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0 "&_
	" ) bbb "&_
	" group by a6 "&_
	" order by b6 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>" & vbCrLf
	  d1=(rs("b1")-rs("b2"))*100/rs("b1")
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  if rs("b6")="包材" then
		  b1=rs("b1")
		  b2=rs("b2")
		  b3=rs("b3")
		  b4=rs("b4")
		  b5=rs("b5")
	  end if
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  a5=a5+rs("b5")
	  tcategories(i)=rs("b6")
	  tvalues(i)=formatnumber(d1,2)
	  i=i+1
	  rs.movenext
    wend
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>" & vbCrLf
	  d1=(a1-a2)*100/a1
	  tcategories(i)="合计"
	  tvalues(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>包材除外</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a1-b1)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a2-b2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a3-b3)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a4-b4)&"</td>" & vbCrLf
	  d1=((a1-b1)-(a2-b2))*100/(a1-b1)
	  tcategories(i+1)="包材除外"
	  tvalues(i+1)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a5-b5)&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <table style="display:none" id="datatb1">
    <tr><td colspan="3"></td></tr>
    <tr><td><%=qctego(0)%></td><td><%=qcval(0)%></td><td><%=formatnumber((qcval(0)*100/(qcval(0)+qcval(1))),2)%>%</td></tr>
    <tr><td><%=qctego(1)%></td><td><%=qcval(1)%></td><td><%=formatnumber((qcval(1)*100/(qcval(0)+qcval(1))),2)%>%</td></tr>
    </table>
<div id="container" style="width: 600px; height: 300px; margin: 0 auto; display:inline;"></div>	
<div id="container1" style="width: 500px; height: 300px; margin: 0 auto; display:inline;"></div>
		</td>
  </tr>
</table>
</div>
<%
elseif Rtype="OneMonth" and Rclass="MN1" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
lastzhouqi
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0
b
for b=1 to zhouqi
  start_date=split(getDateRangebyMonth(b&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(b&"#"&years),"###")(1)
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='10' and d.fnumber<>'10.04' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"月"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"月</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where left(部门代码,2)='10' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and left(d.fnumber,2)='10' and d.fnumber<>'10.04') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneMonth" and Rclass="MN2" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
lastzhouqi
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0
b
for b=1 to zhouqi
  start_date=split(getDateRangebyMonth(b&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(b&"#"&years),"###")(1)
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='11' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"月"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"月</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 部门代码='KD01.0001.0007' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and left(d.fnumber,2)='11') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneMonth" and Rclass="MN3" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
lastzhouqi
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0
b
for b=1 to zhouqi
  start_date=split(getDateRangebyMonth(b&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(b&"#"&years),"###")(1)
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='12' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"月"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"月</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 部门代码='KD01.0001.0008' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and left(d.fnumber,2)='12') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	tcategories2(i)=rs("ftext4")
	tvalues(i)=rs("a1")
	i=i+1
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneMonth" and Rclass="MN4" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
lastzhouqi
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0
b
for b=1 to zhouqi
  start_date=split(getDateRangebyMonth(b&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(b&"#"&years),"###")(1)
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where d.fnumber='10.04' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"月"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"月</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where left(部门代码,2)='10' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and d.fnumber='10.04') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneMonth" and Rclass="MN" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>月份</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
lastzhouqi
if zhouqi>=4 then 
lastzhouqi=zhouqi-3
else
lastzhouqi=1
end if
i=0
b
for b=1 to zhouqi
  start_date=split(getDateRangebyMonth(b&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(b&"#"&years),"###")(1)
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"月"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"月</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid) aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Dtype="IQC" then
start_date=request("start_date")
end_date=request("end_date")
if request("print_tag")=1 then
	response.ContentType("application/vnd.ms-excel")
	response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
%>
 <div id="Detailslisttable" style="width:1190px; height:380px; z-index:600">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closeadDetails()" >&nbsp;<strong>IQC不合格物料处置状况</strong></font> <input type="button" value="导出" onclick="OutDetails('IQC','<%=start_date%>','<%=end_date%>')" /></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>序号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>供方</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物料名称</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>品号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格原因</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>处置结果</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>扣款额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>确认人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格分类</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物料分类</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客户等级</strong></font></td>
  </tr>
 <%
    sql="select a.fdate as 日期,b.fnumber as 供方,c.fname as 物料名称,a.ftext as 订单号,c.fnumber as 品号,a.FDecimal as 进料数量,f.FDefectQty as 不合格数量,a.FText1 as 不合格原因,"&_
	" case when a.FInteger=0 then '未处理' when FDefectHandlingID=1077 then '让步接收' else '退货' end as 处置结果,a.FDecimal2 as 扣款额,d.fname as 确认人,a.FText2 as 不合格分类,a.FText3 as 物料分类 "&_
	" from t_supplycheck a left join t_Supplier b on a.fbase=b.fitemid left join  "&_
	" t_ICItem c on a.fbase1=c.fitemid left join t_emp d on a.FBase4=d.fitemid left join "&_
	" qmreject e on a.fbillno=e.FBillNo_SRC and a.fbase1=e.fitemid left join qmrejectentry f on e.fid=f.fid "&_
	" where a.fuser>0 and a.FComboBox2='不合格' "&_
	" and datediff(d,e.fdate,'"&start_date&"')<=0 and datediff(d,e.fdate,'"&end_date&"')>=0 "
'	" where (a.fuser>0 and a.FComboBox2='不合格' and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0) "&_
'	" or (datediff(d,f.fdealdate,'"&start_date&"')<=0 and datediff(d,f.fdealdate,'"&end_date&"')>=0) "
dim zzzz:zzzz=0
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
		zzzz=zzzz+1
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zzzz&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("日期")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("供方")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("物料名称")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("订单号")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("品号")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("进料数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("不合格数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("不合格原因")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("处置结果")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("扣款额")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("确认人")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("不合格分类")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("物料分类")&"</td>" & vbCrLf
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend
  %>
  </table>
	</td>
  </tr>
</table>
</div>
<%
elseif Dtype="OQC" then
start_date=request("start_date")
end_date=request("end_date")
%>
 <div id="Detailslisttable" style="width:1190px; height:380px; z-index:600">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closeadDetails()" >&nbsp;<strong>OQC出货检查状况</strong></font><input type="button" value="导出" onclick="OutDetails('OQC','<%=start_date%>','<%=end_date%>')" /></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>序号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>业务员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品型号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>判定</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不良描述</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>处理方式</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>备注</strong></font></td>
  </tr>
 <%
    sql="select a.fdate as 日期,b.fname as 业务员,a.FText as 订单号,c.fname as 产品型号,a.FQty as 数量, "&_
	" a.FDecimal as 件数,a.FComboBox as 判定,a.FText1 as 不良描述,a.FComboBox1 as 处理方式,a.FNOTE as 备注 "&_
	"  from t_DeliverCheck a left join t_emp b on a.FBase=b.fitemid left join "&_
	" t_ICItem c on a.FBase1= c.fitemid "&_
	" where a.fuser>0 and a.FComboBox='不合格' and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		zzzz=0
    while (not rs.eof)
		zzzz=zzzz+1
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zzzz&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("日期")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("业务员")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("订单号")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("产品型号")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("件数")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("判定")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("不良描述")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("处理方式")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("备注")&"</td>" & vbCrLf
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend
  %>
  </table>
	</td>
  </tr>
</table>
</div>
<%
elseif Dtype="UQF" then
start_date=request("start_date")
end_date=request("end_date")
%>
 <div id="Detailslisttable" style="width:1190px; height:380px; z-index:600">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closeadDetails()" >&nbsp;<strong>不合格明细</strong></font>
&nbsp;<font style="background-color:#ffff66">一分厂</font>&nbsp;
<font style="background-color:#ff99ff">二分厂</font>&nbsp;
<font style="background-color:#B9BBC7">三分厂</font>&nbsp;<input type="button" value="导出" onclick="OutDetails('UQF','<%=start_date%>','<%=end_date%>')" />
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>序号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格比例</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>原因分析</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>对策(5W1H)</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>负责人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>预定完成</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>确认人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>备注</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>现象分析</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>原因分析</strong></font></td>
  </tr>
 <%
    sql="select a.*,b.fname,left(c.fnumber,2) as partment,t_emp.fname as tuser "&_
	"  from t_unqualified a left join t_emp b on a.FBase=b.fitemid left join t_department c on c.fitemid=a.FBase1 left join t_emp on a.fbase8=t_emp.fitemid "&_
	" where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		zzzz=0
    while (not rs.eof)
		zzzz=zzzz+1
	dim uqfbg:uqfbg="#EBF2F9"
	if rs("partment")=10 then
	  uqfbg="#ffff66"
	elseif rs("partment")=11 then
	  uqfbg="#ff99ff"
	elseif rs("partment")=12 then
	  uqfbg="#B9BBC7"
	end if
	  Response.Write "<tr bgcolor='"&uqfbg&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zzzz&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FDate1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FDecimal")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FDecimal1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fname")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FDate3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("tuser")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FNOTE")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText5")&"</td>" & vbCrLf
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend
  %>
  </table>
	</td>
  </tr>
</table>
</div>
<%

elseif Rtype="OneSeason" and Rclass="QC" then
  if zhouqi=1 then
  start_date=split(getDateRangebyMonth(1&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(3&"#"&years),"###")(1)
  elseif zhouqi=2 then
  start_date=split(getDateRangebyMonth(4&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(6&"#"&years),"###")(1)
  elseif zhouqi=3 then
  start_date=split(getDateRangebyMonth(7&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(9&"#"&years),"###")(1)
  elseif zhouqi=4 then
  start_date=split(getDateRangebyMonth(10&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(12&"#"&years),"###")(1)
  end if
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、OQC出货批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>季度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>出货批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>返工</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>评审放行</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
    sql="select sum(unFinishQOne) as 出货批次,sum(FinishQOne) as 合格批次 from qcsys where uptdate>='"&start_date&"' and uptdate<='"&end_date&"'"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    if (not rs.eof) then
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zhouqi&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("出货批次")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("合格批次")&"</td>"
	  if rs("出货批次") = 0 then
	  d1=0
	  else
	  d1=rs("合格批次")*100/rs("出货批次")
	  end if
    sql="select count(1) as a from t_DeliverCheck where fuser>0 and FComboBox='不合格' and FComboBox1='返工' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(0)=rs("a")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('OQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("a")&"</font></a></td>" & vbCrLf
    sql="select count(1) as b from t_DeliverCheck where fuser>0 and FComboBox='不合格' and FComboBox1='出货' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(1)=rs("b")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('OQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("b")&"</font></a></td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
    end if
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF">&nbsp;<strong>二、IQC批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>季度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>让步接收</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退货</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率包材除外</strong></font></td>
  </tr>
 <%
  qctego(0)="让不接受"
  qctego(1)="退货"
    sql="select count(1) as a1 from t_supplycheck where fuser>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c1=rs("a1")
    end if
    sql="select count(1) as a2 from t_supplycheck where fuser>0 and FComboBox2='合格' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c2=rs("a2")
    end if
    sql="select count(1) as a1 from t_supplycheck where fuser>0 and FText3<>'包材' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c3=rs("a1")
    end if
    sql="select count(1) as a2 from t_supplycheck where fuser>0 and FComboBox2='合格' and FText3<>'包材' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c4=rs("a2")
    end if
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zhouqi&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c2&"</td>"
	  if c1 = 0 then
	  d1=0
	  else
	  d1=c2*100/c1
	  end if
	  if c3 = 0 then
	  d2=0
	  else
	  d2=c4*100/c3
	  end if
    sql="select count(1) as a From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1077 and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(0)=rs("a")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("a")&"</font></a></td>" & vbCrLf
    sql="select count(1) as b From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1036 and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(1)=rs("b")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("b")&"</font></a></td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF">&nbsp;<strong>三、IQC各类原材料进料情况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
 <tr style="display:none"><td colspan="7"></td></tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>特采批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退货批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格扣款额</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
i=0
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,a6 as b6 "&_
	" from (select 1 as a1, "&_
	" case when FComboBox2='不合格' then 1 else 0 end as a2, "&_
	" 0 as a3, "&_
	" 0 as a4, "&_
	" FDecimal2 as a5,FText3 as a6 "&_
	" from t_supplycheck where fuser>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0 "&_
	" union all "&_
	" select 0 as a1,0 as a2, "&_
	" case when b.FDefectHandlingID=1077 then 1 else 0 end as a3, "&_
	" case when b.FDefectHandlingID=1036 then 1 else 0 end as a4,0 as a5,FText3 as a6 "&_
	"  from qmreject a,qmrejectentry b,t_supplycheck c where a.FID_SRC=c.FID and a.fid=b.fid and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0 "&_
	" ) bbb "&_
	" group by a6 "&_
	" order by b6 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>" & vbCrLf
	  d1=(rs("b1")-rs("b2"))*100/rs("b1")
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  if rs("b6")="包材" then
		  b1=rs("b1")
		  b2=rs("b2")
		  b3=rs("b3")
		  b4=rs("b4")
		  b5=rs("b5")
	  end if
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  a5=a5+rs("b5")
	  tcategories(i)=rs("b6")
	  tvalues(i)=formatnumber(d1,2)
	  i=i+1
	  rs.movenext
    wend
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>" & vbCrLf
	  d1=(a1-a2)*100/a1
	  tcategories(i)="合计"
	  tvalues(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>包材除外</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a1-b1)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a2-b2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a3-b3)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a4-b4)&"</td>" & vbCrLf
	  d1=((a1-b1)-(a2-b2))*100/(a1-b1)
	  tcategories(i+1)="包材除外"
	  tvalues(i+1)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a5-b5)&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <table style="display:none" id="datatb1">
    <tr><td colspan="3"></td></tr>
    <tr><td><%=qctego(0)%></td><td><%=qcval(0)%></td><td><%=formatnumber((qcval(0)*100/(qcval(0)+qcval(1))),2)%>%</td></tr>
    <tr><td><%=qctego(1)%></td><td><%=qcval(1)%></td><td><%=formatnumber((qcval(1)*100/(qcval(0)+qcval(1))),2)%>%</td></tr>
    </table>
<div id="container" style="width: 600px; height: 300px; margin: 0 auto; display:inline;"></div>	
<div id="container1" style="width: 500px; height: 300px; margin: 0 auto; display:inline;"></div>

	</td>
  </tr>
</table>
</div>
<%
elseif Rtype="OneSeason" and Rclass="MN1" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>

  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>季度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=4 then 
zhouqi=4
else
lastzhouqi=1
end if
i=0
for b=1 to zhouqi
  if b=1 then
  start_date=split(getDateRangebyMonth(1&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(3&"#"&years),"###")(1)
  elseif b=2 then
  start_date=split(getDateRangebyMonth(4&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(6&"#"&years),"###")(1)
  elseif b=3 then
  start_date=split(getDateRangebyMonth(7&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(9&"#"&years),"###")(1)
  elseif b=4 then
  start_date=split(getDateRangebyMonth(10&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(12&"#"&years),"###")(1)
  end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='10' and d.fnumber<>'10.04' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"季"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"季</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 部门代码='KD01.0001.0010' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and left(d.fnumber,2)='10' and d.fnumber<>'10.04') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0

	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneSeason" and Rclass="MN2" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>季度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=4 then 
zhouqi=4
else
lastzhouqi=1
end if
i=0
for b=1 to zhouqi
  if b=1 then
  start_date=split(getDateRangebyMonth(1&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(3&"#"&years),"###")(1)
  elseif b=2 then
  start_date=split(getDateRangebyMonth(4&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(6&"#"&years),"###")(1)
  elseif b=3 then
  start_date=split(getDateRangebyMonth(7&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(9&"#"&years),"###")(1)
  elseif b=4 then
  start_date=split(getDateRangebyMonth(10&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(12&"#"&years),"###")(1)
  end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='11' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"季"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"季</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where left(部门代码,2)='07' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and left(d.fnumber,2)='11') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneSeason" and Rclass="MN3" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>季度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=4 then 
zhouqi=4
else
lastzhouqi=1
end if
i=0
for b=1 to zhouqi
  if b=1 then
  start_date=split(getDateRangebyMonth(1&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(3&"#"&years),"###")(1)
  elseif b=2 then
  start_date=split(getDateRangebyMonth(4&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(6&"#"&years),"###")(1)
  elseif b=3 then
  start_date=split(getDateRangebyMonth(7&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(9&"#"&years),"###")(1)
  elseif b=4 then
  start_date=split(getDateRangebyMonth(10&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(12&"#"&years),"###")(1)
  end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='12' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"季"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"季</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where left(部门代码,2)='08' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and left(d.fnumber,2)='12') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  


elseif Rtype="OneSeason" and Rclass="MN4" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>

  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>季度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=4 then 
zhouqi=4
else
lastzhouqi=1
end if
i=0
for b=1 to zhouqi
  if b=1 then
  start_date=split(getDateRangebyMonth(1&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(3&"#"&years),"###")(1)
  elseif b=2 then
  start_date=split(getDateRangebyMonth(4&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(6&"#"&years),"###")(1)
  elseif b=3 then
  start_date=split(getDateRangebyMonth(7&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(9&"#"&years),"###")(1)
  elseif b=4 then
  start_date=split(getDateRangebyMonth(10&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(12&"#"&years),"###")(1)
  end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where d.fnumber='10.04' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"季"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"季</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 部门代码='KD01.0001.0010' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and d.fnumber='10.04') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0

	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneSeason" and Rclass="MN" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>

  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>季度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=4 then 
zhouqi=4
else
lastzhouqi=1
end if
i=0
for b=1 to zhouqi
  if b=1 then
  start_date=split(getDateRangebyMonth(1&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(3&"#"&years),"###")(1)
  elseif b=2 then
  start_date=split(getDateRangebyMonth(4&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(6&"#"&years),"###")(1)
  elseif b=3 then
  start_date=split(getDateRangebyMonth(7&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(9&"#"&years),"###")(1)
  elseif b=4 then
  start_date=split(getDateRangebyMonth(10&"#"&years),"###")(0)
  end_date=split(getDateRangebyMonth(12&"#"&years),"###")(1)
  end if
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"季"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"季</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid) aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0

	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneYear" and Rclass="QC" then
  start_date=zhouqi&"-01-01"
  end_date=zhouqi&"-12-31"
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、OQC出货批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>年度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>出货批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>返工</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>评审放行</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
  </tr>
 <%
    sql="select sum(unFinishQOne) as 出货批次,sum(FinishQOne) as 合格批次 from qcsys where uptdate>='"&start_date&"' and uptdate<='"&end_date&"'"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    if (not rs.eof) then
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zhouqi&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("出货批次")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("合格批次")&"</td>"
	  if rs("出货批次") = 0 then
	  d1=0
	  else
	  d1=rs("合格批次")*100/rs("出货批次")
	  end if
    sql="select count(1) as a from t_DeliverCheck where fuser>0 and FComboBox='不合格' and FComboBox1='返工' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(0)=rs("a")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('OQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("a")&"</font></a></td>" & vbCrLf
    sql="select count(1) as b from t_DeliverCheck where fuser>0 and FComboBox='不合格' and FComboBox1='出货' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(1)=rs("b")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('OQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("b")&"</font></a></td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
    end if
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF">&nbsp;<strong>二、IQC批次合格率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>年度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>让步接收</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退货</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率包材除外</strong></font></td>
  </tr>
 <%
  qctego(0)="让不接受"
  qctego(1)="退货"
    sql="select count(1) as a1 from t_supplycheck where fuser>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")

    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c1=rs("a1")
    end if
    sql="select count(1) as a2 from t_supplycheck where fuser>0 and FComboBox2='合格' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c2=rs("a2")
    end if
    sql="select count(1) as a1 from t_supplycheck where fuser>0 and FText3<>'包材' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c3=rs("a1")
    end if
    sql="select count(1) as a2 from t_supplycheck where fuser>0 and FComboBox2='合格' and FText3<>'包材' and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	c4=rs("a2")
    end if
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&zhouqi&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&c2&"</td>"
	  if c1 = 0 then
	  d1=0
	  else
	  d1=c2*100/c1
	  end if
	  if c3 = 0 then
	  d2=0
	  else
	  d2=c4*100/c3
	  end if
    sql="select count(1) as a From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1077 and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(0)=rs("a")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("a")&"</font></a></td>" & vbCrLf
    sql="select count(1) as b From qmreject a,qmrejectentry b where a.fid=b.fid and b.FDefectHandlingID=1036 and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	qcval(1)=rs("b")
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('IQC','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs("b")&"</font></a></td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF">&nbsp;<strong>三、IQC各类原材料进料情况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
 <tr style="display:none"><td colspan="7"></td></tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>进料批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>特采批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>退货批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格扣款额</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
i=0
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,a6 as b6 "&_
	" from (select 1 as a1, "&_
	" case when FComboBox2='不合格' then 1 else 0 end as a2, "&_
	" 0 as a3, "&_
	" 0 as a4, "&_
	" FDecimal2 as a5,FText3 as a6 "&_
	" from t_supplycheck where fuser>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate,'"&end_date&"')>=0 "&_
	" union all "&_
	" select 0 as a1,0 as a2, "&_
	" case when b.FDefectHandlingID=1077 then 1 else 0 end as a3, "&_
	" case when b.FDefectHandlingID=1036 then 1 else 0 end as a4,0 as a5,FText3 as a6 "&_
	"  from qmreject a,qmrejectentry b,t_supplycheck c where a.FID_SRC=c.FID and a.fid=b.fid and datediff(d,a.fdate,'"&start_date&"')<=0 and datediff(d,a.fdate,'"&end_date&"')>=0 "&_
	" ) bbb "&_
	" group by a6 "&_
	" order by b6 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>" & vbCrLf
	  d1=(rs("b1")-rs("b2"))*100/rs("b1")
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  if rs("b6")="包材" then
		  b1=rs("b1")
		  b2=rs("b2")
		  b3=rs("b3")
		  b4=rs("b4")
		  b5=rs("b5")
	  end if
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  a5=a5+rs("b5")
	  tcategories(i)=rs("b6")
	  tvalues(i)=formatnumber(d1,2)
	  i=i+1
	  rs.movenext
    wend
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>" & vbCrLf
	  d1=(a1-a2)*100/a1
	  tcategories(i)="合计"
	  tvalues(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>包材除外</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a1-b1)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a2-b2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a3-b3)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a4-b4)&"</td>" & vbCrLf
	  d1=((a1-b1)-(a2-b2))*100/(a1-b1)
	  tcategories(i+1)="包材除外"
	  tvalues(i+1)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&(a5-b5)&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
    <table style="display:none" id="datatb1">
    <tr><td colspan="3"></td></tr>
    <tr><td><%=qctego(0)%></td><td><%=qcval(0)%></td><td><%=formatnumber((qcval(0)*100/(qcval(0)+qcval(1))),2)%>%</td></tr>
    <tr><td><%=qctego(1)%></td><td><%=qcval(1)%></td><td><%=formatnumber((qcval(1)*100/(qcval(0)+qcval(1))),2)%>%</td></tr>
    </table>
<div id="container" style="width: 600px; height: 300px; margin: 0 auto; display:inline;"></div>	
<div id="container1" style="width: 500px; height: 300px; margin: 0 auto; display:inline;"></div>
	</td>
  </tr>
</table>
</div>
<%
elseif Rtype="OneYear" and Rclass="MN1" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>年度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=2015 then 
lastzhouqi=zhouqi-4
else
lastzhouqi=2011
end if
i=0
for b=lastzhouqi to zhouqi
  start_date=b&"-01-01"
  end_date=b&"-12-31"
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='10' and d.fnumber<>'10.04' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"年"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"年</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where left(部门代码,2)='10' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and left(d.fnumber,2)='10' and d.fnumber<>'10.04') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='10' and c.fnumber<>'10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneYear" and Rclass="MN2" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>年度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=2015 then 
lastzhouqi=zhouqi-4
else
lastzhouqi=2011
end if
i=0
for b=lastzhouqi to zhouqi
  start_date=b&"-01-01"
  end_date=b&"-12-31"
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='11' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"年"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"年</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where left(部门代码,2)='07' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and left(d.fnumber,2)='11') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='11' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneYear" and Rclass="MN3" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>年度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=2015 then 
lastzhouqi=zhouqi-4
else
lastzhouqi=2011
end if
i=0
for b=lastzhouqi to zhouqi
  start_date=b&"-01-01"
  end_date=b&"-12-31"
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where left(d.fnumber,2)='12' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"年"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"年</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where left(部门代码,2)='08' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and left(d.fnumber,2)='12') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where left(c.fnumber,2)='12' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneYear" and Rclass="MN4" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>年度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=2015 then 
lastzhouqi=zhouqi-4
else
lastzhouqi=2011
end if
i=0
for b=lastzhouqi to zhouqi
  start_date=b&"-01-01"
  end_date=b&"-12-31"
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where d.fnumber='10.04' and datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"年"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"年</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where left(部门代码,2)='10' and 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid and d.fnumber='10.04') aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where c.fnumber='10.04' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

elseif Rtype="OneYear" and Rclass="MN" then
%>
 <div id="listtable" style="width:790px; height:480px; z-index:500">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>一、总体合格状况</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
 <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="4" align="center"><font color="#FFFFFF"><strong>首件</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>转序</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>巡检</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="3" align="center"><font color="#FFFFFF"><strong>成品</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>客诉</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2" align="center"><font color="#FFFFFF"><strong>补料(金额)</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>年度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提交合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格次数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>合格批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检验批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次合格率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客诉件数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损失金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工次补料</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>料次补料</strong></font></td>
  </tr>
 <%
ReDim Preserve tcategories(9)
ReDim Preserve tvalues(9)
ReDim Preserve tvalues2(9)
ReDim Preserve tvalues3(9)
ReDim Preserve tvalues4(9)
ReDim Preserve tvalues5(9)
if zhouqi>=2015 then 
lastzhouqi=zhouqi-4
else
lastzhouqi=2011
end if
i=0
for b=lastzhouqi to zhouqi
  start_date=b&"-01-01"
  end_date=b&"-12-31"
    sql="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2,isnull(sum(a3),0) as b3,isnull(sum(a4),0) as b4,isnull(sum(a5),0) as b5,isnull(sum(a6),0) as b6 "&_
	"from ( "&_
	"select case when b.FComboBox='首件' then b.finteger1 else 0 end as a1, "&_
	"case when b.FComboBox='首件' then b.Finteger else 0 end as a2, "&_
	"case when b.FComboBox='转序' then b.finteger1 else 0 end as a3, "&_
	"case when b.FComboBox='转序' then b.Finteger else 0 end as a4, "&_
	"case when b.FComboBox='成品' then b.finteger1 else 0 end as a5, "&_
	"case when b.FComboBox='成品' then b.Finteger else 0 end as a6,c.fname,b.FBase1 "&_
	"from t_t_productqualitycountEntry b inner join  "&_
	"t_productqualitycount a on a.FID=b.FID left join  "&_
	"t_workcenter c on c.fitemid=b.FBase1 left join "&_
	"t_item d on c.fdeptid=d.fitemid and d.fitemclassid=2 "&_
	"where datediff(d,b.fdate1,'"&start_date&"')<=0 and datediff(d,b.fdate1,'"&end_date&"')>=0) aaa "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    if (not rs.eof) then
	  tcategories(i)=b&"年"
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&b&"年</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b2")&"</td>"
	  a1=a1+rs("b1")
	  a2=a2+rs("b2")
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where FText='首件' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
	  if rs("b2") = 0 then
	  d1=0
	  d2=0
	  else
	  d2=rs("b1")*100/rs("b2")
	  d1=(rs("b2")-rs2("a1"))*100/rs("b2")
	  end if
	  a7=a7+rs2("a1")
	  tvalues(i)=formatnumber(d1,2)
	  tvalues2(i)=formatnumber(d2,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b3")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b4")&"</td>"
	  a3=a3+rs("b3")
	  a4=a4+rs("b4")
	  if rs("b4") = 0 then
	  d1=0
	  else
	  d1=rs("b3")*100/rs("b4")
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where FText='制程' and datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
	  rs2.open sql2,connk3,0,1
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&rs2("a1")&"</font></a></td>"
      Response.Write "<td nowrap>"&rs("b5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("b6")&"</td>"
	  a5=a5+rs("b5")
	  a6=a6+rs("b6")
	  a8=a8+rs2("a1")
	  if rs("b6") = 0 then
	  d1=0
	  else
	  d1=rs("b5")*100/rs("b6")
	  end if
	  tvalues3(i)=formatnumber(d1,2)
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      set rs2=server.createobject("adodb.recordset")
	  sql2="select count(1) as a1,isnull(sum(损失金额),0) as a2 from [Q-顾客抱怨调查处理报告单单头表] where 客诉日期>='"&start_date&"' and 客诉日期<='"&end_date&"'"
	  rs2.open sql2,conn,0,1
	  a9=a9+rs2("a1")
	  a10=a10+rs2("a2")
	  tvalues4(i)=rs2("a1")
	  tvalues5(i)=rs2("a2")
      Response.Write "<td nowrap>"&rs2("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs2("a2")&"</td>"
      set rs2=server.createobject("adodb.recordset")
	  sql2="select isnull(sum(a1),0) as b1,isnull(sum(a2),0) as b2 "&_
		"from ( "&_
		"select case when c.fname like '%工次%' then b.fauxqtysupply*e.forderprice else 0 end as a1, "&_
		"case when c.fname like '%料次%' then b.fauxqtysupply*e.forderprice else 0 end as a2 "&_
		"from icitemscrap a,icitemscrapentry b ,t_submessage c,t_department d,t_ICItem e "&_
		"where a.finterid=b.finterid and fcheckerid>0 and datediff(d,fdate,'"&start_date&"')<=0 and datediff(d,fdate ,'"&end_date&"')>=0 "&_
		"and b.fentryselfz0633=c.finterid and c.ftypeid=10007 and b.fentryselfz0626=d.fitemid  "&_
		"and b.fitemid=e.fitemid ) aaa "
	  rs2.open sql2,connk3,0,1
	  a11=a11+rs2("b1")
	  a12=a12+rs2("b2")
      Response.Write "<td nowrap>"&formatnumber(rs2("b1"),2)&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(rs2("b2"),2)&"</td>"
      Response.Write "</tr>" & vbCrLf
	  i=i+1
	  end if
next
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>合计</td>" & vbCrLf
      Response.Write "<td nowrap>"&a1&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a2&"</td>"
	  if a2 = 0 then
	  d1=0
	  d2=0
	  else
	  d2=a1*100/a2
	  d1=(a2-a7)*100/a2
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&formatnumber(d2,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a3&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a4&"</td>"
	  if a4 = 0 then
	  d1=0
	  else
	  d1=a3*100/a4
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap><a href=""javascript:ShowDetails('UQF','"&start_date&"','"&end_date&"')""><font color='#FF00FF'>"&a8&"</font></a></td>"
      Response.Write "<td nowrap>"&a5&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a6&"</td>"
	  if a6 = 0 then
	  d1=0
	  else
	  d1=a5*100/a6
	  end if
      Response.Write "<td nowrap>"&formatnumber(d1,2)&"%</td>" & vbCrLf
      Response.Write "<td nowrap>"&a9&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a10&"</td>"
      Response.Write "<td nowrap>"&a11&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&a12&"</td>"
      Response.Write "</tr>" & vbCrLf
  %>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
<div id="container" style="width: 800px; height: 400px; margin: 0 auto;"></div>
	</td>
  </tr>
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead()" >&nbsp;<strong>二、过程不合格现象及原因分析</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
  i=1
	sql="select count(1) as a1 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
		i=rs("a1")
	sql="select count(1) as a1,ftext4 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext4 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container1" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="80%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <table width="30%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="datatb2">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>不合格现象</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>批次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>比例</strong></font></td>
  </tr>
<%  
	sql="select count(1) as a1,ftext5 from t_unqualified a left join t_item c on c.fitemid=a.FBase1 where datediff(d,a.FDate1,'"&start_date&"')<=0 and datediff(d,a.FDate1,'"&end_date&"')>=0 group by ftext5 order by a1 desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while (not rs.eof)
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("a1")&"</td>" & vbCrLf
			Response.Write "<td nowrap>"&formatnumber((cdbl(rs("a1"))*100/i),2)&"%</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
	wend

%>
  </table>
	</td>
    <td height="36" align="left" nowrap  bgcolor="#EBF2F9">
  <div id="container2" style="width: 500px; height: 300px; margin: 0 auto;"></div>
	</td>
  </tr>
  </table>
	</td>
  </tr>
</table>
  </div>

<%  

end if
  rs.close
  set rs=nothing
%>
</body>
</html>
