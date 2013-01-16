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
if Instr(session("AdminPurviewFLW"),"|104,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
  dim replyflag,seachword
'  replyflag=request("replyflag")
  seachword=request("seachword")
'  if replyflag="" then
  replyflag=""
'  else
'  replyflag=" and replyFlag="&replyflag
'  end if
  if seachword="" then
  seachword=""
  else
  seachword=" and (a.fbillno like '%"&seachword&"%' or d.fname like '%"&seachword&"%')"
  end if
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工厂订单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品型号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>业务员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>实际交期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>出货样</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>提供确认</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>品保确认</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生管接受</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>业务接受</strong></font></td>
  </tr>
 <%
  dim page'页码
      page=clng(request("Page"))
  dim datafrom'数据表名
      datafrom=" seorder a,seorderentry b,t_emp c,t_ICItem d "
  dim datawhere'数据条件
		 datawhere=" where b.fentryselfs0165>0 and a.finterid=b.finterid and b.fstockqty=0 "&_
		"and a.fcheckerid>0 and a.fcancellation=0 and b.fqty>0 and b.FEntrySelfS0182 <> 1 "&_
		"and c.fitemid=a.fempid and d.fitemid=b.fitemid and a.fdate >'2011-01-01' "&seachword
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by b.finterid,b.fentryid asc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
'-----------------------------------------------------------
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    dim sql2,rs2
	dim formdata(3),bgcolors
    sql="select a.fbillno,d.fname as 产品型号,c.fname as 业务员,b.fdate,b.fentryselfs0165,b.finterid,b.fentryid from "& datafrom &" "& datawhere &" "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = 20 '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
  	  bgcolors="#EBF2F9"
	  sql2="select SampleNum as a1,QCReplyer as a2,MCReplyer as a3,SEReplyer as a4,replyFlag from Flw_SamplesHandover where FInterID="&rs("finterid")&" and FEntryID="&rs("FEntryID")&replyflag
	  set rs2=server.createobject("adodb.recordset")
	  rs2.open sql2,connzxpt,0,1
	  if rs2.eof and rs2.bof then
	    formdata(0)=""
	    formdata(1)=""
	    formdata(2)=""
	    formdata(3)=""
	  else
	    formdata(0)=rs2("a1")
	    formdata(1)=rs2("a2")
	    formdata(2)=rs2("a3")
	    formdata(3)=rs2("a4")
		if rs2("replyFlag")=1 then
		  bgcolors="#ffff66"'黄色
		elseif rs2("replyFlag")=2 then
		  bgcolors="#ff99ff"'粉色
		elseif rs2("replyFlag")=3 then
		  bgcolors="#66ff66"'绿色
		elseif rs2("replyFlag")=4 then
		  bgcolors="#6666ff"'紫色
		else
		  bgcolors="#EBF2F9"
		end if
	  end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fbillno")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("产品型号")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("业务员")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fdate")&"</td>"
      Response.Write "<td nowrap>"&rs("fentryselfs0165")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return SHDClickTd(this,'SHDSPLreply',"&rs("finterid")&","&rs("FEntryID")&")"">"&formdata(0)&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return SHDClickTd(this,'SHDQCreply',"&rs("finterid")&","&rs("FEntryID")&")"">"&formdata(1)&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return SHDClickTd(this,'SHDMCreply',"&rs("finterid")&","&rs("FEntryID")&")"">"&formdata(2)&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return SHDClickTd(this,'SHDSEreply',"&rs("finterid")&","&rs("FEntryID")&")"">"&formdata(3)&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    next
	  rs2.close
	  set rs2=nothing
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>

</body>
</html>
