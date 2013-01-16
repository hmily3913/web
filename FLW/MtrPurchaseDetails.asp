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
<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<BODY>
<%
if Instr(session("AdminPurviewFLW"),"|102,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>采购单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>供应商</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品编号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品名称</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>规格型号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>采购数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>入库数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>采购日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>交货日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生管回复</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物控回复</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>仓库回复</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>业务回复</strong></font></td>
  </tr>
 <%
  dim page'页码
      page=clng(request("Page"))
  dim datafrom'数据表名
      datafrom=" vwICBill_26 "
  dim datawhere'数据条件
		 datawhere="where fcheckflag='※' and FStatusEx='' and FCommitdate<='"&date()&"' "
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by FCommitdate"
  dim i'用于循环的整数
  dim rs,sql'sql语句
'-----------------------------------------------------------
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    dim sql2,rs2
	dim formdata(3),bgcolors
    sql="select * from "& datafrom &" "& datawhere &" "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = 20 '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
  	  bgcolors="#EBF2F9"
	  sql2="select left(PCReplyText,10) as a1,PCReplyFlag,left(MCReplyText,10) as a2,MCReplyFlag,left(STReplyText,10) as a3,STReplyFlag,left(SEReplyText,10) as a4,SEReplyFlag from Flw_vwICBill_26 where FInterID="&rs("finterid")&" and FEntryID="&rs("FEntryID")
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
		if rs2("PCReplyFlag")>0 then
		  bgcolors="#ffff66"'黄色
		end if
		if rs2("MCReplyFlag")>0 then
		  bgcolors="#ff99ff"'粉色
		end if
		if rs2("STReplyFlag")>0 then
		  bgcolors="#66ff66"'绿色
		end if
		if rs2("SEReplyFlag")>0 then
		  bgcolors="#6666ff"'紫色
		end if
	  end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fbillno")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fentryselfp0250")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FSupplyID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fnumber")&"</td>"
      Response.Write "<td nowrap>"&rs("fitemid")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fmodel")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fqty")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FStockQty")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fdate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FCommitdate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return MtrClickTd(this,'MtrPCreply',"&rs("finterid")&","&rs("FEntryID")&")"">"&formdata(0)&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return MtrClickTd(this,'MtrMCreply',"&rs("finterid")&","&rs("FEntryID")&")"">"&formdata(1)&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return MtrClickTd(this,'MtrSTreply',"&rs("finterid")&","&rs("FEntryID")&")"">"&formdata(2)&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return MtrClickTd(this,'MtrSEreply',"&rs("finterid")&","&rs("FEntryID")&")"">"&formdata(3)&"</td>" & vbCrLf
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
