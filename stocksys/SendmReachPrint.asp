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
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="../Script/Rpt.js"></script>
<link rel="stylesheet" href="../Images/jquery.datepick.css">
<script language="javascript" src="../Script/jquery.datepick.pack.js"></script>
<script language="javascript" src="../Script/jquery.datepick-zh-CN.js"></script>
<script language="javascript">
$(function(){
$('#start_date').datepick({dateFormat: 'yyyy-mm-dd'});
$('#end_date').datepick({dateFormat: 'yyyy-mm-dd'});
});
</script>
<script language="javascript">

var xPos; var yPos; 
$(document).bind('mousemove',function(e){ 
            xPos= e.pageX ;
			yPos= e.pageY; 
});
</script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<%

response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
'========判断是否具有管理权限
%>
<BODY>
<%
dim Result,StartDate,EndDate,Keyword,Reachsum,unReachsum,Reachper,sqlstr
Result=request("Result")
StartDate=request("start_date")
if StartDate="" then StartDate=date()
EndDate=request("end_date")
if EndDate="" then EndDate=date()
Keyword=request("Keyword")
sqlstr="SELECT     TOP (100) PERCENT v1.FDate AS 日期, Department.FName AS 领料部门, v1.FBillNo AS 单据单号, u1.FEntryID AS 行号, u1.FBatchNo AS 批号,  "&_
"                      ICItem.FNumber AS 物料代码, ICItem.FName AS 物料名称, Unit1.FName AS 单位, u1.FAuxQtyMust AS 应发数量, u1.FAuxQty AS 实发数量, "&_
"                      Stock.FName AS 仓库, u1.FMTONo AS 计划跟踪号, u1.FEntrySelfB0449 AS 工厂订单号, u1.FEntrySelfB0450 AS 型号, u1.FEntrySelfB0447 AS 仓库员, "&_
"                      v1.FCheckDate AS 审核日期, u1.FSourceTranType AS 源单类型, u1.FSourceBillNo AS 源单编号, u1.FNote AS 备注, USER1.FName AS 制单, "&_
"                      USER2.FName AS 审核, u1.FEntrySelfB0451 AS 制单库存数量, (LTRIM(RTRIM(STR(u1.FInterID))) +'t'+ LTRIM(RTRIM(STR(u1.FEntryID)))) AS id, "&_
"					  datediff(day,v1.FDate,v1.FCheckDate) as diffdays "&_
"FROM         dbo.ICStockBillEntry AS u1 INNER JOIN "&_
"                      dbo.ICStockBill AS v1 ON u1.FInterID = v1.FInterID AND v1.FTranType = 24 LEFT OUTER JOIN "&_
"                      dbo.t_ICItemCore AS ICItem ON u1.FItemID = ICItem.FItemID LEFT OUTER JOIN "&_
"                      dbo.t_MeasureUnit AS Unit1 ON u1.FUnitID = Unit1.FMeasureUnitID LEFT OUTER JOIN "&_
"                      dbo.t_Base_User AS USER1 ON v1.FBillerID = USER1.FUserID LEFT OUTER JOIN "&_
"                      dbo.t_Base_User AS USER2 ON v1.FCheckerID = USER2.FUserID LEFT OUTER JOIN "&_
"                      dbo.t_Stock AS Stock ON u1.FDCStockID = Stock.FItemID LEFT OUTER JOIN "&_
"                      dbo.t_Department AS Department ON v1.FDeptID = Department.FItemID "&_
"where v1.FDate>='"&StartDate&"' and v1.FDate<='"&EndDate&"' and fstatus>0 "&_
" union "&_
"SELECT     TOP (100) PERCENT v1.FDate AS 日期, Department.FName AS 领料部门, v1.FBillNo AS 单据单号, u1.FEntryID AS 行号, u1.FBatchNo AS 批号,  "&_
"                      ICItem.FNumber AS 物料代码, ICItem.FName AS 物料名称, Unit1.FName AS 单位, u1.FAuxQtyMust AS 应发数量, u1.FAuxQty AS 实发数量, "&_
"                      Stock.FName AS 仓库, u1.FMTONo AS 计划跟踪号, u1.FEntrySelfB0449 AS 工厂订单号, u1.FEntrySelfB0450 AS 型号, u1.FEntrySelfB0447 AS 仓库员, "&_
"                      '' AS 审核日期, u1.FSourceTranType AS 源单类型, u1.FSourceBillNo AS 源单编号, u1.FNote AS 备注, USER1.FName AS 制单, "&_
"                      USER2.FName AS 审核, u1.FEntrySelfB0451 AS 制单库存数量, (LTRIM(RTRIM(STR(u1.FInterID))) +'t'+ LTRIM(RTRIM(STR(u1.FEntryID)))) AS id, "&_
"					  datediff(day,v1.FDate,getdate()) as diffdays "&_
"FROM         dbo.ICStockBillEntry AS u1 INNER JOIN "&_
"                      dbo.ICStockBill AS v1 ON u1.FInterID = v1.FInterID AND v1.FTranType = 24 LEFT OUTER JOIN "&_
"                      dbo.t_ICItemCore AS ICItem ON u1.FItemID = ICItem.FItemID LEFT OUTER JOIN "&_
"                      dbo.t_MeasureUnit AS Unit1 ON u1.FUnitID = Unit1.FMeasureUnitID LEFT OUTER JOIN "&_
"                      dbo.t_Base_User AS USER1 ON v1.FBillerID = USER1.FUserID LEFT OUTER JOIN "&_
"                      dbo.t_Base_User AS USER2 ON v1.FCheckerID = USER2.FUserID LEFT OUTER JOIN "&_
"                      dbo.t_Stock AS Stock ON u1.FDCStockID = Stock.FItemID LEFT OUTER JOIN "&_
"                      dbo.t_Department AS Department ON v1.FDeptID = Department.FItemID "&_
"where v1.FDate>='"&StartDate&"' and v1.FDate<='"&EndDate&"' and fstatus=0 "&_
"ORDER BY 单据单号, 行号 "


'response.Write(sqlstr)
function PlaceFlag()
  dim rs,sql,sqlstr2'sql语句
  if Result="Search" then
	sqlstr2="select sum(SendmROne) as idCount1,sum(unSendmROne) as idCount2 from stocksys where  UPTdate>='"&StartDate&"' and UPTdate<='"&EndDate&"'"
    set rs=server.createobject("adodb.recordset")
    rs.open sqlstr2,connzxpt,0,1
    Reachsum=rs("idCount1")
    unReachsum=rs("idCount2")
    if unReachsum=0 then
	Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，总收发料数为0"
	else
    Reachper=Reachsum/unReachsum*100
    Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，总发料笔数[<font color='red'>"&unReachsum&"</font>]，及时发料次数[<font color='red'>"&Reachsum&"</font>]，发料及时率为：[<font color='red'>"&formatnumber(Reachper,2)&"%</font>]"
	end if
  else
	sqlstr2="select top 1 SendmROne as idCount1,unSendmROne as idCount2,SendmReachOne from stocksys order by SerialNum desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sqlstr2,connzxpt,0,1
    Reachsum=rs("idCount1")
    unReachsum=rs("idCount2")
    Reachper=rs("SendmReachOne")
    Response.Write "统计时间：[<font color='red'>昨日</font>]，总发料笔数[<font color='red'>"&unReachsum&"</font>]，及时发料次数[<font color='red'>"&Reachsum&"</font>]，发料及时率为：[<font color='red'>"&formatnumber(Reachper,2)&"%</font>]"
  end if
  rs.close
  set rs=nothing
end function  
 
%>

  <% ProductsList() %>

</BODY>
</HTML>
<%
'-----------------------------------------------------------
function ProductsList()
 %>
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>制单日期</strong></font></td>
    <td width="80" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>领料部门</strong></font></td>
    <td width="80" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单据单号</strong></font></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品编号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品名称</strong></font></td>
    <td width="30" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单位</strong></font></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">应发数量</font></strong></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">实发数量</font></strong></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">发料日期</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">工作回复</font></strong></td>
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
      datafrom="stocksys_SendmReach"
  dim datawhere'数据条件
      if Keyword="list" then
	     datawhere="where  UPTdate>='"&StartDate&"' and UPTdate<='"&EndDate&"'"
	  else
		 datawhere="where datediff(d,UPTdate,getdate())=1"
 	  end if
  dim sqlid'本页需要用到的id
  dim Myself,PATH_INFO,QUERY_STRING'本页地址和参数
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" or Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself = PATH_INFO & "?"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis'排序的语句 asc,desc
      taxis=" order by SerialNum"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(id) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
  idCount=rs("idCount")
  '获取记录总数
  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select SerialNum from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    rs.pagesize = pages '每页显示记录数
    if page < 1 then page = 1
    if page > pagec then page = pagec
    if pagec > 0 then rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("SerialNum")
	  else
	    sqlid=sqlid &","&rs("SerialNum")
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  if(idcount>0 and sqlid<>"") then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    sql="select *,left(ReplyText,10) as a20 from "& datafrom &" " &datawhere&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
	dim bgcolors
    while(not rs.eof)'填充数据到表格
	  bgcolors="#EBF2F9"
		if rs("replyFlag")>0 then
		  bgcolors="#ffff66"'
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("日期")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("领料部门")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("单据单号")&"</td>"
      Response.Write "<td nowrap>"&rs("物料代码")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("物料名称")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("单位")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("应发数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("实发数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("审核日期")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ReplyText")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
  else
    response.write "<tr><td height='50' align='center' colspan='10' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  <%
end function 

%>


