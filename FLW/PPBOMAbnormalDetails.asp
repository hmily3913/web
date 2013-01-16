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

if Instr(session("AdminPurviewFLW"),"|106,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag
showType=request("showType")
print_tag=request("print_tag")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="DetailsList" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>投料单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生产车间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品代码</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品名称</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生产任务单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>销售订单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品入库数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>行号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物料代码</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物料名称</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单位</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>计划投料数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单位用量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>应发数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>已领数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>损耗数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>补料数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>计划跟踪号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>最后领料日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>最后入库日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>状态</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>结束</strong></font></td>
  </tr>
 <%
  dim Depart,wherestr
  wherestr=""
  Depart=request("s6partment")
  start_date=request("start_date")
  end_date=request("end_date")
  if start_date<>""  then wherestr=wherestr&" and isnull(ICStockBill2.fdate,'') >='"&start_date&"'"
  if end_date<>""  then	wherestr=wherestr&" and isnull(ICStockBill2.fdate,'')<='"&end_date&"'"

		  
  dim page'页码
      page=clng(request("Page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" dbo.PPBOMEntry AS PPBOMEntry INNER JOIN "&_
               "       dbo.PPBOM AS PPBOM ON PPBOMEntry.FInterID = PPBOM.FInterID INNER JOIN "&_
               "       dbo.ICMO AS ICMO ON PPBOM.FICMOInterID = ICMO.FInterID AND ICMO.FTranType = 85 LEFT OUTER JOIN "&_
               "       dbo.t_ICItemCore AS t_ICItem_1 ON PPBOMEntry.FItemID = t_ICItem_1.FItemID LEFT OUTER JOIN "&_
               "       dbo.t_Base_User AS User1 ON PPBOM.FBillerID = User1.FUserID LEFT OUTER JOIN "&_
               "       dbo.t_Base_User AS User2 ON PPBOM.FCheckerID = User2.FUserID LEFT OUTER JOIN "&_
               "       dbo.t_MeasureUnit AS MeasureUnit ON PPBOMEntry.FUnitID = MeasureUnit.FMeasureUnitID LEFT OUTER JOIN "&_
               "       dbo.t_ICItemCore AS t_ICItem_2 ON PPBOM.FItemID = t_ICItem_2.FItemID LEFT OUTER JOIN "&_
               "       dbo.t_Department AS Department ON PPBOM.FWorkSHop = Department.FItemID LEFT OUTER JOIN "&_
               "           (SELECT     dbo.ICStockBillEntry.FSourceEntryID, dbo.ICStockBillEntry.FSourceInterId, MAX(ICStockBill_1.FDate) AS fdate "&_
               "             FROM          dbo.ICStockBillEntry INNER JOIN "&_
               "                                    dbo.ICStockBill AS ICStockBill_1 ON ICStockBill_1.FInterID = dbo.ICStockBillEntry.FInterID "&_
               "             WHERE      (ICStockBill_1.FTranType = 24) AND (dbo.ICStockBillEntry.FSourceTranType = 85) "&_
               "             GROUP BY dbo.ICStockBillEntry.FSourceEntryID, dbo.ICStockBillEntry.FSourceInterId) AS ICStockBill ON  "&_
               "       ICStockBill.FSourceEntryID = PPBOMEntry.FEntryID AND ICStockBill.FSourceInterId = ICMO.FInterID LEFT OUTER JOIN "&_
               "           (SELECT     dbo.ICStockBillEntry.FSourceEntryID, dbo.ICStockBillEntry.FSourceInterId, MAX(ICStockBill.FDate) AS fdate "&_
               "             FROM          dbo.ICStockBillEntry INNER JOIN "&_
               "                                    dbo.ICStockBill ON ICStockBill.FInterID = dbo.ICStockBillEntry.FInterID "&_
               "             WHERE      (ICStockBill.FTranType = 2) AND (dbo.ICStockBillEntry.FSourceTranType = 85) "&_
               "             GROUP BY dbo.ICStockBillEntry.FSourceEntryID, dbo.ICStockBillEntry.FSourceInterId) AS ICStockBill2 ON  "&_
               "        ICStockBill2.FSourceInterId = ICMO.FInterID "
  dim datawhere'数据条件
		 datawhere="where ICMO.FAuxStockQty >=ICMO.FAuxQty and PPBOMEntry.FAuxStockQty<PPBOMEntry.FAuxQtyMust "+wherestr
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" ORDER BY 投料单号 desc, 行号"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  idCount=rs("idCount")
  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
  end if
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="SELECT     TOP (100) PERCENT PPBOM.FInterID as id,PPBOMEntry.FEntrySelfY0266 as 结束,PPBOM.FBillNo AS 投料单号, Department.FName AS 生产车间, t_ICItem_2.FNumber AS 产品代码,  "&_
        "              t_ICItem_2.FName AS 产品名称, ICMO.FBillNo AS 生产任务单号, PPBOM.FOrderBillNo AS 销售订单号, PPBOM.FAuxQty AS 产品数量, ICMO.FAuxStockQty AS 产品入库数量,  "&_
        "              PPBOMEntry.FEntryID AS 行号, t_ICItem_1.FNumber AS 物料代码, t_ICItem_1.FName AS 物料名称, MeasureUnit.FName AS 单位,  "&_
        "              PPBOMEntry.FAuxQtyMust AS 计划投料数量, PPBOMEntry.FAuxQtyScrap AS 单位用量, PPBOMEntry.FAuxQtyPick AS 应发数量,  "&_
        "              PPBOMEntry.FAuxStockQty AS 已领数量, PPBOMEntry.FAuxQtyLoss AS 损耗数量, PPBOMEntry.FAuxQtySupply AS 补料数量,  PPBOMEntry.FMTONo AS 计划跟踪号,  "&_
        "              ICStockBill.fdate AS 最后领料日期,ICStockBill2.fdate AS 产品最后入库日期,case when ICMO.FStatus=0 then '计划' when ICMO.FStatus=1 then '下达' when ICMO.FStatus=2 then '完全入库' when ICMO.FStatus=3 then '结案' when ICMO.FStatus=5 then '确认' else '作废' end as 任务单状态 "&_
		" from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = 20 '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
  '获取本页需要用到的id结束============================================
'-----------------------------------------------------------
'-----------------------------------------------------------
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
	dim bgcolors
  	  bgcolors="#EBF2F9"
	  if rs("结束")="1" then bgcolors="#ff99ff"
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("投料单号")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("生产车间")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("产品代码")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("产品名称")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("生产任务单号")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("销售订单号")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("产品数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("产品入库数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("行号")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("物料代码")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("物料名称")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("单位")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("计划投料数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("单位用量")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("应发数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("已领数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("损耗数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("补料数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("计划跟踪号")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("最后领料日期")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("产品最后入库日期")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("任务单状态")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("id")&","&rs("行号")&")"">"&rs("结束")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    next
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
	response.Write("###"&pagec&"###"&idcount&"###")
elseif showType="AddEditShow" then 
  dim FInterId:FInterId=request("FInterId")
  dim FEntryId:FEntryId=request("FEntryId")
    sql="SELECT PPBOM.FInterID as id,PPBOM.FBillNo AS 投料单号, Department.FName AS 生产车间, t_ICItem_2.FNumber AS 产品代码,  "&_
        "              t_ICItem_2.FName AS 产品名称, ICMO.FBillNo AS 生产任务单号, PPBOM.FOrderBillNo AS 销售订单号, PPBOM.FAuxQty AS 产品数量, ICMO.FAuxStockQty AS 产品入库数量,  "&_
        "              PPBOMEntry.FEntryID AS 行号, t_ICItem_1.FNumber AS 物料代码, t_ICItem_1.FName AS 物料名称, MeasureUnit.FName AS 单位,  "&_
        "              PPBOMEntry.FAuxQtyMust AS 计划投料数量, PPBOMEntry.FAuxQtyScrap AS 单位用量, PPBOMEntry.FAuxQtyPick AS 应发数量,  "&_
        "              PPBOMEntry.FAuxStockQty AS 已领数量, PPBOMEntry.FAuxQtyLoss AS 损耗数量, PPBOMEntry.FAuxQtySupply AS 补料数量,  PPBOMEntry.FMTONo AS 计划跟踪号,  "&_
        "              case when ICMO.FStatus=0 then '计划' when ICMO.FStatus=1 then '下达' when ICMO.FStatus=2 then '完全入库' when ICMO.FStatus=3 then '结案' when ICMO.FStatus=5 then '确认' else '作废' end as 任务单状态, "&_
		"			   PPBOMEntry.FEntrySelfY0266,PPBOMEntry.FEntrySelfY0267,PPBOMEntry.FEntrySelfY0268,PPBOMEntry.FEntrySelfY0269 "&_
       " from  dbo.PPBOMEntry AS PPBOMEntry INNER JOIN "&_
	   "       dbo.PPBOM AS PPBOM ON PPBOMEntry.FInterID = PPBOM.FInterID INNER JOIN "&_
	   "       dbo.ICMO AS ICMO ON PPBOM.FICMOInterID = ICMO.FInterID AND ICMO.FTranType = 85 LEFT OUTER JOIN "&_
	   "       dbo.t_ICItemCore AS t_ICItem_1 ON PPBOMEntry.FItemID = t_ICItem_1.FItemID LEFT OUTER JOIN "&_
	   "       dbo.t_Base_User AS User1 ON PPBOM.FBillerID = User1.FUserID LEFT OUTER JOIN "&_
	   "       dbo.t_Base_User AS User2 ON PPBOM.FCheckerID = User2.FUserID LEFT OUTER JOIN "&_
	   "       dbo.t_MeasureUnit AS MeasureUnit ON PPBOMEntry.FUnitID = MeasureUnit.FMeasureUnitID LEFT OUTER JOIN "&_
	   "       dbo.t_ICItemCore AS t_ICItem_2 ON PPBOM.FItemID = t_ICItem_2.FItemID LEFT OUTER JOIN "&_
	   "       dbo.t_Department AS Department ON PPBOM.FWorkSHop = Department.FItemID "&_
	   " where PPBOM.FInterID="&FInterId&" and PPBOMEntry.FEntryID="&FEntryId
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
%>
 <div id="AddandEditdiv" style="width:790px; height:460px; overflow:auto; ">
 <form id="AddForm" name="AddForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>已入库未领料管理</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews width="100%">
      <tr>
        <td height="20" align="left">投料单号：</td>
        <td>
		<%= rs("投料单号") %></td>
        <td width="120" height="20" align="left">生产车间：</td>
        <td><%= rs("生产车间") %></td>
        <td width="120" height="20" align="left">产品代码：</td>
        <td><%= rs("产品代码") %></td>
      </tr>
      <tr>
        <td height="20" align="left">产品名称：</td>
        <td>
		<%= rs("产品名称") %></td>
        <td width="120" height="20" align="left">生产任务单号：</td>
        <td><%= rs("生产任务单号") %></td>
        <td width="120" height="20" align="left">销售订单号：</td>
        <td><%= rs("销售订单号") %></td>
      </tr>
      <tr>
        <td height="20" align="left">产品数量：</td>
        <td>
		<%= rs("产品数量") %></td>
        <td width="120" height="20" align="left">产品入库数量：</td>
        <td><%= rs("产品入库数量") %></td>
        <td width="120" height="20" align="left">物料代码：</td>
        <td><%= rs("物料代码") %></td>
      </tr>
      <tr>
        <td height="20" align="left">物料名称：</td>
        <td>
		<%= rs("物料名称") %></td>
        <td width="120" height="20" align="left">单位：</td>
        <td><%= rs("单位") %></td>
        <td width="120" height="20" align="left">计划投料数量：</td>
        <td><%= rs("计划投料数量") %></td>
      </tr>
      <tr>
        <td height="20" align="left">单位用量：</td>
        <td>
		<%= rs("单位用量") %></td>
        <td width="120" height="20" align="left">应发数量：</td>
        <td><%= rs("应发数量") %></td>
        <td width="120" height="20" align="left">已领数量：</td>
        <td><%= rs("已领数量") %></td>
      </tr>
      <tr>
        <td height="20" align="left">损耗数量：</td>
        <td>
		<%= rs("损耗数量") %></td>
        <td width="120" height="20" align="left">补料数量：</td>
        <td><%= rs("补料数量") %></td>
        <td width="120" height="20" align="left">任务单状态：</td>
        <td><%= rs("任务单状态") %></td>
      </tr>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="1" colspan="6">  </td></tr>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td width="60"> 结束人： </td>
 <td width="60">
 <%= rs("FEntrySelfY0267") %></td>
 <td width="60"> 结束日期： </td>
 <td width="60">
 <%= rs("FEntrySelfY0268") %></td>
 <td width="60"> 结束标志： </td>
 <td width="60">
 <select name="FEntrySelfY0266" id="FEntrySelfY0266">
 <option value="1" <% If rs("FEntrySelfY0266")="1" Then Response.Write("selected")%>>结束</option>
 <option value="0">未结束</option>
 </select>
 </td>
</tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 结束说明： </td>
<td colspan="5">
  <textarea name="FEntrySelfY0269" id="FEntrySelfY0269" style="width:500px; height:100px; "><%= rs("FEntrySelfY0269") %></textarea>
</td>
</tr> 
  <tr>  <td height="1" colspan="6">  </td> </tr>
	<tr>
	  <td align="center" colspan="6">
	  <input type="hidden" name="FInterId" id="FInterId" value="<%= rs("id") %>">
	  <input type="hidden" name="FEntryId" id="FEntryId" value="<%= rs("行号") %>">
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="结束" style="WIDTH: 80;"  onClick="toSubmit();">&nbsp;
	  </td>
	</tr>
  <tr>  <td height="1">  </td></tr>
  </table>
</td>
  </tr>
</table>
</form>
</div>
<%
  rs.close
  set rs=nothing
elseif showType="CheckProcess" and Instr(session("AdminPurviewFLW"),"|106.1,")>0 then 
  FInterId=request("FInterId")
  FEntryId=request("FEntryId")
	set rs = server.createobject("adodb.recordset")
	sql="select * from PPBOMEntry where FInterId="&FInterId&" and FEntryId="&FEntryId
	rs.open sql,connk3,1,3	
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs("FEntrySelfY0266")=request("FEntrySelfY0266")
	rs("FEntrySelfY0267")=session("AdminName")
	rs("FEntrySelfY0268")=now()
	rs("FEntrySelfY0269")=request("FEntrySelfY0269")
	rs.update
  rs.close
  set rs=nothing
  response.Write("###")
end if
 %>
</body>
</html>
