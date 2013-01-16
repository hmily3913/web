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
<!--#include file="../CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|301,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<%
dim Result,StartDate,EndDate,Keyword,Reachsum,unReachsum,Reachper,sqlstr,department,Depart,anywhere,urlstr
Result=request("Result")
department=request("department")
StartDate=request("start_date")
EndDate=request("end_date")
Keyword=request("Keyword")
Depart=session("Depart")
urlstr=""
if Instr(request.ServerVariables("QUERY_STRING"),department) = 0 then
urlstr="end_date="&EndDate&"&start_date="&StartDate&"&department="&department&"&"
end if
'response.Write(QUERY_STRING&"$$$"&urlstr)
sqlstr="SELECT     TOP (100) PERCENT SEOrder.FBillNo AS 销售订单号, ICMO.FBillNo AS 生产任务单编号, ICItem.FNumber AS 产品代码, ICItem.FName AS 产品名称, "&_
"                      Item1.FName AS 产品类别1, Item2.FName AS 产品类别2, Department.FName AS 生产车间, Unit.FName AS 单位, ICMO.FQty AS 生产数量, "&_
"                      ICMO.FCommitQty AS 完工数量, SUM(ICStockBillEntry.FQty) AS 入库数量, SEOrderEntry.FAdviceConsignDate AS 生管回复交期, MAX(ICStockBill.FDate) "&_
"                      AS 最后入库日期, ICMO.FCloseDate AS 结案日期, ICMO.FCheckDate AS 制单日期, User3.FName AS 制单人, "&_
"                      ICMO.FNote AS 备注, ICMO.FMTONo AS 计划跟踪号, ICMO.FHandworkClose AS 手工结案, SEOrderEntry.FEntrySelfS0174 AS 颜色, "&_
"                      emp.FName AS 业务员, SEOrderEntry.FEntrySelfS0165 AS 出货样, "&_
"                      CASE WHEN (SEOrderEntry.FDate >= MAX(ICStockBill.FDate) AND "&_
"                      SEOrderEntry.FDate <= getdate() AND SUM(ICStockBillEntry.FQty) >= ICMO.FQty) OR"&_
"                      (SUM(ICStockBillEntry.FQty) IS NULL AND SEOrderEntry.FDate >= ICMO.FCloseDate) THEN '按期完成' ELSE '超交期' END AS 是否按期, "&_
"                      SEOrderEntry.FDate,ICMO.FInterID as id,case when Item3.FParentID=0 then Item3.Fname else Item4.Fname end as 分厂"&_
"                      "&_
"                      "&_
"FROM         dbo.ICMO AS ICMO LEFT OUTER JOIN"&_
"                      dbo.t_Department AS Department ON ICMO.FWorkShop = Department.FItemID LEFT OUTER JOIN"&_
"                      dbo.SEOrder AS SEOrder ON ICMO.FOrderInterID = SEOrder.FInterID LEFT OUTER JOIN"&_
"                      dbo.t_MeasureUnit AS Unit ON ICMO.FUnitID = Unit.FMeasureUnitID LEFT OUTER JOIN"&_
"                      dbo.t_User AS User3 ON ICMO.FBillerID = User3.FUserID LEFT OUTER JOIN"&_
"                      dbo.SEOrderEntry AS SEOrderEntry ON ICMO.FOrderInterID = SEOrderEntry.FInterID AND "&_
"                      ICMO.FSourceEntryID = SEOrderEntry.FEntryID  LEFT OUTER JOIN"&_
"                      dbo.t_ICItem AS ICItem ON ICItem.FItemID = ICMO.FItemID LEFT OUTER JOIN"&_
"                      dbo.ICStockBill AS ICStockBill INNER JOIN"&_
"                      dbo.ICStockBillEntry AS ICStockBillEntry ON ICStockBill.FInterID = ICStockBillEntry.FInterID ON ICMO.FInterID = ICStockBillEntry.FSourceInterId AND "&_
"                      ICStockBill.FTranType = 2 LEFT OUTER JOIN"&_
"                      dbo.t_Emp AS emp ON SEOrder.FEmpID = emp.FItemID LEFT OUTER JOIN"&_
"                      dbo.t_Item AS Item1 ON LEFT(ICItem.FNumber, 4) = Item1.FNumber AND Item1.FItemClassID = 4 LEFT OUTER JOIN"&_
"                      dbo.t_Item AS Item2 ON ICItem.FParentID = Item2.FItemID left join"&_
"                      dbo.t_Item AS Item3 ON ICMO.fworkshop = Item3.FItemID  AND Item3.FItemClassID = 2 left join"&_
"                      dbo.t_Item AS Item4 ON Item3.FParentID = Item4.FItemID "&_
"WHERE     (ICMO.FTranType = 85 and SEOrderEntry.FDate <= getdate()"&_
"                      and SEOrderEntry.FDate >= '"&StartDate&"' and SEOrderEntry.FDate <= '"&EndDate&"' "&_
"					"&_
")"&_
"GROUP BY SEOrder.FBillNo, ICMO.FMTONo, ICMO.FBillNo, ICItem.FNumber, ICItem.FName, Department.FName, ICMO.FQty, ICMO.FCommitQty, Unit.FName, "&_
"                      SEOrderEntry.FAdviceConsignDate, ICMO.FCloseDate, ICMO.FConveyerID, "&_
"                      ICMO.FCheckerID, ICMO.FCheckDate, ICMO.FBillerID, User3.FName, ICMO.FNote, ICMO.FUnitID, ICMO.FOrderInterID, "&_
"                      ICMO.FHandworkClose, emp.FName,  SEOrderEntry.FEntrySelfS0174, "&_
"                      SEOrderEntry.FEntrySelfS0165,  Item1.FName, Item2.FName,SEOrderEntry.FDate,ICMO.FInterID, Item3.Fname , Item4.Fname,Item3.FParentID "&_
"ORDER BY 生产任务单编号"



function PlaceFlag()
  dim rs,sql,sqlstr2'sql语句
  dim a1,a2,a3,a4,b1,b2,b3,b4,c1,c2,c3,c4
  
  if Result="Search" then
	sqlstr2="select sum(unDeliveryR1One) as a1,sum(DeliveryR1One) as b1,sum(unDeliveryR2One) as a2,sum(DeliveryR2One) as b2,sum(unDeliveryR3One) as a3,sum(DeliveryR3One) as b3,sum(unDeliveryR4One) as a4,sum(DeliveryR4One) as b4 from manusys where  UPTdate>='"&StartDate&"' and UPTdate<='"&EndDate&"'"
    set rs=server.createobject("adodb.recordset")
    rs.open sqlstr2,connzxpt,0,1
    a1=rs("a1")
    b1=rs("b1")
    a2=rs("a2")
    b2=rs("b2")
    a3=rs("a3")
    b3=rs("b3")
    a4=rs("a4")
    b4=rs("b4")
	if department="10" then
	Reachper="一分厂"
    if b1=0 then
	Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，一分厂需完成笔数0"
	else
    c1=a1/b1*100
    Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，一分厂需完成笔数[<font color='red'>"&b1&"</font>]，交期达成笔数[<font color='red'>"&a1&"</font>]，一分厂交期达成率为：[<font color='red'>"&formatnumber(c1,2)&"%</font>]"
	end if
	end if
	if department="11" then
	Reachper="二分厂"
    if b2=0 then
	Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，二分厂需完成笔数0"
	else
    c2=a2/b2*100
    Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，二分厂需完成笔数[<font color='red'>"&b2&"</font>]，交期达成笔数[<font color='red'>"&a2&"</font>]，二分厂交期达成率为：[<font color='red'>"&formatnumber(c2,2)&"%</font>]"
	end if
	end if
	if department="12" then
	Reachper="三分厂"
    if b3=0 then
	Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，三分厂需完成笔数0"
	else
    c3=a3/b3*100
    Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，三分厂需完成笔数[<font color='red'>"&b3&"</font>]，交期达成笔数[<font color='red'>"&a3&"</font>]，三分厂交期达成率为：[<font color='red'>"&formatnumber(c3,2)&"%</font>]"
	end if
	end if
	if department="19" then
	Reachper="娄桥分厂"
    if b4=0 then
	Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，娄桥分厂需完成笔数0"
	else
    c4=a4/b4*100
    Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，娄桥分厂需完成笔数[<font color='red'>"&b4&"</font>]，交期达成笔数[<font color='red'>"&a4&"</font>]，娄桥分厂交期达成率为：[<font color='red'>"&formatnumber(c4,2)&"%</font>]"
	end if
	end if
  else
	sqlstr2="select top 1 unDeliveryR1One as a1,DeliveryR1One as b1,DeliveryReach1One as c1,unDeliveryR2One as a2,DeliveryR2One as b2,DeliveryReach2One as c2,unDeliveryR3One as a3,DeliveryR3One as b3,DeliveryReach3One as c3,unDeliveryR4One as a4,DeliveryR4One as b4,DeliveryReach4One as c4 from manusys order by SerialNum desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sqlstr2,connzxpt,0,1
    a1=rs("a1")
    b1=rs("b1")
    a2=rs("a2")
    b2=rs("b2")
    a3=rs("a3")
    b3=rs("b3")
    a4=rs("a4")
    b4=rs("b4")
	c1=rs("c1")
	c2=rs("c2")
	c3=rs("c3")
	c4=rs("c4")
if Depart="10" then
anywhere = " and 分厂='一分厂'"
    Response.Write "统计时间：[<font color='red'>昨日</font>]，一分厂需完成笔数[<font color='red'>"&b1&"</font>]，交期达成笔数[<font color='red'>"&a1&"</font>]，一分厂交期达成率为：[<font color='red'>"&formatnumber(c1,2)&"%</font>]"
elseif Depart="11" then
anywhere = " and 分厂='二分厂'"
    Response.Write "统计时间：[<font color='red'>昨日</font>]，二分厂需完成笔数[<font color='red'>"&b2&"</font>]，交期达成笔数[<font color='red'>"&a2&"</font>]，二分厂交期达成率为：[<font color='red'>"&formatnumber(c2,2)&"%</font>]"
elseif Depart="12" then
anywhere = " and 分厂='三分厂'"
    Response.Write "统计时间：[<font color='red'>昨日</font>]，三分厂需完成笔数[<font color='red'>"&b3&"</font>]，交期达成笔数[<font color='red'>"&a3&"</font>]，三分厂交期达成率为：[<font color='red'>"&formatnumber(c3,2)&"%</font>]"
elseif Depart="19" then
anywhere = " and 分厂='娄桥分厂'"
    Response.Write "统计时间：[<font color='red'>昨日</font>]，娄桥分厂需完成笔数[<font color='red'>"&b4&"</font>]，交期达成笔数[<font color='red'>"&a4&"</font>]，娄桥分厂交期达成率为：[<font color='red'>"&formatnumber(c4,2)&"%</font>]"
else
anywhere = ""
    Response.Write "统计时间：[<font color='red'>昨日</font>]，一分厂需完成笔数[<font color='red'>"&b1&"</font>]，交期达成笔数[<font color='red'>"&a1&"</font>]，一分厂交期达成率为：[<font color='red'>"&formatnumber(c1,2)&"%</font>]，二分厂需完成笔数[<font color='red'>"&b2&"</font>]，交期达成笔数[<font color='red'>"&a2&"</font>]，二分厂交期达成率为：[<font color='red'>"&formatnumber(c2,2)&"%</font>]，三分厂需完成笔数[<font color='red'>"&b3&"</font>]，交期达成笔数[<font color='red'>"&a3&"</font>]，三分厂交期达成率为：[<font color='red'>"&formatnumber(c3,2)&"%</font>]，娄桥分厂需完成笔数[<font color='red'>"&b4&"</font>]，交期达成笔数[<font color='red'>"&a4&"</font>]，娄桥分厂交期达成率为：[<font color='red'>"&formatnumber(c4,2)&"%</font>]"
end if
  end if
  rs.close
  set rs=nothing
end function  
 
%>
<div id="ReplyDiv" style="width:590px;height:180px;top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
<form name="ReplyForm" id="ReplyForm" action="test1.asp">
<table id="ReplyTable" border="0" width="100%" cellspacing="0" cellpadding="1" align="center" bgcolor="black" height="100%">
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td width="60"> 回复人 </td>
 <td width="60">
 <input name="Replyer" type="text" id="Replyer" ></td>
 <td width="60"> 回复日期 </td>
 <td width="60">
 <input name="ReplyDate" type="text" id="ReplyDate" ></td>
 <td width="20" align="right"><img src="../images/close.jpg" onClick="javascript:closead()"></td>
</tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 回复内容 </td>
<td colspan="4">
  <textarea name="ReplyText" id="ReplyText" style="width:500px; height:100px; "></textarea>
</td>
</tr> 
<tr style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%;border-bottom: 1px solid #000000;">
<td valign="bottom" colspan="5" align="center">
<input type="hidden" name="FItemid" id="FItemid" value="">
<input type="hidden" name="Keyword" id="Keyword" value="">
&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="保存" style="WIDTH: 80;" onClick="SaveEdit()">
</td>
</tr>
</table>
</form>
</div>


<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>交期达成率信息</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="DeliveryReach.asp?Result=Search&Keyword=list&Page=1">
          <td nowrap> 产品检索：从<input type="text" class="textfield" style="width:80px" id="start_date" name="start_date" value="<%=StartDate%>"/>

          &nbsp;到<input type="text" class="textfield" style="width:80px" id="end_date" name="end_date" value="<%=EndDate%>" /><select id="department" name="department">
		  <option value="10">一分厂</option>
		  <option value="11">二分厂</option>
		  <option value="12">三分厂</option>
		  <option value="19">娄桥分厂</option>
		  </select>
		  <input name="submitSearch" type="submit" class="button" value="检索">
          </td>
        </form>
		<%if Result="Search" then%>
        <td align="right" nowrap>查看：
		<a href="DeliveryReach.asp?Result=Search&Keyword=list&start_date=<%=StartDate%>&end_date=<%=EndDate%>&department=<%=department%>&Page=1" onClick='changeAdminFlag("超交期列表")'>超交期列表</a>
		</td>
		<%end if%>
      </tr>
    </table>      </td>    
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td height="30"><%PlaceFlag()%></td>
  </tr>
</table>


  <% ProductsList() %>

</BODY>
</HTML>
<%
'-----------------------------------------------------------
function ProductsList()
 %>
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td width="80" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">分厂</font></strong></td>
    <td width="80" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">要求交期</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>客户排行</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>订单号</strong></font></td>
    <td width="80" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生产单号</strong></font></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品编号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>产品名称</strong></font></td>
    <td width="30" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>分类1</strong></font></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">分类2</font></strong></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">生产数量</font></strong></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">完工数量</font></strong></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">入库数量</font></strong></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">入库日期</font></strong></td>
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
      datafrom="manusys_DeliveryReach"
  dim datawhere'数据条件
      if Keyword="list" then
	     datawhere="where  UPTdate>='"&StartDate&"' and UPTdate<='"&EndDate&"' and 分厂='"&Reachper&"'"
	  else
		 datawhere="where datediff(d,UPTdate,getdate())=1 "&anywhere
 	  end if
  dim sqlid'本页需要用到的id
  dim Myself,PATH_INFO,QUERY_STRING'本页地址和参数
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if QUERY_STRING = "" or Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself = PATH_INFO & "?"&urlstr
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)&urlstr
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
    sql="select *,left(ReplyText,10) as a20 from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
	dim bgcolors
    while(not rs.eof)'填充数据到表格
	  bgcolors="#EBF2F9"
		if rs("replyFlag")>0 then
		  bgcolors="#ffff66"'
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("分厂")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("客户排行")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("销售订单号")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("生产任务单编号")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("产品代码")&"</td>"
      Response.Write "<td nowrap>"&rs("产品名称")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("产品类别1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("产品类别2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("生产数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("完工数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("入库数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("最后入库日期")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return SAClickTd(this,'DRreply',"&rs("SerialNum")&")"">"&rs("a20")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='12' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='14' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='14' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
  Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td>共计：<font color='#ff6600'>"&idcount&"</font>条记录&nbsp;页次：<font color='#ff6600'>"&page&"</font></strong>/"&pagec&"&nbsp;每页：<font color='#ff6600'>"&pages&"</font>条</td>" & vbCrLf
  Response.Write "<td align='right'>" & vbCrLf
  '设置分页页码开始===============================
  pagenmin=page-pagenc '计算页码开始值
  pagenmax=page+pagenc '计算页码结束值
  if(pagenmin<1) then pagenmin=1 '如果页码开始值小于1则=1
  if(page>1) then response.write ("<a href='"& myself &"Page=1'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>9</font></a>&nbsp;") '如果页码大于1则显示(第一页)
  if(pagenmin>1) then response.write ("<a href='"& myself &"Page="& page-(pagenc*2+1) &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>7</font></a>&nbsp;") '如果页码开始值大于1则显示(更前)
  if(pagenmax>pagec) then pagenmax=pagec '如果页码结束值大于总页数,则=总页数
  for i = pagenmin to pagenmax'循环输出页码
	if(i=page) then
	  response.write ("&nbsp;<font color='#ff6600'>"& i &"</font>&nbsp;")
	else
	  response.write ("[<a href="& myself &"Page="& i &">"& i &"</a>]")
	end if
  next
  if(pagenmax<pagec) then response.write ("&nbsp;<a href='"& myself &"Page="& page+(pagenc*2+1) &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>8</font></a>&nbsp;") '如果页码结束值小于总页数则显示(更后)
  if(page<pagec) then response.write ("<a href='"& myself &"Page="& pagec &"'><font style='FONT-SIZE: 14px; FONT-FAMILY: Webdings'>:</font></a>&nbsp;") '如果页码小于总页数则显示(最后页)	
  '设置分页页码结束===============================
  Response.Write "跳到：第&nbsp;<input name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' onchange=""if(/\D/.test(this.value)){alert('只能在跳转目标页框内输入整数！');this.value='"&Page&"';}"" style='HEIGHT: 18px;WIDTH: 40px;'  type='text' class='textfield' value='"&Page&"'>&nbsp;页" & vbCrLf
  Response.Write "<input style='HEIGHT: 18px;WIDTH: 20px;' name='submitSkip' type='button' class='button' onClick='GoPage("""&Myself&""")' value='GO'>" & vbCrLf
  Response.Write "</td>" & vbCrLf
  Response.Write "</tr>" & vbCrLf
  Response.Write "</table>" & vbCrLf

  Response.Write "</td>" & vbCrLf  
  Response.Write "</tr>" & vbCrLf
'-----------------------------------------------------------
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  <%
end function 

%>


