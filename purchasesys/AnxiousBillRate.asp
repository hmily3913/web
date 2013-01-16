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
if Instr(session("AdminPurview"),"|202,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
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
sqlstr="SELECT     TOP (100) PERCENT dbo.PORequest.FBillNo AS 采购申请号, dbo.PORequest.FDate AS 采购申请日期, "&_
"		FCheckTime as 审核日期, dbo.PORequestEntry.FEntrySelfP0132 AS 工厂订单号,  "&_
"    dbo.PORequestEntry.FSourceBillNo AS 销售订单号, dbo.t_Item.FName AS 物料类别,  "&_
"    ICItem2.FNumber AS 物料代码, ICItem2.FName AS 物料名称,  "&_
"     dbo.PORequestEntry.FQty AS 基本采购申请数量,  "&_
"     dbo.t_MeasureUnit.FName AS 基本单位, dbo.PORequestEntry.FFetchTime AS 交货日期,  "&_
"     t2.FFixLeadTime as 前置周期, t_user.FName AS 制单人, emp2.FName AS 申请人, "&_
"    (t2.FFixLeadTime-datediff(dd,dbo.PORequest.FDate ,dbo.PORequestEntry.FFetchTime)) as diffdays, "&_
"    (LTRIM(RTRIM(STR(PORequestEntry.FInterID)))+ 't' +LTRIM(RTRIM(STR(PORequestEntry.FEntryID)))) AS id "&_
"FROM         dbo.PORequest INNER JOIN "&_
"    dbo.PORequestEntry ON dbo.PORequest.FInterID = dbo.PORequestEntry.FInterID LEFT OUTER JOIN "&_
"    dbo.t_Base_Emp AS emp2 ON dbo.PORequest.FRequesterID = emp2.FItemID LEFT OUTER JOIN "&_
"    dbo.t_ICItemCore AS ICItem2 ON dbo.PORequestEntry.FItemID = ICItem2.FItemID LEFT OUTER JOIN "&_
"    dbo.t_ICItemBase AS t1 ON ICItem2.FItemID = t1.FItemID LEFT OUTER JOIN "&_
"    dbo.t_ICItemplan AS t2 ON ICItem2.FItemID = t2.FItemID LEFT OUTER JOIN "&_
"    dbo.t_MeasureUnit ON t1.FUnitID = dbo.t_MeasureUnit.FMeasureUnitID LEFT OUTER JOIN "&_
"    dbo.t_Base_User AS t_user ON t_user.FUserID = dbo.PORequest.FBillerID LEFT OUTER JOIN "&_
"    dbo.t_Item ON ICItem2.FParentID = dbo.t_Item.FItemID "&_
"where  PORequest.FDate>='"&StartDate&"' and PORequest.FDate<='"&EndDate&"' and dbo.PORequest.FcheckerID>0"


'response.Write(sqlstr)
function PlaceFlag()
  dim rs,sql,sqlstr2'sql语句
  if Result="Search" then
	sqlstr2="select sum(AnxiousBillOne) as idCount1,sum(unAnxiousBillOne) as idCount2 from purchasesys where  UPTdate>='"&StartDate&"' and UPTdate<='"&EndDate&"'"
    set rs=server.createobject("adodb.recordset")
    rs.open sqlstr2,connzxpt,0,1
    Reachsum=rs("idCount1")
    unReachsum=rs("idCount2")
    if unReachsum=0 then
	Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，共有采购笔数为0"
	else
    Reachper=Reachsum/unReachsum*100
    Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，共有采购笔数[<font color='red'>"&unReachsum&"</font>]，急单笔数[<font color='red'>"&Reachsum&"</font>]，急单率为：[<font color='red'>"&formatnumber(Reachper,2)&"%</font>]"
	end if
  else
	sqlstr2="select top 1 AnxiousBillOne as idCount1,unAnxiousBillOne as idCount2,AnxiousBillRateOne from purchasesys order by SerialNum desc"
    set rs=server.createobject("adodb.recordset")
    rs.open sqlstr2,connzxpt,0,1
    Reachsum=rs("idCount1")
    unReachsum=rs("idCount2")
    Reachper=rs("AnxiousBillRateOne")
    Response.Write "统计时间：[<font color='red'>昨日</font>]，共有采购笔数[<font color='red'>"&unReachsum&"</font>]，急单笔数[<font color='red'>"&Reachsum&"</font>]，急单率为：[<font color='red'>"&Reachper&"%</font>]"
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
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>采购急单率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="AnxiousBillRate.asp?Result=Search">
          <td nowrap> 产品检索：从<input type="text" class="textfield" style="width:80px" id="start_date" name="start_date" value="<%=StartDate%>"/>

          &nbsp;到<input type="text" class="textfield" style="width:80px" id="end_date" name="end_date" value="<%=EndDate%>" /><input name="submitSearch" type="submit" class="button" value="检索">
          </td>
        </form>
		<%if Result="Search" then%>
        <td align="right" nowrap>查看：
		<a href="AnxiousBillRate.asp?Result=Search&Keyword=list&start_date=<%=StartDate%>&end_date=<%=EndDate%>&Page=1" onClick='changeAdminFlag("采购急单列表")'>采购急单列表</a><font color="#0000FF">
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
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td width="80" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请单号</strong></font></td>
    <td width="80" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>审核日期</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物料代码</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物料名称</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">物料类别</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">基本单位</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">申请数量</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>交货日期</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">前置周期</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">申请人</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">使用部门</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">急单类别</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">急单原因</font></strong></td>
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
      datafrom="purchase_AnxiousBillRate"
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
    sql="select *,left(ReplyText,10) as a1 from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
	dim bgcolors
    while(not rs.eof)'填充数据到表格
	  bgcolors="#EBF2F9"
		if rs("replyFlag")>0 then
		  bgcolors="#ffff66"'
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("采购申请日期")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("采购申请号")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("审核日期")&"</td>"
      Response.Write "<td nowrap>"&rs("物料代码")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("物料名称")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("物料类别")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("基本单位")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("基本采购申请数量")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("交货日期")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("前置周期")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("申请人")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("使用部门")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("急单类别")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("急单原因")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return SAClickTd(this,'ABRreply',"&rs("SerialNum")&")"">"&rs("a1")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='13' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='15' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='15' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
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


