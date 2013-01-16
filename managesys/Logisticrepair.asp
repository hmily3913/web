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
if Instr(session("AdminPurview"),"|1001,")=0 then 
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
sqlstr="select a.fbillno,a.fdate5,a.fdate2,a.fdate3,t_Department.FName as name1,t_Emp.FName as name2,b.ftext,b.ftext1,b.ftext4, "&_
"(LTRIM(RTRIM(STR(b.fid)))+ 't' +LTRIM(RTRIM(STR(b.findex)))) AS id ,datediff(d,a.fdate2,isnull(a.fdate3,getdate())) as diffdays "&_
"from t_repair_apply a,t_BOS257709027Entry2 b, "&_
"dbo.t_Department,t_Emp  "&_
"where a.fbase=t_Department.FItemID and a.fid=b.fid "&_
"and a.fbase1=t_Emp.FItemID "&_
"and a.fcombobox1 = '行政检修' and a.FBillStatus='维修'  "&_
"and a.fdate5>='"&StartDate&"' and a.fdate5>='"&EndDate&"'"


function PlaceFlag()
  dim rs,sql,sqlstr2'sql语句
  if Result="Search" then
	sqlstr2="select sum(当日有效数) as idCount1,sum(当日总数) as idCount2,个人 from managesys_LogisticrepairSum where  UPTdate>='"&StartDate&"' and UPTdate<='"&EndDate&"' group by 个人"
    set rs=server.createobject("adodb.recordset")
    rs.open sqlstr2,connzxpt,0,1
	while (not rs.eof)
    Reachsum=rs("idCount1")
    unReachsum=rs("idCount2")
    if unReachsum=0 then
	Response.Write "[<font color='red'>"&rs("个人")&"</font>]统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，行政维修笔数为0<br>"
	else
    Reachper=Reachsum/unReachsum*100
    Response.Write "[<font color='red'>"&rs("个人")&"</font>]统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，行政维修笔数[<font color='red'>"&unReachsum&"</font>]，行政维修及时笔数[<font color='red'>"&Reachsum&"</font>]，行政维修及时率为：[<font color='red'>"&formatnumber(Reachper,2)&"%</font>]<br>"
	end if
	rs.movenext
	wend
  else
	sqlstr2="select * from managesys_LogisticrepairSum where datediff(d,UPTdate,getdate())=1"
    set rs=server.createobject("adodb.recordset")
    rs.open sqlstr2,connzxpt,1,1
	while(not rs.eof)
	if rs("当日总数")=0 then
	 Reachper=0
	else
	 Reachper=rs("当日有效数")*100/rs("当日总数")
	end if
	if rs("本月总数")=0 then
	 Reachsum=0
	else
	 Reachsum=rs("本月有效数")*100/rs("本月总数")
	end if
    Response.Write "[<font color='red'>"&rs("个人")&"</font>]统计时间：[<font color='red'>昨日</font>]，行政维修笔数[<font color='red'>"&rs("当日总数")&"</font>]，行政维修及时笔数[<font color='red'>"&rs("当日有效数")&"</font>]，行政维修及时率为：[<font color='red'>"&formatnumber(Reachper,2)&"%</font>]<br>"
    Response.Write "　　　　统计时间：[<font color='red'>本月</font>]，行政维修笔数[<font color='red'>"&rs("本月总数")&"</font>]，行政维修及时笔数[<font color='red'>"&rs("本月有效数")&"</font>]，行政维修及时率为：[<font color='red'>"&formatnumber(Reachsum,2)&"%</font>]<br>"
	rs.movenext
	wend
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
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>行政维修及时率信息</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="Logisticrepair.asp?Result=Search">
          <td nowrap> 产品检索：从
          <script language=javascript> 
          var myDate=new dateSelector(); 
          myDate.year; 
		  myDate.date; 
          myDate.inputName='start_date';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。 
          myDate.display(); 
          </script>
          &nbsp;到
          <script language=javascript> 
          myDate.year; 
          myDate.inputName='end_date';  //注意这里设置输入框的name，同一页中的日期输入框，不能出现重复的name。 
          myDate.display(); 
          </script>
		  <input name="submitSearch" type="submit" class="button" value="检索">
          </td>
        </form>
		<%if Result="Search" then%>
        <td align="right" nowrap>查看：
		<a href="Logisticrepair.asp?Result=Search&Keyword=list&start_date=<%=StartDate%>&end_date=<%=EndDate%>&Page=1" onClick='changeAdminFlag("行政维修延误")'>行政维修延误</a>
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
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>请修部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>请修类型</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>设备名称</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>问题描述</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>维修进度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>处理情况</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>计划时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>处理时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>结果反馈</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈时间</strong></font></td>
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
      datafrom="Bill_RepairApplication"
  dim datawhere'数据条件
      if Keyword="list" then
	     datawhere="where CheckFlag>1 and RepairType='行政维修' and isnull(datediff(hh,ReplyFinishDate,RReplyDate),1)>12 and datediff(d,ReplyFinishDate,'" & StartDate & "')<=0 and datediff(d,ReplyFinishDate,'" & EndDate & "')>=0"
	  else
		 datawhere="where CheckFlag>1 and RepairType='行政维修' and isnull(datediff(hh,ReplyFinishDate,RReplyDate),1)>12 and datediff(d,ReplyFinishDate,getdate())=1"
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
  sql="select count(SerialNum) as idCount from "& datafrom &" " & datawhere
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
    sql="select *,left(SituationDescrib,10) as a1,left(RReplyText,10) as a2,left(AReplyText,10) as a3  from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
	dim bgcolors
	dim formdata(3)
    while(not rs.eof)'填充数据到表格
	  bgcolors="#EBF2F9"
		if rs("CheckFlag")="1" then
		  formdata(0)="已审核"
		  bgcolors="#ffff66"'黄色
		elseif rs("CheckFlag")="2" then
		  formdata(0)="维修中"
		  bgcolors="#ff99ff"'粉色
		elseif rs("CheckFlag")="3" then
		  formdata(0)="已完成"
		  bgcolors="#66ff66"'绿色
		else
		  formdata(0)="未审核"
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap >"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("RegisterName")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("Departmentname")&"</td>"
      Response.Write "<td nowrap >"&rs("ReceivDepartment")&"</td>"
      Response.Write "<td nowrap >"&rs("RepairType")&"</td>"
      Response.Write "<td nowrap >"&rs("DeviceName")&"</td>"
      Response.Write "<td nowrap >"&rs("a1")&"</td>"
      Response.Write "<td nowrap >"&formdata(0)&"</td>"
      Response.Write "<td nowrap >"&rs("a2")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("ReplyFinishDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("RReplyDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("a3")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("AReplyDate")&"</td>" & vbCrLf
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


