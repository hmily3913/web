<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="报表系统" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
<link rel="stylesheet" href="Images/jquery.datepick.css">
<style type="text/css">
.item {
  width: 300px;
	height:200px;
  margin: 10px;
  float: left;
	background:#E8EFFF;
}


</style>
<script language="javascript" src="Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="Script/jquery.masonry.min.js"></script>
<script language="javascript" src="Script/jquery.easydrag.js"></script>
<script language="javascript" src="Script/jquery.datepick.pack.js"></script>
<script language="javascript" src="Script/jquery.datepick-zh-CN.js"></script>
<script language="javascript">
$(function(){
  $('#container').masonry({
    // options
    itemSelector : '.item',
    columnWidth : 320
  });
	messFresh();
	partstateFresh();
	$('#datepicker').datepick({dateFormat: 'yyyy-mm-dd'});
});
function messFresh(){
	$('#mess').html('<div style="width:100%;height:155px"><img src="Images/images/flexigrid/load.gif" /></div>');
  jQuery.get("CheckMessage.asp", { "key": "AllNeed"},
   function(data){
		if(data.length>0){
				$('#mess').html(data);
		}else{
			$('#mess').html('暂无信息');
		}
   });
}
function partstateFresh(){
	$('#partstate').html('<div style="width:100%;height:155px"><img src="Images/images/flexigrid/load.gif" /></div>');
  jQuery.get("CheckMessage.asp", { "key": "PartnerState"},
   function(data){
		if(data.length>0){
				$('#partstate').html(data);
		}else{
			$('#partstate').html('暂无信息');
		}
   });
}
function showNotice(sn,tt){
	jQuery.get("CheckMessage.asp", { "key": "showNotice","SerialNum":sn},
	function(data){
		var sHTML='<html><head><meta content="text/html; charset=UTF-8" http-equiv="Content-Type"/><title>'+tt+'</title></head><body>' + data + '</body></html>';
		var screen=window.screen,oWindow=window.open('', 'xhePreview', 'toolbar=yes,location=no,status=yes,menubar=yes,scrollbars=yes,resizable=yes,width='+Math.round(screen.width*0.9)+',height='+Math.round(screen.height*0.8)+',left='+Math.round(screen.width*0.05)),oDoc=oWindow.document;
		oDoc.open();
		oDoc.write(sHTML);
		oDoc.close();
		oWindow.focus();
	});
}
</script>
</HEAD>
<!--#include file="CheckAdmin.asp"-->
<!--#include file="Include/ConnSiteData.asp" -->
<BODY>
<div id="container">
  <div class="item">
	<div>
	<div class="tablemenu" style="height:24px;"><img src="Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>公告通知</strong></div>
	<ul>
  <%
	sql="select top 5 * from oa_Announce order by BillDate desc"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,Connzxpt,1,1
	if rs.eof and rs.bof then
		response.Write("<li> 暂无公告！</li>")
	else
	while (not rs.eof)
		response.Write("<li style='cursor:hand;' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" onClick='showNotice("&rs("SerialNum")&","""&rs("Title")&""")' title='点击查看对应公告内容'>"&rs("Title"))
		if Instr(rs("Reader"),UserName)=0 then response.Write("<font color='#FF0000'>(新)</font>")
		response.Write("<font style='position: absolute;right:10px'>"&rs("BillDate")&"</font></li>")
		rs.movenext
	wend
	end if
	%>
	</ul>
	</div>
	</div>
  <div class="item">
	<div >
	<div class="tablemenu" style="height:24px;"><img src="Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>近一周考勤</strong><a href="Attendance/AttQuery.asp" target="_self" style="position: absolute;right: 5px;top:4px;">更多</a></div>
  <div style="overflow:auto;height:175px; padding-left:10px;">
<table border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
<tr class="TitleRow">
<td bgcolor="#8DB5E9">日期</td>
<td bgcolor="#8DB5E9">班次</td>
<td bgcolor="#8DB5E9">打卡一</td>
<td bgcolor="#8DB5E9">打卡二</td>
<td bgcolor="#8DB5E9">异常</td>
</tr>

<%
sql=""
for i=0 to 6
sql=sql&"select convert(varchar(10),dateadd(d,-"&i&",getdate()),120) as date,a.userid,a.ssn,c.num_runid,c.name,c.units,d.sdays,d.edays,e.schclassid,e.schName, "
sql=sql&"e.starttime,e.endtime,e.checkin,e.checkout,e.checkintime1,e.checkintime2,e.checkouttime1,e.checkouttime2,e.workday "
sql=sql&"from USERINFO a,USER_OF_RUN b,NUM_RUN c,NUM_RUN_DEIL d,SchClass e,(select max(Order_run) as Order_run_num,UserID from USER_OF_RUN group by UserID) i "
sql=sql&"where a.userid=b.userid and b.num_of_run_id=c.num_runid and b.Order_run=i.Order_run_num and b.Userid=i.Userid "
sql=sql&"and c.num_runid=d.num_runid and d.schclassid=e.schclassid and b.startdate<=dateadd(d,-"&i&",getdate()) and b.enddate>=dateadd(d,-"&i&",getdate()) and a.ssn='"&session("UserName")&"' and ((DATEPART(weekday,dateadd(d,-"&i&",getdate()))-1=d.sdays%7 and c.units=1) or ((day(dateadd(d,-"&i&",getdate()))+(datediff(m,c.startdate,dateadd(d,-"&i&",getdate()))%c.cyle)*31)=d.sdays and c.units=2)) "
if i<6 then sql=sql&" union all "
next
sql=sql&"order by date desc,e.schclassid "
  set rs=server.createobject("adodb.recordset")
  rs.open sql,ConnStrkq,0,1
	while (not rs.eof)
	dim yclb:yclb=""
	response.Write("<tr bgcolor='#EBF2F9'>")
	response.Write("<td>"&rs("date")&"</td>")
	response.Write("<td>"&rs("schName")&"</td>")
		sql2="select min(checktime) as ttime from CHECKINOUT where datediff(d,checktime,'"&rs("date")&"')=0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkintime1")&"',114))<0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkintime2")&"',114))>0 and userid="&rs("userid")
		
			sql2="select min(checktime) as ttime from CHECKINOUT where datediff(d,checktime,'"&rs("date")&"')=0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkintime1")&"',114))<0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkintime2")&"',114))>0 and userid="&rs("userid")
			set rs2=server.createobject("adodb.recordset")
			rs2.open sql2,ConnStrkq,0,1
			if isnull(rs2("ttime")) then
			if rs("checkin")=1 then yclb="旷工"
	response.Write("<td></td>")
			else
				if datediff("s",rs2("ttime"),rs("date")&" "&right(rs("starttime"),8))<0 then
	response.Write("<td>"&right(rs2("ttime"),8)&"</td>")
	if rs("checkin")=1 then yclb="迟到"
				else
	response.Write("<td>"&right(rs2("ttime"),8)&"</td>")
				end if
			end if
			sql2="select max(checktime) as ttime from CHECKINOUT where datediff(d,checktime,'"&rs("date")&"')=0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkouttime1")&"',114))<0 and datediff(s,checktime,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("checkouttime2")&"',114))>0 and userid="&rs("userid")
			set rs2=server.createobject("adodb.recordset")
			rs2.open sql2,ConnStrkq,0,1
			if isnull(rs2("ttime")) then
			if rs("checkout")=1 then yclb="旷工"
	response.Write("<td></td>")
			else
				if datediff("s",rs2("ttime"),rs("date")&" "&left(rs("endtime"),12))>0 then
	response.Write("<td>"&right(rs2("ttime"),8)&"</td>")
		if rs("checkout")=1 then yclb="早退"
				else
	response.Write("<td>"&right(rs2("ttime"),8)&"</td>")
				end if
			end if
		sql2="select b.leavename from USER_SPEDAY a,LeaveClass b where a.dateid=b.leaveid and datediff(d,startspecday,'"&rs("date")&"')=0 and (datediff(s,startspecday,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("endtime")&"',114))>=0 and datediff(s,endspecday,'"&rs("date")&"'+' '+convert(varchar(8),'"&rs("starttime")&"',114))<=0) and userid="&rs("userid")
		set rs2=server.createobject("adodb.recordset")
		rs2.open sql2,ConnStrkq,0,1
		if rs2.eof and rs2.bof then
			response.Write("<td>"&yclb&"</td>")
		else
			response.Write("<td>"&rs2("leavename")&"</td>")
		end if
		response.Write("</tr>" & vbCrLf)
		rs.movenext
	wend
  rs2.close
  set rs2=nothing
  rs.close
  set rs=nothing
%>
	</table>
	</div>
	</div>
	</div>
  <div class="item">
	<div>
	<div class="tablemenu" style="height:24px;"><img src="Images/images/flexigrid/load.png" width="16" height="16" border="0" align="absmiddle" onClick="messFresh()" alt="刷新">&nbsp;<strong>待办事项</strong></div>
	<div id="mess" style="padding:20px; overflow:auto;height:175px;">
	</div>
	</div>
	</div>
  <div class="item">
	<div>
	<div class="tablemenu" style="height:24px;"><img src="Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>日历</strong></div>
	<div id="datepicker" align="center" style="overflow:hidden;height:175px;"></div>
	</div>
	</div>
  <div class="item">
	<div>
	<div class="tablemenu" style="height:24px;"><img src="Images/images/flexigrid/load.png" width="16" height="16" border="0" align="absmiddle" onClick="partstateFresh()" alt="刷新">&nbsp;<strong>同事状态</strong></div>
	<div id="partstate" style="padding:20px; overflow:auto;height:175px;">
	</div>
	</div>
	</div>
  <div class="item">
	<div>
	<div class="tablemenu" style="height:24px;"><img src="Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>天气预报</strong></div>
	<iframe src="/Include/weather.htm" width="100%" height="175px" marginwidth="0" marginheight="0" hspace="0" vspace="0" frameborder="0" scrolling="no"></iframe>
	</div>
	</div>
	</div>
</BODY>
</HTML>