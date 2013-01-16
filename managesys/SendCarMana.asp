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
<script language="javascript" src="../Script/CustomAjax.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
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
//处理添加按钮
function showpadd(obj,sid){
	$('#addShowDiv').load("SendCarManaDetails.asp #listtable",{
	  showType:'DetailsList'
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#addDiv").show("slow");
		$("#maindiv").hide("slow");
	  }	
    })
}
function closead1(){
	$("#addDiv").hide("slow");
	$("#maindiv").show("slow");
}
function closead12(){
	$("#addDetailDiv").hide("slow");
	$("#addShowDiv").show("slow");
}
function toAdd(obj,sid){
	$('#addDetailDiv').load("SendCarManaDetails.asp #AddandEditdiv",{
	  showType:'AddEditShow',
	  detailType:obj,
	  SerialNum:sid
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#addDetailDiv").css('display',"block");
		$("#addShowDiv").hide("slow");
	  }	
    })
}
function getInfo(obj){
  if($("#"+obj).val()==''){alert("对应编号不能为空！");return false;}
	$.get("SendCarManaDetails.asp", { showType: "getInfo",detailType: obj, InfoID: $("#"+obj).val() },
	   function(data){
		 if(data.indexOf("###")==-1){
		   alert("对应编号不存在，请检查！");
		   $("#"+obj).val('');
		   $("#Driver").val('');
		 }
		 else{
		   if(obj=="DriverID"){
		   $("#Driver").val(data.split('###')[1]);
		   }
		 }
	   });
}
function toSubmit(){
  $.post('SendCarManaDetails.asp?showType=DataProcess',$("#AddForm").serialize(),function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	else showpadd();
  });
	$("#addDetailDiv").hide("slow");
	$("#addShowDiv").show("slow");
}

</script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
'if Instr(session("AdminPurview"),"|1003,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
%>
<BODY>
<%
dim Result,StartDate,EndDate,Keyword,Reachsum,unReachsum,Reachper,sqlstr,checkst,checkdx,useCarFlag
Result=request("Result")
StartDate=request("start_date")
if StartDate="" then StartDate=date()
EndDate=request("end_date")
if EndDate="" then EndDate=date()
Keyword=request("Keyword")
Reachsum=request("Reachsum")
checkst=request("checkst")
checkdx=request("checkdx")
useCarFlag=request("useCarFlag")
sqlstr="t_item a,z_SendCar "
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>外出单信息</strong></font>
	<font style="background-color:#EBF2F9">未审核、已结束</font>
	<font style="background-color:#ffff66">主管审核</font>
	<font style="background-color:#FFDAB9">车队一审</font>
	<font style="background-color:#ff99ff">门卫一审</font>
	<font style="background-color:#66ff66">门卫二审</font>
	<font style="background-color:#808080">车队驳回</font>
	</td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="SendCarMana.asp?Result=Search&Keyword=list&Page=1">
          <td nowrap> 产品检索：从<input type="text" class="textfield" style="width:80px" id="start_date" name="start_date" value="<%=StartDate%>"/>

          &nbsp;到<input type="text" class="textfield" style="width:80px" id="end_date" name="end_date" value="<%=EndDate%>" /><input name="submitSearch" type="submit" class="button" value="检索">
		  &nbsp;关键字：<input type="text" name="Reachsum" id="Reachsum" value="<%=Reachsum%>" size="10">
		  &nbsp;状态<select id="checkdx" name="checkdx">
		  <option value="=" <% If checkdx="=" Then Response.Write("selected")%>>=</option>
		  <option value="<" <% If checkdx="<" Then Response.Write("selected")%>><</option>
		  <option value="<=" <% If checkdx="<=" Then Response.Write("selected")%>><=</option>
		  <option value=">" <% If checkdx=">" Then Response.Write("selected")%>>></option>
		  <option value=">=" <% If checkdx=">=" Then Response.Write("selected")%>>>=</option>
		  </select>
		  <select id="checkst" name="checkst">
		  <option value=""></option>
		  <option value="0" <% If checkst="0" Then Response.Write("selected")%>>未审核</option>
		  <option value="1" <% If checkst="1" Then Response.Write("selected")%>>已审核</option>
		  <option value="2" <% If checkst="2" Then Response.Write("selected")%>>车队一</option>
		  <option value="3" <% If checkst="3" Then Response.Write("selected")%>>门卫一</option>
		  <option value="4" <% If checkst="4" Then Response.Write("selected")%>>门卫二</option>
		  <option value="5" <% If checkst="5" Then Response.Write("selected")%>>车队二</option>
		  </select>
		  <select id="useCarFlag" name="useCarFlag">
		  <option value=""></option>
		  <option value="是" <% If useCarFlag="是" Then Response.Write("selected")%>>用车</option>
		  <option value="否" <% If useCarFlag="否" Then Response.Write("selected")%>>不用车</option>
		  </select>
		  <select id="queryFlag" name="queryFlag">
		  <option value=""></option>
		  <option value="kqgl" >考勤关联</option>
		  </select>
		  <input name="submitSearch" type="submit" class="button" value="检索">
          </td>
        </form>
      </tr>
      <tr>
		<td>
		<a href="SendCarEdit.asp?Result=SendCar&Action=Add" onClick='changeAdminFlag("外出单登记")'>外出单登记</a>
		<font color="#0000FF">&nbsp;|&nbsp;</font>
<a href="SendCarMana.asp?Result=Search&Keyword=all&Page=1" onClick='changeAdminFlag("全部外出信息")'>全部外出信息</a>
		<font color="#0000FF">&nbsp;|&nbsp;</font>
<a href="javascript:showpadd();" onClick='changeAdminFlag("车辆信息")'>查看车辆信息</a><font color="#0000FF">&nbsp;|&nbsp;</font>
		  <a href="SendCarMana.asp?Result=Search&Keyword=all&queryFlag=none&Page=1" onClick='changeAdminFlag("外出单查询")'>查未审核</a>&nbsp;|&nbsp;
		  <a href="SendCarMana.asp?Result=Search&Keyword=all&queryFlag=chedui&Page=1" onClick='changeAdminFlag("外出单查询")'>车队待审</a>&nbsp;|&nbsp;
		  <a href="SendCarMana.asp?Result=Search&Keyword=all&queryFlag=menwei&Page=1" onClick='changeAdminFlag("外出单查询")'>门卫待审</a>&nbsp;|&nbsp;
		</td>
      </tr>
    </table>      </td>    
  </tr>
</table>
<div id="addDiv" style="width:100%;height:'480px';display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
<font color="#FF0000"><strong><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<input type="button" name="toAdd" id="toAdd" onClick="toAdd('AddNew','')" value="添加" style='HEIGHT: 18px;WIDTH: 65px;font-size:12px;'>
</font>
<div id="addShowDiv"></div>
<div id="addDetailDiv"></div>
</div>

  <% ProductsList() %>

</BODY>
</HTML>
<%
'-----------------------------------------------------------
function ProductsList()
 if Result="Search" then
 %>
 <div id="maindiv">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <form action="SendCarPrint.asp" method="post" name="formPrint" target="new_window"   onsubmit="window.open('SendCarPrint.asp', 'new_window')">
  <tr>
    <td width="76" colspan="2" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong>
	<input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
    <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">	
	</td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单号</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td width="60" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请部门</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>外出工号</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>外出人员</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>外出分类</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>用车</strong></font></td>
    <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>事由及内容</strong></font></td>
    <td width="120" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">目的地</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">计划出发</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">计划回来</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">出发时间</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">回来时间</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">里程</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">状态</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">打印</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">重要性</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">车牌号</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">驾驶员</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">预计里程</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">使用时间</font></strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">备注</font></strong></td>
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
      datafrom=sqlstr
  dim datawhere'数据条件
  dim Myself,PATH_INFO,QUERY_STRING'本页地址和参数
      PATH_INFO = request.servervariables("PATH_INFO")
	  QUERY_STRING = request.ServerVariables("QUERY_STRING")'
      if Keyword="list" then
	     datawhere=" where DeleteFlag<1 and RegistDepartment=a.fitemid and (SendReason like '%"&Reachsum&"%' or GoodsName like '%"&Reachsum&"%' or DeliveryAddr like '%"&Reachsum&"%' or Register like '%"&Reachsum&"%' or OutPeron like '%"&Reachsum&"%' or CarNumber like '%"&Reachsum&"%') and RegDate>='"&StartDate&"' and RegDate<='"&EndDate&"'"
		 if Instr(QUERY_STRING,"Page=1")>1 then QUERY_STRING="Reachsum="&server.urlencode(Reachsum)&"&start_date="&StartDate&"&end_date="&EndDate&"&"&QUERY_STRING
		 if checkst<>"" then
		   datawhere=datawhere&" and checkflag"&checkdx&checkst
		   if Instr(QUERY_STRING,"Page=1")>1 then QUERY_STRING="checkdx="&server.urlencode(checkdx)&"&checkst="&checkst&"&"&QUERY_STRING
		 end if
		 if useCarFlag<>"" then
		   datawhere=datawhere&" and useCarFlag='"&useCarFlag&"'"
		   if Instr(QUERY_STRING,"Page=1")>1 then QUERY_STRING="useCarFlag="&server.urlencode(useCarFlag)&"checkdx="&server.urlencode(checkdx)&"&checkst="&checkst&"&"&QUERY_STRING
		 end if
		 if request("queryFlag")="kqgl" then
		   datawhere=datawhere&" and ISCQ=1 "
		   if Instr(QUERY_STRING,"Page=1")>1 then QUERY_STRING="useCarFlag="&server.urlencode(useCarFlag)&"checkdx="&server.urlencode(checkdx)&"&checkst="&checkst&"&queryFlag="&request("queryFlag")&"&"&QUERY_STRING
		 end if
	  else
		 datawhere=" where DeleteFlag<1 and RegistDepartment=a.fitemid "
		 dim queryFlag:queryFlag=request("queryFlag")
		 if queryFlag<>"" then
		   if queryFlag="none" then
		     datawhere=datawhere&" and checkflag=0"
		   elseif queryFlag="chedui" then
		     datawhere=datawhere&" and (checkflag=1 or checkFlag=4) and UseCarFlag='是' "
		   elseif queryFlag="menwei" then
		     datawhere=datawhere&" and ((checkflag=1 and UseCarFlag='否') or checkFlag=2 or checkFlag=3) "
		   elseif queryFlag="kqgl" then
		     datawhere=datawhere&" and ISCQ=1 "
		   end if
		 end if
 	  end if
	if Instr(session("AdminPurview"),"|1003.2,")=0 and Instr(session("AdminPurview"),"|1003.3,")=0 and Instr(session("AdminPurview"),"|1003.9,")=0 then
	dim Depart:Depart=session("Depart")
	  if Depart="KD01.0001.0001"  then
		datawhere=datawhere&" and (a.fnumber like '06%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0002" then
		datawhere=datawhere&" and (a.fnumber like '03%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0003" then
		datawhere=datawhere&" and (a.fnumber like '05%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0005.0004" then
		datawhere=datawhere&" and (a.fnumber like '02%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0005" then
		datawhere=datawhere&" and (a.fnumber like '08%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0006" then
		datawhere=datawhere&" and (a.fnumber like '07%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0007" then
		datawhere=datawhere&" and (a.fnumber like '11%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0008" then
		datawhere=datawhere&" and (a.fnumber like '12%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0009" then
		datawhere=datawhere&" and (a.fnumber like '04%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0010" then
		datawhere=datawhere&" and (a.fnumber like '10%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0011" then
		datawhere=datawhere&" and (a.fnumber like '09%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0018" then
		datawhere=datawhere&" and (a.fnumber like '10.04%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0019" then
		datawhere=datawhere&" and (a.fnumber like '23%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  elseif Depart="KD01.0001.0012" AND Instr(session("AdminPurview"),"|1003.5,")=0 then
		datawhere=datawhere&" and (a.fnumber like '01%' or register='"&session("UserName")&"' or fbiller='"&session("UserName")&"')"
	  end if
  end if
  dim sqlid'本页需要用到的id
      if QUERY_STRING = "" or Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")=0 then
	    Myself = PATH_INFO & "?"
	  else
	    Myself = Left(PATH_INFO & "?" & QUERY_STRING,Instr(PATH_INFO & "?" & QUERY_STRING,"Page=")-1)
	  end if
  dim taxis'排序的语句 asc,desc
      taxis=" order by SerialNum desc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
	datawhere=datawhere&Session("AllMessage20")&Session("AllMessage51")
	session.contents.remove "AllMessage20"
	session.contents.remove "AllMessage51"
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
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
    rs.open sql,connk3,1,1
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
	dim CheckFlag,PrintFlag,bgcolors
    sql="select z_SendCar.*,t_item.fname as name1,t_emp.fname as name2 from z_SendCar left join t_item on t_item.fitemid=RegistDepartment left join t_emp on t_emp.fnumber=z_SendCar.Register where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while(not rs.eof)'填充数据到表格
	dim outpersonname,OutPeron
	outpersonname=""
	  if rs("CheckFlag")=1 then
	  CheckFlag="主管审核"
	bgcolors="#ffff66"
	  elseif rs("CheckFlag")=2 then
	bgcolors="#FFDAB9"
	  CheckFlag="车队一审"
	  elseif rs("CheckFlag")=3 then
	bgcolors="#ff99ff"
	  CheckFlag="门卫一审"
	  elseif rs("CheckFlag")=4 then
	bgcolors="#66ff66"
	  CheckFlag="门卫二审"
	  elseif rs("CheckFlag")=5 then
	bgcolors="#EBF2F9"
	  CheckFlag="已结案"
	  else
	bgcolors="#EBF2F9"
	  CheckFlag="未审核"
	  if rs("RejecteFlag")=1 then bgcolors="#808080"'作废标志
	  end if

	  OutPeron=rs("OutPeron")
	  if OutPeron<>"" then
	  dim ooo
	  ooo=0
	  while (ooo<=UBound(split(OutPeron,",")))
	  if ooo<>UBound(split(OutPeron,",")) then
	  outpersonname=outpersonname+getUser(split(OutPeron,",")(ooo))+","
	  else
	  outpersonname=outpersonname+getUser(split(OutPeron,",")(ooo))
	  end if
	  ooo=ooo+1
	  wend
	  end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td width='65' nowrap><a href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"' >查看</a>|<b onClick='window.open(""SendCarPrint.asp?SerialNum="&rs("SerialNum")&""",""Print"","""",""false"")' >打印</b></td>" & vbCrLf
      Response.Write "<td width='22' nowrap><input name='SerialNum' type='checkbox' value='"&rs("SerialNum")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("name2")&"</td>"
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("name1")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&OutPeron&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&outpersonname&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("SendReason")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("UseCarFlag")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"" width='150'>"&rs("GoodsName")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"" width='150' >"&rs("DeliveryAddr")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("PlanStarteDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("PlanEndDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("StarteDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("EndDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("mileage")&"</td>" & vbCrLf
	  if rs("PrintFlag")=1 then
	  PrintFlag="√"
	  else
	  PrintFlag="×"
	  end if
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&CheckFlag&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&PrintFlag&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("Importance")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("CarNumber")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("Driver")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("Planmileage")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("totalTime")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""location.href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'"">"&rs("Remark")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='23' nowrap  bgcolor='#EBF2F9' align='left'><input name='submitPrintSelect' type='submit' class='button'  id='submitPrintSelect' value='打印所选' >&nbsp;<input name='output' type='button' class='button'  id='output' value='引出' onClick='window.open(""SendCarOutPut.asp?"&Left(QUERY_STRING,Instr(QUERY_STRING,"Page=")-2)&""",""Print"","""",""false"")'>&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='25' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='25' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
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
    </form>
  </table>
  </div>
  <%
  elseif Result="CarRPT" then
%>
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td colspan="11" height="30" bgcolor="#8DB5E9" align="center"><font color="#FFFFFF" size="+1"><strong>车辆状态表</strong></font></td>
  </tr>
  <tr>
    <td width="20" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>序号</strong></font></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>车牌号</strong></font></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>驾驶员</strong></font></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>驾驶手机</strong></font></td>
    <td width="40" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>当前状态</strong></font></td>
    <td width="40" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>剩余载货量</strong></font></td>
    <td width="40" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>剩余载人数</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>当前位置</strong></font></td>
    <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>出发时间</strong></font></td>
    <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>预计回来时间</strong></font></td>
    <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>计划下次派车时间</strong></font></td>
  </tr>
<%
    sql="select CarID,CarSatus, "&_
"case when CarSatus='出行' then DeliveryAddr else '车队' end as weizhi,z_SendCar.SerialNum,z_SendCar.StarteDate, "&_
"case when CarSatus='出行' then PlanEndDate end as returntime, "&_
"nextstime,z_Car.Driver,DriverPhone,z_Car.CarryGoods,z_Car.CarryMans "&_
"from z_Car left join z_SendCar on z_SendCar.CarNumber=z_Car.CarID and CheckFlag=3 "&_
"left join (select CarNumber,min(PlanStarteDate) as nextstime from z_SendCar where CheckFlag=1 and usecarflag='是' and datediff(n,getdate(),PlanStarteDate)>0 group by CarNumber) as aaa "&_
"on aaa.CarNumber=z_Car.CarID "
	dim iii,temcar
	iii=1
	temcar=""
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&iii&"</td>" & vbCrLf
	  if temcar<>rs("CarID") then
      Response.Write "<td nowrap>"&rs("CarID")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Driver")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("DriverPhone")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("CarSatus")&"</td>"
	  temcar=rs("CarID")
	  else
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap></td>" & vbCrLf
      Response.Write "<td nowrap></td>" & vbCrLf
	  end if
      Response.Write "<td nowrap>"&rs("CarryGoods")&"</td>"
      Response.Write "<td nowrap>"&rs("CarryMans")&"</td>"
      Response.Write "<td nowrap width=""320"">"
	  if rs("CarSatus")="出行" then
	    Response.Write "<a href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'>"&rs("weizhi")&"</a>"
	  else
	    Response.Write rs("weizhi")
	  end if
	  Response.Write "</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("StarteDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("returntime")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("nextstime")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  iii=iii+1
	  rs.movenext
    wend
Response.Write "</table>" & vbCrLf


  end if
end function 
Function getUser(ID)
  Dim rs,sql
  Set rs=server.CreateObject("adodb.recordset")
  sql="Select * From t_emp where fnumber='"&ID&"'"
  rs.open sql,connk3,1,1
  if rs.bof and rs.eof then
  getUser=""
  else
  getUser=rs("Fname")
  end if
  rs.close
  set rs=nothing
End Function    

%>


