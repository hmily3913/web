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
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
'if Instr(session("AdminPurview"),"|1006,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
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
Reachsum=request("Reachsum")
sqlstr="z_GoodsCarryOutMain a,z_GoodsCarryOutDetails b"
 
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>物品携出信息</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="GoodsOutPassMana.asp?Result=Search&Keyword=list">
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
        <td align="right" nowrap>
		<a href="GoodsOutPassEdit.asp?Result=GoodsOut&Action=Add" onClick='changeAdminFlag("物品携出登记")'>物品携出登记</a>
		<font color="#0000FF">&nbsp;|&nbsp;</font>
<a href="GoodsOutPassMana.asp?Result=Search&Keyword=all&Page=1" onClick='changeAdminFlag("全部物品携出信息")'>全部物品携出信息</a>
		</td>
      </tr>
    </table>      </td>    
  </tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td height="30"></td>
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
  <form action="GoodsCarryOutPrint.asp" method="post" name="formPrint" target="new_window"   onsubmit="window.open('GoodsCarryOutPrint.asp', 'new_window')">
  <tr>
    <td width="76" colspan="2" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong>
	<input onClick="CheckAll(this.form)" name="buttonAllSelect" type="button" class="button"  id="submitAllSearch" value="全" style="HEIGHT: 18px;WIDTH: 16px;">
    <input onClick="CheckOthers(this.form)" name="buttonOtherSelect" type="button" class="button"  id="submitOtherSelect" value="反" style="HEIGHT: 18px;WIDTH: 16px;">	
	</td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>单据编号</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td width="60" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>携出日期</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>取回日期</strong></font></td>
    <td width="100" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>物品名称</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>数量</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">用途说明</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">是否回厂</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">放行状态</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">取回状态</font></strong></td>
    <td width="40" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">打印</font></strong></td>
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
      if Keyword="list" then
	     datawhere=" where a.SerialNum=b.SerialNum and a.RegDate >='"&StartDate&"' and a.RegDate <='"&EndDate&"' "
	  else
		 datawhere=" where a.SerialNum=b.SerialNum "
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
      taxis=" order by a.SerialNum desc, b.Findex asc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(b.Fentryid) as idCount from "& datafrom &" " & datawhere
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
    sql="select b.Fentryid from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = pages '每页显示记录数
    if page < 1 then page = 1
    if page > pagec then page = pagec
    if pagec > 0 then rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("Fentryid")
	  else
	    sqlid=sqlid &","&rs("Fentryid")
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  if(idcount>0 and sqlid<>"") then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
	dim OutCheckFlag,InCheckFlag,PrintFlag,ReturnFlag
    sql="select * from "& datafrom &" where a.SerialNum=b.SerialNum and b.Fentryid in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
	dim tempbill
	tempbill=""
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
	  if tempbill=rs("SerialNum") then
	  Response.Write "<td nowrap></td>" & vbCrLf
	  Response.Write "<td nowrap></td>" & vbCrLf
	  Response.Write "<td nowrap></td>" & vbCrLf
	  Response.Write "<td nowrap></td>" & vbCrLf
	  Response.Write "<td nowrap></td>" & vbCrLf
	  Response.Write "<td nowrap></td>" & vbCrLf
	  Response.Write "<td nowrap></td>" & vbCrLf
	  else
      Response.Write "<td width='65' nowrap><a href='GoodsOutPassEdit.asp?Result=GoodsOut&Action=Modify&SerialNum="&rs("SerialNum")&"' >查看</a>|<b onClick='window.open(""GoodsCarryOutPrint.asp?SerialNum="&rs("SerialNum")&""",""Print"","""",""false"")' >打印</b></td>" & vbCrLf
      Response.Write "<td width='22' nowrap><input name='SerialNum' type='checkbox' value='"&rs("SerialNum")&"' style='HEIGHT: 13px;WIDTH: 13px;'></td>" & vbCrLf
	  Response.Write "<td nowrap>"&rs("SerialNum")&"</td>" & vbCrLf
	  Response.Write "<td nowrap>"&getUser(rs("Register"))&"</td>" & vbCrLf
	  Response.Write "<td nowrap>"&rs("RegDate")&"</td>" & vbCrLf
	  Response.Write "<td nowrap>"&rs("GetOutDate")&"</td>" & vbCrLf
	  Response.Write "<td nowrap>"&rs("GetInDate")&"</td>" & vbCrLf
	  tempbill=rs("SerialNum")
	  end if
      Response.Write "<td nowrap>"&rs("Goods")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FQty")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("UseState")&"</td>"
	  if rs("ReturnFlag")=1 then
	  ReturnFlag="是"
	  else
	  ReturnFlag="否"
	  end if
      Response.Write "<td nowrap>"&ReturnFlag&"</td>" & vbCrLf
	  if rs("OutCheckFlag")=1 then
	  OutCheckFlag="主管审核"
	  elseif rs("OutCheckFlag")=2 then
	  OutCheckFlag="门卫审核"
	  else
	  OutCheckFlag="未审核"
	  end if
	  if rs("InCheckFlag")=1 then
	  InCheckFlag="主管审核"
	  elseif rs("InCheckFlag")=2 then
	  InCheckFlag="门卫审核"
	  else
	  InCheckFlag="未审核"
	  end if
      Response.Write "<td nowrap>"&OutCheckFlag&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&InCheckFlag&"</td>" & vbCrLf
	  if rs("PrintFlag")=1 then
	  PrintFlag="√"
	  else
	  PrintFlag="×"
	  end if
      Response.Write "<td nowrap>"&PrintFlag&"</td>" & vbCrLf
	  Response.Write "</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='12' nowrap  bgcolor='#EBF2F9'><input name='submitPrintSelect' type='submit' class='button'  id='submitPrintSelect' value='打印所选' >&nbsp;</td>" & vbCrLf
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
      </form>
  </table>
  <%

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

