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
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|802,")=0 then 
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
sqlstr="select sum(a4)+sum(a2)-sum(a1)-sum(a3) as b1,fname as id from ("&_
"select sum(本位币应付金额) as a1,sum(本位币实付金额) as a2,0 as a3,0 as a4, fname from ("&_
"select t_Supplier.fnumber,t_Supplier.fname,t_PayColCondition.fname as 付款条件,                 "&_
"t_rp_contact.fyear as 年,t_rp_contact.fperiod as 期间,t_rp_contact.ftype as 类别,               "&_
"t_rp_contact.fdate as 单据日期,t_rp_contact.ffincdate as 财务日期,                              "&_
"t_rp_contact.fnumber as 单据编号, Department1.fname as 部门,emp.fname as 业务员,                "&_
"t_rp_contact.FExchangeRate as 汇率,                                                             "&_
"case when t_rp_contact.ftype=6 then 0 else t_rp_contact.FAmountfor end as 应付金额,             "&_
"case when t_rp_contact.ftype=6 then 0 else t_rp_contact.FAmount end 本位币应付金额,             "&_
"case when t_rp_contact.ftype<>6 then 0 else t_rp_contact.FAmountfor end as 实付金额,            "&_
"case when t_rp_contact.ftype<>6 then 0 else t_rp_contact.FAmount end 本位币实付金额,         "&_
"frpdate as 应付日期                                                                             "&_
"from t_rp_contact inner join                                                                    "&_
" t_Supplier on fcustomer =t_Supplier.FItemID left outer join                                    "&_
" t_PayColCondition on t_Supplier.fcreditdays=t_PayColCondition.fid LEFT OUTER JOIN              "&_
"dbo.t_Department AS Department1 ON Department1.FItemID =t_rp_contact.fdepartment LEFT OUTER JOIN "&_
"dbo.t_Emp AS emp ON t_rp_contact.femployee = emp.FItemID                                        "&_
"where frp=0 and frpdate>='"&StartDate&"' and frpdate<='"&EndDate&"'                                 "&_
") as aaa                                                                                        "&_
"group by aaa.fname                                                                              "&_
"union all                                                                                       "&_
"select 0 as a1,0 as a2,sum(本位币应付金额) as a3,sum(本位币实付金额) as a4, fname from (        "&_
"select t_Supplier.fnumber,t_Supplier.fname,t_PayColCondition.fname as 付款条件,                 "&_
"t_rp_contact.fyear as 年,t_rp_contact.fperiod as 期间,t_rp_contact.ftype as 类别,               "&_
"t_rp_contact.fdate as 单据日期,t_rp_contact.ffincdate as 财务日期,                              "&_
"t_rp_contact.fnumber as 单据编号, Department1.fname as 部门,emp.fname as 业务员,                "&_
"t_rp_contact.FExchangeRate as 汇率,                                                             "&_
"case when t_rp_contact.ftype=6 then 0 else t_rp_contact.FAmountfor end as 应付金额,             "&_
"case when t_rp_contact.ftype=6 then 0 else t_rp_contact.FAmount end 本位币应付金额,             "&_
"case when t_rp_contact.ftype<>6 then 0 else t_rp_contact.FAmountfor end as 实付金额,            "&_
"case when t_rp_contact.ftype<>6 then 0 else t_rp_contact.FAmount end 本位币实付金额,            "&_
"frpdate as 应付日期                                                                             "&_
"from t_rp_contact inner join                                                                    "&_
" t_Supplier on fcustomer =t_Supplier.FItemID left outer join                                    "&_
" t_PayColCondition on t_Supplier.fcreditdays=t_PayColCondition.fid LEFT OUTER JOIN              "&_
"dbo.t_Department AS Department1 ON Department1.FItemID =t_rp_contact.fdepartment LEFT OUTER JOIN "&_
"dbo.t_Emp AS emp ON t_rp_contact.femployee = emp.FItemID                                        "&_
"where frp=0 and frpdate<'"&StartDate&"'                                                            "&_
") as aaa                                                                                        "&_
"group by aaa.fname) bbb                                                                         "&_
"group by fname                                                                                  "&_
"having  sum(a4)+sum(a2)-sum(a1)-sum(a3) <0"   


'response.Write(sqlstr)
function PlaceFlag()
  if Result="Search" then
    dim rs,sql,sqlstr2'sql语句
	dim tem1,tem2
	sqlstr2="select count(id) as idCount from ("&sqlstr&") aaa "
  set rs=server.createobject("adodb.recordset")
'  response.Write(sqlstr2)
  rs.open sqlstr2,connk3,0,1
  Reachsum=rs("idCount")'本期应收
    if Reachsum=0 then
	Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，应付数为0"
	else
	tem2=tem1-unReachsum
    Reachper=tem2/Reachsum*100
    Response.Write "统计时间：[<font color='red'>"&StartDate&"至"&EndDate&"</font>]，未及时付款供应商数[<font color='red'>"&Reachsum&"</font>]"
	end if
  else
      Response.Write "请选择日期进行统计!"
  end if
  rs.close
  set rs=nothing
end function  
 
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>付款及时率</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="PayPrompt.asp?Result=Search">
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
		  <input name="submitSearch" type="submit" class="button" value="检索">(为保证查询速度时间尽量不超过3个月)
          </td>
        </form>
		<%if Result="Search" then%>
        <td align="right" nowrap>查看：
		<a href="PayPrompt.asp?Result=Search&Keyword=list&start_date=<%=StartDate%>&end_date=<%=EndDate%>&Page=1" onClick='changeAdminFlag("未及时付款供应商")'>未及时付款供应商</a><font color="#0000FF"></a>
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
 if Keyword<>"" then
 %>
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td width="120" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>供应商名称</strong></font></td>
    <td width="80" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>未付金额</strong></font></td>
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
	     datawhere=""
	  else
		 datawhere=""
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
      taxis=""
  dim i'用于循环的整数
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(id) as idCount from ("& datafrom &") as aaa " & datawhere
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
    sql="select id from ("& datafrom &")  as aaa " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = pages '每页显示记录数
    if page < 1 then page = 1
    if page > pagec then page = pagec
    if pagec > 0 then rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid="'"&rs("id")&"'"
	  else
	    sqlid=sqlid &",'"&rs("id")&"'"
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  if(idcount>0 and sqlid<>"") then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    sql="select * from ("& datafrom &")  as aaa where aaa.id in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while(not rs.eof)'填充数据到表格
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("id")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&-rs("b1")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
'    Response.Write "<td colspan='5' nowrap  bgcolor='#EBF2F9'>&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='2' nowrap  bgcolor='#EBF2F9'></td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
  else
    response.write "<tr><td height='50' align='center' colspan='2' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
'-----------------------------------------------------------
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td colspan='2' nowrap  bgcolor='#D7E4F7'>" & vbCrLf
  Response.Write "<table width='100%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
  Response.Write "<tr>" & vbCrLf
  Response.Write "<td>共计：<font color='#ff6600'>"&Reachsum&"</font>条订单<font color='#ff6600'>"&idcount&"</font>条记录&nbsp;页次：<font color='#ff6600'>"&page&"</font></strong>/"&pagec&"&nbsp;每页：<font color='#ff6600'>"&pages&"</font>条</td>" & vbCrLf
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
end if
end function 

%>


