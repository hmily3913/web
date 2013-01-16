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

if Instr(session("AdminPurviewFLW"),"|201,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
dim showType,print_tag,UserName,AdminName,detailType
UserName=session("UserName")
AdminName=session("AdminName")
dim sql,rs
showType=request("showType")
print_tag=request("print_tag")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="DetailsList" then 
  if print_tag<>"1" then
%>
 <div id="listtable" style="width:100%; height:'100%'; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请工号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>职务</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>职等</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>入职日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请事项</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生效日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>审核进度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>部门主管审核</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>相关部门审核</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总监副总审核</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总经理审核</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>是否实施</strong></font></td>
  </tr>
 <%
  dim Depart,wherestr,checkflag4seach,leixing,startdate,enddate
  wherestr=""
  checkflag4seach=request("checkflag4seach")
  leixing=request("leixing")
  Depart=session("Depart")
	if UserName="A09837" then Depart="KD01.0005.0004"
	startdate=request("startdate")
	enddate=request("enddate")
  dim page'页码
      page=clng(request("Page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Bill_WelfareAdjust inner join LDERP.dbo.[N-基本资料单头] on Register=员工代号 "
  dim datawhere,datawhere2'数据条件
  '老总及实施单位（人资部）
  if Instr(session("AdminPurviewFLW"),"|201.3,")>0 or Instr(session("AdminPurviewFLW"),"|201.6,")>0 or Instr(session("AdminPurviewFLW"),"|201.5,")>0 or Instr(session("AdminPurviewFLW"),"|201.4,")>0 then
    datawhere="where (1=1) "
	if Instr(session("AdminPurviewFLW"),"|201.4,")>0 then
		datawhere=" where (left(Department,9)=left('"&Depart&"',9) or (left('"&Depart&"',9)='KD01.0005' and left(Department,9)<>'KD01.0001') )"
		datawhere2=" and (Register='"&UserName&"' or Biller='"&UserName&"' "
	elseif Instr(session("AdminPurviewFLW"),"|201.2,")>0 then 
	  datawhere2=" and (Department='"&Depart&"' "
		if Depart="KD01.0001.0012" then datawhere2=" and ((Department='"&Depart&"' or Department='KD01.0001.0005') "
	else
	  datawhere2=" and (Register='"&UserName&"' or Biller='"&UserName&"' "
	end if
  '部门主管，只能看本部门
  elseif Instr(session("AdminPurviewFLW"),"|201.2,")>0 or Instr(session("AdminPurviewFLW"),"|201.1,")>0 then
    datawhere="where (Department='"&Depart&"' or Register='"&UserName&"' or Biller='"&UserName&"') "
		if left(Depart,9)="KD01.0004" then datawhere="where (Department like 'KD01.0004%' or Register='"&UserName&"' or Biller='"&UserName&"') "
	datawhere2=" and (Register='"&UserName&"' or Biller='"&UserName&"' "
  else
    datawhere="where (Register='"&UserName&"' or Biller='"&UserName&"') "
	datawhere2=" and (Register='"&UserName&"' or Biller='"&UserName&"' "
  end if
  if Instr(session("AdminPurviewFLW"),"|201.7,")>0 then    datawhere2=datawhere2&" or shenqxm='工资调薪'"
  if Instr(session("AdminPurviewFLW"),"|201.8,")>0 then    datawhere2=datawhere2&" or shenqxm='话费补贴'"
  if Instr(session("AdminPurviewFLW"),"|201.9,")>0 then    datawhere2=datawhere2&" or shenqxm='住房补贴'"
  if Instr(session("AdminPurviewFLW"),"|201.10,")>0 then    datawhere2=datawhere2&" or shenqxm='岗位补贴'"
  if Instr(session("AdminPurviewFLW"),"|201.11,")>0 then    datawhere2=datawhere2&" or shenqxm='其他补贴'"
  if Instr(session("AdminPurviewFLW"),"|201.12,")>0 then    datawhere2=datawhere2&" or shenqxm='职等调整'"
  if Instr(session("AdminPurviewFLW"),"|201.13,")>0 then    datawhere2=datawhere2&" or shenqxm='工龄恢复'"
  datawhere=datawhere&datawhere2&") "
  if checkflag4seach="999" then 
    wherestr=" and checkFlag<100 "
  else
    wherestr=" and CheckFlag="&checkflag4seach
  end if
  if leixing<>"" then wherestr=wherestr&" and shenqxm='"&leixing&"'"
  if startdate<>"" then wherestr=wherestr&" and EffectiveDate>='"&startdate&"'"
  if enddate<>"" then wherestr=wherestr&" and EffectiveDate<='"&enddate&"'"
  '拼装条件
  datawhere=datawhere&wherestr&Session("AllMessage2")&Session("AllMessage3")&Session("AllMessage4")&Session("AllMessage41")
	session.contents.remove "AllMessage2"
	session.contents.remove "AllMessage3"
	session.contents.remove "AllMessage4"
	session.contents.remove "AllMessage41"
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by SerialNum desc"
  dim i'用于循环的整数
  '获取记录总数
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
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
    sql="select SerialNum from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    rs.pagesize = 20 '每页显示记录数
	rs.absolutepage = page  
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
'-----------------------------------------------------------
'-----------------------------------------------------------
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    dim sql2,rs2
	dim formdata(3),bgcolors
    sql="select *,left(DepartReplyText,10) as a1,left(RelatedReplyText,10) as a2,left(DirectorReplyText,10) as a3,left(CEOReplyText,10) as a4,left(ImplementReplyText,10) as a5,职等,到职日 from "& datafrom &" where  SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
		if rs("CheckFlag")=100 then
		  bgcolors="#B9BBC7"'灰色
		elseif rs("CheckFlag")=99 then
		  bgcolors="#ff99ff"'黄色
		elseif rs("CheckFlag")>0 then
		  bgcolors="#ffff66"'粉色
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegisterName")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Register")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Departmentname")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Position")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("职等")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("到职日")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("shenqxm")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("EffectiveDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"
	  dim checkstat
	  checkstat="未审核"
	  if rs("CheckFlag")="1" then
	    checkstat="主管审核"
	  elseif rs("CheckFlag")="2" then
	    checkstat="相关部门审核"
	  elseif rs("CheckFlag")="3" then
	    checkstat="副总审核"
	  elseif rs("CheckFlag")="4" then
	    checkstat="总经理审核"
	  elseif rs("CheckFlag")="99" then
	    checkstat="确认执行"
	  elseif rs("CheckFlag")="100" then
	    checkstat="作废"
	  end if
	  Response.Write checkstat&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return ClickTd(this,'a1reply',"&rs("SerialNum")&")"">"&rs("DepartFlag")&":"&rs("DepartReplyer")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return ClickTd(this,'a2reply',"&rs("SerialNum")&")"">"&rs("RelatedFlag")&":"&rs("RelatedReplyer")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return ClickTd(this,'a3reply',"&rs("SerialNum")&")"">"&rs("DirectorFlag")&":"&rs("DirectorReplyer")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return ClickTd(this,'a4reply',"&rs("SerialNum")&")"">"&rs("CEOFlag")&":"&rs("CEOReplyer")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return ClickTd(this,'a5reply',"&rs("SerialNum")&")"">"&rs("ImplementFlag")&":"&rs("ImplementReplyer")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  else
    response.write "<tr><td height='50' align='center' colspan='14' nowrap  bgcolor='#EBF2F9'>暂无申请信息</td></tr>"
  end if
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
	response.Write("###"&pagec&"###"&idcount&"###")
  else
%>
 <table width="100%" border="1" cellpadding="0" cellspacing="0" bgcolor="#99BBE8">
  <tr align="center">
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>申请单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>申请工号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>申请部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>职务</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>申请事项</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>生效日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2"><font color="#FFFFFF"><strong>薪资调整</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2"><font color="#FFFFFF"><strong>话费补贴</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2"><font color="#FFFFFF"><strong>住房补贴</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" ><font color="#FFFFFF"><strong>岗位补贴</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" ><font color="#FFFFFF"><strong>其他补贴</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" colspan="2"><font color="#FFFFFF"><strong>职等调整</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" ><font color="#FFFFFF"><strong>工龄恢复</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>审核进度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>部门主管审核</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>相关部门审核</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>总监副总审核</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>总经理审核</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" rowspan="2"><font color="#FFFFFF"><strong>是否实施</strong></font></td>
  </tr>
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>原薪资</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>拟调薪资</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>话费补贴标准</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>特殊话费补贴</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>住房补贴金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>配偶住房补贴</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>原职等</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请职等</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>历史工龄</strong></font></td>
  </tr>
 <%
  wherestr=""
  checkflag4seach=request("checkflag4seach")
  leixing=request("leixing")
  Depart=session("Depart")
	startdate=request("startdate")
	enddate=request("enddate")
  
  if Instr(session("AdminPurviewFLW"),"|201.3,")>0 or Instr(session("AdminPurviewFLW"),"|201.6,")>0 or Instr(session("AdminPurviewFLW"),"|201.5,")>0 or Instr(session("AdminPurviewFLW"),"|201.4,")>0 then
    datawhere="where (1=1) "
	if Instr(session("AdminPurviewFLW"),"|201.2,")>0 then 
	  datawhere2=" and (Department='"&Depart&"' "
	else
	  datawhere2=" and (Register='"&UserName&"' or Biller='"&UserName&"' "
	end if
  '部门主管，只能看本部门
  elseif Instr(session("AdminPurviewFLW"),"|201.2,")>0 or Instr(session("AdminPurviewFLW"),"|201.1,")>0 then
    datawhere="where (Department='"&Depart&"') "
	datawhere2=" and (Register='"&UserName&"' or Biller='"&UserName&"' "
  else
    datawhere="where (Register='"&UserName&"' or Biller='"&UserName&"') "
	datawhere2=" and (Register='"&UserName&"' or Biller='"&UserName&"' "
  end if
  if Instr(session("AdminPurviewFLW"),"|201.7,")>0 then    datawhere2=datawhere2&" or shenqxm='工资调薪'"
  if Instr(session("AdminPurviewFLW"),"|201.8,")>0 then    datawhere2=datawhere2&" or shenqxm='话费补贴'"
  if Instr(session("AdminPurviewFLW"),"|201.9,")>0 then    datawhere2=datawhere2&" or shenqxm='住房补贴'"
  if Instr(session("AdminPurviewFLW"),"|201.10,")>0 then    datawhere2=datawhere2&" or shenqxm='岗位补贴'"
  if Instr(session("AdminPurviewFLW"),"|201.11,")>0 then    datawhere2=datawhere2&" or shenqxm='其他补贴'"
  if Instr(session("AdminPurviewFLW"),"|201.12,")>0 then    datawhere2=datawhere2&" or shenqxm='职等调整'"
  if Instr(session("AdminPurviewFLW"),"|201.13,")>0 then    datawhere2=datawhere2&" or shenqxm='工龄恢复'"
  datawhere=datawhere&datawhere2&") "
  if checkflag4seach="999" then 
    wherestr=" and CheckFlag<100 "
  else
    wherestr=" and CheckFlag="&checkflag4seach
  end if
  '拼装条件
  if leixing<>"" then wherestr=wherestr&" and shenqxm='"&leixing&"'"
  if startdate<>"" then wherestr=wherestr&" and EffectiveDate>='"&startdate&"'"
  if enddate<>"" then wherestr=wherestr&" and EffectiveDate<='"&enddate&"'"
  datawhere=datawhere&wherestr
  taxis=" order by SerialNum desc"
      datafrom=" Bill_WelfareAdjust "
    sql="select * from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
		if rs("CheckFlag")=100 then
		  bgcolors="#B9BBC7"'灰色
		elseif rs("CheckFlag")=99 then
		  bgcolors="#ff99ff"'黄色
		elseif rs("CheckFlag")>0 then
		  bgcolors="#ffff66"'粉色
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap >"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("RegisterName")&"</td>"
      Response.Write "<td nowrap >"&rs("Register")&"</td>"
      Response.Write "<td nowrap >"&rs("Departmentname")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("Position")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("shenqxm")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("EffectiveDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("FormerSalary")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("WantSalary")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("FeeStandards")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("SpecialFee")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("HousingFee")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("PartHousingFee")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("ApplicFee")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("ApplicFee")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("Employment")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("ApplicEmployment")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("Oldlength")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"
	  checkstat="未审核"
	  if rs("CheckFlag")="1" then
	    checkstat="主管审核"
	  elseif rs("CheckFlag")="2" then
	    checkstat="相关部门审核"
	  elseif rs("CheckFlag")="3" then
	    checkstat="副总审核"
	  elseif rs("CheckFlag")="4" then
	    checkstat="总经理审核"
	  elseif rs("CheckFlag")="99" then
	    checkstat="确认执行"
	  elseif rs("CheckFlag")="100" then
	    checkstat="作废"
	  end if
	  Response.Write checkstat&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("DepartFlag")&":"&rs("DepartReplyer")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("RelatedFlag")&":"&rs("RelatedReplyer")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("DirectorFlag")&":"&rs("DirectorReplyer")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("CEOFlag")&":"&rs("CEOReplyer")&"</td>" & vbCrLf
      Response.Write "<td nowrap >"&rs("ImplementFlag")&":"&rs("ImplementReplyer")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend

  end if
elseif showType="AddEditShow" then 
  dim SerialNum,Register,RegisterName,RegDate,Department,Departmentname,Position,shenqxm,CheckFlag
  dim FormerSalary,CWantSalary,WantSalary,ApplicReasons,Explain
  dim EffectiveDate,VirtualNetFlag,PositionDate,FeeStandards,SpecialFee
  dim HusbandWifeFlag,IntoCompanyDate,Employment,HousingFee,SpecialHousing
  dim PartID,PartName,PartIntoCompanyDate,PartEmployment,PartHousingFee,PartSpecialHousing,TotalHousingFee,PartPosition
  dim ApplicFee
  dim ApplicEmployment,OldPosition,NewPosition
  dim OldIntoDate,OldOutDate,DueType,Oldlength
  dim style1,style2,style3
  detailType=request("detailType")
  dim checkType,Replyer,ReplyText,ReplyDate
  if detailType="Add" then
	sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&UserName&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
	Register=UserName
	RegisterName=AdminName
	RegDate=date()
	Department=rs("部门别")
	Departmentname=rs("部门名称")
	Position=rs("工作岗位")
	OldPosition=rs("工作岗位")
	EffectiveDate=date()
	IntoCompanyDate=rs("到职日")
	PositionDate=rs("到职日")
	Employment=rs("职等")
	CheckFlag=0
	if Month(now())=12 then
	  EffectiveDate=(Year(now())+1)&"-01-01"
	else
	  EffectiveDate=Year(now())&"-"&(Month(now())+1)&"-01"
	end if
	'计算住房补贴
	'计算公式：一职等员工：100元+20元×工龄
	'          二职等以上人员：60元×（职等×60%+工龄×40%）
	
	dim emplength
	'计算工龄，从入职日期开始，满年才算1
	emplength=datediff("yy",rs("到职日"),date())
	if rs("职等") = 1 then 
	  HousingFee=100+20*emplength
	else
	  HousingFee=60*(rs("职等")*0.6+emplength*0.4)
	end if
	if HousingFee > 400 then HousingFee=400
	style1="none"
	style2="block"
  elseif detailType="Edit" then
    SerialNum=request("SerialNum")
	sql="select * from Bill_WelfareAdjust where SerialNum="&SerialNum
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
    Register=rs("Register")
	RegisterName=rs("RegisterName")
	RegDate=rs("RegDate")
	Department=rs("Department")
	Departmentname=rs("Departmentname")
	Position=rs("Position")
	shenqxm=rs("shenqxm")
    FormerSalary=rs("FormerSalary")
		CWantSalary=rs("CWantSalary")
	WantSalary=rs("WantSalary")
	ApplicReasons=rs("ApplicReasons")
	Explain=rs("Explain")
    EffectiveDate=rs("EffectiveDate")
	VirtualNetFlag=rs("VirtualNetFlag")
	PositionDate=rs("PositionDate")
	FeeStandards=rs("FeeStandards")
	SpecialFee=rs("SpecialFee")
    HusbandWifeFlag=rs("HusbandWifeFlag")
	IntoCompanyDate=rs("IntoCompanyDate")
	Employment=rs("Employment")
	HousingFee=rs("HousingFee")
	SpecialHousing=rs("SpecialHousing")
    ApplicFee=rs("ApplicFee")
    ApplicEmployment=rs("ApplicEmployment")
    OldIntoDate=rs("OldIntoDate")
	OldOutDate=rs("OldOutDate")
	DueType=rs("DueType")
	Oldlength=rs("Oldlength")
	CheckFlag=rs("CheckFlag")
	OldPosition=rs("OldPosition")
	NewPosition=rs("NewPosition")
	PartID=rs("PartID")
	PartName=rs("PartName")
	PartIntoCompanyDate=rs("PartIntoCompanyDate")
	PartEmployment=rs("PartEmployment")
	PartHousingFee=rs("PartHousingFee")
	PartSpecialHousing=rs("PartSpecialHousing")
	TotalHousingFee=rs("TotalHousingFee")
	PartPosition=rs("PartPosition")
	checkType=request("checkType")
	style3="hidden"
	If checkType="a1reply" Then 
		Replyer=rs("DepartReplyer")
		ReplyText=rs("DepartReplyText")
		ReplyDate=rs("DepartReplyDate")
		style1="block"
		style2="none"
		if CheckFlag<=1 and Instr(session("AdminPurviewFLW"),"|201.2,")>0 then style3="visible"
	elseif checkType="a2reply" Then 
		Replyer=rs("RelatedReplyer")
		ReplyText=rs("RelatedReplyText")
		ReplyDate=rs("RelatedReplyDate")
		style1="block"
		style2="none"
		if CheckFlag<=2 and (Instr(session("AdminPurviewFLW"),"|201.3,")>0 or Instr(session("AdminPurviewFLW"),"|201.3,")>0) then style3="visible"
	elseif checkType="a3reply" Then 
		Replyer=rs("DirectorReplyer")
		ReplyText=rs("DirectorReplyText")
		ReplyDate=rs("DirectorReplyDate")
		style1="block"
		style2="none"
		if CheckFlag<=3 and Instr(session("AdminPurviewFLW"),"|201.4,")>0 then style3="visible"
	elseif checkType="a4reply" Then 
		Replyer=rs("CEOReplyer")
		ReplyText=rs("CEOReplyText")
		ReplyDate=rs("CEOReplyDate")
		style1="block"
		style2="none"
		if CheckFlag<=4 and Instr(session("AdminPurviewFLW"),"|201.5,")>0 then style3="visible"
	elseif checkType="a5reply" Then 
		Replyer=rs("ImplementReplyer")
		ReplyText=rs("ImplementReplyText")
		ReplyDate=rs("ImplementReplyDate")
		style1="block"
		style2="none"
		if CheckFlag<=99 and Instr(session("AdminPurviewFLW"),"|201.6,")>0 then style3="visible"
	Else 
		style2="block"
		style1="none"
	end if
  end if
  rs.close
  set rs=nothing
%>
 <div id="AddandEditdiv" style="width:790px; height:460px; overflow:auto; ">
 <form id="AddForm" name="AddForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" id="t1">
  <tr>
    <td height="24" nowrap id="formove"><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>薪资、福利、津贴变动表申请单</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews" width="100%">
      <tr>
        <td height="20" align="left">单据号：</td>
        <td>
		<input name="SerialNum" type="text" class="textfield" id="SerialNum" style="WIDTH: 140;" value="<%= SerialNum %>" maxlength="100" readonly="true"></td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left">申请人工号：</td>
        <td><input name="Register" type="text" class="textfield" id="Register" style="WIDTH: 140;" value="<%= Register %>" maxlength="100" onBlur="changeEmp()"></td>
        <td height="20" align="left">申请人姓名：</td>
        <td><input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 140;" value="<%= RegisterName %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">申请日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 140;" value="<%= RegDate %>" maxlength="100"  readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">部门编号：</td>
        <td>
		<input name="Department" type="text" class="textfield" id="Department" style="WIDTH: 140;" value="<%= Department %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">部门名称：</td>
        <td>
		<input name="Departmentname" type="text" class="textfield" id="Departmentname" style="WIDTH: 140;" value="<%= Departmentname %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">岗位职务：</td>
        <td><input name="Position" type="text" class="textfield" id="Position" style="WIDTH: 140;" value="<%= Position %>" maxlength="100"  readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">申请项目：</td>
        <td colspan="5">
		<input type="radio" name="shenqxm" id="shenqxm1" value="工资调薪" onClick="changexm()" <% If shenqxm="工资调薪" Then Response.Write("checked") %>><label for="shenqxm1">工资调薪</label>
		<input type="radio" name="shenqxm" id="shenqxm2" value="话费补贴" onClick="changexm()" <% If shenqxm="话费补贴" Then Response.Write("checked") %>><label for="shenqxm2">话费补贴</label>
		<input type="radio" name="shenqxm" id="shenqxm3" value="住房补贴" onClick="changexm()" <% If shenqxm="住房补贴" Then Response.Write("checked") %>><label for="shenqxm3">住房补贴</label>
		<input type="radio" name="shenqxm" id="shenqxm4" value="岗位补贴" onClick="changexm()" <% If shenqxm="岗位补贴" Then Response.Write("checked") %>><label for="shenqxm4">岗位补贴</label>
		<input type="radio" name="shenqxm" id="shenqxm5" value="其他补贴" onClick="changexm()" <% If shenqxm="其他补贴" Then Response.Write("checked") %>><label for="shenqxm5">其他补贴</label>
		<input type="radio" name="shenqxm" id="shenqxm6" value="职等调整" onClick="changexm()" <% If shenqxm="职等调整" Then Response.Write("checked") %>><label for="shenqxm6">职等调整</label>
		<input type="radio" name="shenqxm" id="shenqxm7" value="工龄恢复" onClick="changexm()" <% If shenqxm="工龄恢复" Then Response.Write("checked") %>><label for="shenqxm7">工龄恢复</label>
		</td>
      </tr>
   </table>
<div id="shenqxm1div" style="display:
<% 
If shenqxm="工资调薪" Then 
Response.Write("block;")
Else 
Response.Write("none;")
end if
%>  
"> 
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews2" >
      <tr>
        <td height="20" align="left">原薪资：</td>
        <td>
		<input name="FormerSalary" type="text" class="textfield" id="FormerSalary" style="WIDTH: 140;" value="<%= FormerSalary %>" maxlength="100" onBlur="return mycheckNum(this)"></td>
        <td width="120" height="20" align="left">拟调薪资：</td>
        <td>
		<input name="CWantSalary" type="text" class="textfield" id="CWantSalary" style="WIDTH: 140;" value="<%= CWantSalary %>" maxlength="100" onBlur="return mycheckNum(this)"></td>
        <td width="120" height="20" align="left">调后薪资：</td>
        <td>
		<input name="WantSalary" type="text" class="textfield" id="WantSalary" style="WIDTH: 140;" value="<%= WantSalary %>" readonly="readonly" onBlur="return checkNum(this)"></td>
      </tr>
      <tr>
        <td width="120" height="20" align="left">生效日期：</td>
        <td>
		<input name="EffectiveDate" type="text" class="textfield" id="EffectiveDate" style="WIDTH: 140;" <%if Instr(checkType,"reply")=0 then response.Write("readonly")%> value="<%= EffectiveDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td height="20" align="left">申请理由：</td>
        <td colspan="3">
		<input name="ApplicReasons" type="text" class="textfield" id="ApplicReasons" value="<%= ApplicReasons %>" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left">相关事项说明：</td>
        <td colspan="5">
		<input name="Explain" type="text" class="textfield" id="Explain" value="<%= Explain %>" maxlength="500" ></td>
      </tr>
	  </table>
</div>
<div id="shenqxm2div" style="display:
<% 
If shenqxm="话费补贴" Then 
Response.Write("block;")
Else 
Response.Write("none;")
end if
%>  
"> 
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews2" >
      <tr>
        <td height="20" align="left">已开虚拟网：</td>
        <td>
		<select name="VirtualNetFlag" id="VirtualNetFlag">
		<option value="是" <% If VirtualNetFlag="是" Then Response.Write("selected") %>>是</option>
		<option value="否" <% If VirtualNetFlag="否" Then Response.Write("selected") %>>否</option>
		</select></td>
        <td width="120" height="20" align="left">担任本职日期：</td>
        <td>
		<input name="PositionDate" type="text" class="textfield" id="PositionDate" style="WIDTH: 140;" value="<%= PositionDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td width="120" height="20" align="left" onMouseOut="$('#hfbz').hide('show');" onMouseOver="$('#hfbz').show('show');$('#hfbz').animate({width: '75%',opacity: 0.8,top:10,left:6})">话费补贴标准：</td>
        <td>
		<select name="FeeStandards" id="FeeStandards">
		<option value="5" <% If FeeStandards="5" Then Response.Write("selected") %>>5</option>
		<option value="10" <% If FeeStandards="10" Then Response.Write("selected") %>>10</option>
		<option value="15" <% If FeeStandards="15" Then Response.Write("selected") %>>15</option>
		<option value="25" <% If FeeStandards="25" Then Response.Write("selected") %>>25</option>
		<option value="30" <% If FeeStandards="30" Then Response.Write("selected") %>>30</option>
		<option value="50" <% If FeeStandards="50" Then Response.Write("selected") %>>50</option>
		<option value="80" <% If FeeStandards="80" Then Response.Write("selected") %>>80</option>
		<option value="100" <% If FeeStandards="100" Then Response.Write("selected") %>>100</option>
		<option value="150" <% If FeeStandards="150" Then Response.Write("selected") %>>150</option>
		<option value="200" <% If FeeStandards="200" Then Response.Write("selected") %>>200</option>
		<option value="250" <% If FeeStandards="250" Then Response.Write("selected") %>>250</option>
		</select></td>
      </tr>
      <tr>
        <td height="20" align="left">特殊话费补贴：</td>
        <td >
		<input name="SpecialFee" type="text" class="textfield" id="SpecialFee" style="WIDTH: 140;" value="<%= SpecialFee %>" maxlength="100" onBlur="return checkNum(this)"></td>
        <td width="120" height="20" align="left">生效日期：</td>
        <td>
		<input name="EffectiveDate2" type="text" class="textfield" id="EffectiveDate2" style="WIDTH: 140;" <%if Instr(checkType,"reply")=0 then response.Write("readonly")%> value="<%= EffectiveDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left">相关事项说明：</td>
        <td colspan="5">
		<input name="Explain2" type="text" class="textfield" id="Explain2" style="WIDTH: 540;" value="<%= Explain %>" maxlength="500" ></td>
      </tr>
	  </table>
</div>
<div id="shenqxm3div" style="display:
<% 
If shenqxm="住房补贴" Then 
Response.Write("block;")
Else 
Response.Write("none;")
end if
%>  
">  
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews2" >
      <tr>
<!--        <td height="20" align="left">是否是夫妻：</td>
        <td>
		<select name="HusbandWifeFlag" id="HusbandWifeFlag">
		<option value="是" <% If HusbandWifeFlag="是" Then Response.Write("selected") %>>是</option>
		<option value="否" <% If HusbandWifeFlag="否" Then Response.Write("selected") %>>否</option>
		</select></td>-->
        <td width="120" height="20" align="left">入司日期：</td>
        <td>
		<input name="IntoCompanyDate" type="text" class="textfield" id="IntoCompanyDate" style="WIDTH: 140;" value="<%= IntoCompanyDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td width="120" height="20" align="left">职等：</td>
        <td>
		<input name="Employment" type="text" class="textfield" id="Employment" style="WIDTH: 140;" value="<%= Employment %>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">住房补贴金额：</td>
        <td >
		<input name="HousingFee" type="text" class="textfield" id="HousingFee" style="WIDTH: 140;" value="<%= HousingFee %>" maxlength="100" onBlur="return actcheckNum(this)"></td>
      </tr>
      <tr>
        <td width="120" height="20" align="left">配偶员工编号：</td>
        <td>
		<input name="PartID" type="text" class="textfield" id="PartID" style="WIDTH: 140;" value="<%= PartID %>" maxlength="100" onBlur="return checkEmp2(this)"></td>
        <td width="120" height="20" align="left">配偶姓名：</td>
        <td>
		<input name="PartName" type="text" class="textfield" id="PartName" style="WIDTH: 140;" value="<%= PartName %>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">配偶入司日期：</td>
        <td >
		<input name="PartIntoCompanyDate" type="text" class="textfield" id="PartIntoCompanyDate" style="WIDTH: 140;" value="<%= PartIntoCompanyDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
      </tr>
      <tr>
        <td width="120" height="20" align="left">配偶岗位：</td>
        <td>
		<input name="PartPosition" type="text" class="textfield" id="PartPosition" style="WIDTH: 140;" value="<%= PartPosition %>" maxlength="100" readonly="true"></td>
        <td width="120" height="20" align="left">配偶职等：</td>
        <td>
		<input name="PartEmployment" type="text" class="textfield" id="PartEmployment" style="WIDTH: 140;" value="<%= PartEmployment %>" maxlength="100" readonly="true"></td>
        <td height="20" align="left">配偶住房补贴：</td>
        <td >
		<input name="PartHousingFee" type="text" class="textfield" id="PartHousingFee" style="WIDTH: 140;" value="<%= PartHousingFee %>" maxlength="100" onBlur="return actcheckNum(this)"></td>
      </tr>
      <tr>
        <td width="120" height="20" align="left">特殊补贴申请：</td>
        <td>
		<input name="SpecialHousing" type="text" class="textfield" id="SpecialHousing" style="WIDTH: 140;" value="<%= SpecialHousing %>" maxlength="100" onBlur="return checkNum(this)"></td>
        <td width="120" height="20" align="left">合计住房补贴：</td>
        <td>
		<input name="TotalHousingFee" type="text" class="textfield" id="TotalHousingFee" style="WIDTH: 140;" value="<%= TotalHousingFee %>" maxlength="100" onBlur="return checkNum(this)"></td>
        <td width="120" height="20" align="left">生效日期：</td>
        <td>
		<input name="EffectiveDate3" type="text" class="textfield" id="EffectiveDate3" style="WIDTH: 140;" value="<%= EffectiveDate %>" maxlength="100" onBlur="return checkDate(this)" <%if Instr(checkType,"reply")=0 then response.Write("readonly")%>></td>
      </tr>
      <tr>
        <td height="20" align="left">相关事项说明：</td>
        <td colspan="5">
		<input name="Explain3" type="text" class="textfield" id="Explain3" style="WIDTH: 540;" value="<%= Explain %>" maxlength="500" ></td>
      </tr>
	  </table>
</div>
<div id="shenqxm4div" style="display:
<% 
If shenqxm="岗位补贴" Then 
Response.Write("block;")
Else 
Response.Write("none;")
end if
%>  
">  
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews2" >
	  <tr>
        <td width="120" height="20" align="left">担任本职日期：</td>
        <td>
		<input name="PositionDate4" type="text" class="textfield" id="PositionDate4" style="WIDTH: 140;" value="<%= PositionDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td width="120" height="20" align="left">申请金额：</td>
        <td>
		<input name="ApplicFee" type="text" class="textfield" id="ApplicFee" style="WIDTH: 140;" value="<%= ApplicFee %>" maxlength="100"  onBlur="return checkNum(this)"></td>
        <td width="120" height="20" align="left">生效日期：</td>
        <td>
		<input name="EffectiveDate4" type="text" class="textfield" id="EffectiveDate4" style="WIDTH: 140;" <%if Instr(checkType,"reply")=0 then response.Write("readonly")%> value="<%= EffectiveDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
      </tr>
      <tr>
        <td height="20" align="left">申请理由：</td>
        <td colspan="5">
		<input name="ApplicReasons4" type="text" class="textfield" id="ApplicReasons4" style="WIDTH: 540;" value="<%= ApplicReasons %>" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left">相关事项说明：</td>
        <td colspan="5">
		<input name="Explain4" type="text" class="textfield" id="Explain4" style="WIDTH: 540;" value="<%= Explain %>" maxlength="500" ></td>
      </tr>
	  </table>
</div>
<div id="shenqxm5div" style="display:
<% 
If shenqxm="其他补贴" Then 
Response.Write("block;")
Else 
Response.Write("none;")
end if
%>  
">   
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews2" >
	  <tr>
        <td width="120" height="20" align="left">申请金额：</td>
        <td>
		<input name="ApplicFee5" type="text" class="textfield" id="ApplicFee5" style="WIDTH: 140;" value="<%= ApplicFee %>" maxlength="100"  onBlur="return checkNum(this)"></td>
        <td width="120" height="20" align="left">生效日期：</td>
        <td>
		<input name="EffectiveDate5" type="text" class="textfield" id="EffectiveDate5" style="WIDTH: 140;" <%if Instr(checkType,"reply")=0 then response.Write("readonly")%> value="<%= EffectiveDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left">申请理由：</td>
        <td colspan="5">
		<input name="ApplicReasons5" type="text" class="textfield" id="ApplicReasons5" style="WIDTH: 540;" value="<%= ApplicReasons %>" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left">相关事项说明：</td>
        <td colspan="5">
		<input name="Explain5" type="text" class="textfield" id="Explain5" style="WIDTH: 540;" value="<%= Explain %>" maxlength="500" ></td>
      </tr>
	  </table>
</div>
<div id="shenqxm6div" style="display:
<% 
If shenqxm="职等调整" Then 
Response.Write("block;")
Else 
Response.Write("none;")
end if
%>  
">  
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews2" >
	  <tr>
        <td width="120" height="20" align="left">原职等：</td>
        <td>
		<input name="Employment6" type="text" class="textfield" id="Employment6" style="WIDTH: 140;" value="<%= Employment %>" maxlength="100" onBlur="return checkNum(this)"></td>
        <td width="120" height="20" align="left">原职位：</td>
        <td>
		<input name="OldPosition" type="text" class="textfield" id="OldPosition" style="WIDTH: 140;" value="<%= OldPosition %>" maxlength="100"></td>
        <td width="120" height="20" align="left">新职位：</td>
        <td>
		<input name="NewPosition" type="text" class="textfield" id="NewPosition" style="WIDTH: 140;" value="<%= NewPosition %>" maxlength="100"></td>
      </tr>
	  <tr>
        <td width="120" height="20" align="left">申请职等：</td>
        <td>
		<input name="ApplicEmployment" type="text" class="textfield" id="ApplicEmployment" style="WIDTH: 140;" value="<%= ApplicEmployment %>" maxlength="100" onBlur="return checkNum(this)"></td>
        <td width="120" height="20" align="left">入司日期：</td>
        <td>
		<input name="IntoCompanyDate6" type="text" class="textfield" id="IntoCompanyDate6" style="WIDTH: 140;" value="<%= IntoCompanyDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td width="120" height="20" align="left">生效日期：</td>
        <td>
		<input name="EffectiveDate6" type="text" class="textfield" id="EffectiveDate6" style="WIDTH: 140;" <%if Instr(checkType,"reply")=0 then response.Write("readonly")%> value="<%= EffectiveDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
      </tr>
      <tr>
        <td height="20" align="left">申请理由：</td>
        <td colspan="5">
		<input name="ApplicReasons6" type="text" class="textfield" id="ApplicReasons6" style="WIDTH: 540;" value="<%= ApplicReasons %>" maxlength="300" ></td>
      </tr>
      <tr>
        <td height="20" align="left">相关事项说明：</td>
        <td colspan="5">
		<input name="Explain6" type="text" class="textfield" id="Explain6" style="WIDTH: 540;" value="<%= Explain %>" maxlength="500" ></td>
      </tr>
	  </table>
</div>
<div id="shenqxm7div" style="display:
<% 
If shenqxm="工龄恢复" Then 
Response.Write("block;")
Else 
Response.Write("none;")
end if
%>  
">    
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews2" >
	    <tr>
        <td width="120" height="20" align="left">上次入司日期：</td>
        <td>
		<input name="OldIntoDate" type="text" class="textfield" id="OldIntoDate" style="WIDTH: 140;" value="<%= OldIntoDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td width="120" height="20" align="left">上次离职日期：</td>
        <td>
		<input name="OldOutDate" type="text" class="textfield" id="OldOutDate" style="WIDTH: 140;" value="<%= OldOutDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td width="120" height="20" align="left">离职性质：</td>
        <td>
		<select name="DueType" id="DueType">
		<option value="辞职" <% If DueType="辞职" Then Response.Write("selected")%>>辞职</option>
		<option value="自离" <% If DueType="自离" Then Response.Write("selected")%>>自离</option>
		<option value="辞退" <% If DueType="辞退" Then Response.Write("selected")%>>辞退</option>
		<option value="开除" <% If DueType="开除" Then Response.Write("selected")%>>开除</option>
		</select></td>
      </tr>
      <tr>
        <td height="20" align="left">历史工龄：</td>
        <td>
		<input name="Oldlength" type="text" class="textfield" id="Oldlength" style="WIDTH: 140;" value="<%= Oldlength %>" maxlength="100" ></td>
        <td height="20" align="left">本次入职日期：</td>
        <td>
		<input name="IntoCompanyDate7" type="text" class="textfield" id="IntoCompanyDate7" style="WIDTH: 140;" value="<%= IntoCompanyDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
        <td height="20" align="left">生效日期：</td>
        <td>
		<input name="EffectiveDate7" type="text" class="textfield" id="EffectiveDate7" style="WIDTH: 140;" <%if Instr(checkType,"reply")=0 then response.Write("readonly")%> value="<%= EffectiveDate %>" maxlength="100" onBlur="return checkDate(this)"></td>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">相关事项说明：</td>
        <td colspan="5">
		<input name="Explain7" type="text" class="textfield" id="Explain7" style="WIDTH: 540;" value="<%= Explain %>" maxlength="500" ></td>
      </tr>
	  </table>
</div>
<div id="Buttondiv" style="display:<%= style2 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews3" bgcolor="#99BBE8">
  <tr>  <td height="5">  </td>
  </tr>
	<tr>
	  <td align="center">
	  <input type="hidden" name="CheckFlag" id="CheckFlag" value="<%= CheckFlag %>">
	  <input type="hidden" name="detailType" id="detailType" value="<%= detailType %>">
			<input type="button" class="button"  value="关闭" style="WIDTH: 80;"  onClick="closead1()">&nbsp;
			<input name="submitSaveAdd" type="button" class="button"  id="submitSaveAdd" value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">&nbsp;
			<input name="submitDelete" type="button" class="button"  id="submitDelete" value="删除" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('Delete');toSubmit(this);">
	  </td>
	</tr>
  <tr>  <td height="5">  </td>
  </table>
</div>
<div id="ReplyDiv" style="display:<%= style1 %>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id="editNews3" bgcolor="#99BBE8">
  <tr>  <td height="5" colspan="4">  </td>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td width="60"> 审核人 </td>
 <td width="60">
 <input name="Replyer" type="text" id="Replyer" value="<%= Replyer %>"></td>
 <td width="60"> 审核日期 </td>
 <td width="60">
 <input name="ReplyDate" type="text" id="ReplyDate" value="<%= ReplyDate %>"></td>
</tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 审核意见 </td>
<td colspan="3">
  <textarea name="ReplyText" id="ReplyText" style="width:500px; height:100px; "><%= ReplyText %></textarea>
</td>
</tr> 
  <tr>  <td height="5" colspan="4">  </td>
  </tr>
	<tr>
	  <td align="center" colspan="4">
	  <input type="hidden" name="checkType" id="checkType" value="<%= checkType %>">
	  <input type="hidden" name="checkValue" id="checkValue" value="">
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="同意" style="WIDTH: 80; visibility:<%= style3 %>;" title="同意时请输入同意内容！"  onClick="javascript:$('#checkValue').val(this.value);toSubmit4Check(this);">&nbsp;
			<input name="submitCheckno" type="button" class="button"  id="submitCheckno" value="不同意" style="WIDTH: 80; visibility: <%= style3 %>;" title="不同意完直接作废，请同时输入不同意原因！" onClick="javascript:$('#checkValue').val(this.value);toSubmit4Check(this);">&nbsp;
			<input name="submitunCheck" type="button" class="button"  id="submitunCheck" value="反审核" style="WIDTH: 80; visibility: <%= style3 %>;" title="只能对已经同意的单据进行反审核" onClick="javascript:$('#checkValue').val(this.value);toSubmit4Check(this);">&nbsp;	  </td>
	</tr>
  <tr>  <td height="5">  </td>
  </table>
</div>
</td>
  </tr>
</table>
</form>
</div>
<%
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="Add" then
			if Request("Department")<>session("Depart") and left(session("Depart"),9)<>"KD01.0004" and session("Depart")<>"KD01.0001.0012" then
				response.Write("只能登记本部门的薪资福利调整单！")
				response.End()
			end if
    shenqxm=request("shenqxm")
			SerialNum=getBillNo("Bill_WelfareAdjust",3,date())
	set rs = server.createobject("adodb.recordset")
	sql="select * from Bill_WelfareAdjust"
	rs.open sql,connzxpt,1,3
	rs.addnew
	rs("SerialNum")=SerialNum
	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("Position")=Request("Position")
	rs("shenqxm")=Request("shenqxm")
	rs("Biller")=UserName
	rs("BillDate")=now()
	if shenqxm="工资调薪" then
	  rs("FormerSalary")=Request("FormerSalary")
	  rs("CWantSalary")=Request("CWantSalary")
	  rs("WantSalary")=Request("WantSalary")
	  rs("EffectiveDate")=Request("EffectiveDate")
	  rs("ApplicReasons")=Request("ApplicReasons")
	  rs("Explain")=Request("Explain")
	elseif shenqxm="话费补贴" then
	  rs("VirtualNetFlag")=Request("VirtualNetFlag")
	  rs("PositionDate")=Request("PositionDate")
	  rs("FeeStandards")=Request("FeeStandards")
	  if Request("SpecialFee")<>"" then rs("SpecialFee")=Request("SpecialFee")
	  rs("EffectiveDate")=Request("EffectiveDate2")
	  rs("Explain")=Request("Explain2")
	elseif shenqxm="住房补贴" then
	  rs("HusbandWifeFlag")=Request("HusbandWifeFlag")
	  rs("IntoCompanyDate")=Request("IntoCompanyDate")
	  rs("Employment")=Request("Employment")
	  rs("HousingFee")=Request("HousingFee")
	  rs("SpecialHousing")=Request("SpecialHousing")
	  rs("EffectiveDate")=Request("EffectiveDate3")
	  rs("Explain")=Request("Explain3")
	  rs("PartID")=request("PartID")
	  rs("PartName")=request("PartName")
	  rs("PartIntoCompanyDate")=request("PartIntoCompanyDate")
	  rs("PartEmployment")=request("PartEmployment")
	  rs("PartHousingFee")=request("PartHousingFee")
	  rs("PartSpecialHousing")=request("PartSpecialHousing")
	  rs("TotalHousingFee")=request("TotalHousingFee")
	  rs("PartPosition")=request("PartPosition")
	elseif shenqxm="岗位补贴" then
	  rs("PositionDate")=Request("PositionDate4")
	  rs("ApplicFee")=Request("ApplicFee")
	  rs("ApplicReasons")=Request("ApplicReasons4")
	  rs("EffectiveDate")=Request("EffectiveDate4")
	  rs("Explain")=Request("Explain4")
	elseif shenqxm="其他补贴" then
	  rs("ApplicFee")=Request("ApplicFee5")
	  rs("ApplicReasons")=Request("ApplicReasons5")
	  rs("EffectiveDate")=Request("EffectiveDate5")
	  rs("Explain")=Request("Explain5")
	elseif shenqxm="职等调整" then
	  rs("ApplicEmployment")=Request("ApplicEmployment")
	  rs("IntoCompanyDate")=Request("IntoCompanyDate6")
	  rs("ApplicReasons")=Request("ApplicReasons6")
	  rs("OldPosition")=Request("OldPosition")
	  rs("Employment")=Request("Employment6")
	  rs("NewPosition")=Request("NewPosition")
	  rs("EffectiveDate")=Request("EffectiveDate6")
	  rs("Explain")=Request("Explain6")
	elseif shenqxm="工龄恢复" then
	  rs("OldIntoDate")=Request("OldIntoDate")
	  rs("OldOutDate")=Request("OldOutDate")
	  rs("DueType")=Request("DueType")
	  rs("Oldlength")=Request("Oldlength")
	  rs("IntoCompanyDate")=Request("IntoCompanyDate7")
	  rs("EffectiveDate")=Request("EffectiveDate7")
	  rs("Explain")=Request("Explain7")
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Edit" then
    shenqxm=request("shenqxm")
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_WelfareAdjust where CheckFlag<1 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Request("Department")<>session("Depart") and left(session("Depart"),9)<>"KD01.0004" then
		response.Write("只能登记本部门的薪资福利调整单！")
		response.End()
	end if

	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("Position")=Request("Position")
	rs("shenqxm")=Request("shenqxm")
	if shenqxm="工资调薪" then
	  rs("FormerSalary")=Request("FormerSalary")
	  rs("CWantSalary")=Request("CWantSalary")
	  rs("WantSalary")=Request("WantSalary")
	  rs("EffectiveDate")=Request("EffectiveDate")
	  rs("ApplicReasons")=Request("ApplicReasons")
	  rs("Explain")=Request("Explain")
	elseif shenqxm="话费补贴" then
	  rs("VirtualNetFlag")=Request("VirtualNetFlag")
	  rs("PositionDate")=Request("PositionDate")
	  rs("FeeStandards")=Request("FeeStandards")
	  rs("SpecialFee")=Request("SpecialFee")
	  rs("EffectiveDate")=Request("EffectiveDate2")
	  rs("Explain")=Request("Explain2")
	elseif shenqxm="住房补贴" then
	  rs("HusbandWifeFlag")=Request("HusbandWifeFlag")
	  rs("IntoCompanyDate")=Request("IntoCompanyDate")
	  rs("Employment")=Request("Employment")
	  rs("HousingFee")=Request("HousingFee")
	  rs("SpecialHousing")=Request("SpecialHousing")
	  rs("EffectiveDate")=Request("EffectiveDate3")
	  rs("Explain")=Request("Explain3")
	  
	  rs("PartID")=request("PartID")
	  rs("PartName")=request("PartName")
	  rs("PartIntoCompanyDate")=request("PartIntoCompanyDate")
	  rs("PartEmployment")=request("PartEmployment")
	  rs("PartHousingFee")=request("PartHousingFee")
	  rs("PartPosition")=request("PartPosition")
	  rs("PartSpecialHousing")=request("PartSpecialHousing")
	  rs("TotalHousingFee")=request("TotalHousingFee")
	elseif shenqxm="岗位补贴" then
	  rs("PositionDate")=Request("PositionDate4")
	  rs("ApplicFee")=Request("ApplicFee")
	  rs("ApplicReasons")=Request("ApplicReasons4")
	  rs("EffectiveDate")=Request("EffectiveDate4")
	  rs("Explain")=Request("Explain4")
	elseif shenqxm="其他补贴" then
	  rs("ApplicFee")=Request("ApplicFee5")
	  rs("ApplicReasons")=Request("ApplicReasons5")
	  rs("EffectiveDate")=Request("EffectiveDate5")
	  rs("Explain")=Request("Explain5")
	elseif shenqxm="职等调整" then
	  rs("ApplicEmployment")=Request("ApplicEmployment")
	  rs("IntoCompanyDate")=Request("IntoCompanyDate6")
	  rs("ApplicReasons")=Request("ApplicReasons6")
	  rs("OldPosition")=Request("OldPosition")
	  rs("NewPosition")=Request("NewPosition")
	  rs("Employment")=Request("Employment6")
	  rs("EffectiveDate")=Request("EffectiveDate6")
	  rs("Explain")=Request("Explain6")
	elseif shenqxm="工龄恢复" then
	  rs("OldIntoDate")=Request("OldIntoDate")
	  rs("OldOutDate")=Request("OldOutDate")
	  rs("DueType")=Request("DueType")
	  rs("Oldlength")=Request("Oldlength")
	  rs("IntoCompanyDate")=Request("IntoCompanyDate7")
	  rs("EffectiveDate")=Request("EffectiveDate7")
	  rs("Explain")=Request("Explain7")
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_WelfareAdjust where CheckFlag<1 and (Register='"&UserName&"' or Biller='"&UserName&"') and SerialNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_WelfareAdjust where SerialNum="&SerialNum)
	response.write "###"
  end if
elseif showType="CheckProcess" then 
  checkType=request("checkType")
  SerialNum=request("SerialNum")
  dim checkValue:checkValue=request("checkValue")
  if checkType="a1reply" and Instr(session("AdminPurviewFLW"),"|201.2,")>0 then
    set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	Depart=session("Depart")
	if UserName="A09837" then Depart="KD01.0005.0004"
	if left(Depart,9)="KD01.0004" then Depart="KD01.0004%"
	sql="select * from Bill_WelfareAdjust where CheckFlag<=1 and Department like '"&Depart&"' and SerialNum="&SerialNum
	if Depart="KD01.0001.0012" then
	sql="select * from Bill_WelfareAdjust where CheckFlag<=1 and (Department = '"&Depart&"' or Department = 'KD01.0001.0005') and SerialNum="&SerialNum
	end if
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if rs("shenqxm")="工资调薪" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("CWantSalary")=Request("CWantSalary")
	  rs("WantSalary")=Request("WantSalary")
	  rs("EffectiveDate")=Request("EffectiveDate")
	elseif rs("shenqxm")="话费补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate2"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate2")
	elseif rs("shenqxm")="住房补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate3"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate3")
	elseif rs("shenqxm")="岗位补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate4"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate4")
	elseif rs("shenqxm")="其他补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate5"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate5")
	elseif rs("shenqxm")="职等调整" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate6"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate6")
	elseif rs("shenqxm")="工龄恢复" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate7"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate7")
	end if
	rs("DepartReplyer")=AdminName
	rs("DepartReplyDate")=now()
	rs("DepartReplyText")=request("ReplyText")
	rs("DepartFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=1
	elseif checkValue="不同意" then
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	elseif checkValue="反审核" then
	rs("CheckFlag")=0
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif checkType="a2reply" and Instr(session("AdminPurviewFLW"),"|201.3,")>0 then
    SerialNum=request("SerialNum")
	sql="select * from Bill_WelfareAdjust where (CheckFlag=1 or  CheckFlag=2)  and SerialNum="&SerialNum
    set rs = server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if rs("shenqxm")="工资调薪" and Instr(session("AdminPurviewFLW"),"|201.7,")>0 then
		if datediff("m",rs("RegDate"),Request("EffectiveDate"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("CWantSalary")=Request("CWantSalary")
	  rs("WantSalary")=Request("WantSalary")
	  rs("EffectiveDate")=Request("EffectiveDate")
	rs("RelatedReplyer")=AdminName
	rs("RelatedReplyDate")=now()
	rs("RelatedReplyText")=request("ReplyText")
	rs("RelatedFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=2
	elseif checkValue="反审核" then
	rs("CheckFlag")=1
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	elseif rs("shenqxm")="话费补贴" and Instr(session("AdminPurviewFLW"),"|201.8,")>0 then
		if datediff("m",rs("RegDate"),Request("EffectiveDate2"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate2")
	rs("RelatedReplyer")=AdminName
	rs("RelatedReplyDate")=now()
	rs("RelatedReplyText")=request("ReplyText")
	rs("RelatedFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=2
	elseif checkValue="反审核" then
	rs("CheckFlag")=1
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	elseif rs("shenqxm")="住房补贴" and Instr(session("AdminPurviewFLW"),"|201.9,")>0 then
		if datediff("m",rs("RegDate"),Request("EffectiveDate3"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate3")
	rs("RelatedReplyer")=AdminName
	rs("RelatedReplyDate")=now()
	rs("RelatedReplyText")=request("ReplyText")
	rs("RelatedFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=2
	elseif checkValue="反审核" then
	rs("CheckFlag")=1
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	elseif rs("shenqxm")="岗位补贴" and Instr(session("AdminPurviewFLW"),"|201.10,")>0 then
		if datediff("m",rs("RegDate"),Request("EffectiveDate4"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate4")
	rs("RelatedReplyer")=AdminName
	rs("RelatedReplyDate")=now()
	rs("RelatedReplyText")=request("ReplyText")
	rs("RelatedFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=2
	elseif checkValue="反审核" then
	rs("CheckFlag")=1
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	elseif rs("shenqxm")="其他补贴" and Instr(session("AdminPurviewFLW"),"|201.11,")>0 then
		if datediff("m",rs("RegDate"),Request("EffectiveDate5"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate5")
	rs("RelatedReplyer")=AdminName
	rs("RelatedReplyDate")=now()
	rs("RelatedReplyText")=request("ReplyText")
	rs("RelatedFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=2
	elseif checkValue="反审核" then
	rs("CheckFlag")=1
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	elseif rs("shenqxm")="职等调整" and Instr(session("AdminPurviewFLW"),"|201.12,")>0 then
		if datediff("m",rs("RegDate"),Request("EffectiveDate6"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate6")
	rs("RelatedReplyer")=AdminName
	rs("RelatedReplyDate")=now()
	rs("RelatedReplyText")=request("ReplyText")
	rs("RelatedFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=2
	elseif checkValue="反审核" then
	rs("CheckFlag")=1
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	elseif rs("shenqxm")="工龄恢复" and Instr(session("AdminPurviewFLW"),"|201.13,")>0 then
		if datediff("m",rs("RegDate"),Request("EffectiveDate7"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate7")
	rs("RelatedReplyer")=AdminName
	rs("RelatedReplyDate")=now()
	rs("RelatedReplyText")=request("ReplyText")
	rs("RelatedFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=2
	elseif checkValue="反审核" then
	rs("CheckFlag")=1
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif checkType="a3reply" and Instr(session("AdminPurviewFLW"),"|201.4,")>0 then
    SerialNum=request("SerialNum")
	sql="select * from Bill_WelfareAdjust where (CheckFlag=3 or CheckFlag=2 ) and SerialNum="&SerialNum'or (CheckFlag=0 and left(Department,9)='KD01.0004')
    set rs = server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if rs("shenqxm")="工资调薪" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("CWantSalary")=Request("CWantSalary")
	  rs("WantSalary")=Request("WantSalary")
	  rs("EffectiveDate")=Request("EffectiveDate")
	elseif rs("shenqxm")="话费补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate2"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate2")
	elseif rs("shenqxm")="住房补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate3"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate3")
	elseif rs("shenqxm")="岗位补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate4"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate4")
	elseif rs("shenqxm")="其他补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate5"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate5")
	elseif rs("shenqxm")="职等调整" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate6"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate6")
	elseif rs("shenqxm")="工龄恢复" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate7"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate7")
	end if
	rs("DirectorReplyer")=AdminName
	rs("DirectorReplyDate")=now()
	rs("DirectorReplyText")=request("ReplyText")
	rs("DirectorFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=3
	elseif checkValue="反审核" then
	rs("CheckFlag")=2
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif checkType="a4reply" and Instr(session("AdminPurviewFLW"),"|201.5,")>0 then
    SerialNum=request("SerialNum")
	sql="select * from Bill_WelfareAdjust where CheckFlag<=4 and CheckFlag>=2 and SerialNum="&SerialNum
    set rs = server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("@@@数据库读取记录出错！@@@")
		response.end
	end if
	if rs("shenqxm")="工资调薪" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("CWantSalary")=Request("CWantSalary")
	  rs("WantSalary")=Request("WantSalary")
	  rs("EffectiveDate")=Request("EffectiveDate")
	elseif rs("shenqxm")="话费补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate2"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate2")
	elseif rs("shenqxm")="住房补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate3"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate3")
	elseif rs("shenqxm")="岗位补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate4"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate4")
	elseif rs("shenqxm")="其他补贴" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate5"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate5")
	elseif rs("shenqxm")="职等调整" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate6"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate6")
	elseif rs("shenqxm")="工龄恢复" then
		if datediff("m",rs("RegDate"),Request("EffectiveDate7"))<0 then
			response.Write("生效日期不能早于申请日期所在月份！")
			response.End()
		end if
	  rs("EffectiveDate")=Request("EffectiveDate7")
	end if
	rs("CEOReplyer")=AdminName
	rs("CEOReplyDate")=now()
	rs("CEOReplyText")=request("ReplyText")
	rs("CEOFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=4
	elseif checkValue="反审核" then
	rs("CheckFlag")=3
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif checkType="a5reply" and Instr(session("AdminPurviewFLW"),"|201.6,")>0 then
    SerialNum=request("SerialNum")
	sql="select * from Bill_WelfareAdjust where CheckFlag<=99 and CheckFlag>=2 and SerialNum="&SerialNum
    set rs = server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("@@@数据库读取记录出错！@@@")
		response.end
	end if
	if rs("shenqxm")="工资调薪" then
	  if rs("CheckFlag")<3 then
		response.write ("@@@需副总以上审核才能执行！@@@")
		response.end
	  end if
'		if datediff("m",rs("RegDate"),Request("EffectiveDate"))<0 then
'			response.Write("生效日期不能早于申请日期所在月份！")
'			response.End()
'		end if
	  rs("CWantSalary")=Request("CWantSalary")
	  rs("WantSalary")=Request("WantSalary")
	  rs("EffectiveDate")=Request("EffectiveDate")
	elseif rs("shenqxm")="话费补贴" then
'		if datediff("m",rs("RegDate"),Request("EffectiveDate2"))<0 then
'			response.Write("生效日期不能早于申请日期所在月份！")
'			response.End()
'		end if
	  rs("EffectiveDate")=Request("EffectiveDate2")
	elseif rs("shenqxm")="住房补贴" then
'		if datediff("m",rs("RegDate"),Request("EffectiveDate3"))<0 then
'			response.Write("生效日期不能早于申请日期所在月份！")
'			response.End()
'		end if
	  rs("EffectiveDate")=Request("EffectiveDate3")
	elseif rs("shenqxm")="岗位补贴" then
'		if datediff("m",rs("RegDate"),Request("EffectiveDate4"))<0 then
'			response.Write("生效日期不能早于申请日期所在月份！")
'			response.End()
'		end if
	  rs("EffectiveDate")=Request("EffectiveDate4")
	elseif rs("shenqxm")="其他补贴" then
	  if rs("CheckFlag")<3 then
		response.write ("@@@需副总以上审核才能执行！@@@")
		response.end
	  end if
'		if datediff("m",rs("RegDate"),Request("EffectiveDate5"))<0 then
'			response.Write("生效日期不能早于申请日期所在月份！")
'			response.End()
'		end if
	  rs("EffectiveDate")=Request("EffectiveDate5")
	elseif rs("shenqxm")="职等调整" then
	  if rs("CheckFlag")<3 and rs("ApplicEmployment")>4 then
		response.write ("@@@需副总以上审核才能执行！@@@")
		response.end
	  end if
'		if datediff("m",rs("RegDate"),Request("EffectiveDate6"))<0 then
'			response.Write("生效日期不能早于申请日期所在月份！")
'			response.End()
'		end if
	  rs("EffectiveDate")=Request("EffectiveDate6")
	elseif rs("shenqxm")="工龄恢复" then
'		if datediff("m",rs("RegDate"),Request("EffectiveDate7"))<0 then
'			response.Write("生效日期不能早于申请日期所在月份！")
'			response.End()
'		end if
	  rs("EffectiveDate")=Request("EffectiveDate7")
	end if
	rs("ImplementReplyer")=AdminName
	rs("ImplementReplyDate")=now()
	rs("ImplementReplyText")=request("ReplyText")
	rs("ImplementFlag")=checkValue
	if checkValue="同意" then
	rs("CheckFlag")=99
	elseif checkValue="反审核" then
	rs("CheckFlag")=4
	else
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif checkType="CheckNo" then
    SerialNum=request("SerialNum")
	sql="select * from Bill_WelfareAdjust where CheckFlag<=99 and SerialNum="&SerialNum
    set rs = server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs("CheckFlag")=100
	rs("Canceler")=AdminName
	rs("CancelDate")=now()
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  else
	response.write ("你没有权限！")
	response.end
  end if
elseif showType="getInfo" then 
  dim EmpID
  detailType=request("detailType")
  if detailType="Emp1" then
    EmpID=request("EmpID")
	sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&EmpID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
	response.write("###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###"&rs("工作岗位")&"###"&rs("到职日")&"###"&rs("职等")&"###")
	rs.close
	set rs=nothing 
  end if
end if
 %>
</body>
</html>
