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
'if Instr(session("AdminPurviewFLW"),"|203,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
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
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>申请部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>维修部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>请修类型</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>设备名称</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>问题描述</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>维修进度</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>计划完成时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>处理时间</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>处理情况</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>结果反馈</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>反馈时间</strong></font></td>
  </tr>
 <%
  dim Depart,wherestr,seachword,flag4search,type4search
  wherestr=""
  seachword=request("seachword")
  if seachword<>"" then
  wherestr=" and (SerialNum like '%"&seachword&"%' or DeviceName like '%"&seachword&"%')"
  end if
  flag4search=request("flag4search")
  if flag4search<>"" then wherestr=wherestr&" and CheckFlag="&flag4search
  type4search=request("type4search")
  if type4search<>"" then wherestr=wherestr&" and RepairType='"&type4search&"'"
  dim page'页码
      page=clng(request("Page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Bill_RepairApplication "
  dim datawhere'数据条件
		 datawhere="where 1=1"&wherestr&Session("AllMessage7")&Session("AllMessage8")&Session("AllMessage9")
		 Session("AllMessage7")=""
		 Session("AllMessage8")=""
		 Session("AllMessage9")=""
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by SerialNum desc"
  dim i'用于循环的整数
  dim rs,sql'sql语句
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
    dim sql2,rs2,hejikoufen
	hejikoufen=0
	dim formdata(3),bgcolors
    sql="select *,left(SituationDescrib,10) as a1,left(RReplyText,10) as a2,left(AReplyText,10) as a3 from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
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
		elseif rs("CheckFlag")="-1" then
		  formdata(0)="被驳回"
		  bgcolors="#B9BBC7"'绿色
		else
		  formdata(0)="未审核"
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("SerialNum")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RegisterName")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("Departmentname")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ReceivDepartment")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RepairType")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("DeviceName")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a1")&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&formdata(0)&"</td>"
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("ReplyFinishDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("RReplyDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a2")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("a3")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return showpadd('Edit',"&rs("SerialNum")&")"">"&rs("AReplyDate")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
  else
    response.write "<tr><td height='50' align='center' colspan='13' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
	response.Write("###"&pagec&"###"&idcount&"###"&datawhere)
elseif showType="AddEditShow" then 
  dim detailType
  detailType=request("detailType")
'数据处理
  dim SerialNum,Register,RegisterName,RegDate,Department,Departmentname,DeviceName,SituationDescrib,ReceivDepartment
  dim ChangedFlag,ChangedParts,SendFlag,ForeignSend,ScrapFlag,RepairType
  dim style1,style2,style3,RReplyer,RReplyText,RReplyDate,AReplyer,AReplyText,AReplyDate
  dim CheckFlag,ReplyFinishDate
  dim Checker,CheckDate,ReceivDate,Receiver,buttonval
  if detailType="AddNew" then
	sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&UserName&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
	Register=UserName
	RegisterName=AdminName
	RegDate=now()
	Department=rs("部门别")
	Departmentname=rs("部门名称")
	style1="block;"
	style2="none;"
  elseif detailType="Edit" or detailType="Check" then
    SerialNum=request("SerialNum")
	sql="select * from Bill_RepairApplication where SerialNum="&SerialNum
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	Register=rs("Register")
	RegisterName=rs("RegisterName")
	RegDate=rs("RegDate")
	Department=rs("Department")
	Departmentname=rs("Departmentname")
	ReceivDepartment=rs("ReceivDepartment")
	RepairType=rs("RepairType")
	DeviceName=rs("DeviceName")
	SituationDescrib=rs("SituationDescrib")
	ChangedFlag=rs("ChangedFlag")
	ChangedParts=rs("ChangedParts")
	SendFlag=rs("SendFlag")
	ForeignSend=rs("ForeignSend")
	ScrapFlag=rs("ScrapFlag")
	CheckFlag=rs("CheckFlag")
	RReplyer=rs("RReplyer")
	RReplyText=rs("RReplyText")
	RReplyDate=rs("RReplyDate")
	AReplyer=rs("AReplyer")
	AReplyText=rs("AReplyText")
	AReplyDate=rs("AReplyDate")
	Checker=rs("Checker")
	CheckDate=rs("CheckDate")
	ReceivDate=rs("ReceivDate")
	Receiver=rs("Receiver")
	ReplyFinishDate=rs("ReplyFinishDate")
	style1="none;"
	style2="block;"
	style3="none;"
	if CheckFlag=1 and Instr(session("AdminPurviewFLW"),"|204.3,")>0 then
	ReplyFinishDate=now()
	style2="block;"
	buttonval="接收"
	style3="block;"
	elseif CheckFlag=2 then
	style2="block;"
	if Instr(session("AdminPurviewFLW"),"|204.3,")>0 then 
	style3="block;"
	buttonval="处理"
	end if
	elseif CheckFlag=3 then
	style2="block;"
	if Register=UserName or rs("Biller")=UserName then 
	style3="block;"
	buttonval="反馈"
	end if
	elseif Register=UserName or rs("Biller")=UserName or Instr(session("AdminPurviewFLW"),"|204.2,")>0 then
	style1="block;"
	end if
  end if
  %>
 <div id="AddandEditdiv" style="width:790px; height:460px; overflow:auto; ">
 <form id="AddForm" name="AddForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>维修单维修反馈处理</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews>
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
        <td><input name="Register" type="text" class="textfield" id="Register" style="WIDTH: 140;" value="<%= Register %>" maxlength="100" onBlur="getInfo('Register')"></td>
        <td height="20" align="left">申请人姓名：</td>
        <td><input name="RegisterName" type="text" class="textfield" id="RegisterName" style="WIDTH: 140;" value="<%= RegisterName %>" maxlength="100"  readonly="true"></td>
        <td height="20" align="left">申请日期：</td>
        <td><input name="RegDate" type="text" class="textfield" id="RegDate" style="WIDTH: 140;" value="<%= RegDate %>" maxlength="100"  readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">申请部门编号：</td>
        <td>
		<input name="Department" type="text" class="textfield" id="Department" style="WIDTH: 140;" value="<%= Department %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">申请部门名称：</td>
        <td>
		<input name="Departmentname" type="text" class="textfield" id="Departmentname" style="WIDTH: 140;" value="<%= Departmentname %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left">维修部门：</td>
        <td>
		<select id="ReceivDepartment" name="ReceivDepartment">
		  <option value="总经办" <% If ReceivDepartment="总经办" Then Response.Write("selected")%>>总经办</option>
		</select>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">维修类型：</td>
        <td colspan="5">
		<input type="radio" name="RepairType" id="RepairType1" value="网络维修" onClick="changexm()" <% If RepairType="网络维修" Then Response.Write("checked") %>><label for="RepairType1">网络维修</label>
		<input type="radio" name="RepairType" id="RepairType2" value="行政维修" onClick="changexm()" <% If RepairType="行政维修" Then Response.Write("checked") %>><label for="RepairType2">行政维修</label>
		<label>&nbsp;注意：维修人员可能不在电脑旁，申请完同时仍需电话通知。(<font color="#FF0000">设备、模具不在此录入</font>)</label>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">设备名称：</td>
        <td id="DeviceName_td"><input name="DeviceName" type="text" class="textfield" id="DeviceName" style="WIDTH: 140;" value="<%= DeviceName %>" maxlength="100"></td>
        <td height="20" align="left">情况描述：</td>
        <td colspan="3">
		<input name="SituationDescrib" type="text" class="textfield" id="SituationDescrib" style="WIDTH: 300;" value="<%= SituationDescrib %>" maxlength="500" >
		</td>
      </tr>
      <tr>
        <td height="20" align="left">审核状态：</td>
        <td >
		<% If CheckFlag =1 Then Response.Write("已审核")%>
		<% If CheckFlag =2 Then Response.Write("维修中")%>
		<% If CheckFlag =3 Then Response.Write("已结束")%>
		</td>
        <td height="20" align="left">审核人：</td>
        <td >		<%= Checker %>		</td>
        <td height="20" align="left">审核日期：</td>
        <td >		<%= CheckDate %>		</td>
      </tr>
      <tr>
        <td height="20" align="left"></td>
        <td >
		</td>
        <td height="20" align="left">接收人：</td>
        <td >		<%= Receiver %>		</td>
        <td height="20" align="left">接收日期：</td>
        <td >		<%= ReceivDate %>		</td>
      </tr>
	  </table>
<div id="Buttondiv" style="display:<%= style1 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="5">  </td>
  </tr>
	<tr>
	  <td align="center">
	  <input type="hidden" name="CheckFlag" id="CheckFlag" value="<%= CheckFlag %>">
	  <input type="hidden" name="detailType" id="detailType" value="<%= detailType %>">
			<input name="submitSaveAdd" type="button" class="button"  id="submitSaveAdd" value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">&nbsp;
			<input name="submitDelete" type="button" class="button"  id="submitDelete" value="删除" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('Delete');toSubmit(this);">
			<input name="submitDelete" type="button" class="button"  id="submitDelete" value="审核" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('Check');toSubmit(this);">
	  </td>
	</tr>
  <tr>  <td height="5">  </td></tr>
  </table>
</div>
<div id="Buttondiv" style="display:<%= style2 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="5" colspan="1">  </td>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td height="20" align="left">计划完成日期：</td>
 <td colspan="3">
<input name="ReplyFinishDate" type="text" class="textfield" style="width:140px " id="ReplyFinishDate" value="<%= ReplyFinishDate %>" onBlur="return checkFullTime(this)">
&nbsp;<label for="ReplyFinishDate" style=" color:#FF0000">(接收时确认计划完成日期。)</label>
 </td>
</tr>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td colspan="4">
 <input type="checkbox" name="ChangedFlag" id="ChangedFlag" value="是" <% If ChangedFlag="是" Then Response.Write("checked")%>><label for="ChangedFlag">更换零件。</label>&nbsp;
 <label for="ChangedParts">零件列表：</label><input name="ChangedParts" type="text" class="textfield" style="width:300px " id="ChangedParts" value="<%= ChangedParts %>">
 </td>
</tr>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td colspan="4">
 <input type="checkbox" name="SendFlag" id="SendFlag" value="是" <% If SendFlag="是" Then Response.Write("checked")%> ><label for="SendFlag">外送厂商，请厂商维护。</label>&nbsp;
 <label for="ForeignSend">厂商名称：</label><input name="ForeignSend" type="text" class="textfield" style="width:300px " id="ForeignSend" value="<%= ForeignSend %>">
 </td>
</tr>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td colspan="4">
 <input type="checkbox" name="ScrapFlag" id="ScrapFlag" value="是" <% If ScrapFlag="是" Then Response.Write("checked")%>><label for="ScrapFlag">无法维修，请该部门依规定流程申请报废。</label>&nbsp;
 </td>
</tr>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td width="60"> 维修人 </td>
 <td width="60">
 <input name="RReplyer" class="textfield" type="text" id="RReplyer" value="<%= RReplyer %>"></td>
 <td width="60"> 维修日期 </td>
 <td width="60">
 <input name="RReplyDate" class="textfield" type="text" id="RReplyDate" value="<%= RReplyDate %>"></td>
</tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 维修意见 </td>
<td colspan="3">
  <textarea name="RReplyText" id="RReplyText" style="width:'500px'; height:'50px'; "><%= RReplyText %></textarea>
</td>
</tr> 
  </table>
</div>
<div id="Buttondiv" style="display:<%= style2 %>">    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="5" colspan="1">  </td>
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td width="60"> 反馈人 </td>
 <td width="60">
 <input name="AReplyer" class="textfield" type="text" id="AReplyer" value="<%= AReplyer %>"></td>
 <td width="60"> 反馈日期 </td>
 <td width="60">
 <input name="AReplyDate" class="textfield" type="text" id="AReplyDate" value="<%= AReplyDate %>"></td>
</tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 反馈意见 </td>
<td colspan="3">
  <textarea name="AReplyText" id="AReplyText" style="width:'500px'; height:'50px'; "><%= AReplyText %></textarea>
</td>
</tr> 
  <tr>  <td height="5" colspan="4">  </td>
  </tr>
	<tr>
	  <td align="center" colspan="4">
			<input name="submitCheck" type="button" class="button"  id="submitCheck" value="<%= buttonval %>" style="WIDTH: 80; display:<%= style3 %>;"  onClick="javascript:$('#detailType').val(this.value);toSubmit(this);">&nbsp;
	  </td>
	</tr>
  <tr>  <td height="5">  </td></tr>
  </table>
</div>
	</td>
  </tr>
</table>
</form>
</div>
  <%
    rs.close
    set rs=nothing
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
	set rs = server.createobject("adodb.recordset")
	sql="select * from Bill_RepairApplication"
	rs.open sql,connzxpt,1,3
	rs.addnew
	rs("Biller")=UserName
	rs("BillDate")=now()
	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("ReceivDepartment")=Request("ReceivDepartment")
	rs("RepairType")=Request("RepairType")
	rs("DeviceName")=Request("DeviceName")
	rs("SituationDescrib")=Request("SituationDescrib")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Edit" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_RepairApplication where (Biller='"&UserName&"' or Register='"&UserName&"') and CheckFlag=0 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许编辑，只能编辑本人的单据！")
		response.end
	end if
	rs("Biller")=UserName
	rs("BillDate")=now()
	rs("Register")=Request("Register")
	rs("RegisterName")=Request("RegisterName")
	rs("RegDate")=Request("RegDate")
	rs("Department")=Request("Department")
	rs("Departmentname")=Request("Departmentname")
	rs("ReceivDepartment")=Request("ReceivDepartment")
	rs("RepairType")=Request("RepairType")
	rs("DeviceName")=Request("DeviceName")
	rs("SituationDescrib")=Request("SituationDescrib")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_RepairApplication where (Biller='"&UserName&"' or Register='"&UserName&"') and CheckFlag<1 and SerialNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("当前状态不允许删除，只能删除本人的单据！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_RepairApplication where SerialNum="&SerialNum)
	response.write "###"
  elseif detailType="Check" then
		Depart=session("Depart")
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_RepairApplication where CheckFlag<1 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许审核！")
		response.end
	end if
	if rs("Department")<>Depart then
		response.write ("只能审核本部门的维修申请单！")
		response.end
	end if
	rs("CheckFlag")="1"
    rs("Checker")=session("AdminName")
    rs("Checkdate")=now()
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&session("AdminName")&"###")
  elseif detailType="接收" and Instr(session("AdminPurviewFLW"),"|204.3,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_RepairApplication where CheckFlag=1 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许此操作！")
		response.end
	end if
	rs("CheckFlag")="2"
    rs("Receiver")=session("AdminName")
    rs("RReplyer")=session("AdminName")
    rs("RReplyText")=Request("RReplyText")
    rs("ReceivDate")=now()
	rs("ReplyFinishDate")=Request("ReplyFinishDate")
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&Request("RReplyText")&"###")
  elseif detailType="驳回" and Instr(session("AdminPurviewFLW"),"|204.3,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_RepairApplication where CheckFlag=1 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许此操作！")
		response.end
	end if
	rs("CheckFlag")="-1"
    rs("Receiver")=session("AdminName")
    rs("RReplyer")=session("AdminName")
    rs("RReplyText")=Request("RReplyText")
    rs("ReceivDate")=now()
	rs("ReplyFinishDate")=Request("ReplyFinishDate")
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&Request("RReplyText")&"###")
  elseif detailType="处理" and Instr(session("AdminPurviewFLW"),"|204.3,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_RepairApplication where CheckFlag=2 and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许此操作！")
		response.end
	end if
	rs("CheckFlag")="3"
    rs("RReplyer")=session("AdminName")
    rs("RReplyText")=Request("RReplyText")
    rs("ChangedFlag")=Request("ChangedFlag")
    rs("ChangedParts")=Request("ChangedParts")
    rs("SendFlag")=Request("SendFlag")
    rs("ForeignSend")=Request("ForeignSend")
    rs("ScrapFlag")=Request("ScrapFlag")
    rs("RReplyDate")=now()
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&Request("RReplyText")&"###")
  elseif detailType="反馈" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_RepairApplication where CheckFlag=3 and (Biller='"&UserName&"' or Register='"&UserName&"') and SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("当前状态不允许此操作，只能反馈本人单据！")
		response.end
	end if
	rs("CheckFlag")="2"
    rs("AReplyer")=session("AdminName")
    rs("AReplyText")=Request("AReplyText")
    rs("AReplyDate")=now()
	rs.update
	rs.close
	set rs=nothing 
	response.write("###"&Request("AReplyText")&"###")
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="Register" then
    InfoID=request("InfoID")
	sql="select a.姓名,a.部门别,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write("###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###")
	end if
	rs.close
	set rs=nothing 
  end if
end if
 %>
</body>
</html>
