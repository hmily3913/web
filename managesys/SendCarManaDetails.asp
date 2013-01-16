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
dim rs,sql
if showType="DetailsList" then 
%>
 <div id="listtable" style="width:100%; height:400px; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
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
    sql="select z_Car.SerialNum as sNum,CarID,CarSatus, "&_
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
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")"">"&rs("CarID")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")"">"&rs("Driver")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")"">"&rs("DriverPhone")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")"">"&rs("CarSatus")&"</td>"
	  temcar=rs("CarID")
	  else
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")""></td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")""></td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")""></td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")""></td>" & vbCrLf
	  end if
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")"">"&rs("CarryGoods")&"</td>"
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")"">"&rs("CarryMans")&"</td>"
      Response.Write "<td nowrap width=""320"">"
	  if rs("CarSatus")="出行" then
	    Response.Write "<a href='SendCarEdit.asp?Result=SendCar&Action=Modify&SerialNum="&rs("SerialNum")&"'>"&rs("weizhi")&"</a>"
	  else
	    Response.Write rs("weizhi")
	  end if
	  Response.Write "</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")"">"&rs("StarteDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")"">"&rs("returntime")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return toAdd('Edit',"&rs("sNum")&")"">"&rs("nextstime")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  iii=iii+1
	  rs.movenext
    wend
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
elseif showType="AddEditShow" then 
  dim detailType
  detailType=request("detailType")
'数据处理
  dim SerialNum,CarID,CarSatus,Driver,DriverPhone,CarryMans,CarryGoods,mileageNum,DriverID
  if detailType="AddNew" then
    CarSatus="空闲"
	CarID="浙C"
	CarryMans=0
	CarryGoods=0
	mileageNum=0
  elseif detailType="Edit" then
    SerialNum=request("SerialNum")
	sql="select * from z_Car where SerialNum="&SerialNum
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,0,1
	CarID=rs("CarID")
	CarSatus=rs("CarSatus")
	Driver=rs("Driver")
	DriverPhone=rs("DriverPhone")
	CarryMans=rs("CarryMans")
	CarryGoods=rs("CarryGoods")
	mileageNum=rs("mileageNum")
	DriverID=rs("DriverID")
  end if
  %>
 <div id="AddandEditdiv" style="width:790px; height:460px; overflow:auto; ">
 <form id="AddForm" name="AddForm">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap><font color="#FFFFFF"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead12()" >&nbsp;<strong>车辆管理</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews width="100%">
      <tr>
        <td height="20" align="left">编号：</td>
        <td>
		<input name="SerialNum" type="text" class="textfield" id="SerialNum" style="WIDTH: 140;" value="<%= SerialNum %>" maxlength="100" readonly="true"></td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
        <td width="120" height="20" align="left">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td height="20" align="left">车牌号：</td>
        <td><input name="CarID" type="text" class="textfield" id="CarID" style="WIDTH: 140;" value="<%= CarID %>" maxlength="100" onChange="getInfo('CarID')"></td>
        <td height="20" align="left">驾驶员编号：</td>
        <td><input name="DriverID" type="text" class="textfield" id="DriverID" style="WIDTH: 140;" value="<%= DriverID %>" maxlength="100" onBlur="getInfo('DriverID')"></td>
        <td height="20" align="left">驾驶员姓名：</td>
        <td><input name="Driver" type="text" class="textfield" id="Driver" style="WIDTH: 140;" value="<%= Driver %>" maxlength="100"  readonly="true"></td>
      </tr>
      <tr>
        <td height="20" align="left">驾驶员电话：</td>
        <td>
		<input name="DriverPhone" type="text" class="textfield" id="DriverPhone" style="WIDTH: 140;" value="<%= DriverPhone %>" maxlength="100">
		</td>
        <td height="20" align="left">车辆状态：</td>
        <td>
		<input name="CarSatus" type="text" class="textfield" id="CarSatus" style="WIDTH: 140;" value="<%= CarSatus %>" maxlength="100"  readonly="true">
		</td>
        <td height="20" align="left"></td>
        <td>
		</td>
      </tr>
      <tr>
        <td height="20" align="left">载货量：</td>
        <td><input name="CarryGoods" type="text" class="textfield" id="CarryGoods" style="WIDTH: 140;" value="<%= CarryGoods %>" maxlength="100" onBlur="return checkNum(this)"></td>
        <td height="20" align="left">载人数：</td>
        <td>
		<input name="CarryMans" type="text" class="textfield" id="CarryMans" style="WIDTH: 140;" value="<%= CarryMans %>" maxlength="100" onBlur="return checkNum(this)">
		</td>
        <td height="20" align="left">当前公里数：</td>
        <td>
		<input name="mileageNum" type="text" class="textfield" id="mileageNum" style="WIDTH: 140;" value="<%= mileageNum %>" maxlength="100" onBlur="return checkNum(this)">
		</td>
      </tr>
	  </table>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews bgcolor="#99BBE8">
  <tr>  <td height="5">  </td>
  </tr>
	<tr>
	  <td align="center">
	  <input type="hidden" name="detailType" id="detailType" value="<%= detailType %>">
			<input name="submitSaveAdd" type="button" class="button"  id="submitSaveAdd" value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">&nbsp;
			<input name="submitDelete" type="button" class="button"  id="submitDelete" value="删除" style="WIDTH: 80;"  onClick="javascript:$('#detailType').val('Delete');toSubmit(this);">
	  </td>
	</tr>
  <tr>  <td height="5">  </td>
  </table>
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
  if detailType="AddNew" and Instr(session("AdminPurview"),"|1008.1,")>0 then
	set rs = server.createobject("adodb.recordset")
	sql="select * from z_Car"
	rs.open sql,connk3,1,3
	rs.addnew
	rs("CarID")=Request("CarID")
	rs("CarSatus")=Request("CarSatus")
	rs("Driver")=Request("Driver")
	rs("DriverPhone")=Request("DriverPhone")
	rs("CarryMans")=Request("CarryMans")
	rs("CarryGoods")=Request("CarryGoods")
	rs("mileageNum")=Request("mileageNum")
	rs("DriverID")=Request("DriverID")
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Edit" and Instr(session("AdminPurview"),"|1008.1,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from z_Car where SerialNum="&SerialNum
	rs.open sql,connk3,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs("CarID")=Request("CarID")
	rs("Driver")=Request("Driver")
	rs("DriverPhone")=Request("DriverPhone")
	rs("DriverID")=Request("DriverID")
	if rs("CarSatus")="空闲" then
	rs("CarryMans")=Request("CarryMans")
	rs("CarryGoods")=Request("CarryGoods")
	rs("mileageNum")=Request("mileageNum")
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Delete" and Instr(session("AdminPurview"),"|1008.1,")>0 then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from z_Car where CarSatus='空闲' and SerialNum="&SerialNum
	rs.open sql,connk3,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connk3.Execute("Delete from z_Car where SerialNum="&SerialNum)
	response.write "###"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="CarID" then
    InfoID=request("InfoID")
	sql="select * from z_car where CarID='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
    if not rs.bof or not rs.eof then
        response.write ("车牌号已经存在！")
        response.end
    end if
	response.write("###")
	rs.close
	set rs=nothing 
  elseif detailType="DriverID" then
    InfoID=request("InfoID")
	sql="select a.姓名 from [N-基本资料单头] a where a.员工代号='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write("###"&rs("姓名")&"###")
	end if
	rs.close
	set rs=nothing 
  end if
end if
 %>
</body>
</html>
