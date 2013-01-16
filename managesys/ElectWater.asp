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
<link rel="stylesheet" href="../Images/jqi.css">
<script language="javascript" src="../Script/jquery-impromptu.3.1.js"></script>
<script language="javascript">
function Count(){
	$.get("ElectWaterDetails.asp", { showType: "Count",start_date:$('#start_date').val(),end_date:$('#end_date').val(),louhao:$('#louhao').val() },function(data){
		if(data.indexOf("计算")>0){
			$.prompt(data,{
				buttons: { 计算: 'doCount'},
				submit:function(v,m,f){ 
					if(v=='doCount'){
						$.get("ElectWaterDetails.asp", { showType: "doCount",start_date:$('#start_date').val(),end_date:$('#end_date').val(),louhao:$('#louhao').val() },function(dt){
							alert(dt);
							});
					}
					$.prompt.close();
					return false; 
				 }
			 });
		}else{
			alert(data);
		}
	});
}
function Export(){
				window.open("ElectWaterDetails.asp?print_tag=1&showType=Export&start_date="+$('#start_date').val()+"&end_date="+$('#end_date').val()+"&Keyword="+encodeURI($('#Keyword').val())+"&louhao="+encodeURI($('#louhao').val()),"Print","","false");
}
</script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
'if Instr(session("AdminPurview"),"|1002,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
%>
<BODY>
<%
dim Result,StartDate,EndDate,Keyword,Reachsum,unReachsum,Reachper,sqlstr,louhao
Result=request("Result")
StartDate=request("start_date")
if StartDate="" then StartDate=date()
EndDate=request("end_date")
if EndDate="" then EndDate=date()
Keyword=request("Keyword")
louhao=request("louhao")

function PlaceFlag()
  if Result="Search" then
      Response.Write "[<font color='red'>"&StartDate&"</font>] 年 [<font color='red'>"&EndDate&"</font>]月 [<font color='red'>"&louhao&"</font>] 用 电 扣 费 明 细 表"
  else
      Response.Write "请选择日期进行统计!"
  end if
end function  
 
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>水电管理信息</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="ElectWater.asp?Result=Search">
          <td nowrap> 年月检索：
		  <select id="start_date" name="start_date">
			<option value="2012" <%if StartDate="2012" then response.write ("selected")%>>2012</option>
			<option value="2013" <%if StartDate="2013" then response.write ("selected")%>>2013</option>
		  </select>
          
          &nbsp;年
		  <select id="end_date" name="end_date">
			<option value="1" <%if EndDate="1" then response.write ("selected")%>>01</option>
			<option value="2" <%if EndDate="2" then response.write ("selected")%>>02</option>
			<option value="3" <%if EndDate="3" then response.write ("selected")%>>03</option>
			<option value="4" <%if EndDate="4" then response.write ("selected")%>>04</option>
			<option value="5" <%if EndDate="5" then response.write ("selected")%>>05</option>
			<option value="6" <%if EndDate="6" then response.write ("selected")%>>06</option>
			<option value="7" <%if EndDate="7" then response.write ("selected")%>>07</option>
			<option value="8" <%if EndDate="8" then response.write ("selected")%>>08</option>
			<option value="9" <%if EndDate="9" then response.write ("selected")%>>09</option>
			<option value="10" <%if EndDate="10" then response.write ("selected")%>>10</option>
			<option value="11" <%if EndDate="11" then response.write ("selected")%>>11</option>
			<option value="12" <%if EndDate="12" then response.write ("selected")%>>12</option>
		  </select>
		<select name="louhao" id="louhao" >
		<option value="办公楼宿舍" <%if louhao="办公楼宿舍" then response.write ("selected")%>>办公楼宿舍</option>
		<option value="旭日小区4号楼" <%if louhao="旭日小区4号楼" then response.write ("selected")%>>旭日小区4号楼</option>
		<option value="旭日小区22号楼" <%if louhao="旭日小区22号楼" then response.write ("selected")%>>旭日小区22号楼</option>
		</select>
		  <input name="submitSearch3" type="button" class="button" value="计算" onclick="Count()">
		  <input name="submitSearch" type="submit" class="button" value="查询">
      <input name="submitSearch2" type="button" class="button" value="导出" onclick="Export()">
          </td>
        </form>
		<td>
		</font><a href="ElectWaterList.asp?Result=Person" onClick='changeAdminFlag("宿舍人员列表")'>宿舍人员列表</a>
		<font color="#0000FF">&nbsp;|&nbsp;
		</font><a href="EWMoneyList.asp" onClick='changeAdminFlag("宿舍水电列表")'>宿舍水电列表</a>
		</td>
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
    <td width="40" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>宿舍</strong></font></td>
    <td width="50" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>部门</strong></font></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>工号</strong></font></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>姓名</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>上月水表</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">本月水表</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>上月电表</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">本月电表</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>上月热水</strong></font></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">本月热水</font></strong></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">用水量</font></strong></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">用电量</font></strong></td>
    <td width="50" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">用热水</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">合计费用</font></strong></td>
    <td width="60" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">应扣费用</font></strong></td>
  </tr>
 <%

  dim rs,sql'sql语句
  '获取记录总数
  sql="select * from managesys_ShuidianRpt "&_
"where year="&StartDate&" and period="&EndDate&" and louhao='"&louhao&"'"&_
"order by ftext"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
  dim totalfee,totalelect '总计
  totalelect=0.0
  totalfee=0.0
    while(not rs.eof)'填充数据到表格
			totalfee=totalfee+cdbl(rs("onefee"))
	  Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("ftext")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("name1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("fnumber")&"</td>"
      Response.Write "<td nowrap>"&rs("name2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("water")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("thiswater")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("elect")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("thiselect")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("lasthotwater")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("thihotwater")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("waterdiff")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("electdiff")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("hotwaterdiff")&"</td>" & vbCrLf
	  
      Response.Write "<td nowrap>"&rs("tempnum")&"</td>" & vbCrLf
	    
      Response.Write "<td nowrap>"&formatnumber(rs("onefee"),1)&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='14' nowrap  bgcolor='#EBF2F9'>&nbsp;合计</td>" & vbCrLf
    Response.Write "<td  nowrap  bgcolor='#EBF2F9'>"&totalfee&"</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
'    response.write "<tr><td height='50' align='center' colspan='9' nowrap  bgcolor='#EBF2F9'>暂无该月水电信息</td></tr>"
'-----------------------------------------------------------

  rs.close
  set rs=nothing
  %>
  </table>
  <%

end function 

%>


