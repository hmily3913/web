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
<script language="javascript" src="../Script/highcharts.js"></script>
<script language="javascript" src="../Script/exporting.js"></script>
<script language="javascript">
function ShowDetails(obj){
	$('#listDiv').load("ReporAllDetails.asp #listtable",{
	  showType:obj
	 },function(response, status, xhr){
	  if (status =="success") {
	  }	
    })
}
function showChart(){
			$('input[type=button][name=Charts]').toggle();
			$('#chartDiv').show();
	options = {
			 chart: {
					renderTo: 'chartDiv',
					defaultSeriesType: 'line'
			 },
			 title: {
					text:'仓库系统报表汇总信息'
			 },
			 subtitle: {
			 	text: '本年度各月份信息'
			 },
			 xAxis: {
			 	gridLineWidth: 1
			 },
			 yAxis: {
					title: {
						 text: '数值'
					}
			 },
			 tooltip: {
			 	shared  :true,
			 	crosshairs  :true
			 }
		};
	$.get('ReporAllDetails.asp',{ showType: "getChart"},
	  function(data){
			
		 	var datajson=jQuery.parseJSON(data);
			options.xAxis.categories = [];
			$.each(datajson,function(i){
				options.xAxis.categories.push(datajson[i].Monthdata[1].value+'月');
			});
			
			options.series = [];
			$.each(datajson,function(i){
				for(var j=0;j<datajson[i].Monthdata.length;j++){
					if (j > 1) { // get the name and init the series
						if(i==0){
							options.series[j-2] = { 
								name: datajson[i].Monthdata[j].name,
								dataLabels:{enabled: true,formatter: function() {return this.y;}},
								data: []
							};
						} 
						options.series[j-2].data.push(parseFloat(datajson[i].Monthdata[j].value));
					}
				}
			});
			var chart = new Highcharts.Chart(options);
		});
	
	
}
</script>
</HEAD>
<!--#include file="../Include/ConnSiteData.asp" -->
<!--#include file="../CheckAdmin.asp"-->
<%
if Instr(session("AdminPurview"),"|50,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<BODY>
<%
dim sdaynum,edaynum
dim stryear,strmonth,strdate,laststrmonth
stryear=Year(now())
if Month(now() <=9) then
strmonth="0"&Month(now())
else
strmonth=Month(now())
end if
if Month(now())= 1 then
laststrmonth=12&"#"&eval(stryear-1)
else
laststrmonth=eval(Month(now())-1)
end if
strdate=date()
sdaynum=1
select case Month(now())
	case 2
	  if ((stryear mod 4=0) and (stryear mod 100>0)) or (stryear mod 400=0) then
	    edaynum=29
	  else
	    edaynum=28
	  end if
    case 4
	  edaynum=30
    case 6
	  edaynum=30
    case 9
	  edaynum=30
    case 11
	  edaynum=30
	case 1
	  edaynum=31
	  stryear=Year(now())-1
	case else
	  edaynum=31
end select
'response.Write(stryear&"-"&strmonth&"-"&edaynum)

dim rs,sql,sqlstr,StartDate,EndDate
dim Reachsum,unReachsum,Reachper
dim num11,num12,num13
dim num21,num22,num23
dim num31,num32,num33
dim num41,num42,num43
dim num51,num52,num53
dim num61,num62,num63
dim num71,num72,num73
StartDate=stryear&"-"&strmonth&"-"&sdaynum
EndDate=stryear&"-"&strmonth&"-"&edaynum
if Weekday(dateadd("d",-1,now()))=1 then
sql="select * from stocksys where datediff(d,UPTdate,getdate())=2" 
else  
sql="select * from stocksys where datediff(d,UPTdate,getdate())=1" 
end if
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num11=rs("ReceivePromOne")	
	num21=rs("SendmReachOne")
	num31=rs("FinishDeliveryRateOne")
	num41=rs("CardAccuracyRateOne")
	num61=rs("PickupPromOne")
	num71=rs("SendCarPromOne")
	num12=rs("ReceiveProm")	
	num22=rs("SendmReach")
	num32=rs("FinishDeliveryRate")
	num42=rs("CardAccuracyRate")
	num62=rs("PickupProm")
	num72=rs("SendCarProm")
  rs.close
  set rs=nothing
sql="select * from stocksys where UPTdate='"&split(getDateRangebyMonth(laststrmonth),"###")(1)&"'" 
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num43=rs("CardAccuracyRate")
	num63=rs("PickupProm")
	num73=rs("SendCarProm")
  rs.close
  set rs=nothing
sql="select * from stocksys where UPTdate='"&split(getDateRangebyMonth(strmonth),"###")(0)&"'" 
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num13=rs("ReceiveProm")	
	num23=rs("SendmReach")
	num33=rs("FinishDeliveryRate")
  rs.close
  set rs=nothing
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap  class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>仓库系统报表汇总信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><font style="left:inherit ">&nbsp;报表周期从<%response.Write(stryear&"-"&strmonth&"-"&sdaynum)%>至<%response.Write(stryear&"-"&strmonth&"-"&edaynum)%></font><input type="button" name="Charts" style="height:18px;" value="显示图表" onClick="return showChart()"><input type="button" name="Charts" value="隐藏图表" style="display:none;height:18px" onClick="$('input[type=button][name=Charts]').toggle();$('#chartDiv').hide();"></td>
  </tr>  <tr>
    <td height="36"  align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="1">
      <tr bgcolor='#99BBE8'>
          <td nowrap width="89"> 名称
          </td>
          <td nowrap width="66"> 上月数据
          </td>
          <td nowrap width="55"> 昨日数据
          </td>
          <td nowrap width="54"> 本月数据
          </td>
          <td nowrap width="33"> 单位
          </td>
          <td nowrap width="54"> 针对部门
          </td>
          <td nowrap width="489"> 计算描述
          </td>
	  </tr>
	  <tr>
          <td nowrap> 收料及时率
          </td>
          <td nowrap> <%=num13%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('RP1')">
          </td>
          <td nowrap> <%=num11%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('RP2')">
          </td>
          <td nowrap> <%=num12%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('RP3')">
          </td>
		  <td nowrap>%</td>
		  <td nowrap>仓储科</td>
          <td nowrap> (及时收料次数/应收料次数)*100% 收料通知单审核日期至外购入库单审核日期</td>
          
      </tr>
      <tr>
          <td nowrap> 发料达成率</td>
          <td nowrap> <%=num23%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('SR1')">
          </td>
          <td nowrap> <%=num21%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('SR2')">
          </td>
          <td nowrap> <%=num22%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('SR3')">
          </td>
		  <td nowrap>%</td>
		  <td nowrap>仓储科</td>
          <td nowrap> (及时发料笔数/应发料笔数)*100% 领料单制单日期至领料单审核日期</td>
      </tr>
      <tr>
          <td nowrap> 成品出货及时率</td>
          <td nowrap> <%=num33%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('FDR1')">
          </td>
          <td nowrap> <%=num31%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('FDR2')">
          </td>
          <td nowrap> <%=num32%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('FDR3')">
          </td>

		  <td nowrap>%</td>
		  <td nowrap>仓储科</td>
          <td nowrap> (及时出货笔数/应出货笔数)*100% 发货通知单发货日期至销售出库单审核日期</td>
      </tr>
      <tr>
          <td nowrap> 帐卡物准确率</td>
          <td nowrap> <%=num43%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('CAR1')">
          </td>
          <td nowrap> <%=num41%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('CAR2')">
          </td>
          <td nowrap> <%=num42%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('CAR3')">
          </td>

		  <td nowrap>%</td>
		  <td nowrap>仓储科</td>
          <td nowrap> 准确笔数/盘点总笔数*100%</td>
      </tr>
      <tr>
          <td nowrap> 存货金额</td>
          <td nowrap> 详情查明细
          </td>
		  <td nowrap></td>
		  <td nowrap></td>
		  <td nowrap></td>
		  <td nowrap>仓储科</td>
          <td nowrap> 按仓库，仓管，物料类别</td>
      </tr>
      <tr>
          <td nowrap> 提货及时率</td>
          <td nowrap> <%=num63%>
          </td>
          <td nowrap> <%=num61%>
          </td>
          <td nowrap> <%=num62%>
          </td>

		  <td nowrap>%</td>
		  <td nowrap>仓储科</td>
          <td nowrap> （及时提货次数/总提货次数）*100%</td>
      </tr>
      <tr>
          <td nowrap> 派车及时率</td>
          <td nowrap> <%=num73%>
          </td>
          <td nowrap> <%=num71%>
          </td>
          <td nowrap> <%=num72%>
          </td>

		  <td nowrap>%</td>
		  <td nowrap>仓储科</td>
          <td nowrap> （及时派车次数/总派车次数）*100%</td>
      </tr>
    </table>      </td>    
  </tr>
</table>
<div id="chartDiv" style="width: 800px; height: 400px; margin: 0 auto; display:none;"></div>
<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td height="30"></td>
  </tr>
</table>
<div id="listDiv"></div>
</BODY>
</HTML>
