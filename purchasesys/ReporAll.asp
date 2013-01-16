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
    });
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
					text:'采购系统报表汇总信息'
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
if Instr(session("AdminPurview"),"|20,")=0 then 
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
dim num11,num12,num13,num14,num15
dim num21,num22,num23,num24,num25
dim num31,num32,num33,num34,num35
dim num41,num42,num43,num44,num45
dim num51,num52,num53,num54,num55
dim num61,num62,num63,num64,num65
dim num71,num72,num73
StartDate=stryear&"-"&strmonth&"-"&sdaynum
EndDate=stryear&"-"&strmonth&"-"&edaynum
sql="select top 1 * from purchasesys order by SerialNum desc" 
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num11=rs("ReturnRateOne")	
	num12=rs("ReturnRate")	
	num14=rs("ReturnRatejidu")	
	num15=rs("ReturnRateniandu")	
	num21=rs("AnxiousBillRateOne")
	num31=rs("SupplyPromOne")
	num41=rs("SpecialAcceptOne")
	num51=rs("SupplyQualifiedOne")
	num61=rs("DevelopReachOne")
	num71=rs("returnOne")
	num22=rs("AnxiousBillRate")
	num32=rs("SupplyProm")
	num42=rs("SpecialAccept")
	num52=rs("SupplyQualified")
	num62=rs("DevelopReach")
	num72=rs("return")
	num24=rs("AnxiousBillRatejidu")	
	num25=rs("AnxiousBillRateniandu")	
	num34=rs("SupplyPromjidu")	
	num35=rs("SupplyPromniandu")	
	num44=rs("SpecialAcceptjidu")	
	num45=rs("SpecialAcceptniandu")	
	num54=rs("SupplyQualifiedjidu")	
	num55=rs("SupplyQualifiedniandu")	
	num64=rs("DevelopReachjidu")	
	num65=rs("DevelopReachniandu")	
  rs.close
  set rs=nothing
sql="select * from purchasesys where UPTdate='"&split(getDateRangebyMonth(laststrmonth),"###")(1)&"'" 
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num13=rs("ReturnRate")	
	num23=rs("AnxiousBillRate")
	num33=rs("SupplyProm")
	num43=rs("SpecialAccept")
	num53=rs("SupplyQualified")
	num63=rs("DevelopReach")
	num73=rs("return")
  rs.close
  set rs=nothing
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap  class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>采购系统报表汇总信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><font style="left:inherit ">&nbsp;报表周期从<%response.Write(stryear&"-"&strmonth&"-"&sdaynum)%>至<%response.Write(stryear&"-"&strmonth&"-"&edaynum)%></font><input type="button" name="Charts" style="height:18px;" value="显示图表" onClick="return showChart()"><input type="button" name="Charts" value="隐藏图表" style="display:none;height:18px" onClick="$('input[type=button][name=Charts]').toggle();$('#chartDiv').hide();"></td>
  </tr>  <tr>
    <td height="36"  align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="1">
      <tr bgcolor='#99BBE8'>
          <td nowrap width="108"> 名称
          </td>
          <td nowrap width="71"> 上月数据
          </td>
          <td nowrap width="71"> 昨日数据
          </td>
          <td nowrap width="57"> 本月数据
          </td>
          <td nowrap width="57"> 季度数据
          </td>
          <td nowrap width="57"> 年度数据
          </td>
          <td nowrap width="48"> 单位
          </td>
          <td nowrap width="71"> 针对部门
          </td>
          <td nowrap width="274"> 计算描述
          </td>
	  </tr>
	  <tr>
          <td nowrap> 采购物料退货率
          </td>
          <td nowrap> <%=num13%>
          </td>
          <td nowrap> <%=num11%>
          </td>
          <td nowrap> <%=num12%>
          </td>
          <td nowrap> <%=num14%>
          </td>
          <td nowrap> <%=num15%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>采购部</td>
          <td nowrap> (退货笔数/采购总笔数)*100%</td>
          
      </tr>
	  <tr>
          <td nowrap> 采购物料退货笔数
          </td>
          <td nowrap> <%=num73%>
          </td>
          <td nowrap> <%=num71%>
          </td>
          <td nowrap> <%=num72%>
          </td>
          <td nowrap> 
          </td>
          <td nowrap> 
          </td>
		  <td nowrap>笔</td>
		  <td nowrap>采购部</td>
          <td nowrap> 退货笔数</td>
          
      </tr>
      <tr>
          <td nowrap> 采购急单率</td>
          <td nowrap> <%=num23%>
          </td>
          <td nowrap> <%=num21%>
          </td>
          <td nowrap> <%=num22%>
          </td>
          <td nowrap> <%=num24%>
          </td>
          <td nowrap> <%=num25%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>生管部</td>
          <td nowrap> (急单笔数/采购总笔数)*100%</td>
      </tr>
      <tr>
          <td nowrap> 供货及时率</td>
          <td nowrap> <%=num33%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('SP1')">
          </td>
          <td nowrap> <%=num31%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('SP2')">
          </td>
          <td nowrap> <%=num32%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('SP3')">
          </td>
          <td nowrap> <%=num34%>
          </td>
          <td nowrap> <%=num35%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>采购部</td>
          <td nowrap> (合格到位笔数/采购总笔数)*100%</td>
      </tr>
      <tr>
          <td nowrap> 特采件数</td>
          <td nowrap> <%=num43%>
          </td>
          <td nowrap> <%=num41%>
          </td>
          <td nowrap> <%=num42%>
          </td>
          <td nowrap> <%=num44%>
          </td>
          <td nowrap> <%=num45%>
          </td>
		  <td nowrap>笔</td>
		  <td nowrap>采购部</td>
          <td nowrap> 特采笔数</td>
      </tr>
      <tr>
          <td nowrap> 供货批次合格率</td>
          <td nowrap> <%=num53%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('SQ1')">
          </td>
          <td nowrap> <%=num51%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('SQ2')">
          </td>
          <td nowrap> <%=num52%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('SQ3')">
          </td>
          <td nowrap> <%=num54%>
          </td>
          <td nowrap> <%=num55%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>采购部</td>
          <td nowrap> 合格批次/总批次</td>
      </tr>
      <tr>
          <td nowrap> 材料开发达成率</td>
          <td nowrap> <%=num63%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('DR1')">
          </td>
          <td nowrap> <%=num61%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('DR2')">
          </td>
          <td nowrap> <%=num62%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('DR3')">
          </td>
          <td nowrap> <%=num64%>
          </td>
          <td nowrap> <%=num65%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>采购部</td>
          <td nowrap> (合格到位笔数/打样采购总笔数)*100%</td>
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

<%	
  rs.close
  set rs=nothing
%>
</BODY>
</HTML>
