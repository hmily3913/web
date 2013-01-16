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
					text:'分厂系统报表汇总信息'
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
if Instr(session("AdminPurview"),"|30,")=0 then 
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
dim num81,num82,num83
dim num91,num92,num93
dim num101,num102,num103
dim num111,num112,num113
dim num121,num122,num123
StartDate=stryear&"-"&strmonth&"-"&sdaynum
EndDate=stryear&"-"&strmonth&"-"&edaynum
sql="select top 1 * from manusys order by SerialNum desc" 
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num11=rs("DeliveryReach1One")	
	num12=rs("DeliveryReach1")	
	num21=rs("DeliveryReach2One")	
	num22=rs("DeliveryReach2")	
	num31=rs("DeliveryReach3One")	
	num32=rs("DeliveryReach3")	
	num41=rs("DeliveryReach4One")	
	num42=rs("DeliveryReach4")	
	num111=rs("DeliveryReach5One")	
	num112=rs("DeliveryReach5")	
	num121=rs("DeliveryReach6One")	
	num122=rs("DeliveryReach6")	
	num51=rs("ProductConsum1One")	
	num52=rs("ProductConsum1")	
	num61=rs("ProductConsum2One")	
	num62=rs("ProductConsum2")	
	num71=rs("ProductConsum3One")	
	num72=rs("ProductConsum3")	
	num81=rs("ProductConsum4One")	
	num82=rs("ProductConsum4")	
	num91=rs("FinishRemakeOne")	
	num92=rs("FinishRemake")	
	num101=rs("PMDReachOne")	
	num102=rs("PMDReach")	
  rs.close
  set rs=nothing
sql="select * from manusys where UPTdate='"&split(getDateRangebyMonth(laststrmonth),"###")(1)&"'" 
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num13=rs("DeliveryReach1")	
	num23=rs("DeliveryReach2")	
	num33=rs("DeliveryReach3")	
	num43=rs("DeliveryReach4")	
	num113=rs("DeliveryReach5")	
	num123=rs("DeliveryReach6")	
	num53=rs("ProductConsum1")	
	num63=rs("ProductConsum2")	
	num73=rs("ProductConsum3")	
	num83=rs("ProductConsum4")	
	num93=rs("FinishRemake")	
	num103=rs("PMDReach")	
  rs.close
  set rs=nothing
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap  class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>分厂系统报表汇总信息</strong></font></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><font style="left:inherit ">&nbsp;报表周期从<%response.Write(stryear&"-"&strmonth&"-"&sdaynum)%>至<%response.Write(stryear&"-"&strmonth&"-"&edaynum)%></font><input type="button" name="Charts" style="height:18px;" value="显示图表" onClick="return showChart()"><input type="button" name="Charts" value="隐藏图表" style="display:none;height:18px" onClick="$('input[type=button][name=Charts]').toggle();$('#chartDiv').hide();"></td>
  </tr>  <tr>
    <td height="36"  align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="1">
      <tr bgcolor='#99BBE8'>
          <td nowrap width="54"> 分厂
          </td>
          <td nowrap width="114"> 名称
          </td>
          <td nowrap width="54"> 上月数据
          </td>
          <td nowrap width="54"> 昨日数据
          </td>
          <td nowrap width="54"> 本月数据
          </td>
          <td nowrap width="30"> 单位
          </td>
          <td nowrap width="54"> 针对部门
          </td>
          <td nowrap width="456"> 计算描述
          </td>
	  </tr>
	  <tr>
          <td nowrap> 一厂
          </td>
          <td nowrap> 一厂交期达成率
          </td>
          <td nowrap> <%=num13%>
          </td>
          <td nowrap> <%=num11%>
          </td>
          <td nowrap> <%=num12%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>一分厂</td>
          <td nowrap> (实际完成的订单批次/应完成的订单批次)*100%</td>
      </tr>
	  <tr>
          <td nowrap> 
          </td>
          <td nowrap> 制程超耗率
          </td>
          <td nowrap> <%=num53%>
          </td>
          <td nowrap> <%=num51%>
          </td>
          <td nowrap> <%=num52%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>一分厂</td>
          <td nowrap> 制程中质量控制不力导致损失金额在200元以上质量事件(当月制程异常件数) </td>
          
      </tr>
	  <tr>
          <td nowrap> 二厂
          </td>
          <td nowrap> 二厂交期达成率
          </td>
          <td nowrap> <%=num23%>
          </td>
          <td nowrap> <%=num21%>
          </td>
          <td nowrap> <%=num22%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>二分厂</td>
          <td nowrap> (实际完成的订单批次/应完成的订单批次)*100%</td>
      </tr>
	  <tr>
          <td nowrap> 
          </td>
          <td nowrap> 制程超耗率
          </td>
          <td nowrap> <%=num63%>
          </td>
          <td nowrap> <%=num61%>
          </td>
          <td nowrap> <%=num62%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>二分厂</td>
          <td nowrap> 制程中质量控制不力导致损失金额在200元以上质量事件(当月制程异常件数) </td>
          
      </tr>
	  <tr>
          <td nowrap> 三厂
          </td>
          <td nowrap> 三厂交期达成率
          </td>
          <td nowrap> <%=num33%>
          </td>
          <td nowrap> <%=num31%>
          </td>
          <td nowrap> <%=num32%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>三分厂</td>
          <td nowrap> (实际完成的订单批次/应完成的订单批次)*100%</td>
      </tr>
	  <tr>
          <td nowrap> 
          </td>
          <td nowrap> 制程超耗率
          </td>
          <td nowrap> <%=num73%>
          </td>
          <td nowrap> <%=num71%>
          </td>
          <td nowrap> <%=num72%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>三分厂</td>
          <td nowrap> 制程中质量控制不力导致损失金额在200元以上质量事件(当月制程异常件数) </td>
          
      </tr>
	  <tr>
          <td nowrap> 眼镜布绳
          </td>
          <td nowrap> 眼镜布绳交期达成率
          </td>
          <td nowrap> <%=num113%>
          </td>
          <td nowrap> <%=num111%>
          </td>
          <td nowrap> <%=num112%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>眼镜布绳</td>
          <td nowrap> (实际完成的订单批次/应完成的订单批次)*100%</td>
      </tr>
	  <tr>
          <td nowrap> 花生盒
          </td>
          <td nowrap> 花生盒交期达成率
          </td>
          <td nowrap> <%=num123%>
          </td>
          <td nowrap> <%=num121%>
          </td>
          <td nowrap> <%=num122%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>花生盒</td>
          <td nowrap> (实际完成的订单批次/应完成的订单批次)*100%</td>
      </tr>
	  <tr>
          <td nowrap> 娄桥厂
          </td>
          <td nowrap> 娄桥厂交期达成率
          </td>
          <td nowrap> <%=num43%>
          </td>
          <td nowrap> <%=num41%>
          </td>
          <td nowrap> <%=num42%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>娄桥</td>
          <td nowrap> (实际完成的订单批次/应完成的订单批次)*100%</td>
      </tr>
	  <tr>
          <td nowrap> 
          </td>
          <td nowrap> 制程超耗率
          </td>
          <td nowrap> <%=num83%>
          </td>
          <td nowrap> <%=num81%>
          </td>
          <td nowrap> <%=num82%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>娄桥</td>
          <td nowrap> 制程中质量控制不力导致损失金额在200元以上质量事件(当月制程异常件数) </td>
          
      </tr>
	  <tr>
          <td nowrap>所有分厂
          </td>
          <td nowrap> 成品返工件数
          </td>
          <td nowrap> <%=num93%>
          </td>
          <td nowrap> <%=num91%>
          </td>
          <td nowrap> <%=num92%>
          </td>
		  <td nowrap>件</td>
		  <td nowrap>分厂</td>
          <td nowrap> 实际返工返修件数</td>
      </tr>
	  <tr>
          <td nowrap> 生管
          </td>
          <td nowrap> 生管交期达成率
          </td>
          <td nowrap> <%=num103%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('PMDR1')">
          </td>
          <td nowrap> <%=num101%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('PMDR2')">
          </td>
          <td nowrap> <%=num102%>
		  <input type="button" value="细" style='HEIGHT: 16px;WIDTH: 20px;font-size:10px;' onClick="return ShowDetails('PMDR3')">
          </td>
		  <td nowrap>%</td>
		  <td nowrap>生管部</td>
          <td nowrap> (生管回复交期满足营销交期笔数/总计划笔数)*100%</td>
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
