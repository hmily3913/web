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
					text:'人资系统报表汇总信息'
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
if Instr(session("AdminPurview"),"|90,")=0 then 
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
dim num11,num12,num13,num111,num121,num131,num112,num122,num132,num113,num123,num133
dim num21,num22,num23
dim num31,num32,num33
dim num41,num42,num43
dim num51,num52,num53
dim num61,num62,num63
dim num71,num72,num73
dim num81,num82,num83
dim num91,num92,num93
dim num141,num142,num143
dim num151,num152,num153
StartDate=stryear&"-"&strmonth&"-"&sdaynum
EndDate=stryear&"-"&strmonth&"-"&edaynum
sql="select top 1 * from hrsys Order by SerialNum desc" 
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num11=rs("PersonnelLossLDOne")	
	num12=rs("PersonnelLossLD")	
	num111=rs("PersonnelLossLD1One")	
	num121=rs("PersonnelLossLD1")	
	num112=rs("PersonnelLossLD2One")	
	num122=rs("PersonnelLossLD2")	
	num113=rs("PersonnelLossLD3One")	
	num123=rs("PersonnelLossLD3")	
	num142=rs("PersonnelLossLD4One")	
	num143=rs("PersonnelLossLD4")	
	num152=rs("PersonnelLossLD5One")	
	num153=rs("PersonnelLossLD5")	
	num21=rs("PersonnelLossLQOne")	
	num22=rs("PersonnelLossLQ")	
	num31=rs("RecruitmentEffRateLDOne")
	num32=rs("RecruitmentEffRateLD")
	num41=rs("RecruitmentEffRateLQOne")
	num42=rs("RecruitmentEffRateLQ")
	num51=rs("RecruitmentTimelyLDOne")
	num52=rs("RecruitmentTimelyLD")
	num61=rs("RecruitmentTimelyLQOne")
	num62=rs("RecruitmentTimelyLQ")
	num71=rs("ClassRateOne")
	num72=rs("ClassRate")
	num81=rs("SalaryCalculAccurOne")
	num82=rs("SalaryCalculAccur")
	num91=rs("SalaryCalculTimelyOne")
	num92=rs("SalaryCalculTimely")
  rs.close
  set rs=nothing
sql="select * from hrsys where UPTdate='"&split(getDateRangebyMonth(laststrmonth),"###")(1)&"'" 
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num13=rs("PersonnelLossLD")	
	num23=rs("PersonnelLossLQ")	
	num131=rs("PersonnelLossLD1")	
	num132=rs("PersonnelLossLD2")	
	num133=rs("PersonnelLossLD3")	
	num141=rs("PersonnelLossLD4")	
	num151=rs("PersonnelLossLD5")	
	num33=rs("RecruitmentEffRateLD")
	num43=rs("RecruitmentEffRateLQ")
	num53=rs("RecruitmentTimelyLD")
	num63=rs("RecruitmentTimelyLQ")
	num73=rs("ClassRate")
	num83=rs("SalaryCalculAccur")
	num93=rs("SalaryCalculTimely")
  rs.close
  set rs=nothing
%>

<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap  class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>人资系统报表汇总信息</strong></font><input type="button" name="Charts" style="height:18px;" value="显示图表" onClick="return showChart()"><input type="button" name="Charts" value="隐藏图表" style="display:none;height:18px" onClick="$('input[type=button][name=Charts]').toggle();$('#chartDiv').hide();"></td>
  </tr>
  <tr>
    <td height="24" nowrap bgcolor="#EBF2F9"><font style="left:inherit ">&nbsp;报表周期从<%response.Write(stryear&"-"&strmonth&"-"&sdaynum)%>至<%response.Write(stryear&"-"&strmonth&"-"&edaynum)%></font></td>
  </tr>  <tr>
    <td height="36"  align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="1">
      <tr bgcolor='#99BBE8'>
          <td nowrap width="40"> 厂区
          </td>
          <td nowrap width="93"> 名称
          </td>
          <td nowrap width="56"> 上月数据
          </td>
          <td nowrap width="56"> 昨日数据
          </td>
          <td nowrap width="54"> 本月数据
          </td>
          <td nowrap width="30"> 单位
          </td>
          <td nowrap width="54"> 针对部门
          </td>
          <td nowrap width="553"> 计算描述
          </td>
	  </tr>
	  <tr>
          <td nowrap> 蓝道
          </td>
          <td nowrap> 人员流失率
          </td>
          <td nowrap> <%=num13%>
          </td>
          <td nowrap> <%=num11%>
          </td>
          <td nowrap> <%=num12%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>人资部</td>
          <td nowrap> (当月流失人员数/公司当月平均人数)*100%(2职等试用期内\普工入职7天内离职人员不在考核之内)</td>
          
      </tr>
	  <tr>
          <td nowrap> 
          </td>
          <td nowrap> 一厂流失率
          </td>
          <td nowrap> <%=num131%>
          </td>
          <td nowrap> <%=num111%>
          </td>
          <td nowrap> <%=num121%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>一分厂</td>
          <td nowrap> (当月流失人员数/公司当月平均人数)*100%(2职等试用期内\普工入职7天内离职人员不在考核之内)</td>
          
      </tr>
	  <tr>
          <td nowrap> 
          </td>
          <td nowrap> 二厂流失率
          </td>
          <td nowrap> <%=num132%>
          </td>
          <td nowrap> <%=num112%>
          </td>
          <td nowrap> <%=num122%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>二分厂</td>
          <td nowrap> (当月流失人员数/公司当月平均人数)*100%(2职等试用期内\普工入职7天内离职人员不在考核之内)</td>
          
      </tr>
	  <tr>
          <td nowrap> 
          </td>
          <td nowrap> 三厂流失率
          </td>
          <td nowrap> <%=num133%>
          </td>
          <td nowrap> <%=num113%>
          </td>
          <td nowrap> <%=num123%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>三分厂</td>
          <td nowrap> (当月流失人员数/公司当月平均人数)*100%(2职等试用期内\普工入职7天内离职人员不在考核之内)</td>
          
      </tr>
	  <tr>
          <td nowrap> 
          </td>
          <td nowrap> 眼镜绳布流失率
          </td>
          <td nowrap> <%=num141%>
          </td>
          <td nowrap> <%=num142%>
          </td>
          <td nowrap> <%=num143%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>眼镜绳布</td>
          <td nowrap> (当月流失人员数/公司当月平均人数)*100%(2职等试用期内\普工入职7天内离职人员不在考核之内)</td>
          
      </tr>
	  <tr>
          <td nowrap> 
          </td>
          <td nowrap> 行政流失率
          </td>
          <td nowrap> <%=num151%>
          </td>
          <td nowrap> <%=num152%>
          </td>
          <td nowrap> <%=num153%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>人资部</td>
          <td nowrap> (当月流失人员数/公司当月平均人数)*100%(2职等试用期内\普工入职7天内离职人员不在考核之内)</td>
          
      </tr>
      <tr>
          <td nowrap> </td>
          <td nowrap> 招聘有效率</td>
          <td nowrap> <%=num33%>
          </td>
          <td nowrap> <%=num31%>
          </td>
          <td nowrap> <%=num32%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>人资部</td>
          <td nowrap> 有效率=试用期后转正人数/招聘试用人数.</td>
      </tr>
      <tr>
          <td nowrap> </td>
          <td nowrap> 招聘及时率</td>
          <td nowrap> <%=num53%>
          </td>
          <td nowrap> <%=num51%>
          </td>
          <td nowrap> <%=num52%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>人资部</td>
          <td nowrap> （实际按时到位人数/需求人数）×100%</td>
      </tr>
	  <tr>
          <td nowrap> 娄桥
          </td>
          <td nowrap> 人员流失率
          </td>
          <td nowrap> <%=num23%>
          </td>
          <td nowrap> <%=num21%>
          </td>
          <td nowrap> <%=num22%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>人资部</td>
          <td nowrap> (当月流失人员数/公司当月平均人数)*100%(2职等试用期内\普工入职7天内离职人员不在考核之内)</td>
          
      </tr>
      <tr>
          <td nowrap> </td>
          <td nowrap> 招聘有效率</td>
          <td nowrap> <%=num43%>
          </td>
          <td nowrap> <%=num41%>
          </td>
          <td nowrap> <%=num42%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>人资部</td>
          <td nowrap> 有效率=试用期后转正人数/招聘试用人数.</td>
      </tr>
      <tr>
          <td nowrap> </td>
          <td nowrap> 招聘及时率</td>
          <td nowrap> <%=num63%>
          </td>
          <td nowrap> <%=num61%>
          </td>
          <td nowrap> <%=num62%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>人资部</td>
          <td nowrap> （实际按时到位人数/需求人数）×100%</td>
      </tr>
      <tr>
          <td nowrap>公同 </td>
          <td nowrap> 开课率</td>
          <td nowrap> <%=num73%>
          </td>
          <td nowrap> <%=num71%>
          </td>
          <td nowrap> <%=num72%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>人资部</td>
          <td nowrap> （开课项目/计划总科目）×100%</td>
      </tr>
      <tr>
          <td nowrap> </td>
          <td nowrap> 薪资计算准确率</td>
          <td nowrap> <%=num83%>
          </td>
          <td nowrap> <%=num81%>
          </td>
          <td nowrap> <%=num82%>
          </td>
		  <td nowrap>笔</td>
		  <td nowrap>人资部</td>
          <td nowrap> 错误笔数</td>
      </tr>
      <tr>
          <td nowrap> </td>
          <td nowrap> 薪资计算及时率</td>
          <td nowrap> <%=num93%>
          </td>
          <td nowrap> <%=num91%>
          </td>
          <td nowrap> <%=num92%>
          </td>
		  <td nowrap>%</td>
		  <td nowrap>人资部</td>
          <td nowrap> 薪资按时完成率</td>
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
