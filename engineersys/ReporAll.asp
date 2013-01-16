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
					text:'工程系统报表汇总信息'
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
if Instr(session("AdminPurview"),"|70,")=0 then 
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
StartDate=stryear&"-"&strmonth&"-"&sdaynum
EndDate=stryear&"-"&strmonth&"-"&edaynum
if Weekday(dateadd("d",-1,now()))=1 then
sql="select * from engineersys where datediff(d,UPTdate,getdate())=2"
else
sql="select * from engineersys where datediff(d,UPTdate,getdate())=1"
end if
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num11=rs("ProofingFinishRateOne")	
	num21=rs("BomFinishOne")
	num31=rs("NewProductDevOne")
	num41=rs("NewProductOrderOne")
	num51=rs("TryQualifiedOne")
	num12=rs("ProofingFinishRate")	
	num22=rs("BomFinish")
	num32=rs("NewProductDev")
	num42=rs("NewProductOrder")
	num52=rs("TryQualified")
  rs.close
  set rs=nothing
sql="select * from engineersys where UPTdate='"&split(getDateRangebyMonth(laststrmonth),"###")(1)&"'" 
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num13=rs("ProofingFinishRate")	
	num33=rs("NewProductDev")
	num43=rs("NewProductOrder")
	num53=rs("TryQualified")
  rs.close
  set rs=nothing
sql="select * from engineersys where UPTdate='"&split(getDateRangebyMonth(strmonth),"###")(0)&"'" 
'response.Write(sql)
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	num23=rs("BomFinish")
  rs.close
  set rs=nothing
%>


<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24"   class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>工程系统报表汇总信息</strong></font></td>
  </tr>
  <tr>
    <td height="24"  bgcolor="#EBF2F9"><font style="left:inherit ">&nbsp;报表周期从<%response.Write(stryear&"-"&strmonth&"-"&sdaynum)%>至<%response.Write(stryear&"-"&strmonth&"-"&edaynum)%></font><input type="button" name="Charts" style="height:18px;" value="显示图表" onClick="return showChart()"><input type="button" name="Charts" value="隐藏图表" style="display:none;height:18px" onClick="$('input[type=button][name=Charts]').toggle();$('#chartDiv').hide();"></td>
  </tr>  <tr>
    <td height="36"  align="center"   bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="1">
      <tr bgcolor='#99BBE8'>
          <td width="113"> 名称
          </td>
          <td width="50"> 上月数据
          </td>
          <td  width="50"> 昨日数据
          </td>
          <td  width="50"> 本月数据
          </td>
          <td  width="30"> 单位
          </td>
          <td  width="60"> 针对部门
          </td>
          <td width="374"> 计算描述
          </td>
	  </tr>
	  <tr>
          <td width="113"> 打样按期完成率
          </td>
          <td width="50"> <%=num13%>
          </td>
          <td width="50"> <%=num11%>
          </td>
          <td width="50"> <%=num12%>
          </td>
		  <td >%</td>
		  <td >工程部</td>
          <td > (完成件数/应打样件数)*100%</td>
          
      </tr>
      <tr>
          <td > BOM表按期完成率</td>
          <td > <%=num23%>
          </td>
          <td > <%=num21%>
          </td>
          <td > <%=num22%>
          </td>
		  <td >%</td>
		  <td >工程部</td>
          <td > (制作完成件数/总制作件数)*100%</td>
      </tr>
      <tr>
          <td > 新产品开发件数</td>
          <td > <%=num33%>
          </td>
          <td > <%=num31%>
          </td>
          <td > <%=num32%>
          </td>
		  <td >件</td>
		  <td >工程部</td>
          <td > 总开发件数</td>
      </tr>
      <tr>
          <td > 新产品接单额</td>
          <td > <%=num43%>
          </td>
          <td > <%=num41%>
          </td>
          <td > <%=num42%>
          </td>
		  <td >元</td>
		  <td >工程部</td>
          <td > 当月接单额</td>
      </tr>
      <tr>
          <td > 产前试做合格率</td>
          <td > <%=num53%>
          </td>
          <td > <%=num51%>
          </td>
          <td > <%=num52%>
          </td>
		  <td >%</td>
		  <td >工程部</td>
          <td > （完成件数/应试做件数）*100%</td>
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
