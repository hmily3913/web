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
<link rel="stylesheet" href="../Images/jquery.datepick.css">
<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="../Script/jquery.datepick.pack.js"></script>
<script language="javascript" src="../Script/jquery.datepick-zh-CN.js"></script>
<script language="javascript" src="../Script/jquery.datepick-zh-CN.js"></script>
<script language="javascript" src="../Script/highcharts.js"></script>
<script language="javascript" src="../Script/exporting.js"></script>
<script language="javascript">
	
function closead(){
  $("#listDiv").hide("slow");
  $('#SearchDiv').show("slow") ;
}
function closeadDetails(){
  $("#DetailslistDiv").hide("slow");
  $('#listDiv').animate({opacity: 1});
//  $('#SearchDiv').show("slow") ;
}

//显示汇总报表
function ReportShow(){
  if(($('#Rtype').val()=="OneDay"&&checkDate(document.getElementById("zhouqi")))||($('#Rtype').val()=="OneWeek"&&checkInt(document.getElementById("zhouqi")))||$('#Rtype').val()=="OneMonth"||$('#Rtype').val()=="OneSeason"||$('#Rtype').val()=="OneYear"){
	//显示加载页面
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#listDiv').load("QCReportQueryDetails.asp #listtable",{
		Rtype:$('#Rtype').val(),
		Rclass:$('#Rclass').val(),
		years:$('#years').val(),
		zhouqi:$('#zhouqi').val()
	  },function(response, status, xhr){
	  if (status =="success") {
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
		$('#SearchDiv').hide("slow");
		$('#listDiv').show("slow");
		//测试图表
		var table = document.getElementById('datatb');
		var table1 = document.getElementById('datatb1');
		var table2 = document.getElementById('datatb2');
		var names='',names1='',names2='';
		names=$('#Rclass option:selected').text()+'总体质检状况';
		if ($('#Rclass').val()=='QC')
		names1=$('#Rclass option:selected').text()+'不合格批次处理';
		else
		names1=$('#Rclass option:selected').text()+'不合格现象分布';
		names2=$('#Rclass option:selected').text()+'不合格原因分布';
			if(table){
			var number =$('tr', table).length-1;
			
			options = {
					 chart: {
							renderTo: 'container',
							defaultSeriesType: 'column'
					 },
					 title: {
							text:names
					 },
					 xAxis: {
					 },
					 yAxis: {
							title: {
								 text: '数值'
							}
					 },
					 tooltip: {
							formatter: function() {
								 return '<b>'+ this.series.name +'</b><br/>'+
										this.y +' '+ this.x;
							}
					 }
				};
			
				options.xAxis.categories = [];
				$('tr:gt(1):lt('+number+')', table).each( function(i) {
					options.xAxis.categories.push($('td:first',this).text());
				});
				
				// the data series
				options.series = [];
				$('tr:gt(0):lt('+number+')', table).each( function(i) {
					var tr = this;
					var n=0;
					$('td', tr).each( function(j) {
						if (j > 0) { // skip first column
							if(($('#Rclass').val()!='QC'&&(j==3||j==4||j==11||j==12||j==13))||($('#Rclass').val()=='QC'&&j==5)){
								if (i == 0) { // get the name and init the series
									options.series[n] = { 
										name: this.innerText,
										dataLabels:{enabled: true,formatter: function() {return this.y;}},
										data: []
									};
								} else { // add values
									options.series[n].data.push(parseFloat(this.innerHTML));
								}
								n++;
							}
						}
					});
				});
				
				var chart = new Highcharts.Chart(options);
			}
			if(table1){
			options1 = {
					 chart: {
							renderTo: 'container1',
							plotBackgroundColor: null,
							plotBorderWidth: null,
							plotShadow: false
					 },
					 title: {
							text:names1
					 },
						plotOptions: {
							pie: {
								allowPointSelect: true,
								cursor: 'pointer',
								dataLabels: {
									enabled: true,
									color: '#000000',
									connectorColor: '#000000',
									formatter: function() {
										return '<b>'+ this.point.name +'</b>: '+ this.y +' %';
									}
								}
							}
						},
						series: [{
						type: 'pie',
						name: '不合格现象',
						data: [
						]
					}],
					 tooltip: {
							formatter: function() {
								 return '<b>'+ this.point.name +'</b>: '+ this.y +' %';
							}
					 }
				};
				$('tr:gt(0)', table1).each( function(i) {
					var tr = this;
					var trvalue=new Array(2);
					$('td', tr).each( function(j) {
						if(j==0)
							trvalue[0]=this.innerHTML;
						else if(j==2)
							trvalue[1]=parseFloat(this.innerHTML);
					});
					options1.series[0].data.push(trvalue);
				});
				var chart2 = new Highcharts.Chart(options1);
			}
			if(table2){
			options2 = {
					 chart: {
							renderTo: 'container2',
							plotBackgroundColor: null,
							plotBorderWidth: null,
							plotShadow: false
					 },
					 title: {
							text:names2
					 },
						plotOptions: {
							pie: {
								allowPointSelect: true,
								cursor: 'pointer',
								dataLabels: {
									enabled: true,
									color: '#000000',
									connectorColor: '#000000',
									formatter: function() {
										return '<b>'+ this.point.name +'</b>: '+ this.y +' %';
									}
								}
							}
						},
						series: [{
						type: 'pie',
						name: '不合格现象',
						data: [
						]
					}],
					 tooltip: {
							formatter: function() {
								 return '<b>'+ this.point.name +'</b>: '+ this.y +' %';
							}
					 }
				};
				$('tr:gt(0)', table2).each( function(i) {
					var tr = this;
					var trvalue=new Array(2);
					$('td', tr).each( function(j) {
						if(j==0)
							trvalue[0]=this.innerHTML;
						else if(j==2)
							trvalue[1]=parseFloat(this.innerHTML);
					});
					options2.series[0].data.push(trvalue);
				});
				var chart3 = new Highcharts.Chart(options2);
			}
		}		
    })
  }
}
//显示明细
function ShowDetails(detailtype,sday,eday){
	//加载Detailslist内容，ajax提交
	$('#DetailslistDiv').load("QCReportQueryDetails.asp #Detailslisttable",{
		Dtype:detailtype,
		start_date:sday,
		end_date:eday
	  },function(response, status, xhr){
	  if (status =="success") {
		$('#DetailslistDiv').show("slow");
		$('#listDiv').animate({opacity: 0.25});
	  }	
    })
}
function OutDetails(detailtype,sday,eday){
	//加载Detailslist内容，ajax提交
	window.open("QCReportQueryDetails.asp?print_tag=1&Dtype="+detailtype+"&start_date="+sday+"&end_date="+eday,"Print","","false");
}

function TypeChange(){
	if($('#Rtype').val()=="OneDay"){
		var txt='<input type="text" id="zhouqi" name="zhouqi">';
		$('#zhouqis').html(txt);
		$('#zhouqi').datepick({dateFormat: 'yyyy-mm-dd'});
	}else	if($('#Rtype').val()=="OneWeek"){
		var txt='<select id="years" name="years">';
		txt+='<option value="2011">2011</option>';
		txt+='<option value="2012">2012</option>';
		txt+='<option value="2013">2013</option>';
		txt+='<option value="2014">2014</option></select>';
		txt+='<select id="zhouqi" name="zhouqi">';
		txt+='<option value="1">01</option>';
		txt+='<option value="2">02</option>';
		txt+='<option value="3">03</option>';
		txt+='<option value="4">04</option>';
		txt+='<option value="5">05</option>';
		txt+='<option value="6">06</option>';
		txt+='<option value="7">07</option>';
		txt+='<option value="8">08</option>';
		txt+='<option value="9">09</option>';
		txt+='<option value="10">10</option>';
		txt+='<option value="11">11</option>';
		txt+='<option value="12">12</option>';
		txt+='<option value="13">13</option>';
		txt+='<option value="14">14</option>';
		txt+='<option value="15">15</option>';
		txt+='<option value="16">16</option>';
		txt+='<option value="17">17</option>';
		txt+='<option value="18">18</option>';
		txt+='<option value="19">19</option>';
		txt+='<option value="20">20</option>';
		txt+='<option value="21">21</option>';
		txt+='<option value="22">22</option>';
		txt+='<option value="23">23</option>';
		txt+='<option value="24">24</option>';
		txt+='<option value="25">25</option>';
		txt+='<option value="26">26</option>';
		txt+='<option value="27">27</option>';
		txt+='<option value="28">28</option>';
		txt+='<option value="29">29</option>';
		txt+='<option value="30">30</option>';
		txt+='<option value="31">31</option>';
		txt+='<option value="32">32</option>';
		txt+='<option value="33">33</option>';
		txt+='<option value="34">34</option>';
		txt+='<option value="35">35</option>';
		txt+='<option value="36">36</option>';
		txt+='<option value="37">37</option>';
		txt+='<option value="38">38</option>';
		txt+='<option value="39">39</option>';
		txt+='<option value="40">40</option>';
		txt+='<option value="41">41</option>';
		txt+='<option value="42">42</option>';
		txt+='<option value="43">43</option>';
		txt+='<option value="44">44</option>';
		txt+='<option value="45">45</option>';
		txt+='<option value="46">46</option>';
		txt+='<option value="47">47</option>';
		txt+='<option value="48">48</option>';
		txt+='<option value="49">49</option>';
		txt+='<option value="50">50</option>';
		txt+='<option value="51">51</option>';
		txt+='<option value="52">52</option></select>';
		$('#zhouqis').html(txt);
	}else	if($('#Rtype').val()=="OneMonth"){
		var txt='<select id="years" name="years">';
		txt+='<option value="2011">2011</option>';
		txt+='<option value="2012">2012</option>';
		txt+='<option value="2013">2013</option>';
		txt+='<option value="2014">2014</option></select>';
		txt+='<select id="zhouqi" name="zhouqi">';
		txt+='<option value="1">01</option>';
		txt+='<option value="2">02</option>';
		txt+='<option value="3">03</option>';
		txt+='<option value="4">04</option>';
		txt+='<option value="5">05</option>';
		txt+='<option value="6">06</option>';
		txt+='<option value="7">07</option>';
		txt+='<option value="8">08</option>';
		txt+='<option value="9">09</option>';
		txt+='<option value="10">10</option>';
		txt+='<option value="11">11</option>';
		txt+='<option value="12">12</option></select>';
		$('#zhouqis').html(txt);
	}else	if($('#Rtype').val()=="OneSeason"){
		var txt='<select id="years" name="years">';
		txt+='<option value="2011">2011</option>';
		txt+='<option value="2012">2012</option>';
		txt+='<option value="2013">2013</option>';
		txt+='<option value="2014">2014</option></select>';
		txt+='<select id="zhouqi" name="zhouqi">';
		txt+='<option value="1">春季</option>';
		txt+='<option value="2">夏季</option>';
		txt+='<option value="3">秋季</option>';
		txt+='<option value="4">冬季</option></select>';
		$('#zhouqis').html(txt);
	}else	if($('#Rtype').val()=="OneYear"){
		var txt='<select id="zhouqi" name="zhouqi">';
		txt+='<option value="2011">2011</option>';
		txt+='<option value="2012">2012</option>';
		txt+='<option value="2013">2013</option>';
		txt+='<option value="2014">2014</option></select>';
		$('#zhouqis').html(txt);
	}

}
</script>
</HEAD>
<BODY>
<%
'if Instr(session("AdminPurview"),"|40,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
%>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden; display:none">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="margin:0 auto; ">
<div id="SearchDiv" style="position:fixed !important;position:absolute;top:5;left:10;height:100%; width:100%;background-color:#ffffff;position:absolute;marginTop:1px;marginLeft:1px; z-index:100">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="24" nowrap class="tablemenu"><font color="#FFFFFF"><img src="../Images/images/flexigrid/grid.png" width="18" height="18" border="0" align="absmiddle">&nbsp;<strong>品保周期报表查询</strong></font></td>
  </tr>
  <tr>
    <td height="36" align="center" nowrap  bgcolor="#EBF2F9"><table width="100%" border="0" cellspacing="0">
      <tr>
        <form name="formSearch" method="post" action="test.asp">
          <td nowrap> 报表类型:&nbsp;<select id="Rtype" name="Rtype" onChange="return TypeChange()">
		  <option value="">请选择类型</option>
		  <option value="OneDay">日报</option>
		  <option value="OneWeek">周报</option>
		  <option value="OneMonth">月报</option>
		  <option value="OneSeason">季报</option>
		  <option value="OneYear">年报</option>
		  </select>
		  &nbsp;分类:&nbsp;<select id="Rclass" name="Rclass">
		  <option value="QC">IQC/OQC</option>
		  <option value="MN1">一分厂</option>
		  <option value="MN2">二分厂</option>
		  <option value="MN3">三分厂</option>
		  <option value="MN4">眼镜布绳</option>
		  <option value="MN">公司汇总</option>
		  </select>
		  &nbsp;报表期号：<div id="zhouqis" style="display:inline"></div>
		  <input name="submitSearch" type="button" class="button" value="检索" onClick="ReportShow()">
          </td>
        </form>
      </tr>
    </table>      </td>    
  </tr>
</table>
</div>
<div id="listDiv" style="z-index:500;position:fixed !important;position:absolute;top:10;left:20;height:95%; width:100%; overflow-y:auto; overflow-x: auto; display:none"></div>
<div id="DetailslistDiv" style="position:fixed !important;position:absolute;top:20;left:30;height:90%; width:100%;z-index:600;overflow-y:auto; overflow-x: auto; display:none"></div>
</div>
</BODY>
</HTML>