<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!--#include file="../CheckAdmin.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="报表系统" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="../Images/CssAdmin.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="../Script/jquery.easydrag.js"></script>
<script language="javascript">
//关闭弹出层
function closead1(){
  $("#addDiv").hide("slow");
}
var arr = new Array();
//分页
function pageN(){
    arr = new Array();
    for(var i = 0 ; i < pageN.arguments.length ; i++){
        arr[i] = pageN.arguments[i];
    }
	//显示加载页面
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#listDiv').load("OrderAbnormalDetails.asp #listtable",{
	  showType:'DetailsList',
	  Page:arr[0],
	  seachword:$('#seachword').val(),
	  flag4search:$('#flag4search').val()
	 },function(response, status, xhr){
	  if (status =="success") {
	    pageNavigation('pageN',arr,response.split('###')[1],response.split('###')[2],'showDiv');
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    });
	//产生分页导航栏
}
//双击弹出回复层
var thistd;
function ORClickTd(obj,reptype,tdid){
	//加载list内容，ajax提交
	thistd=obj;
	$('#ReplyDiv').load("OrderReviewDetails.asp #showReplyDiv",{
	  showType:'showDetails',
	  detailType:reptype,
	  Fentry:tdid
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#ReplyDiv").show("slow");
			$('#ReplyDiv').easydrag(); 
	$("#ReplyDiv").setHandler("formove"); 
	  }	
    });
}
//处理添加按钮
function showpadd(obj,sid){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#addShowDiv').load("OrderAbnormalDetails.asp #AddandEditdiv",{
	  showType:'AddEditShow',
	  detailType:obj,
	  SerialNum:sid
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#addDiv").show("slow");
			$('#addDiv').easydrag(); 
			$('#addDiv').setHandler("formove"); 
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    });
}
//处理提交事务
function toSubmit(){
  $.post('OrderAbnormalDetails.asp?showType=DataProcess',$("#AddForm").serialize(),function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	else pageN(0);
  });
  $("#addDiv").hide("slow");
}
//处理选择打样单号事务
function getInfo(obj){
  if($("#"+obj).val()==''){alert("对应编号不能为空！");return false;}
	$.get("OrderAbnormalDetails.asp", { showType: "getInfo",detailType: obj, InfoID: $("#"+obj).val() },
	   function(data){
		 if(data.indexOf("###")==-1)alert("对应编号不存在，请检查！");
		 else{
		   if(obj=="OrderID"){
		   $("#CustomID").val(data.split('###')[1]);
		   $("#CustomRanke").val(data.split('###')[3]);
		   $("#CustomLevel").val(data.split('###')[4]);
		   $("#Agenter").val(data.split('###')[2]);
		   $("#ProductType").val(data.split('###')[7]);
		   $("#OrderDate").val(data.split('###')[5]);
		   $("#MCReplyDate").val(data.split('###')[6]);
		   $("#OrderQuantity").val(data.split('###')[8]);
		   $("#CustomDate").val(data.split('###')[9]);
		   $("#Product_td").html(data.split('###')[10]);
		   }
		   else if(obj=="Register"){
		   $("#RegisterName").val(data.split('###')[1]);
		   $("#Department").val(data.split('###')[2]);
		   $("#Departmentname").val(data.split('###')[3]);
		   }
		 }
	   });
}
//选择产品触发
function changePro(){
  $("#ProductType").val($("#Product option:selected").text().split('@')[2]);
  $("#MCReplyDate").val($("#Product option:selected").text().split('@')[1]);
  $("#OrderQuantity").val($("#Product option:selected").text().split('@')[3]);
  $("#CustomDate").val($("#Product option:selected").text().split('@')[4]);
}
function output(){
  window.open("OrderAbnormalDetails.asp?print_tag=1&showType=DetailsList&seachword="+$('#seachword').val()+"&flag4search="+$('#flag4search').val(),"Print","","false");
}
</script>
</HEAD>
<BODY>
<%
'if Instr(session("AdminPurviewFLW"),"|203,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
%>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden;">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="top:0;margin:0 auto; ">
<font color="#FF0000"><strong>订单异常反馈处理进度汇总</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
<font style="background-color:#EBF2F9"><a href="javascript:$('#flag4search').val('1');pageN(1);">未处理</a></font>&nbsp;
<font style="background-color:#ff99ff"><a href="javascript:$('#flag4search').val('2');pageN(1);">已处理</a></font>&nbsp;
<font style="background-color:#66ff66"><a href="javascript:$('#flag4search').val('3');pageN(1);">已结案</a></font>&nbsp;
<input type="hidden" name="flag4search" id="flag4search" value="">
周别：<input type="text" name="seachword" id="seachword" style='HEIGHT: 18px;WIDTH: 40px;'>
<input type="button" name="seachbutton" id="seachbutton" onClick="pageN(arr)" value="查找" style='HEIGHT: 18px;WIDTH: 40px;'>
<input type="button" name="addbutton" id="button" onClick="showpadd('AddNew','')" value="添加" style='HEIGHT: 18px;WIDTH: 40px;'>
<input type="button" name="addbutton" id="button" onClick="output()" value="导出" style='HEIGHT: 18px;WIDTH: 40px;'>
</p>
<div id="addDiv" style="width:100%;height:'480px';top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
<div id="addShowDiv"></div>
</div>
<div id="listDiv"></div>
<div id="showDiv"></div>
<script language="javascript">
arr[0] = 1;
pageN(arr);
</script>
</div>
</BODY>
</HTML>