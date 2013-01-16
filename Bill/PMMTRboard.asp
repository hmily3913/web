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
	$('#listDiv').load("PMMTRboardDetails.asp #listtable",{
	  showType:'DetailsList',
	  Page:arr[0],
	  seachword:$('#seachword').val(),
	  flag4search:$('#flag4search').val()
	 },function(response, status, xhr){
	  if (status =="success") {
	    pageNavigation('pageN',arr,response.split('###')[1],response.split('###')[2],'showDiv');
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
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
	  }	
    })
}
//处理添加按钮
function showpadd(obj,sid){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#addShowDiv').load("PMMTRboardDetails.asp #AddandEditdiv",{
	  showType:'AddEditShow',
	  detailType:obj,
	  SerialNum:sid
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#addDiv").show("slow");
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
}
//处理提交事务
function toSubmit(){
  $.post('PMMTRboardDetails.asp?showType=DataProcess',$("#AddForm").serialize(),function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	else pageN(0);
  });
  $("#addDiv").hide("slow");
}
//处理提交事务
function toMainSubmit(obj){
  if($("input:radio:checked").val()===undefined){alert('请先选择一条记录再进行操作！');return false;}
  if($("input:radio:checked").attr('ForCheck').split('_')[0]==1&&obj=='QC')showpadd('Edit',$("input:radio:checked").val());
  else
  $.post('PMMTRboardDetails.asp?showType=DataProcess',{detailType:obj,SerialNum:$("input:radio:checked").val()},function(data){
    if(data.indexOf("@@@")>-1) alert(data.split('@@@')[1]);
	else pageN(0);
	//$("input:radio:checked").parent().parent().attr('bgColor','#ff99ff');
  });
}
//处理选择打样单号事务
function getInfo(obj){
  if($("#"+obj).val()==''){alert("对应编号不能为空！");return false;}
	$.get("PMMTRboardDetails.asp", { showType: "getInfo",detailType: obj, InfoID: $("#"+obj).val() },
	   function(data){
		 if(data.indexOf("###")==-1)alert("对应编号不存在，请检查！");
		 else{
		   if(obj=="OrderID"){
		   $("#ProductId").val(data.split('###')[1]);
		   $("#Model").val(data.split('###')[2]);
		   $("#Unit").val(data.split('###')[3]);
		   $("#Quantity").val(data.split('###')[4]);
		   $("#Quality").val(data.split('###')[5]);
		   $("#Product_td").html(data.split('###')[6]);
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
  $("#ProductId").val($("#ProductName option:selected").text().split('@')[1]);
  $("#Model").val($("#ProductName option:selected").text().split('@')[2]);
  $("#Unit").val($("#ProductName option:selected").text().split('@')[3]);
  $("#Quantity").val($("#ProductName option:selected").text().split('@')[4]);
  $("#Quality").val($("#ProductName option:selected").text().split('@')[5]);
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
<font color="#FF0000"><strong>生管紧急物料看板</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
<font style="background-color:#EBF2F9"><a href="javascript:$('#flag4search').val('0');pageN(0);">未处理</a></font>&nbsp;
<font style="background-color:#ff99ff"><a href="javascript:$('#flag4search').val('1');pageN(0);">验前接收</a></font>&nbsp;
<font style="background-color:#7CFC00"><a href="javascript:$('#flag4search').val('2');pageN(0);">验后确认</a></font>&nbsp;
<input type="hidden" name="flag4search" id="flag4search" value="">
周别：<input type="text" name="seachword" id="seachword" style='HEIGHT: 18px;WIDTH: 40px;'>
<input type="button" name="seachbutton" id="seachbutton" onClick="pageN(arr)" value="查找" style='HEIGHT: 18px;WIDTH: 40px;'>
<input type="button" name="addbutton" id="button" onClick="showpadd('AddNew','')" value="添加" style='HEIGHT: 18px;WIDTH: 40px;'>
<input type="button" name="addbutton" id="button" onClick="toMainSubmit('QC')" value="品保接收" style='HEIGHT: 18px;WIDTH: 80px;'>
<input type="button" name="addbutton" id="button" onClick="toMainSubmit('ST')" value="仓库接收" style='HEIGHT: 18px;WIDTH: 80px;'>
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