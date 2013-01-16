<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!--#include file="../CheckAdmin.asp" -->
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="报表系统" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="../Images/CssAdmin.css">
<link rel="stylesheet" href="../Images/jquery.datepick.css">
<link rel="stylesheet" href="../Images/jquery.autocomplete.css">
<link rel="stylesheet" href="../Images/zTreeStyle.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="../Script/jquery.easydrag.js"></script>
<script language="javascript" src="../Script/jquery.datepick.pack.js"></script>
<script language="javascript" src="../Script/jquery.datepick-zh-CN.js"></script>
<script language="javascript" src="../Script/jquery.autocomplete.pack.js"></script>
<script language="javascript" src="../Script/jquery.ztree-2.6.min.js"></script>
<script language="javascript">
var zTree2;
var setting2;
	setting2 = {
		expandSpeed:"",
		dragCopy:true,
		dragCopy:true,
		keepParent: false,
		keepLeaf: false,
		isSimpleData : true,
		treeNodeParentKey : "PSNum",
		treeNodeKey : "SerialNum",
		rootPID : 0,
		async: true,
		asyncParamOther : {"showType":"DepartMent"},
		callback: {
			click: zTreeOnClick2,
			asyncSuccess:zTreeOnAsyncSuccess
		}
	};
	function zTreeOnAsyncSuccess(event, treeId, treeNode, msg) {
		var tempNode = zTree2.getNodeByParam("SerialNum",$('#DepartID').val());
		if (tempNode)zTree2.selectNode(tempNode);
	}
	function showMenu() {
		setting2.asyncUrl= "AttQueryDetails.asp?cq="+$('#cq').val();
		var nodes = [];
		setting2.async = true;
		zTree2 = $("#dropdownMenu").zTree(setting2, nodes);
		var cityObj = $("#Depart");
		var cityOffset = $("#Depart").offset();
		$("#DropdownMenuBackground").css({left:cityOffset.left + "px", top:cityOffset.top + cityObj.outerHeight() + "px"}).slideDown("fast");
	}
	function hideMenu() {
		$("#DropdownMenuBackground").hide();
	}

	function zTreeOnClick2(event, treeId, treeNode) {
		if (treeNode) {
			var cityObj = $("#Depart");
			cityObj.attr("value", treeNode.name);
			$('#DepartID').val(treeNode.SerialNum);
			$.getJSON('AttQueryDetails.asp?showType=getInfo&DepartID='+treeNode.SerialNum,null,function(data){
				$('#ID').unautocomplete().autocomplete(data,Autooptions);
			});
			hideMenu();
		}
	}
		
var productList;
var Autooptions={
	minChars:1,
	max:20,
	width:250,
	mustMatch: true,
	matchContains:true,
	formatItem:function(row,i,max){
		return "\""+row.ssn+"\""+" ["+row.name+"]";
	},
	formatMatch:function(row,i,max){return row.ssn+" "+row.name;},
	formatResult:function(row){return row.ssn;}
};
function initAutoComplete(data){
	productList=data;
	$('#ID').focus().autocomplete(productList,Autooptions);
}
$(function(){
	$.getJSON('AttQueryDetails.asp?showType=getInfo&cq='+$('#cq').val()+'&DepartID='+$('#DepartID').val(),null,initAutoComplete);
	$('#Sdate').datepick({dateFormat: 'yyyy-mm-dd'});
	$('#Edate').datepick({dateFormat: 'yyyy-mm-dd'});
	var yestoday=showdate(0);
	$('#Sdate').val(yestoday);
	$('#Edate').val(yestoday);
});
//分页
function pageN(){
	//显示加载页面
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#listDiv').load("AttQueryDetails.asp #listtable",{
	  showType:'DetailsList',
		DepartID:$('#DepartID').val(),
	  ID:$('#ID').val(),
	  Sdate:$('#Sdate').val(),
	  cq:$('#cq').val(),
	  Edate:$('#Edate').val()
	 },function(response, status, xhr){
	  if (status =="success") {
			$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    });
	//产生分页导航栏
}

function output(){
  window.open("AttQueryDetails.asp?print_tag=1&showType=DetailsList&cq="+$('#cq').val()+"&DepartID="+$('#DepartID').val()+"&ID="+$('#ID').val()+"&Sdate="+$('#Sdate').val()+"&Edate="+$('#Edate').val(),"Print","","false");
}
function changecq(){
	if($('#cq').val()!=""){
		$('DepartID').val(1);
		$('Depart').val('总公司');
	}
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
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden; display:none;">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="position:fixed !important;position:absolute;top:0;margin:0 auto; ">
<font color="#FF0000"><strong>考勤查询</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
厂区：<select id="cq" name="cq" style="font-size:12px; width:80px;" onChange="return changecq()">
<option value="">请选择</option>
<option value="bh">滨海</option>
<option value="lq">娄桥</option>
</select>
<input type="hidden" id="DepartID" name="DepartID" value="1">
部门：<input type="text" class="textfield" id="Depart" name="Depart" value="总公司" style='WIDTH: 100px;' readonly>&nbsp;<a id="menuBtn" href="#" onClick="showMenu(); return false;">选择</a>&nbsp;
工号：<input type="text" name="ID" id="ID" class="textfield" style='WIDTH: 60px;'>&nbsp;
日期：从<input type="text" name="Sdate" id="Sdate" class="textfield" style='WIDTH: 80px;'>
到<input type="text" name="Edate" id="Edate" class="textfield" style='WIDTH: 80px;'>&nbsp;
<input type="button" name="seachbutton" id="seachbutton" onClick="pageN()" value="查找" style='HEIGHT: 18px;WIDTH: 40px;'>
<input type="button" name="addbutton" id="button" onClick="output()" value="导出" style='HEIGHT: 18px;WIDTH: 40px;'>
</p>
<div id="addDiv" style="width:100%;height:'480px';top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow:auto;">
<div id="addShowDiv" ></div>
</div>
<div id="listDiv" style="height:450px;"></div>
<div id="showDiv"></div>
</div>
<div id="DropdownMenuBackground" style="display:none; position:absolute; height:300px; min-width:150px; background-color:white;border:1px solid;overflow-y:auto;overflow-x:auto; z-index:999999">
	<ul id="dropdownMenu" class="tree"></ul>
</div>

</BODY>
</HTML>