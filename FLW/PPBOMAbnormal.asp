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
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript">
function closead1(){
  $("#addDiv").hide("slow");
}
var edittr;
var arr = new Array();
function showpadd(obj,sid,fid){
	if(obj=='Edit'){
//		evt = evt ? evt : (window.event ? window.event : null);
		edittr=event.srcElement.parentNode;
	}
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#addShowDiv').load("PPBOMAbnormalDetails.asp #AddandEditdiv",{
	  showType:'AddEditShow',
	  detailType:obj,
	  FInterId:sid,
	  FEntryId:fid
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#addDiv").show("slow");
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
}
function toSubmit(){
  $.post('PPBOMAbnormalDetails.asp?showType=CheckProcess',$("#AddForm").serialize(),function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	else edittr.bgColor="#ff99ff";
	  $("#addDiv").hide("slow");
  });
}
//分页
function pageN(){
    arr = new Array();
    for(var i = 0 ; i < pageN.arguments.length ; i++){
        arr[i] = pageN.arguments[i];
    }
	//显示加载页面
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#listDiv').load("PPBOMAbnormalDetails.asp #listtable",{
	  showType:'DetailsList',
	  Page:arr[0],
	  start_date:$("#DS_start_date").val(),
	  end_date:$("#DS_end_date").val()
	 },function(response, status, xhr){
	  if (status =="success") {
	    pageNavigation('pageN',arr,response.split('###')[1],response.split('###')[2],'showDiv');
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
	//产生分页导航栏
}

</script>
</HEAD>
<BODY>
<%
if Instr(session("AdminPurviewFLW"),"|106,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden; display:none">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="margin-top:0; margin-bottom:0; ">
<font color="#FF0000"><strong>已入库未领料处理</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
<table><tr><td>
<font style="background-color:#ff99ff">已结束</font>&nbsp;
<input type="hidden" id="departhide" value="allpart">
		从
          <script language=javascript> 
          var myDate=new dateSelector(); 
          myDate.year; 
		  myDate.date; 
          myDate.inputName='start_date';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。 
          myDate.display(); 
          </script>
          &nbsp;到
          <script language=javascript> 
          myDate.year; 
          myDate.inputName='end_date';  //注意这里设置输入框的name，同一页中的日期输入框，不能出现重复的name。 
          myDate.display(); 
          </script>
<input type="button" name="seachbutton" id="seachbutton" onClick="pageN(1)" value="查询" style='HEIGHT: 18px;WIDTH: 40px;font-size:12px;'>
</td></tr></table>
</p>
<div id="listDiv"></div>
<div id="showDiv"></div>
<div id="addDiv" style="width:'820px';height:'480px';top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
<div id="addShowDiv"></div>
</div>
</div>
</BODY>
</HTML>