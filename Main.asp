<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012 - zbh-STUDIO" />
<META NAME="Author" CONTENT="---honglms.vicp.net" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<TITLE>欢迎使用蓝道报表管理系统</TITLE>
<link rel="stylesheet" href="Images/CssAdmin.css">
<link rel="stylesheet" href="Images/zTreeStyle.css">
<link rel="stylesheet" href="Images/flexigrid.bbit.css">
<link rel="stylesheet" href="Images/jquery.jgrowl.css">
<style type="text/css">
.tree li button.onlineUser {
	background: url("Images/people.gif") repeat scroll 0 0 transparent;
} 
button.messBtn {
	background:url("Images/images/zTreeimg/edit.png") no-repeat scroll 1px 1px transparent;
}
html{
	font: 12px Arial, Helvetica, sans-serif;
}
h1{margin-top:-10px;}
hr{margin:20px 0;}
th{font: 12px;}


</style>
<script language="javascript" src="Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="Script/jquery.messager.js"></script>
<script language="javascript" src="Script/jquery.ztree-2.6.min.js"></script>
<script language="javascript" src="Script/jquery.easydrag.js"></script>
<script language="javascript" src="Script/flexigrid.pack.js"></script>
<script language="javascript" src="Script/jquery.jgrowl_minimized.js"></script>
<SCRIPT language=JavaScript>
var st1=setInterval("showMessager()",30000);
var st2=setInterval("CheckNewMessage()",5000);
var step=0;
function flash_title() 
{ 
step++ ;
var title_new='你有待办事项未完成！'
if (step==7) {step=1} 
if (step==1) {document.title='◆◇◇◇'+title_new+'◇◇◇◆'} 
if (step==2) {document.title='◇◆◇◇'+title_new+'◇◇◆◇'} 
if (step==3) {document.title='◇◇◆◇'+title_new+'◇◆◇◇'} 
if (step==4) {document.title='◇◇◇◆'+title_new+'◆◇◇◇'} 
if (step==5) {document.title='◇◇◆◇'+title_new+'◇◆◇◇'} 
if (step==6) {document.title='◇◆◇◇'+title_new+'◇◇◆◇'} 
} 
showMessager();
function SetSession(sn,url){
  jQuery.get("CheckMessage.asp", { "key": "SetSession","SessionNum": sn},function(data){
		frames["mainFrame"].location.href=url;
	});
}
var mesflag=true;
var mesnum=0;
function CheckNewMessage(){
  jQuery.get("CheckMessage.asp", { "key": "CheckNewMessage"},
   function(data){
		if(data.length>0){
			if(!frames["topFrame"].document.getElementById('newmessage'))
      	$(frames["topFrame"].document.getElementById('messageMana')).append('<img src="Images/announce.gif" border="0" alt="'+data+'条新消息" height="16px" width="16px" id="newmessage">');
				if(data>mesnum)mesflag=true;
				if(data!=mesnum){
					mesnum=data;
				}
				if(mesflag){
					mesflag=false;
					$.jGrowl('<a href="javascript:void(0)" style="color:#FFF" onclick="tempfunc()">'+data+'条新消息</a>',{sticky:true,position:'center'});
				}
		}else{
			mesnum=data;
			if(frames["topFrame"].document.getElementById('newmessage'))
      	$(frames["topFrame"].document.getElementById('newmessage')).remove();
		}
   });
}
function tempfunc(){
	initmessList();
	$('#jGrowl').jGrowl('shutdown');
	$('#jGrowl').remove();
}
function showMessager(){
  jQuery.get("CheckMessage.asp", { "key": "AllNeed"},
   function(data){
	if(data.length>0){
			st3=setInterval("flash_title()",360);
      $.messager.show(0,data);
	}
	else{
			window.clearInterval(st3);
	}
   });
}
function switchSysBar()
{
   if (switchPoint.innerText==3)
   {
      switchPoint.innerText=4
      document.all("frameTitle").style.display="none"
   }
   else
   {
      switchPoint.innerText=3
      document.all("frameTitle").style.display=""
   }
}
var zTree1;
var setting;
var rMenu;
setting = {
	async: true,
	addHoverDom: addHoverDom,
	removeHoverDom: removeHoverDom,
	asyncParam: ["SerialNum"],
	asyncParamOther : {"key":"longinUser"},
	callback: {
		dblclick: zTreeOnDblclick
	}
};
function addHoverDom(treeId, treeNode) {
	var aObj = $("#" + treeNode.tId + "_a");
	if ($("#diyBtn_"+treeNode.UserName).length>0||treeNode.isParent) return;
	var editStr = "<span id='diyBtn_space_" +treeNode.UserName+ "' >&nbsp;</span><button type='button' class='messBtn' id='diyBtn_" +treeNode.UserName+ "' title='"+treeNode.name+"' onfocus='this.blur();'></button>";
	aObj.append(editStr);
	var btn = $("#diyBtn_"+treeNode.UserName);
	if (btn) btn.bind("click", function(){showSendDiv(treeNode.UserName,treeNode.name);});
}
function removeHoverDom(treeId, treeNode) {
	if (treeNode.isParent) return;
	$("#diyBtn_"+treeNode.UserName).unbind().remove();
	$("#diyBtn_space_" +treeNode.UserName).unbind().remove();
}
function zTreeOnDblclick(event, treeId, treeNode) {
	if(treeNode.isParent)return;
	showSendDiv(treeNode.UserName,treeNode.name);
}
function refreshTree() {
	setting.expandSpeed = ($.browser.msie && parseInt($.browser.version)<=6)?"":"fast";
	setting.asyncUrl= "CheckMessage.asp";
	var nodes = [];
	setting.async = true;
	zTree1 = $("#users").zTree(setting, nodes);
	$("#loginusers").show('fast');
	$('#loginusers').easydrag(); 
	$("#loginusers").setHandler("usermove"); 
}
function showSendDiv(uid,uname){
	$('#incept').text(uname);
	$('#sender').text('我');
	$('#senderuserid').val('');
	$('#inceptuserid').val(uid);
	$('#messReply').attr('disabled','disabled');
	$('#messSend').removeAttr("disabled");
	initmessList();
}
</SCRIPT>
</HEAD>
<!--#include file="CheckAdmin.asp"-->
<BODY scroll="no" topmargin="0" bottom="0" leftmargin="0" rightmargin="0">
<div id="loginusers" style="width:170px;display:none;top:75px;right:20px;position:absolute;">
<div class="tablemenu" style="border:1px solid #99BBE8;border-width: 1px;">
	<div style="height:24px; float:left;position: absolute;top:5px;width:100%;" id="usermove"><strong>&nbsp;同事列表</strong></div>
	<div style="height:16px;width:18px;float:right;position: absolute;top:5px;right:24px;background: url('Images/images/flexigrid/load.png') no-repeat;color:#BBBBBB;" title="刷新" onClick="refreshTree()"></div>
	<div style="height:16px;width:18px;float:right;position: absolute;top:5px;right:4px;background: url('Images/panel_tools.gif') no-repeat scroll -16px 0 transparent;color:#BBBBBB;" title="关闭" onClick="$(this).parent().parent().hide('fast');">	</div>
</div>
	<div style="background:#E8EFFF;border:1px solid #99BBE8;border-width:0 1px 1px 1px;height:400px;overflow:auto;">
	<ul id="users" class="tree"></ul>
	</div>
</div>
<div id="messagelist" style="width:400px;height:450px;top:75px;display:none;left:40%;position:absolute;background:#E8EFFF;">
<div class="tablemenu" style="border:1px solid #99BBE8;border-width: 1px;">
	<div style="height:24px; float:left;position: absolute;top:5px;width:100%;" id="messmove"><strong>&nbsp;短信管理</strong></div>
	<div style="height:16px;width:18px;float:right;position: absolute;top:5px;right:4px;background: url('Images/panel_tools.gif') no-repeat scroll -16px 0 transparent;color:#BBBBBB;" title="关闭" onClick="$(this).parent().parent().hide('fast');">	</div>
</div>
<table id="messflex" style="display:none"></table>
<div style="height:20">
	<div style="height:20;width:200px;float:left;">标题：<input type="text" class="textfield" id="title" name="title"></div>
	<div style="height:20;width:200px;float:left;">
	<input type="hidden" id="inceptuserid" name="inceptuserid">
	<input type="hidden" id="senderuserid" name="senderuserid">
	[<div style="display:inline;" id="sender"></div>]
	发送给：
	[<div style="display:inline;" id="incept"></div>]
	</div>
</div>
<div><textarea style="height:75px;width:100%;" id="content" name="content"></textarea></div>
<div style="height:22">
	<div style="right:5px;bottom:2px;position:absolute;border:0px;"><input type="button" class="nBtn" value="回复" onClick="messReply()" id="messReply">&nbsp;<input type="button" class="nBtn" value="发送消息" onClick="messSend()" id="messSend"></div>
</div>
</div>
<TABLE height="100%" cellSpacing="0" cellPadding="0" border="0" width="100%">
  <TR>
    <TD colSpan="3">
	<IFRAME 
      style="Z-INDEX:1; VISIBILITY:inherit; WIDTH:100%; HEIGHT:30" 
      name="topFrame" id="topFrame" marginWidth="0" marginHeight="0"
      src="SysTop.asp" frameBorder="0" noResize scrolling="no">	</IFRAME>	</TD>
  </TR>
  <TR>
    <TD width="170" height="100%" rowspan="2" align="middle" bgcolor="#C6D9F4" style="border:1px solid #99BBE8;border-width: 0 1px 0 0;" id="frameTitle" >
    <IFRAME 
      style="Z-INDEX:2; VISIBILITY:inherit; WIDTH:170; HEIGHT:100%" 
      name="leftFrame" id="leftFrame"  marginWidth="0" marginHeight="0"
      src="<% If Instr(session("AdminPurviewFLW"),"|10,")>0 Then %>FLW_SysLeft.asp<% Else %>
SysLeft.asp<% End If %>
" frameBorder="0" noResize scrolling="no">
	</IFRAME>
	</TD>
    <TD width="10" height="100%" rowspan="2" align="center" bgcolor="#D2E0F2" onClick="switchSysBar()">
	<FONT style="FONT-SIZE: 10px; CURSOR: hand; COLOR: #ffffff; FONT-FAMILY: Webdings">
	  <SPAN id="switchPoint">3</SPAN>
	</FONT>
	</TD>
    <TD height="30" style="border:1px solid #99BBE8;border-width: 1px 0 0 1px;">
	<iframe 
      style="Z-INDEX:3; VISIBILITY:inherit; WIDTH:100%; HEIGHT:30"
	  name="headFrame" id="headFrame" marginwidth="16" marginheight="3"
	  src="SysHead.asp" frameborder="0"  scrolling="no">
	</iframe>
	</TD>
  </TR>
  <TR>
    <TD height="100%" style="border:1px solid #99BBE8;border-width: 0 0 0 1px;">
	<iframe 
      style="Z-INDEX:4; VISIBILITY:inherit; WIDTH:100%; HEIGHT:100%"
	  name="mainFrame" id="mainFrame" marginwidth="16" marginheight="16"
	  src="SysCome.asp" frameborder="0" noresize scrolling="yes">
	</iframe>
	</TD>
  </TR>
</TABLE>
<script language="javascript">
//消息窗口
function messReply(){
	$('#incept').text($('#sender').text());
	$('#inceptuserid').val($('#senderuserid').val());
	$('#messReply').attr('disabled','disabled');
	$('#messSend').removeAttr("disabled");
}
function messSend(){
	if($('#incept').text()==''||$('#content').val()==''){
		alert('接收人、发送内容不能为空！');
		$('#messSend').attr('disabled','disabled');
		return false;
	}
	$('#messReply').attr('disabled','disabled');
	$('#messSend').attr('disabled','disabled');
  $.post('CheckMessage.asp?key=messSend',{
		incept:$('#incept').text(),
		inceptuserid:$('#inceptuserid').val(),
		title:$('#title').val(),
		content:$('#content').val()
	},function(data){
		$("#messflex").flexReload();
		$('#title').val('');
		$('#content').val('');
		$('#messSend').removeAttr("disabled");
  });
}
function initmessList(){
	$("#messflex").flexReload();
	$("#messagelist").show();
	$('#messagelist').easydrag(); 
	$("#messagelist").setHandler("messmove"); 
}
$("#messflex").flexigrid({
	url: 'CheckMessage.asp?key=messagelist',
	dataType: 'json',
	colModel : [
	{display: '新', name : 'newflag', width : 20, sortable : true, align: 'left'},
	{display: '发送', name : 'sender', width : 50, sortable : true, align: 'left'},
	{display: '接收', name : 'incept', width : 50, sortable : true, align: 'left'},
	{display: '时间', name : 'sendtime', width : 70, sortable : true, align: 'left'},
	{display: '标题', name : 'title', width : 100, sortable : true, align: 'left'},
	{display: '内容', name : 'content', width : 130, sortable : true,align: 'left'},
	{display: 'id', name : 'SerialNum', width : 30, hide:true,toggle:false},
	{display: '接收人编号', name : 'inceptuserid', width : 30, hide:true,toggle:false},
	{display: '发送人编号', name : 'senderuserid', width : 30, hide:true,toggle:false}
		],
	buttons : [
		{name: '删除',  onpress : test},
		{separator: true},
		{name: '我接收',  onpress : test},
		{separator: true},
		{name: '我发送',  onpress : test},
		{separator: true}
		],
	onRowDblclick:rowdbclick,
	sortname: "sendtime",
	sortorder: "desc",
	singleSelect: false,
	striped:true,//
	rp: 10,
	usepager: true,
	title: '短信列表',
	showTableToggleBtn: true,
	autoload: false,
	width:400,
	height:200
});
function rowdbclick(rowData){
	$('#incept').text($(rowData).data("incept"));
	$('#sender').text($(rowData).data("sender"));
	$('#inceptuserid').val($(rowData).data("inceptuserid"));
	$('#senderuserid').val($(rowData).data("senderuserid"));
	$('#title').val($(rowData).data("title"));
	$('#content').val($(rowData).data("content"));
	$('#messReply').removeAttr("disabled");
	$('#messSend').attr('disabled','disabled');
	if($(rowData).data("newflag")=="新"){
		$.post('CheckMessage.asp?key=messRead',{
			SerialNum:$(rowData).data("SerialNum")
		});
	}
}
function test(com,grid){
	if (com=='删除')
		{
			if($('.trSelected', grid).length==0){alert('请先选择一条记录再进行操作！');return false;}
			var AllSNum="";
			$('.trSelected', grid).each(function(i,fieldone){
				AllSNum=AllSNum+fieldone.id.replace("row","")+",";
			});
			AllSNum=AllSNum.substr(0,AllSNum.length-1);
			if (confirm('确定要删除单号为：' + AllSNum + ' 的记录?')){
				$.post('CheckMessage.asp?key=messDel',{"SerialNum":AllSNum},function(data){
				if(data.indexOf("###")==-1) alert(data);
				else $("#messflex").flexReload();
				});
			}
		}
	else if (com=='我接收')
		{
			$("#messflex").flexOptions({newp: 1, params:[{name:"undo",value:1}]
			});
			$("#messflex").flexReload();
		}
	else if (com=='我发送')
		{
			$("#messflex").flexOptions({newp: 1, params:[{name:"undo",value:2}]
			});
			$("#messflex").flexReload();
		}
}
</script>
</BODY>
</HTML>