<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<HTML xmlns="http://www.w3.org/1999/xhtml">
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
  $("#type4search").show("slow");
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
	$('#listDiv').load("SoftwareDevelopDetails.asp #listtable",{
	  showType:'DetailsList',
	  Page:arr[0],
	  flag4search:$('#flag4search').val(),
	  type4search:$('#type4search').val(),
	  seachword:$('#seachword').val()
	 },function(response, status, xhr){
	  if (status =="success") {
	    pageNavigation('pageN',arr,response.split('###')[1],response.split('###')[2],'showDiv');
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
	//产生分页导航栏
}
//处理添加按钮
function showpadd(obj,sid){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#addShowDiv').load("SoftwareDevelopDetails.asp #AddandEditdiv",{
	  showType:'AddEditShow',
	  detailType:obj,
	  SerialNum:sid
	},function(response, status, xhr){
	  if (status =="success") {
		$("#type4search").hide("slow");
	    $("#addDiv").show("slow");
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
}
//处理提交事务
function toSubmit(){
  $.post('SoftwareDevelopDetails.asp?showType=DataProcess',$("#AddForm").serialize(),function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	else pageN(0);
  });
  $("#addDiv").hide("slow");
  $("#type4search").show("slow");
}
//处理选择打样单号事务
function getInfo(obj){
  if($("#"+obj).val()==''){alert("对应编号不能为空！");return false;}
	$.get("SoftwareDevelopDetails.asp", { showType: "getInfo",detailType: obj, InfoID: $("#"+obj).val() },
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
		   $("#Product_td").html(data.split('###')[8]);
		   }
		   else if(obj=="Register"){
		   $("#Register").val(data.split('###')[1]);
		   $("#RegisterName").val(data.split('###')[2]);
		   $("#Department").val(data.split('###')[3]);
		   $("#Departmentname").val(data.split('###')[4]);
		   }
		 }
	   });
}
//选择项目触发
function changexm(){
  var shenqxmval=$("input:radio:checked").val();
  if(shenqxmval=="网络维修"){
    var strselect="<select name='DeviceName' id='DeviceName' style='WIDTH: 140;'>";
	strselect+="<option value='电脑主机'>电脑主机</option>";
	strselect+="<option value='键盘'>键盘</option>";
	strselect+="<option value='显示器'>显示器</option>";
	strselect+="<option value='鼠标'>鼠标</option>";
	strselect+="<option value='光驱'>光驱</option>";
	strselect+="<option value='网络连接问题'>网络连接问题</option>";
	strselect+="<option value='电源'>电源</option>";
	strselect+="<option value='打印机'>打印机</option>";
	strselect+="<option value='U盘/硬盘'>U盘/硬盘</option>";
	strselect+="<option value='申请上网'>申请上网</option>";
	strselect+="<option value='其他（需注明）'>其他（需注明）</option>";
	strselect+="</select>";
    $("#DeviceName_td").html(strselect);
  }else if(shenqxmval=="行政维修"){
    $("#DeviceName_td").html("<input name='DeviceName' type='text' class='textfield' id='DeviceName' style='WIDTH: 140;' maxlength='100'>");
  }
}

</script>
</HEAD>
<BODY>
<%
if Instr(session("AdminPurviewFLW"),"|205,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden;">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="position:fixed !important;position:absolute;top:0;margin:0 auto; ">
<font color="#FF0000"><strong>程序开发处理进度汇总</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
<font style="background-color:#EBF2F9"><a href="javascript:$('#flag4search').val('0');pageN(0);">未审核</a></font>&nbsp;
<font style="background-color:#ffff66"><a href="javascript:$('#flag4search').val('1');pageN(0);">已经审核</a></font>&nbsp;
<font style="background-color:#ff99ff"><a href="javascript:$('#flag4search').val('2');pageN(0);">开发中</a></font>&nbsp;
<font style="background-color:#66ff66"><a href="javascript:$('#flag4search').val('3');pageN(0);">开发完毕</a></font>&nbsp;
<input type="hidden" name="flag4search" id="flag4search" value="">
<input type="text" name="seachword" id="seachword" style='HEIGHT: 18px;WIDTH: 80px;'>
<input type="button" name="seachbutton" id="seachbutton" onClick="pageN(arr)" value="查找" style='HEIGHT: 18px;WIDTH: 40px;'>
<input type="button" name="addbutton" id="button" onClick="showpadd('AddNew','')" value="添加" style='HEIGHT: 18px;WIDTH: 40px;'>
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