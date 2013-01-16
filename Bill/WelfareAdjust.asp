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
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="../Script/jquery.datepick.pack.js"></script>
<script language="javascript" src="../Script/jquery.datepick-zh-CN.js"></script>
<script language="javascript" src="../Script/jquery.easydrag.js"></script>
<script language="javascript">
$(function(){
$('#start_date').datepick({dateFormat: 'yyyy-mm-dd'});
$('#end_date').datepick({dateFormat: 'yyyy-mm-dd'});
});
function mycheckNum(obj){
	if(checkNum(obj)){
		var WantSalary=parseFloat($('#FormerSalary').val())+parseFloat($('#CWantSalary').val());
		$('#WantSalary').val(WantSalary);
	}
}
function changeEmp(){
  if($("#Register").val()==''){alert("员工编号不能为空！");return false;}
	$.get("WelfareAdjustDetails.asp", { showType: "getInfo",detailType: "Emp1", EmpID: $("#Register").val() },
	   function(data){
		 if(data.indexOf("###")==-1)alert("员工编号不存在，请检查！");
		 else{
		   $("#RegisterName").val(data.split('###')[1]);
		   $("#Department").val(data.split('###')[2]);
		   $("#Departmentname").val(data.split('###')[3]);
		   $("#Position").val(data.split('###')[4]);
		   $("#OldPosition").val(data.split('###')[4]);
		   $("#PositionDate").val(data.split('###')[5]);
		   $("#PositionDate4").val(data.split('###')[5]);
		   $("#IntoCompanyDate").val(data.split('###')[5]);
		   $("#IntoCompanyDate6").val(data.split('###')[5]);
		   $("#IntoCompanyDate7").val(data.split('###')[5]);
		   $("#Employment").val(data.split('###')[6]);
		   $("#Employment6").val(data.split('###')[6]);
		   var housefeepart=0;
		   var date1 = data.split('###')[5];
		   var date2 = new Date();
		   var fm=date1.split('-')[1];
		   var sm=date2.getMonth();
		   var dd1=(sm-fm)*30+(date2.getDate()-date1.split('-')[2]);
		   var year = date2.getFullYear() - date1.split('-')[0]+parseInt(dd1/360);
		   if(data.split('###')[6]==1)housefeepart=100+20*year;
		   else housefeepart=60*(data.split('###')[6]*0.6+year*0.4);
		   $("#HousingFee").val(housefeepart);
		 }
	   });
}
function actcheckNum(obj){
  checkNum(obj);
  var fee1=parseFloat($("#HousingFee").val());
  var fee2=parseFloat($("#PartHousingFee").val());
  var feet=fee1+fee2;
  if(feet>400){
    $("#TotalHousingFee").val(400);
	fee1=parseInt(40000*fee1/feet)/100;
	$("#HousingFee").val(fee1);
	fee2=parseInt(40000*fee2/feet)/100;
	$("#PartHousingFee").val(fee2);
  }
  else $("#TotalHousingFee").val(feet);
   
}
function checkEmp2(){
  if($("#PartID").val()==''){alert("员工编号不能为空！");return false;}
	$.get("WelfareAdjustDetails.asp", { showType: "getInfo",detailType: "Emp1", EmpID: $("#PartID").val() },
	   function(data){
		 if(data.indexOf("###")==-1)alert("员工编号不存在，请检查！");
		 else{
		   $("#PartName").val(data.split('###')[1]);
		   $("#PartPosition").val(data.split('###')[4]);
		   $("#PartIntoCompanyDate").val(data.split('###')[5]);
		   $("#PartEmployment").val(data.split('###')[6]);
		   var housefeepart=0;
		   var date1 = data.split('###')[5];
		   var date2 = new Date();
		   var fm=date1.split('-')[1];
		   var sm=date2.getMonth();
		   var dd1=(sm-fm)*30+(date2.getDate()-date1.split('-')[2]);
		   var year = date2.getFullYear() - date1.split('-')[0]+parseInt(dd1/360);
		   
		   if(data.split('###')[6]==1)housefeepart=100+20*year;
		   else housefeepart=60*(data.split('###')[6]*0.6+year*0.4);
		   $("#PartHousingFee").val(housefeepart);
		   actcheckNum(document.getElementById("PartHousingFee"));
//		   housefeepart=parseFloat($("#HousingFee").val())+housefeepart;
//		   if (housefeepart>=400)housefeepart=400;
//		   $("#TotalHousingFee").val(housefeepart);
		 }
	   });
}
function closead1(){
  $("#addDiv").hide("slow");
  $("#addDiv").css("z-index","500");
  $("#changecheck").show("slow");
  $("#leixing").show("slow");
}
function toSubmit(obj){
  if($("#CheckFlag").val()>0){alert("当前状态不允许此操作，请检查！");return false;}
  $.post('WelfareAdjustDetails.asp?showType=DataProcess',$("#AddForm").serialize(),function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	  $("#addDiv").hide("slow");
  $("#addDiv").css("z-index","500");
	  $("#changecheck").show("slow");
      $("#leixing").show("slow");
	  pageN(0);
  });
}
function toSubmit4Check(obj){
//  if($("#CheckFlag").val()>0){alert("当前状态不允许此操作，请检查！");return false;}
  $.post('WelfareAdjustDetails.asp?showType=CheckProcess',$("#AddForm").serialize(),function(data){
    if(data.indexOf("###")==-1) alert("数据异常，请检查！");
	  $("#addDiv").hide("slow");
  $("#addDiv").css("z-index","500");
	  $("#changecheck").show("slow");
	  $("#leixing").show("slow");
	  pageN(0);
  });
}
function showpadd(obj,sid){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#addShowDiv').load("WelfareAdjustDetails.asp #AddandEditdiv",{
	  showType:'AddEditShow',
	  detailType:obj,
	  SerialNum:sid
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#changecheck").hide("slow");
	    $("#leixing").hide("slow");
	    $("#addDiv").show("slow");
			$('#addDiv').easydrag(); 
			$("#addDiv").setHandler("formove"); 
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
}
function closead(){
  $("#ReplyDiv").hide("slow");
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
	$('#listDiv').load("WelfareAdjustDetails.asp #listtable",{
	  showType:'DetailsList',
	  Page:arr[0],
	  checkflag4seach:$('#checkflag4seach').val(),
	  leixing:$('#leixing').val(),
	  startdate:$('#start_date').val(),
	  enddate:$('#end_date').val()
	 },function(response, status, xhr){
	  if (status =="success") {
	    //产生分页导航栏
	    pageNavigation('pageN',arr,response.split('###')[1],response.split('###')[2],'showDiv');
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
}

function changexm(){
  $('#shenqxm1div').hide();
  $('#shenqxm2div').hide();
  $('#shenqxm3div').hide();
  $('#shenqxm4div').hide();
  $('#shenqxm5div').hide();
  $('#shenqxm6div').hide();
  $('#shenqxm7div').hide();
  var shenqxmval=$("input:radio:checked").val();
  if (shenqxmval=="工资调薪"){
    $('#shenqxm1div').slideDown('slow');
  }else if (shenqxmval=="话费补贴"){
    $('#shenqxm2div').slideDown('slow');
  }else if (shenqxmval=="住房补贴"){
    $('#shenqxm3div').slideDown('slow');
  }else if (shenqxmval=="岗位补贴"){
    $('#shenqxm4div').slideDown('slow');
  }else if (shenqxmval=="其他补贴"){
    $('#shenqxm5div').slideDown('slow');
  }else if (shenqxmval=="职等调整"){
    $('#shenqxm6div').slideDown('slow');
  }else if (shenqxmval=="工龄恢复"){
    $('#shenqxm7div').slideDown('slow');
  }
}
function changecheck(){
  $('#checkflag4seach').val($('#changecheck').val());
  pageN(1);
}

function ClickTd(obj,reptype,tdid){
	$('#loading-one').empty().append('页面载入中...').parent().fadeIn('slow');
	//加载list内容，ajax提交
	$('#addShowDiv').load("WelfareAdjustDetails.asp #AddandEditdiv",{
	  showType:'AddEditShow',
	  detailType:'Edit',
	  checkType:reptype,
	  SerialNum:tdid
	},function(response, status, xhr){
	  if (status =="success") {
	    $("#changecheck").hide("slow");
	    $("#leixing").hide("slow");
	    $("#addDiv").show("slow");
			$('#addDiv').easydrag(); 
			$("#addDiv").setHandler("formove"); 
			$('#EffectiveDate').datepick({dateFormat: 'yyyy-mm-dd'});
			$('#EffectiveDate2').datepick({dateFormat: 'yyyy-mm-dd'});
			$('#EffectiveDate3').datepick({dateFormat: 'yyyy-mm-dd'});
			$('#EffectiveDate4').datepick({dateFormat: 'yyyy-mm-dd'});
			$('#EffectiveDate5').datepick({dateFormat: 'yyyy-mm-dd'});
			$('#EffectiveDate6').datepick({dateFormat: 'yyyy-mm-dd'});
			$('#EffectiveDate7').datepick({dateFormat: 'yyyy-mm-dd'});
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
}

function output(){
  window.open("WelfareAdjustDetails.asp?print_tag=1&showType=DetailsList&checkflag4seach="+$('#checkflag4seach').val()+"&leixing="+encodeURI($('#leixing').val())+"&startdate="+encodeURI($('#start_date').val())+"&enddate="+encodeURI($('#end_date').val()),"Print","","false");
}

</script>
</HEAD>
<BODY>
<%
if Instr(session("AdminPurviewFLW"),"|201,")=0 then 
  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
  response.end
end if
'========判断是否具有管理权限
%>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden;">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="margin:0 auto; ">
<font color="#FF0000"><strong>薪资、福利、津贴变动表</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
<font style="background-color:#ffff66">已审核</font>&nbsp;
<font style="background-color:#ff99ff">已确认</font>&nbsp;
<font style="background-color:#B9BBC7">已作废</font>&nbsp;
<input type="hidden" id="checkflag4seach" value="999">

<select id="leixing" name="leixing"  onChange="pageN(1)" style="font-size:12px;height:15px; width:70px; z-index:1">
<option value="">查看全部</option>
<option value="工资调薪">工资调薪</option>
<option value="话费补贴">话费补贴</option>
<option value="住房补贴">住房补贴</option>
<option value="岗位补贴">岗位补贴</option>
<option value="其他补贴">其他补贴</option>
<option value="职等调整">职等调整</option>
<option value="工龄恢复">工龄恢复</option>
</select>
<select id="changecheck" name="changecheck" onChange="changecheck()" style="font-size:12px;height:15px; width:70px; z-index:1">
<option value="999">查看全部</option>
<option value="0">查未审核</option>
<option value="1">直属部门审核</option>
<option value="2">相关部门审核</option>
<option value="3">副总审核</option>
<option value="4">总经理审核</option>
<option value="99">确定实施</option>
<option value="100">作废</option>
</select>

从<input type="text" id="start_date" style="width:80px;height:18px;">
至<input type="text" id="end_date" style="width:80px;height:18px;">
<input type="button" onClick="pageN(1)" value="查询" style='HEIGHT: 18px;WIDTH: 40px;font-size:12px;'>
<input type="button" name="seachbutton" id="seachbutton" onClick="showpadd('Add')" value="提交申请" style='HEIGHT: 18px;WIDTH: 65px;font-size:12px;'>
<input type="button" name="output" id="output" onClick="output()" value="引出" style='HEIGHT: 18px;WIDTH: 40px;font-size:12px;'>
</p>
<div id="ReplyDiv" style="width:590px;height:180px;top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
<form name="ReplyForm" id="ReplyForm" action="test1.asp">
<table id="ReplyTable" border="0" width="100%" cellspacing="0" cellpadding="1" align="center" bgcolor="black" height="100%">
<tr height="24" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
 <td width="60"> 审核人 </td>
 <td width="60">
 <input name="Replyer" type="text" id="Replyer" ></td>
 <td width="60"> 审核日期 </td>
 <td width="60">
 <input name="ReplyDate" type="text" id="ReplyDate" ></td>
 <td width="20" align="right"><img src="../images/close.jpg" onClick="javascript:closead()"></td>
</tr>
<tr width="574" style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%">
<td width="60"> 审核意见 </td>
<td colspan="4">
  <textarea name="ReplyText" id="ReplyText" style="width:500px; height:100px; "></textarea>
</td>
</tr> 
<tr style="background-color: #FFFFFF; background-repeat: repeat; background-attachment: scroll; border-right: 1px solid #000000; border-bottom: 1px solid #000000; background-position: 0%;border-bottom: 1px solid #000000;">
<td valign="bottom" colspan="5" align="center">
<input type="hidden" name="SerialNum" id="SerialNum" value="">
<input type="hidden" name="Keyword" id="Keyword" value="">
&nbsp;<input name="submitSaveEdit" type="button" class="button"  id="submitSaveEdit" value="审核" style="WIDTH: 80;"  >
</td>
</tr>
</table>
</form>
</div>
<div id="listDiv"></div>
<div id="showDiv"></div>
<div id="addDiv" align="left" style="width:'820px';height:'480px';top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
<div id="addShowDiv"></div>
</div>
<div id="hfbz" style="display:none; font-size:9px; z-index:99999;position:fixed !important;position:absolute;top:0;left:0;">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8" >
  <tr height="15">
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>标准</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>总经办</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>管理部</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>财务部</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>营销部</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>品保工程</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>采购部</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>分厂</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>仓储科</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>生产中心</strong></font></td>
    <td width="8%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>综合管理部</strong></font></td>
  </tr>
  <tr bgcolor='#EBF2F9'  height="12">
  <td>250</td>
  <td>总经理</td>
  <td></td>
  <td></td>
  <td></td>
  <td></td>
  <td></td>
  <td></td>
  <td></td>
  <td></td>
  <td></td>
  </tr>
  <tr bgcolor='#EBF2F9' height="12" >
  <td>200</td>
  <td></td>
  <td></td>
  <td></td>
  <td>经理</td>
  <td>副总</td>
  <td></td>
  <td></td>
  <td></td>
  <td>副总</td>
  <td></td>
  </tr>
  <tr bgcolor='#EBF2F9' height="12" >
  <td>150</td>
  <td></td>
  <td></td>
  <td>经理</td>
  <td></td>
  <td></td>
  <td>采购外</td>
  <td></td>
  <td></td>
  <td></td>
  <td>采购</td>
  </tr>
  <tr bgcolor='#EBF2F9' height="12" >
  <td>100</td>
  <td>副总</td>
  <td>经理</td>
  <td></td>
  <td></td>
  <td>经理</td>
  <td>经理</td>
  <td></td>
  <td></td>
  <td></td>
  <td></td>
  </tr>
  <tr bgcolor='#EBF2F9' height="12" >
  <td>80</td>
  <td>总经理助理</td>
  <td>经理</td>
  <td>经理</td>
  <td></td>
  <td>经理</td>
  <td>经理</td>
  <td>厂长</td>
  <td></td>
  <td></td>
  <td></td>
  </tr>
  <tr bgcolor='#EBF2F9' height="12" >
  <td>50</td>
  <td></td>
  <td>科长</td>
  <td></td>
  <td>业务组长</td>
  <td>科长</td>
  <td>科长</td>
  <td>厂长</td>
  <td>车队主管/科长/司机/外协员</td>
  <td>科长</td>
  <td>信息主管</td>
  </tr>
  <tr bgcolor='#EBF2F9' height="12" >
  <td>30</td>
  <td>设备科长 </td>
  <td>招聘专员</td>
  <td>出纳</td>
  <td>司机</td>
  <td>工程师</td>
  <td></td>
  <td></td>
  <td>司 机/外协员/废料处理员</td>
  <td></td>
  <td></td>
  </tr>
  <tr bgcolor='#EBF2F9' height="12" >
  <td>25</td>
  <td>技改工程师</td>
  <td></td>
  <td>科长/erp主管</td>
  <td>外贸/内贸/商务助理</td>
  <td>助理</td>
  <td>模具师</td>
  <td></td>
  <td></td>
  <td></td>
  <td></td>
  </tr>
  <tr bgcolor='#EBF2F9' height="12" >
  <td>15</td>
  <td>电(焊)工/IE专员</td>
  <td>薪资专员/保安组长/网管员/维修工</td>
  <td>会计</td>
  <td>单证员</td>
  <td>线切割/模具工</td>
  <td>采购员(内)</td>
  <td>段 长/车间主任/技术员/刀版师/计划员/物控员</td>
  <td></td>
  <td>外协/成本报价员/采购员内</td>
  <td>网管</td>
  </tr>
  <tr bgcolor='#EBF2F9'  height="12">
  <td>10</td>
  <td></td>
  <td>人事专员/宿管员</td>
  <td>ERP编程员</td>
  <td>文 员/样品管理员</td>
  <td>文 员/打样员/工艺员/制作员/设计员/组长</td>
  <td></td>
  <td>组 长/总统计/备料员</td>
  <td>收料员/仓库员/叉车司机/输单员</td>
  <td></td>
  <td>文员</td>
  </tr>
  <tr bgcolor='#EBF2F9' height="12" >
  <td>5</td>
  <td></td>
  <td>保安/清洁工</td>
  <td></td>
  <td>样品员</td>
  <td></td>
  <td></td>
  <td>统计员</td>
  <td>搬运工</td>
  <td>BOM制表员</td>
  <td></td>
  </tr>
 </table>
</div>
<script language="javascript">
arr[0] = 1;
pageN(arr);
</script>
</div>
</BODY>
</HTML>