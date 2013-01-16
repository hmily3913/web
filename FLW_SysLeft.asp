<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<HTML>
<HEAD>
<TITLE>后台管理导航</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012 - zbh-STUDIO" />
<META NAME="Author" CONTENT="---zbh" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="Images/CssAdmin.css">
<style type="text/css">
.panel-tool div {
    cursor: pointer;
    display: block;
    float: right;
    height: 16px;
    margin-left: 2px;
    opacity: 0.6;
    width: 16px;
}
.panel-tool {
    position: absolute;
    right: 5px;
    top: 4px;
}
.accordion-collapse {
    background: url("images/layout_button_down.gif") no-repeat scroll 0 0 transparent;
}
.accordion-expand {
    background: url("images/layout_button_up.gif") no-repeat scroll 0 0 transparent;
}
</style>
<script language="javascript" src="Script/Admin.js"></script>
<script language="javascript" src="Script/jquery-1.5.2.min.js"></script>
<script>
function closewin() {
   if (opener!=null && !opener.closed) {
      opener.window.newwin=null;
      opener.openbutton.disabled=false;
      opener.closebutton.disabled=true;
   }
}

var count=0;//做计数器
var limit=new Array();//用于记录当前显示的哪几个菜单
var countlimit=1;//同时打开菜单数目，可自定义

function expandIt(el) {
   obj = eval("sub" + el);
	 mobj= eval("main" + el);
   if (obj.style.display == "none") {
      obj.style.display = "block";//显示子菜单
			$('.accordion-collapse',$(mobj)).addClass('accordion-expand');
//   if(el<11){
//     rep = "ReportCome.asp?sub="+el;
 ///    parent.frames["mainFrame"].location.href=rep;
  /// }
      if (count<countlimit) {//限制2个
         limit[count]=el;//录入数组
         count++;
      }
      else {
         eval("sub" + limit[0]).style.display = "none";
				 $('.accordion-collapse',$(eval("main" + limit[0]))).removeClass('accordion-expand');
         for (i=0;i<limit.length-1;i++) {limit[i]=limit[i+1];}//数组去掉头一位，后面的往前挪一位
         limit[limit.length-1]=el;
      }
   }
   else {
      obj.style.display = "none";
			$('.accordion-collapse',$(mobj)).removeClass('accordion-expand');
      var j;
      for (i=0;i<limit.length;i++) {if (limit[i]==el) j=i;}//获取当前点击的菜单在limit数组中的位置
      for (i=j;i<limit.length-1;i++) {limit[i]=limit[i+1];}//j以后的数组全部往前挪一位
      limit[limit.length-1]=null;//删除数组最后一位
      count--;
   }
}
</script>
</HEAD>
<!--#include file="CheckAdmin.asp"-->

<BODY background="Images/SysLeft_bg.gif" onmouseover="self.status='全心全意为您打造!';return true">
<%'if Instr(session("AdminPurview"),"|10,")>0 then%>
<div id="main1" onclick=expandIt(1)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;">
      <td class="SystemLeft-header">工作回复平台
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
			</td>
    </tr>
  </table>
</div>
<div id="sub1" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="FLW/IcomFlw.asp" target="mainFrame" onClick='changeAdminFlag("生产任务情况")'>生产任务情况</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="FLW/MtrPurchase.asp" target="mainFrame" onClick='changeAdminFlag("物料采购情况")'>物料采购情况</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="FLW/6sExecute.asp" target="mainFrame" onClick='changeAdminFlag("６Ｓ执行情况")'>６Ｓ执行情况</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="FLW/SamplesHandover.asp" target="mainFrame" onClick='changeAdminFlag("出货样交接表")'>出货样交接表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="FLW/OrderReview.asp" target="mainFrame" onClick='changeAdminFlag("订单评审")'>订单评审</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="FLW/PPBOMAbnormal.asp" target="mainFrame" onClick='changeAdminFlag("已入库未领料处理")'>已入库未领料处理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="FLW/ExpressTransce.html" target="mainFrame" onClick='changeAdminFlag("公文快递收发管制表")'>公文快递收发管制表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="FLW/PrenatalTest.html" target="mainFrame" onClick='changeAdminFlag("产前试做单")'>产前试做单</a></td>
    </tr>
  </table>
</div>
<div id="main2" onclick=expandIt(2)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;">
      <td class="SystemLeft-header">工作单据平台
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
			</td>
    </tr>
  </table>
</div>
<div id="sub2" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/WelfareAdjust.asp" target="mainFrame" onClick='changeAdminFlag("薪资福利津贴变动")'>薪资福利津贴变动</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/ProofingAbnormal.asp" target="mainFrame" onClick='changeAdminFlag("打样异常处理")'>打样异常反馈处理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/OrderAbnormal.asp" target="mainFrame" onClick='changeAdminFlag("订单异常处理")'>订单异常反馈处理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/RepairApplication.asp" target="mainFrame" onClick='changeAdminFlag("维修申请处理")'>维修申请处理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/SoftwareDevelop.asp" target="mainFrame" onClick='changeAdminFlag("程序开发进度汇总")'>程序开发进度汇总</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/InternalWorkLetter.asp" target="mainFrame" onClick='changeAdminFlag("内部工作联络函")'>内部工作联络函</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/PMMTRboard.asp" target="mainFrame" onClick='changeAdminFlag("生管物料看板")'>生管紧急物料看板</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/FinishDelivAbnormal.html" target="mainFrame" onClick='changeAdminFlag("成品发货异常反馈")'>成品发货异常反馈</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/WorkInjury.html" target="mainFrame" onClick='changeAdminFlag("工伤事故报告书")'>工伤事故报告书</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/SoftOperateRule.html" target="mainFrame" onClick='changeAdminFlag("软件操作规程管理")'>软件操作规程管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/AllPersonalCtrl.html" target="mainFrame" onClick='changeAdminFlag("各部门人员出勤情况")'>各部门人员出勤情况</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/ICMOUnEnd.html" target="mainFrame" onClick='changeAdminFlag("生产任务单反结案申请")'>生产任务单反结案申请</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/LeaveApplication.html" target="mainFrame" onClick='changeAdminFlag("离职意向申请表")'>离职意向申请表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/StampUse.html" target="mainFrame" onClick='changeAdminFlag("印章使用申请表")'>印章使用申请表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/VirtualNetwork.html" target="mainFrame" onClick='changeAdminFlag("虚拟网变更单")'>虚拟网变更单</a></td>
    </tr>
  </table>
</div>
<div id="main3" onclick=expandIt(3)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;">
      <td class="SystemLeft-header">文件数据平台
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
			</td>
    </tr>
  </table>
</div>
<div id="sub3" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Offer/Products.html" target="mainFrame" onClick='changeAdminFlag("报价信息数据管理")'>报价信息数据管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Offer/PrintPaper.html" target="mainFrame" onClick='changeAdminFlag("转印纸工艺管理")'>转印纸工艺管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Offer/ReportSubmit.html" target="mainFrame" onClick='changeAdminFlag("报表报送执行进度")'>报表报送执行进度</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Offer/NewMeetingDeal.html" target="mainFrame" onClick='changeAdminFlag("新员工会议意见")'>新员工会议意见</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Offer/MedicalRecord.html" target="mainFrame" onClick='changeAdminFlag("员工体检记录")'>员工体检记录</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Offer/KnifePlate.html" target="mainFrame" onClick='changeAdminFlag("刀板规格统计表")'>刀板规格统计表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Offer/Honor.html" target="mainFrame" onClick='changeAdminFlag("荣誉证书管理")'>荣誉证书管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Offer/Moju.html" target="mainFrame" onClick='changeAdminFlag("模具统计表")'>模具统计表</a></td>
    </tr>
  </table>
</div>
<div id="main5" onclick=expandIt(5)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;">
      <td class="SystemLeft-header">协同办公平台
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
			</td>
    </tr>
  </table>
</div>
<div id="sub5" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OA/Meeting.html" target="mainFrame" onClick='changeAdminFlag("会议管理")'>会议管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Javascript:alert('制作中')" target="mainFrame" onClick='changeAdminFlag("计划管理")'>计划管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OA/Knowledge.html" target="mainFrame" onClick='changeAdminFlag("知识管理")'>知识管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Javascript:alert('制作中')" target="mainFrame" onClick='changeAdminFlag("制度管理")'>制度管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OA/Document.html" target="mainFrame" onClick='changeAdminFlag("文档管理")'>文档管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OA/Signwas.html" target="mainFrame" onClick='changeAdminFlag("签呈管理")'>签呈管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OA/Announce.html" target="mainFrame" onClick='changeAdminFlag("公告通知")'>公告通知</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="OA/Contacts.html" target="mainFrame" onClick='changeAdminFlag("公司通讯录")'>公司通讯录</a></td>
    </tr>
  </table>
</div>
<div id="main6" onclick=expandIt(6)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;">
      <td class="SystemLeft-header">食堂管理
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
			</td>
    </tr>
  </table>
</div>
<div id="sub6" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Dining/Material.html" target="mainFrame" onClick='changeAdminFlag("物料管理")'>物料管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Dining/SellPriceChange.html" target="mainFrame" onClick='changeAdminFlag("菜价调整管理")'>菜价调整管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Dining/StockInOut.html" target="mainFrame" onClick='changeAdminFlag("物料出入管理")'>物料出入管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Dining/Question.html" target="mainFrame" onClick='changeAdminFlag("问卷试题管理")'>问卷试题管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Javascript:alert('制作中')" target="mainFrame" onClick='changeAdminFlag("问卷调查管理")'>问卷调查管理</a></td>
    </tr>
  </table>
</div>
<div id="main7" onclick=expandIt(7)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;">
      <td class="SystemLeft-header">考勤管理平台
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
			</td>
    </tr>
  </table>
</div>
<div id="sub7" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/SendCarMana.asp?Result=Search&Page=1" target="mainFrame" onClick='changeAdminFlag("外出管理")'>外出管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/Overtime.html" target="mainFrame" onClick='changeAdminFlag("加班管理")'>加班管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/UnCardProof.html" target="mainFrame" onClick='changeAdminFlag("未打卡管理")'>未打卡管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/Annualleave.html" target="mainFrame" onClick='changeAdminFlag("调休管理")'>调休管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Bill/Leave.html" target="mainFrame" onClick='changeAdminFlag("请假管理")'>请假管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Attendance/Travel.html" target="mainFrame" onClick='changeAdminFlag("出差管理")'>出差管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="Attendance/AttQuery.asp" target="mainFrame" onClick='changeAdminFlag("考勤查询")'>考勤查询</a></td>
    </tr>
<%
if Instr(session("AdminPurview"),"|1009,")>0 then
%>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/YCDataInput.html" target="mainFrame" onClick='changeAdminFlag("验厂考勤数据导入")'>验厂考勤数据导入</a></td>
    </tr>
<%end if%>
  </table>
</div>

<div id="main4" onclick=expandIt(4)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;">
      <td class="SystemLeft-header">基础资料
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
			</td>
    </tr>
  </table>
</div>
<div id="sub4" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/DepartReport.html" target="mainFrame" onClick='changeAdminFlag("各部门相关报表")'>各部门相关报表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/PersonalCtrl.html" target="mainFrame" onClick='changeAdminFlag("各单位人员管制")'>各单位人员管制</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/EmpEmail.html" target="mainFrame" onClick='changeAdminFlag("员工邮箱维护")'>员工邮箱维护</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/RewardPunish.html" target="mainFrame" onClick='changeAdminFlag("奖惩基础资料")'>奖惩基础资料</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/JopPrice.html" target="mainFrame" onClick='changeAdminFlag("岗位计时单价")'>岗位计时单价</a></td>
    </tr>
  </table>
</div>

<%'end if
if Instr(session("AdminPurviewFLW"),"|110,")>0 then
%>

<div id="main11" onclick=expandIt(11)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;">
      <td class="SystemLeft-header">系统管理
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
			</td>
    </tr>
  </table>
</div>
<div id="sub11" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="smmsys/PassModify.asp" target="mainFrame" onClick='changeAdminFlag("修改密码")'>修改密码</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="smmsys/FLW_PurviewSet.asp" target="mainFrame" onClick='changeAdminFlag("权限分配")'>权限分配</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="smmsys/DateUpdate.asp" target="mainFrame" onClick='changeAdminFlag("数据更新")'>数据更新</a></td>
    </tr>
  </table>
</div>
<%end if%>
<table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
  <tr style="cursor: hand;">
    <td class="SystemLeft-header"><a href="smmsys/PassUpdate.asp" target="mainFrame" onClick='changeAdminFlag("密码修改")'><font color="#15428b">密码修改</font></a></td>
  </tr>
</table>
<table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
  <tr style="cursor: hand;">
    <td class="SystemLeft-header"><a href="javascript:AdminOut()"><font color="#15428b">退出登录</font></a></td>
  </tr>
</table>

</BODY>
</HTML>