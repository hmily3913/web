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
   if(el<11){
     rep = "ReportCome.asp?sub="+el;
     parent.frames["mainFrame"].location.href=rep;
   }
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
    <tr style="cursor: hand;"><td class="SystemLeft-header">营销报表系统
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
      <td class="SystemLeft"><a href="salesys/OrderChange.asp" target="mainFrame" onClick='changeAdminFlag("订单变更率")'>订单变更率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="salesys/OrderDeliverRate.asp" target="mainFrame" onClick='changeAdminFlag("订单出货达成率")'>订单出货达成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="salesys/GatherPrompt.asp" target="mainFrame" onClick='changeAdminFlag("收款完成率")'>收款完成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="salesys/Custinterview.asp" target="mainFrame" onClick='changeAdminFlag("拜访/接待客户")'>拜访/接待客户</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/CustomInspect.html" target="mainFrame" onClick='changeAdminFlag("客户验货处理")'>客户验货处理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="salesys/ChangeOrderDeal.html" target="mainFrame" onClick='changeAdminFlag("订单变更后续处理")'>订单变更后续处理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="salesys/OrderForecast.html" target="mainFrame" onClick='changeAdminFlag("业务接单预测")'>业务接单预测</a></td>
    </tr>
  </table>
</div>
<%'end if
'if Instr(session("AdminPurview"),"|20,")>0 then
%>
<div id="main2" onclick=expandIt(2)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">采购报表系统
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
      <td class="SystemLeft"><a href="purchasesys/ReturnRate.asp" target="mainFrame" onClick='changeAdminFlag("采购物料退货率")'>采购物料退货率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="purchasesys/AnxiousBillRate.asp" target="mainFrame" onClick='changeAdminFlag("采购急单率")'>采购急单率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="purchasesys/SupplyProm.asp" target="mainFrame" onClick='changeAdminFlag("供货及时率")'>供货及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="purchasesys/SpecialAccept.asp" target="mainFrame" onClick='changeAdminFlag("特采件数")'>特采件数</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="purchasesys/SupplyQualified.asp" target="mainFrame" onClick='changeAdminFlag("供货批次合格率")'>供货批次合格率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="purchasesys/DevelopReach.asp" target="mainFrame" onClick='changeAdminFlag("材料开发达成率")'>材料开发达成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="purchasesys/SupplierEvaluat.html" target="mainFrame" onClick='changeAdminFlag("供应商评估")'>供应商评估</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="purchasesys/MaterialQuote.html" target="mainFrame" onClick='changeAdminFlag("物料报价")'>物料报价</a></td>
    </tr>
  </table>
</div>
<%'end if
'if Instr(session("AdminPurview"),"|30,")>0 then
%>

<div id="main3" onclick=expandIt(3)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">分厂报表系统
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
      <td class="SystemLeft"><a href="manusys/DeliveryReach.asp" target="mainFrame" onClick='changeAdminFlag("分厂生产交期达成率")'>分厂生产交期达成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/PMDReach.asp" target="mainFrame" onClick='changeAdminFlag("生管交期达成率")'>生管交期达成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/ProductConsum.asp" target="mainFrame" onClick='changeAdminFlag("制程超耗率")'>制程超耗率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/FinishRemake.asp" target="mainFrame" onClick='changeAdminFlag("成品返工件数")'>成品返工件数</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/YieldReach.asp" target="mainFrame" onClick='changeAdminFlag("产量目标达成率")'>产量目标达成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/QCReportQuery.asp" target="mainFrame" onClick='changeAdminFlag("品保周期报表查询")'>品保周期报表查询</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/WorkshopBadProd.html" target="mainFrame" onClick='changeAdminFlag("车间良品及不良品统计")'>车间良品及不良品统计</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/CopperFilm.html" target="mainFrame" onClick='changeAdminFlag("铜模菲林登记表")'>铜模菲林登记表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/InnerProductPrice.html" target="mainFrame" onClick='changeAdminFlag("相互加工单价")'>相互加工单价</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/InnerProduct.html" target="mainFrame" onClick='changeAdminFlag("相互加工明细")'>相互加工明细</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/OutProduct.html" target="mainFrame" onClick='changeAdminFlag("受托加工明细")'>受托加工明细</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/Jishi.html" target="mainFrame" onClick='changeAdminFlag("员工计时单")'>员工计时单</a></td>
    </tr>
  </table>
</div>
<div id="main13" onclick=expandIt(13)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">生管报表系统
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
		</td>
    </tr>
  </table>
</div>
<div id="sub13" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="manusys/PMDReach.asp" target="mainFrame" onClick='changeAdminFlag("生管交期达成率")'>生管交期达成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="pmsys/PutProductPlan.html" target="mainFrame" onClick='changeAdminFlag("生管投产计划表")'>生管投产计划表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="pmsys/OrderDetails.html" target="mainFrame" onClick='changeAdminFlag("生管订单状态表")'>生管订单状态表</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="pmsys/ProductCycle.html" target="mainFrame" onClick='changeAdminFlag("生产前置周期")'>生产前置周期</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="FLW/AbnomalOrder.html" target="mainFrame" onClick='changeAdminFlag("物料起订量管理追踪")'>物料起订量管理追踪</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/PMProductCycle.html" target="mainFrame" onClick='changeAdminFlag("生管产品生产计划周期")'>生管产品生产计划周期</a></td>
    </tr>
  </table>
</div>
<%'end if
'if Instr(session("AdminPurview"),"|40,")>0 then
%>

<div id="main4" onclick=expandIt(4)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">品保报表系统
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
      <td class="SystemLeft"><a href="purchasesys/SupplyQualified.asp" target="mainFrame" onClick='changeAdminFlag("供货批次合格率")'>供货批次合格率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/FinishQualified.asp" target="mainFrame" onClick='changeAdminFlag("成品出货检验合格率")'>成品出货检验合格率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/UnQualifiMtrDeal.asp" target="mainFrame" onClick='changeAdminFlag("不合格来料处理率")'>不合格来料处理率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/ComeCheckProm.asp" target="mainFrame" onClick='changeAdminFlag("进料检验的及时率")'>进料检验的及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/ComeCheckAccur.asp" target="mainFrame" onClick='changeAdminFlag("进料检验错误率")'>进料检验错误率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/CustomComplain.asp" target="mainFrame" onClick='changeAdminFlag("客户投诉")'>客户投诉</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/CheckProm.asp" target="mainFrame" onClick='changeAdminFlag("实验室检测及时率")'>实验室检测及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/FirstCheckAccur.asp" target="mainFrame" onClick='changeAdminFlag("首检错误")'>首检错误</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/QCReportQuery.asp" target="mainFrame" onClick='changeAdminFlag("品保周期报表查询")'>品保周期报表查询</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/SupplyQCSort.html" target="mainFrame" onClick='changeAdminFlag("供方来料品质名次")'>供方来料品质名次</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/SupplyQCTally.html" target="mainFrame" onClick='changeAdminFlag("品保来料点检记录查询")'>品保来料点检记录查询</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/IQCDeduction.html" target="mainFrame" onClick='changeAdminFlag("IQC扣款记录查询")'>IQC扣款记录查询</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/CustomInspect.html" target="mainFrame" onClick='changeAdminFlag("客户验货处理")'>客户验货处理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="qcsys/EnvironTest.html" target="mainFrame" onClick='changeAdminFlag("环保检测")'>环保检测</a></td>
    </tr>
  </table>
</div>
<%'end if
'if Instr(session("AdminPurview"),"|50,")>0 then
%>

<div id="main5" onclick=expandIt(5)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">仓库报表系统
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
      <td class="SystemLeft"><a href="stocksys/ReceiveProm.asp" target="mainFrame" onClick='changeAdminFlag("收料及时率")'>收料及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="stocksys/SendmReach.asp" target="mainFrame" onClick='changeAdminFlag("发料达成率")'>发料达成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="stocksys/FinishDeliveryRate.asp" target="mainFrame" onClick='changeAdminFlag("成品出货及时率")'>成品出货及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="stocksys/CardAccuracyRate.asp" target="mainFrame" onClick='changeAdminFlag("帐卡物准确率")'>帐卡物准确率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="stocksys/InventoryAmount.asp" target="mainFrame" onClick='changeAdminFlag("存货金额")'>存货金额</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="stocksys/PickupProm.asp" target="mainFrame" onClick='changeAdminFlag("提货及时率")'>提货及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="stocksys/SendCarProm.asp" target="mainFrame" onClick='changeAdminFlag("派车及时率")'>派车及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="stocksys/OrderStockMtr.html" target="mainFrame" onClick='changeAdminFlag("库存订单物料")'>库存订单物料</a></td>
    </tr>
  </table>
</div>
<%'end if
'if Instr(session("AdminPurview"),"|60,")>0 then
%>

<div id="main6" onclick=expandIt(6)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">生技报表系统
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
      <td class="SystemLeft"><a href="technologysys/Moldrepair.asp" target="mainFrame" onClick='changeAdminFlag("模具维修及时率")'>模具维修及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="technologysys/Moldmake.asp" target="mainFrame" onClick='changeAdminFlag("模具制作及时率")'>模具制作及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="technologysys/Devicerepair.asp" target="mainFrame" onClick='changeAdminFlag("设备维修及时率")'>设备维修及时率</a></td>
    </tr>
  </table>
</div>
<%'end if
'if Instr(session("AdminPurview"),"|70,")>0 then
%>

<div id="main7" onclick=expandIt(7)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">工程报表系统
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
      <td class="SystemLeft"><a href="engineersys/ProofingFinishRate.asp" target="mainFrame" onClick='changeAdminFlag("打样按期完成率")'>打样按期完成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="engineersys/BomFinish.asp" target="mainFrame" onClick='changeAdminFlag("BOM表按期完成率")'>BOM表按期完成率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="engineersys/NewProductDev.asp" target="mainFrame" onClick='changeAdminFlag("新产品开发件数")'>新产品开发件数</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="engineersys/NewProductOrder.asp" target="mainFrame" onClick='changeAdminFlag("新产品接单额")'>新产品接单额</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="engineersys/TryQualified.asp" target="mainFrame" onClick='changeAdminFlag("产前试做合格率")'>产前试做合格率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="engineersys/ProofingBill.html" target="mainFrame" onClick='changeAdminFlag("打样单统计表")'>打样单统计表</a></td>
    </tr>
  </table>
</div>
<%'end if
'if Instr(session("AdminPurview"),"|80,")>0 then
%>

<div id="main8" onclick=expandIt(8)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">财务报表系统
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
		</td>
    </tr>
  </table>
</div>
<div id="sub8" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="financesys/MaterialPurchaseRate.asp" target="mainFrame" onClick='changeAdminFlag("同种物料的采购频率")'>同种物料的采购频率</a></td>
    </tr>
<!--    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="financesys/PayPrompt.asp" target="mainFrame" onClick='changeAdminFlag("付款及时率")'>付款及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="financesys/GatherPrompt.asp" target="mainFrame" onClick='changeAdminFlag("收款到账率")'>收款到账率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="financesys/PayCheckPrompt.asp" target="mainFrame" onClick='changeAdminFlag("收款到账率")'>应付对账及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="financesys/GatherCheckPrompt.asp" target="mainFrame" onClick='changeAdminFlag("收款到账率")'>应收对账及时率</a></td>
    </tr>-->
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="financesys/ClientBillInfo.html" target="mainFrame" onClick='changeAdminFlag("客户开票资料管理")'>客户开票资料管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="financesys/BillCount.html" target="mainFrame" onClick='changeAdminFlag("公司开票统计管理")'>公司开票统计管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="financesys/SupplyBill.html" target="mainFrame" onClick='changeAdminFlag("供方开票管理")'>供方开票管理</a></td>
    </tr>
  </table>
</div>
<%'end if
'if Instr(session("AdminPurview"),"|90,")>0 then
%>

<div id="main9" onclick=expandIt(9)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">人资报表系统
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
		</td>
    </tr>
  </table>
</div>
<div id="sub9" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="hrsys/PersonnelLoss.asp" target="mainFrame" onClick='changeAdminFlag("人员流失率")'>人员流失率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="hrsys/RecruitmentEffRate.asp" target="mainFrame" onClick='changeAdminFlag("招聘有效率")'>招聘有效率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="hrsys/RecruitmentTimely.asp" target="mainFrame" onClick='changeAdminFlag("招聘及时率")'>招聘及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="hrsys/ClassRate.asp" target="mainFrame" onClick='changeAdminFlag("开课率")'>开课率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="hrsys/SalaryCalculAccur.asp" target="mainFrame" onClick='changeAdminFlag("薪资计算准确率")'>薪资计算准确率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="hrsys/SalaryCalculTimely.asp" target="mainFrame" onClick='changeAdminFlag("薪资计算及时率")'>薪资计算及时率</a></td>
    </tr>
  </table>
</div>
<%'end if
'if Instr(session("AdminPurview"),"|100,")>0 then
%>

<div id="main10" onclick=expandIt(10)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">行政报表系统
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
		</td>
    </tr>
  </table>
</div>
<div id="sub10" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/Networkrepair.asp" target="mainFrame" onClick='changeAdminFlag("网络检修及时率")'>网络检修及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/Logisticrepair.asp" target="mainFrame" onClick='changeAdminFlag("后勤维修及时率")'>后勤维修及时率</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/ElectWater.asp" target="mainFrame" onClick='changeAdminFlag("水电管理")'>宿舍水电管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/SendCarMana.asp?Result=Search&Page=1" target="mainFrame" onClick='changeAdminFlag("外出管理")'>外出管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/SecurityMana.asp" target="mainFrame" onClick='changeAdminFlag("保安管理")'>保安管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/ImproveProposal.asp" target="mainFrame" onClick='changeAdminFlag("改善提案")'>改善提案</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/GoodsOutPassMana.asp" target="mainFrame" onClick='changeAdminFlag("物品携出管理")'>物品携出管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="managesys/ProjectApplicate.asp" target="mainFrame" onClick='changeAdminFlag("企业项目申请进度")'>企业项目申请进度</a></td>
    </tr>
  </table>
</div>
<div id="main12" onclick=expandIt(12)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">基础参数设置
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
		</td>
    </tr>
  </table>
</div>
<div id="sub12" style="display:none">
  <table width="160" border="0" cellspacing="0" cellpadding="0" background="Images/SysLeft_bg_link.gif">
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/PerformancePost.html" target="mainFrame" onClick='changeAdminFlag("绩效考核岗位设置")'>绩效岗位设置</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/PerformanceItem.html" target="mainFrame" onClick='changeAdminFlag("绩效考核项目设置")'>绩效项目设置</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/PostToPerson.html" target="mainFrame" onClick='changeAdminFlag("岗位人员设置")'>岗位人员设置</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/PostToItem.html" target="mainFrame" onClick='changeAdminFlag("岗位项目设置")'>岗位项目设置</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/WageBase.html" target="mainFrame" onClick='changeAdminFlag("职等绩效工资基数")'>职等绩效工资基数</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/PersonWageBase.html" target="mainFrame" onClick='changeAdminFlag("个人绩效工资基数")'>个人绩效工资基数</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="parametersys/PerformanceSum.html" target="mainFrame" onClick='changeAdminFlag("职员绩效汇总表")'>职员绩效汇总表</a></td>
    </tr>
  </table>
</div>
<%'end if
if Instr(session("AdminPurview"),"|110,")>0 then
%>

<div id="main11" onclick=expandIt(11)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">系统管理
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
      <td class="SystemLeft"><a href="smmsys/Purview.asp" target="mainFrame" onClick='changeAdminFlag("权限管理")'>权限管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="smmsys/PurviewSet.asp" target="mainFrame" onClick='changeAdminFlag("权限分配")'>权限分配</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="smmsys/PermissionGroup.asp" target="mainFrame" onClick='changeAdminFlag("权限组管理")'>权限组管理</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="smmsys/DateUpdate.asp" target="mainFrame" onClick='changeAdminFlag("数据更新")'>数据更新</a></td>
    </tr>
    <tr>
      <td width="36" height="22"></td>
      <td class="SystemLeft"><a href="smmsys/ReporAll.asp" target="mainFrame" onClick='changeAdminFlag("图表分析")'>图表分析</a></td>
    </tr>
  </table>
</div>
<%end if%>
<table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
  <tr style="cursor: hand;"><td class="SystemLeft-header"><a href="smmsys/PassUpdate.asp" target="mainFrame" onClick='changeAdminFlag("密码修改")'><font color="#15428b">密码修改</font></a></td>
  </tr>
</table>
<table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
  <tr style="cursor: hand;"><td class="SystemLeft-header"><a href="javascript:AdminOut()"><font color="#15428b">退出登录</font></a></td>
  </tr>
</table>
<table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
  <tr style="cursor: hand;"><td align="right"><a href="help.html" target="mainFrame" onClick='changeAdminFlag("环境检查")' style="text-decoration:underline"><font color="#15428B">环境检查</font></a></td>
  </tr>
</table>
</BODY>
</HTML>