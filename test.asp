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
/* ------------- 右键菜单 -----------------  */

div#rMenu {
	background-color:#555555;
	text-align: left;
	padding:2px;
}

div#rMenu ul {
	margin:1px 0;
	padding:0 5px;
	cursor: pointer;
	list-style: none outside none;
	background-color:#DFDFDF;
}
div#rMenu ul li {
	margin:0;
	padding:2px 0;
}
</style>
<script language="javascript" src="Script/Admin.js"></script>
<script language="javascript" src="Script/jquery-1.5.2.min.js"></script>
<link rel="stylesheet" href="../Images/zTreeStyle.css">
<script language="javascript" src="Script/jquery.ztree-2.6.min.js"></script>
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
/*   if(el<11){
     rep = "ReportCome.asp?sub="+el;
     parent.frames["mainFrame"].location.href=rep;
   }*/
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
var zTree1;
var setting;
var rMenu;
	setting = {
		expandSpeed:"",
		dragCopy:true,
		dragCopy:true,
		keepParent: false,
		keepLeaf: false,
		async: true,
		asyncParam: ["SerialNum"],
		asyncParamOther : {"showType":"getInfo","detailType":"EditPurview"},
		callback: {
			rightClick: zTreeOnRightClick
		}
	};
	$(document).ready(function(){
		refreshTree();
		rMenu = document.getElementById("rMenu");
		$("body").bind("mousedown", 
			function(event){
				if (!(event.target.id == "rMenu" || $(event.target).parents("#rMenu").length>0)) {
					rMenu.style.visibility = "hidden";
				}
			});
	});
	function refreshTree() {
//		var paraUrl=$('#PurviewType').length==0?'RPT':$('#PurviewType').val();
//		setting.asyncUrl= "PurviewDetails.asp?PurviewType="+paraUrl;
		setting.asyncUrl= "data.asp";
		hideRMenu();
		$('#editOne').hide();
		var nodes = [];
		setting.async = true;
		zTree1 = $("#FLWTrue").zTree(setting, nodes);
	}
	//右键菜单
	function zTreeOnRightClick(event, treeId, treeNode) {
		if (!treeNode && event.target.tagName.toLowerCase() != "button" && $(event.target).parents("a").length == 0) {
			zTree1.cancelSelectedNode();
			showRMenu("root", event.clientX, event.clientY);
		} else if (treeNode && !treeNode.noR) {
			zTree1.selectNode(treeNode);
			showRMenu("node", event.clientX, event.clientY);
		}
	}
	function showRMenu(type, x, y) {
		$("#rMenu ul").show();
		if (type=="root") {
			$("#m_edit").hide();
			$("#m_del").hide();
			$("#m_check").hide();
			$("#m_unCheck").hide();
		}
		$("#rMenu").css({"top":y+"px", "left":x+"px", "visibility":"visible"});
	}
	function hideRMenu() {
		if (rMenu) rMenu.style.visibility = "hidden";
	}
	var addCount = 1;
	function addTreeNode() {
		hideRMenu();
		var srcNode=zTree1.getSelectedNode();
		var p_s=0;
		var p_pname=($('#PurviewType').val()=='RPT')?'报表系统':'工作流系统';
		var p_pid="";
		var p_lv=0;
		if(srcNode){
		  p_s=srcNode.SerialNum;
		  p_pname=srcNode.LongName+"-"+srcNode.name;
		  p_pid=srcNode.PermissionID;
		  p_lv=srcNode.level;
		}
		$.post('PurviewDetails.asp?showType=DataProcess&detailType=AddNew',{
		  PermissionClass:$('#PurviewType').val(),
		  PermissionID:p_pid,
		  LongName:p_pname,
		  PSNum:p_s,
		  N_lv:p_lv
		},function(data){
		  var newNode=jQuery.parseJSON(data);
		  zTree1.addNodes(srcNode, newNode);
		});
	}
	function removeTreeNode() {
		hideRMenu();
		var node = zTree1.getSelectedNode();
		if (node) {
			if (node.nodes && node.nodes.length > 0) {
				alert("要删除的节点是父节点，不允许执行删除操作，请先删除子节点！");
				return ;
			} else {
				$.post('PurviewDetails.asp?showType=DataProcess&detailType=Delete',{
				  SerialNum:node.SerialNum
				},function(data){
				  if(data.indexOf("###")==-1) alert(data);
				  else {
					zTree1.removeNode(node);
				  }
				});
			}
		}
	}
	function submitSave(){
		var srcNode = zTree1.getSelectedNode();
		if (!srcNode) {
			alert("请先选中一个节点");
			return;
		}
		$.post('PurviewDetails.asp?showType=DataProcess&detailType=Edit',{
		  PermissionID:$('#PermissionID').val(),
		  PermissionName:$('#PermissionName').val(),
		  SerialNum:srcNode.SerialNum
		},function(data){
		  if(data.indexOf("###")==-1) alert(data);
		  else {
			$('#editOne').hide();
			srcNode.PermissionID=$('#PermissionID').val();
			srcNode.name=$('#PermissionName').val();
			zTree1.updateNode(srcNode, true);
		  }
		});
	}

</script>
</HEAD>
<!--#include file="CheckAdmin.asp"-->

<BODY background="Images/SysLeft_bg.gif" onmouseover="self.status='全心全意为您打造!';return true">
<div id="rMenu" style="position:absolute; visibility:hidden;">
<li>
<ul id="m_add" onclick="addTreeNode();"><li>增加节点</li></ul>
<ul id="m_edit" onclick="renameTreeNode(event);"><li>编辑节点</li></ul>
<ul id="m_del" onclick="removeTreeNode();"><li>删除节点</li></ul>
<ul id="m_reset" onclick="refreshTree();"><li>权限重新载入</li></ul>
</li>
</div>
<div id="editOne" style="position:absolute; visibility:hidden;">
<table width="auto" style="margin:0;text-align: left; padding:0; background-color:#EBF2F9; border:0; color:#FFFFFF;">
<tr>
<td>权限代码</td><td><input type="text" id="PermissionID" name="PermissionID" /></td>
</tr>
<tr>
<td>权限名称</td><td><input type="text" id="PermissionName" name="PermissionName" /></td>
</tr>
<tr>
<td colspan="2">
<input type="button" value="确定" style="width:60px; height:16px; font-size:12px " onclick="submitSave()"/>
<input type="button" value="取消" style="width:60px; height:16px; font-size:12px " onclick="$('#editOne').hide()"/></td>
</tr>
</table>
</div>
<%'if Instr(session("AdminPurview"),"|10,")>0 then%>
<div id="main1" onclick=expandIt(1)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">报表系统
			<div class="panel-tool">
			<div class="accordion-collapse"></div>
			</div>
		</td>
    </tr>
  </table>
</div>
<div id="sub1">
  <table width="160" border="0" cellspacing="0" cellpadding="0" >
    <tr>
      <td width="36" height="22">
      <ul id="FLWTrue" class="tree"></ul>
      </td>
    </tr>
	</table>
</div>
<div id="main2" onclick=expandIt(2)>
  <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
    <tr style="cursor: hand;"><td class="SystemLeft-header">工作平台
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