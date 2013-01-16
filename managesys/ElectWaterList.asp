<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<head>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<link rel="stylesheet" href="../Images/CssAdmin.css">
<link rel="stylesheet" href="../Images/flexigrid.bbit.css">
<link rel="stylesheet" href="../Images/jquery.datepick.css">
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<script language="javascript" src="../Script/flexigrid.pack.js"></script>
<script language="javascript" src="../Script/jquery.form.js"></script>
<script language="javascript" src="../Script/jquery.easydrag.js"></script>
<script language="javascript" src="../Script/jquery.datepick.pack.js"></script>
<script language="javascript" src="../Script/jquery.datepick-zh-CN.js"></script>
<script language="javascript" src="../Script/jquery.validate.js"></script>
<script language="javascript" src="../Script/CustomAjax.js"></script>
<link rel="stylesheet" href="../Images/jqi.css">
<script language="javascript" src="../Script/jquery-impromptu.3.1.js"></script>
<style type="text/css">
html{
margin:0;
padding:0;
background-color:#333;
font: 12px Arial, Helvetica, sans-serif;}
body{
margin:0 auto 0 auto;
padding:20px;
background-color:#eee;}
h1{margin-top:-10px;}
hr{margin:20px 0;}
</style>
</head>
<body>
<table id="flex1" style="display:none"></table>
<script language="javascript">
function toSubmit(obj){
	$.post('ElectWaterEdit.asp?showType=DataProcess',$("#AddForm").serialize(),function(data){
		if(data.indexOf("###")==-1) alert(data);
	else $("#flex1").flexReload();
	});
	$("#addDiv").hide("slow");
	$("#addDiv").css("z-index","500");
}
$(function(){
	$('#addDiv').easydrag(); 
	$("#addDiv").setHandler("formove"); 
		$('#checkdate').datepick({dateFormat: 'yyyy-mm-dd'});
});
$("#flex1").flexigrid
	(
	{
	url: 'ElectWaterEdit.asp?showType=DetailsList',
	dataType: 'json',
	colModel : [
	{display: '单号', name : 'fbillno', width : 50, sortable : true, align: 'left'},
	{display: '宿舍号', name : 'Ftext', width : 50, sortable : true, align: 'left'},
	{display: '楼号', name : 'Ftext1', width : 50, sortable : true, align: 'left'},
	{display: '楼层', name:'louceng',width : 50, sortable : true, align: 'left'},
	{display: '最大入住数', name : 'maxperson', width : 50, sortable : true, align: 'left'},
	{display: '已入住人数', name : 'sumperson', width : 50, sortable : true, align: 'left'},
	{display: '房间面积', name : 'FDecimal', width : 50, sortable : true, align: 'left'},
	{display: '水表度数', name : 'waternum', width : 50, sortable : true, align: 'left'},
	{display: '电表度数', name : 'electnum', width : 50, sortable : true, align: 'left'},
	{display: '热水度数', name : 'HotWater', width : 50, sortable : true, align: 'left'},
	{display: 'a.fid', name : 'a.fid', width : 20, sortable : true, align: 'left',hide:true,toggle:false}
		],
	buttons : [
		{name: '增加',  onpress : test},
		{separator: true},
		{name: '维护',  onpress : test},
		{separator: true},
		{name: '删除',  onpress : test},
		{separator: true},
		{name: '查询',  onpress : test},
		{separator: true},
		{name: '返回',  onpress : test},
		{separator: true}
		],
	searchitems : [
		{display: '宿舍号', name : 'Ftext', isdefault: true},
		{display: '单号', name : 'fbillno'}
		],
	onRowDblclick:rowdbclick,
	sortname: "Ftext",
	sortorder: "desc",
	singleSelect: true,
	striped:true,//
	rp: 20,
	usepager: true,
	title: '宿舍员工信息',
	showTableToggleBtn: true,
	width:'100%',
	height: 420
	}
	);
	
	function rowdbclick(rowData){
		$('#AddForm').resetForm();
		$('#fid').val($(rowData).data("a.fid"));
		getInfo('fid');
		$("#addDiv").show("slow");
		$('#detailType').val('Edit');
	}
	function test(com,grid)
	{
		if (com=='删除')
			{
				if($('.trSelected', grid).length==0){alert('请先选择一条记录再进行操作！');return false;}
				var SNum=$('.trSelected', grid).attr("id").replace("row","");
				if (confirm('确定要删除选择的整张单据?')){
				  $.post('ElectWaterEdit.asp?showType=DataProcess&detailType=Delete',{"fid":SNum},function(data){
					if(data.indexOf("###")==-1) alert(data);
					else $("#flex1").flexReload();
				  });
				}
			}
		else if (com=='增加')
			{
				$('#AddForm').resetForm();
				$('#TbDetails tr:gt(1)').remove();
				var myDate=new Date();
				var yestoday=showdate(0);
				$('#checkdate').val(yestoday);
				$('#year').val(myDate.getFullYear());
				$('#period').val((myDate.getMonth()+1));
				$('#detailType').val('AddNew');
				$("#addDiv").show("slow");
			}			
		else if (com=='维护')
			{
				$('#AddForm').resetForm();
				if($('.trSelected', grid).length==0){alert('请先选择一条记录再进行操作！');return false;}
				$('#fid').val($('.trSelected', grid).attr("id").replace("row",""));
				getInfo('fid');
				$("#addDiv").show("slow");
				$('#detailType').val('Edit');
			}			
		else if (com=='返回')
			{
				history.go(-1);
			}			
		else if(com=='查询'){
				var txt='';
				txt+='楼号：<select id="lh" name="lh" class="textfield" style="width:30%"><option value="">全部</option><option value="办公楼宿舍">办公楼宿舍</option><option value="旭日小区4号楼">旭日小区4号楼</option><option value="旭日小区22号楼">旭日小区22号楼</option><option value="食堂宿舍">食堂宿舍</option></select><br/>';
				txt+='工号或姓名：<input id="yg" name="yg" type="text" class="textfield" style="width:30%"><br/>';
				$.prompt(txt,{
					buttons: { 查看: '0',导出: '1'},
					submit:function(v,m,f){ 
			if(v==0){
				$("#flex1").flexOptions({newp: 1, params:[
					{name:"lh",value:f.lh},
					{name:"yg",value:f.yg}
					]
				});
				$("#flex1").flexReload();
				$.prompt.close();
			}else{
				window.open("ElectWaterEdit.asp?print_tag=1&showType=Export&lh="+encodeURI(f.lh)+"&yg="+encodeURI(f.yg),"Print","","false");
				$.prompt.close();
			}
							return false; 
					 }
				 });
		}
	}

function closead1(){
		$("#addDiv").hide("slow");
		$("#addDiv").css("z-index","500");
}

function getInfo(obj){
  if($("#"+obj).val()==''){alert("对应编号不能为空！");return false;}
	$.get("ElectWaterEdit.asp", { showType: "getInfo",detailType: obj, InfoID: $("#"+obj).val()},function(data){
		if(data.indexOf("###")==-1){alert(data);$("#"+obj).val('');}
		else{
			var datajson=jQuery.parseJSON(data);//转换后的JSON对象
			if(obj=="fid"){
			   var datajson=jQuery.parseJSON(data);//转换后的JSON对象
			   $(':input',$('#AddForm')).each(function(i,fieldone){
			     if(fieldone.id){
					   if(fieldone.id!='detailType')$('#'+fieldone.id).val('');
					   if(datajson.fieldValue[0][fieldone.id])
					     $('#'+fieldone.id).val(datajson.fieldValue[0][fieldone.id]);
					 }
			   });
				//子表信息
				$('#TbDetails tr:gt(1)').remove();
				for(var i=0;i<datajson.fieldValue[0]['Entrys'].length;i++){
					$('#TbDetails').append($('#CloneNodeTr').clone().show());
					$(':input',$('#TbDetails tr:last')).each(function(n,fieldone){
						if(datajson.fieldValue[0]['Entrys'][i][fieldone.name])
							fieldone.value=datajson.fieldValue[0]['Entrys'][i][fieldone.name]
					});
				}
			}
		}
	});
}
function getEmp(obj){
  if($(obj).val()=='')return false;
	$.get("ElectWaterEdit.asp", { showType: "getInfo",detailType: 'Emp', InfoID: $(obj).val()},
		function(data){
			if(data.indexOf("###")>-1){
				alert('对应员工不存在，请联系人资部确认！');
				$(obj).val('');
			}
			else{
				var datajson=jQuery.parseJSON(data);//转换后的JSON对象
				$('input[name=person]',obj.parentNode.parentNode).val(datajson.Fitemid);
				$('input[name=personid]',obj.parentNode.parentNode).val(datajson.Fnumber);
				$('input[name=personname]',obj.parentNode.parentNode).val(datajson.fname);
				if($(':input[value='+datajson.Fnumber+']',$('#editDetails')).length>1){
					alert('此单据中已经存在该员工,不允许重复录入！');
					$(obj).parent().parent().remove();
				}
			}
		});
}
function getDepart(obj){
  if($(obj).val()=='')return false;
	$.get("ElectWaterEdit.asp", { showType: "getInfo",detailType: 'Depart', InfoID: $(obj).val()},
		function(data){
			if(data.indexOf("###")>-1){
				alert('对应员工不存在，请联系人资部确认！');
				$(obj).val('');
			}
			else{
				var datajson=jQuery.parseJSON(data);//转换后的JSON对象
				$('input[name=fbase1]',obj.parentNode.parentNode).val(datajson.Fitemid);
				$('input[name=depart]',obj.parentNode.parentNode).val(datajson.Fnumber+'/'+datajson.fname);
			}
		});
}
</script>
<div id="addDiv" style="width:96%;height:'420px';top:0;left:20;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
 <form id="AddForm" name="AddForm" style="margin:0; padding:0 ">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="20" width="100%" class="tablemenu" id="formove"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>宿舍员工信息</strong></font></td>
  </tr>
  <tr>
    <td height="20" width="100%" bgcolor="#EBF2F9">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews>
      <tr>
        <td height="20" align="left" width="10%">单据号：</td>
        <td width="20%"><input name="fid" type="hidden"  id="fid"  value="">
		<input name="fbillno" type="text" class="textfield" id="fbillno" value="" maxlength="100" readonly="true"></td>
        <td height="20" align="left" width="10%">宿舍号：</td>
        <td width="20%"><input name="ftext" type="text" class="textfield" id="ftext" value="" maxlength="100" onBlur="checkDorm(this)"></td>
        <td height="20" align="left" width="10%">楼号：</td>
        <td width="20%"><select name="ftext1" id="ftext1" >
		<option value="办公楼宿舍" >办公楼宿舍</option>
		<option value="旭日小区4号楼" >旭日小区4号楼</option>
		<option value="旭日小区22号楼" >旭日小区22号楼</option>
		<option value="食堂宿舍" >食堂宿舍</option>
		</select></td>
      </tr>
      <tr>
        <td height="20" align="left">所在楼层：</td>
        <td><input name="louceng" type="text" class="textfield" id="louceng" value="0" maxlength="100" onBlur="return checkInt(this)"></td>
        <td height="20" align="left">最多入住人数：</td>
        <td><input name="maxperson" type="text" class="textfield" id="maxperson" value="0" maxlength="100" onBlur="return checkInt(this)"></td>
        <td height="20" align="left">电表数：</td>
        <td><input name="electnum" type="text" class="textfield" id="electnum" value="0" maxlength="100" onChange="return checkNum(this)"></td>
      </tr>
      <tr>
        <td height="20" align="left">面积大小：</td>
        <td><input name="fdecimal" type="text" class="textfield" id="fdecimal" value="0" maxlength="100"></td>
        <td height="20" align="left">当前入住人数：</td>
        <td><input name="sumperson" type="text" class="textfield" id="sumperson" value="0" maxlength="100" readonly></td>
        <td height="20" align="left">水表数：</td>
        <td><input name="waternum" type="text" class="textfield" id="waternum" value="0" maxlength="100" onChange="return checkNum(this)"></td>
      </tr>
      <tr>
        <td height="20" align="left">是否使用：</td>
        <td><select id="useflag" name="useflag">
        <option value="1">是</option>
        <option value="0">否</option>
        </select>
				</td>
        <td height="20" align="left">是否统计：</td>
        <td><select id="showflag" name="showflag">
        <option value="1">是</option>
        <option value="0">否</option>
        </select>
				</td>
        <td height="20" align="left">热水表数</td>
        <td><input name="hotwater" type="text" class="textfield" id="hotwater" value="0" maxlength="100" onChange="return checkNum(this)"></td>
      </tr>
      <tr>
        <td height="20" colspan="6">
		<table width="100%" border="0" id="editDetails" cellpadding="0" cellspacing="1" bgcolor="#99BBE8">
		  <tbody id="TbDetails">
		  <tr height="24">
			<td width="4%" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>床位号</strong></font></td>
			<td width="20%" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>员工编号</strong></font></td>
			<td width="20%" height="24" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>员工姓名</strong></font></td>
			<td width="20%" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>部门</strong></font></td>
			<td width="20%" nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>入住日期</strong></font></td>
			<td colspan="2" width="100%" nowrap bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong></td>
		  </tr>
		  <tr bgColor="#EBF2F9" id="CloneNodeTr" style="display:none;" >
			<td><input type="text" name="finteger3" class="textdetail" onblur="checkInt(this)"/></td>
			<td><input type="hidden" name="fentryid" /><input type="hidden" name="person"/><input type="text" name="personid" class="textdetail" onblur="getEmp(this)"/></td>
			<td><input type="text" name="personname" class="textdetail" readonly="readonly"/></td>
			<td><input type="hidden" name="fbase1"/><input type="text" name="depart" class="textdetail" onblur="getDepart(this)"/></td>
			<td><input type="text" name="fdate1" class="textdetail" onfocus="$(this).datepick({dateFormat: 'yyyy-mm-dd'})"/></td>
			<td align="left"><input type="button" onClick="$(this).parent().parent().hide();$(this).next().val(1);" style="background:no-repeat center url(../Images/delete.gif);width:20px; height:20px;" class="button"/>
      <input type="hidden" name="DeleteFlag" id="DeleteFlag" value="0">
      </td>
		  </tr>
		  </tbody>
		</table>
		</td>
      </tr>

	<tr bgcolor="#99BBE8" >
	  <td align="center" colspan="6" class="toolbar">
	  <input type="hidden" name="detailType" id="detailType" value="">
    <input name="addrow" type="button" class="button"  value="增加一行" style="WIDTH: 80;" onClick="$('#TbDetails').append($('#CloneNodeTr').clone().show());">&nbsp;
			<input name="submitSaveAdd" type="submit" class="submit button"  value="保存" style="WIDTH: 80;"  onClick="toSubmit(this)">&nbsp;
			<input type="button" class="button"  value="关闭" style="WIDTH: 80;"  onClick="closead1()">&nbsp;
	  </td>
	</tr>
	</table>
	</td>
  </tr>
</table>
</form>
</div>
</BODY>
</HTML>