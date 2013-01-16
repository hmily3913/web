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
			$.post('EWMoneyEdit.asp?showType=DataProcess',$("#AddForm").serialize(),function(data){
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
	url: 'EWMoneyEdit.asp?showType=DetailsList',
	dataType: 'json',
	colModel : [
	{display: '单号', name : 'a.fbillno', width : 50, sortable : true, align: 'left'},
	{display: '宿舍号', name : 'sushehao', width : 50, sortable : true, align: 'left'},
	{display: '查表日期', name : 'checkdate', width : 50, sortable : true, align: 'left'},
	{display: '年份', name:'year',width : 50, sortable : true, align: 'left'},
	{display: '月份', name : 'period', width : 50, sortable : true, align: 'left'},
	{display: '水价', name : 'waterprice', width : 50, sortable : true, align: 'left'},
	{display: '电价', name : 'fdecimal2', width : 50, sortable : true, align: 'left'},
	{display: '上月水表度数', name : 'n5', width : 50, sortable : true, align: 'left'},
	{display: '上月电表度数', name : 'n6', width : 50, sortable : true, align: 'left'},
	{display: '上月热水度数', name : 'n7', width : 50, sortable : true, align: 'left'},
	{display: '本月水表度数', name : 'n8', width : 50, sortable : true, align: 'left'},
	{display: '本月电表度数', name : 'n9', width : 50, sortable : true, align: 'left'},
	{display: '本月热水度数', name : 'n10', width : 50, sortable : true, align: 'left'},
	{display: 'FID', name : 'FID', width : 20, sortable : true, align: 'left',hide:true,toggle:false}
		],
	buttons : [
		{name: '增加',  onpress : test},
		{separator: true},
		{name: '维护',  onpress : test},
		{separator: true},
		{name: '整单删除',  onpress : test},
		{separator: true},
		{name: '查询',  onpress : test},
		{separator: true},
		{name: '返回',  onpress : test},
		{separator: true}
		],
	searchitems : [
		{display: '年份', name : 'year'},
		{display: '月份', name : 'period', isdefault: true},
		{display: '宿舍号', name : 'b.sushehao'},
		{display: '单号', name : 'a.fbillno'}
		],
	onRowDblclick:rowdbclick,
	sortname: "a.fbillno",
	sortorder: "desc",
	singleSelect: true,
	striped:true,//
	rp: 20,
	usepager: true,
	title: '宿舍水电信息',
	showTableToggleBtn: true,
	width:'100%',
	height: 420
	}
	);
	
	function rowdbclick(rowData){
		$('#AddForm').resetForm();
		$('#FID').val($(rowData).data("FID"));
		getInfo('FID');
		$("#addDiv").show("slow");
		$('#detailType').val('Edit');
	}
	function test(com,grid)
	{
		if (com=='整单删除')
			{
				if($('.trSelected', grid).length==0){alert('请先选择一条记录再进行操作！');return false;}
				var SNum=$('.trSelected', grid).attr("id").replace("row","");
				if (confirm('确定要删除选择的整张单据?')){
				  $.post('EWMoneyEdit.asp?showType=DataProcess&detailType=Delete',{"FID":SNum},function(data){
					if(data.indexOf("###")==-1) alert(data);
					else $("#flex1").flexReload();
				  });
				}
			}
		else if (com=='增加')
			{
				$('#AddForm').resetForm();
				$('#TbDetails tr:gt(0)').remove();
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
				$('#FID').val($('.trSelected', grid).attr("id").replace("row",""));
				getInfo('FID');
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
				txt+='宿舍号：<input id="ss" name="ss" type="text" class="textfield" style="width:30%"><br/>';
				txt+='年份：<select name="nf" id="nf" class="textfield" style="width:30%">';
				txt+='<option value="" >全部</option>';
        txt+='<option value="2011">2011</option>';
        txt+='<option value="2012">2012</option>';
        txt+='<option value="2013">2013</option>';
        txt+='</select><br/>';
				txt+='月份：<select name="yf" id="yf"  class="textfield" style="width:30%">';
				txt+='<option value="" >全部</option>';
				txt+='<option value="1" >1</option>';
				txt+='<option value="2" >2</option>';
				txt+='<option value="3" >3</option>';
				txt+='<option value="4" >4</option>';
				txt+='<option value="5" >5</option>';
				txt+='<option value="6" >6</option>';
				txt+='<option value="7" >7</option>';
				txt+='<option value="8" >8</option>';
				txt+='<option value="9" >9</option>';
				txt+='<option value="10" >10</option>';
				txt+='<option value="11" >11</option>';
				txt+='<option value="12" >12</option>';
				txt+='</select><br/>';
				$.prompt(txt,{
					buttons: { 查看: '0',导出: '1'},
					submit:function(v,m,f){ 
			if(v==0){
				$("#flex1").flexOptions({newp: 1, params:[
					{name:"nf",value:f.nf},
					{name:"yf",value:f.yf},
					{name:"lh",value:f.lh},
					{name:"ss",value:f.ss}
					]
				});
				$("#flex1").flexReload();
				$.prompt.close();
			}else{
				window.open("EWMoneyEdit.asp?print_tag=1&showType=Export&nf="+encodeURI(f.nf)+"&yf="+encodeURI(f.yf)+"&lh="+encodeURI(f.lh)+"&ss="+encodeURI(f.ss),"Print","","false");
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
	$.get("EWMoneyEdit.asp", { showType: "getInfo",detailType: obj, InfoID: $("#"+obj).val()},function(data){
		if(data.indexOf("###")==-1){alert(data);$("#"+obj).val('');}
		else{
			var datajson=jQuery.parseJSON(data);//转换后的JSON对象
			if(obj=="FID"){
				//主表信息
				$(':input',$('#AddForm')).each(function(i,fieldone){
					if(fieldone.id){
						if(fieldone.id!='detailType')$('#'+fieldone.id).val('');
						if(datajson.fieldValue[0][fieldone.id]){
							$('#'+fieldone.id).val(datajson.fieldValue[0][fieldone.id]);
						}
					}
				});
				//子表信息
				$('#TbDetails tr:gt(0)').remove();
				for(var i=0;i<datajson.fieldValue[0]['Entrys'].length;i++){
					var tr=document.createElement('tr');
					tr.setAttribute("bgColor","#EBF2F9");
					for(var n=0;n<NameArr.length;n++){
						if(NameArr[n].setFlag==1){
							var hiddenfield='<input type="hidden" name="'+NameArr[n].name+'" id="'+NameArr[n].name+'" value="'+datajson.fieldValue[0]['Entrys'][i][NameArr[n].name]+'">';
							$(tr).append(hiddenfield);
						}else{
							var td = document.createElement('td');
							td.innerHTML='<input type="text" class="textdetail" value="'+datajson.fieldValue[0]['Entrys'][i][NameArr[n].name]+'" name="'+NameArr[n].name+'" '+(NameArr[n].setFlag==3?'onchange="return checkNum(this)"':'onchange="return getEW(this)"')+'>';
							$(tr).append(td);
							td=null;
						}
					}
					var td = document.createElement('td');
					td.setAttribute("align","center");
					td.innerHTML='<input type="button" onClick="$(this).parent().parent().hide();$(this).next().val(1);" style="background:no-repeat center url(../Images/delete.gif);width:40px;" /><input type="hidden" name="DeleteFlag" id="DeleteFlag" value="0">';
					$(tr).append(td);
					td=null;
					$('#TbDetails').append(tr);
					tr=null;
				}
			}
		}
	});
}

var NameArr=[{name:"FEntryID",value:"",setFlag:1},{name:"sushehao",value:"",setFlag:2},{name:"water",value:0,setFlag:3},{name:"elect",value:0,setFlag:3},{name:"lastHotWater",value:0,setFlag:3},{name:"thiswater",value:0,setFlag:3},{name:"thiselect",value:0,setFlag:3},{name:"thiHotWater",value:0,setFlag:3}];
function AddRowEW(){
	var tr=document.createElement('tr');
	tr.setAttribute("bgColor","#EBF2F9");
	for(var n=0;n<NameArr.length;n++){
		if(NameArr[n].setFlag==1){
			var hiddenfield='<input type="hidden" name="'+NameArr[n].name+'" id="'+NameArr[n].name+'">';
			$(tr).append(hiddenfield);
		}else{
			var td = document.createElement('td');
			td.innerHTML='<input type="text" class="textdetail" value="'+NameArr[n].value+'" name="'+NameArr[n].name+'" '+(NameArr[n].setFlag==3?'onchange="return checkNum(this)"':'onchange="return getEW(this)"')+'>';
			$(tr).append(td);
			td=null;
		}
	}
	var td = document.createElement('td');
	td.setAttribute("align","center");
	td.innerHTML='<input type="button" onClick="$(this).parent().parent().remove()" style="background:no-repeat center url(../Images/delete.gif);width:40px;"  /><input type="hidden" name="DeleteFlag" id="DeleteFlag" value="0">';
	$(tr).append(td);
	td=null;
	$('#TbDetails').append(tr);
	tr=null;
}
function AddRowAllEW(){
	$.get("EWMoneyEdit.asp", { showType: "getInfo",detailType: 'AllEW',louhao: $('#louhao').val()},function(data){
		if(data.length>0){
			var datajson=jQuery.parseJSON(data);//转换后的JSON对象
				//子表信息
			$('#TbDetails tr:gt(0)').remove();
			for(var i=0;i<datajson.Entrys.length;i++){
				var tr=document.createElement('tr');
				tr.setAttribute("bgColor","#EBF2F9");
				for(var n=0;n<NameArr.length;n++){
					if(NameArr[n].setFlag==1){
						var hiddenfield='<input type="hidden" name="'+NameArr[n].name+'" id="'+NameArr[n].name+'" value="'+(datajson.Entrys[i][NameArr[n].name]?datajson.Entrys[i][NameArr[n].name]:'')+'">';
						$(tr).append(hiddenfield);
					}else{
						var td = document.createElement('td');
						td.innerHTML='<input type="text" class="textdetail" value="'+(datajson.Entrys[i][NameArr[n].name]?datajson.Entrys[i][NameArr[n].name]:'')+'" name="'+NameArr[n].name+'" '+(NameArr[n].setFlag==3?'onchange="return checkNum(this)"':'onchange="return getEW(this)"')+'>';
						$(tr).append(td);
						td=null;
					}
				}
				var td = document.createElement('td');
				td.setAttribute("align","center");
				td.innerHTML='<input type="button" onClick="$(this).parent().parent().hide();$(this).next().val(1);" style="background:no-repeat center url(../Images/delete.gif);width:40px;" /><input type="hidden" name="DeleteFlag" id="DeleteFlag" value="0">';
				$(tr).append(td);
				td=null;
				$('#TbDetails').append(tr);
				tr=null;
			}
		}
	});
}
</script>
<div id="addDiv" style="width:96%;height:'420px';top:0;left:20;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
 <form id="AddForm" name="AddForm" style="margin:0; padding:0 ">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td height="20" width="100%" class="tablemenu" id="formove"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:closead1()" >&nbsp;<strong>宿舍水电信息</strong></font></td>
  </tr>
  <tr>
    <td height="20" width="100%" bgcolor="#EBF2F9">
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" id=editNews>
      <tr>
        <td height="20" align="left" width="10%">单据号：</td>
        <td width="20%"><input name="FID" type="hidden"  id="FID"  value="">
		<input name="FBillNo" type="text" class="textfield" id="FBillNo" value="" maxlength="100" readonly="true"></td>
        <td height="20" align="left" width="10%">查表日期：</td>
        <td width="20%"><input name="checkdate" type="text" class="required dateISO textfield" id="checkdate"  value="" maxlength="100"></td>
        <td height="20" align="left" width="10%">年份：</td>
        <td width="20%"><select name="year" id="year" class="textfield">
        <option value="2010">2010</option>
        <option value="2011">2011</option>
        <option value="2012">2012</option>
        </select></td>
      </tr>
      <tr>
        <td height="20" align="left">月份：</td>
        <td><select name="period" id="period"  class="textfield">
		<option value="1" >1</option>
		<option value="2" >2</option>
		<option value="3" >3</option>
		<option value="4" >4</option>
		<option value="5" >5</option>
		<option value="6" >6</option>
		<option value="7" >7</option>
		<option value="8" >8</option>
		<option value="9" >9</option>
		<option value="10" >10</option>
		<option value="11" >11</option>
		<option value="12" >12</option>
		</select>
        <td height="20" align="left">水价/吨：</td>
        <td><input name="waterprice" type="text" class="required number textfield" id="waterprice"  value="" maxlength="100" ></td>
        <td height="20" align="left">电价/度：</td>
        <td><input name="FDecimal2" type="text" class="required number textfield" id="FDecimal2" value="" maxlength="100" ></td>
      </tr>
      <tr>
        <td height="20" align="left">热水价格/吨：</td>
        <td><input name="HotWaterPrice" type="text" class="required number textfield" id="HotWaterPrice"  value="" maxlength="100"></td>
        <td height="20" align="left">制单人：</td>
        <td><input type="hidden" name="FBiller" id="FBiller" value=""><input name="FBillerName" type="text" id="FBillerName"  class="textfield" value="" maxlength="100" readonly></td>
        <td height="20" align="left">制单日期：</td>
        <td><input name="FDate1" type="text" class="textfield" id="FDate1"  value="" maxlength="100" readonly></td>
      </tr>
      <tr>
        <td height="20" colspan="6">
		<table width="100%" border="0" id="editDetails" cellpadding="0" cellspacing="1" bgcolor="#99BBE8">
		  <tbody id="TbDetails">
		  <tr height="24">
			<td width="12.5%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>宿舍号</strong></font></td>
			<td width="12.5%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>上月水表度数</strong></font></td>
			<td width="12.5%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>上月电表度数</strong></font></td>
			<td width="12.5%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>上月热水度数</strong></font></td>
			<td width="12.5%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>本月水表度数</strong></font></td>
			<td width="12.5%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>本月电表度数</strong></font></td>
			<td width="12.5%" bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>本月热水度数</strong></font></td>
			<td width="12.5%" bgcolor="#8DB5E9"><strong><font color="#FFFFFF">操作</font></strong></td>
		  </tr>
		  </tbody>
		</table>
		</td>
      </tr>

	<tr bgcolor="#99BBE8" >
	  <td align="center" colspan="6" class="toolbar">
	  <input type="hidden" name="detailType" id="detailType" value="">
    <select id="louhao" name="louhao">
		<option value="办公楼宿舍" >办公楼宿舍</option>
		<option value="旭日小区4号楼" >旭日小区4号楼</option>
		<option value="旭日小区22号楼" >旭日小区22号楼</option>
		<option value="食堂宿舍" >食堂宿舍</option>
    </select>
    <input name="addrows" type="button" class="button" value="增加全部" style="WIDTH: 80;" onClick="AddRowAllEW()">&nbsp;
    <input name="addrow" type="button" class="button"  value="增加一行" style="WIDTH: 80;" onClick="AddRowEW()">&nbsp;
			<input name="submitSaveAdd" type="submit" class="submit button"  value="保存" style="WIDTH: 80;" onClick="toSubmit(this)">&nbsp;
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