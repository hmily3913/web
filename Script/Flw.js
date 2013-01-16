// FLW工作流平台js
// creat by zbh 2011-04-07

//var xPos; var yPos; 
//$(document).bind('mousemove',function(e){ 
//            xPos= e.pageX ;
//			yPos= e.pageY; 
//});
//
//双击表单对应列
//obj 当前元素
//reptype 提交类型
//tdid 当前行的数据内码
var curTd;
function ClickTd(obj,reptype,tdid){
	//判断权限
	//1.有编辑权限返回
	//弹出div，用来输入信息，保存数据
	//2.有查看权限返回
	//弹出div，用来查看相关信息，不允许编辑
	//3。无权限返回
	curTd=obj;
	$('#Finterid').val(tdid);
	$('#Keyword').val(reptype);
	jQuery.get("FlwAjaxFunction.asp", { "key": reptype, "FItemid": tdid},
	   function(data){
		 if(data.indexOf("###")>-1){
			 var arryreply=data.split("###");
			 $('#Replyer').val(arryreply[0]);
			$('#ReplyDate').val(arryreply[1]);
			$('#ReplyText').val(arryreply[2]);
			if(arryreply[3]==0){
				$('#Replyer').attr('readonly',true);
				$('#ReplyDate').attr('readonly',true);
				$('#ReplyText').attr('readonly',true);
				$('#submitSaveEdit').attr('disabled',true);
			}
		 }
		 $("#ReplyDiv").show('slow');
		 var width=590;//$("#ReplyDiv").css('width');
		 var height=180;//$("#ReplyDiv").css('height');
		var Position = getPosition();
		var leftadd = (Position.left +(Position.width-width)/2)+ "px";
		var topadd = (Position.top +(Position.height-height)/2)+ "px";
//		leftadd = ( Position.left + leftadd+((xPos-Position.width)>0?(xPos-Position.width):0)) + "px";
//		topadd = ( Position.top + topadd+((yPos-Position.height)>0?(yPos-Position.height):0)) + "px";
//alert(Position.scrleft+"#"+Position.scrtop+"#"+leftadd+"#"+topadd+"#"+width+"#"+height);
		 $("#ReplyDiv").animate({top:topadd,left:leftadd,opacity:1,marginLeft:"0.6in",fontSize:"3em"});
	 });
}
//采购双击表单对应列
//obj 当前元素
//reptype 提交类型
//tdid 当前行的数据内码
//tdidd 当前行数据分录内码
function MtrClickTd(obj,reptype,tdid,tdidd){
	//判断权限
	//1.有编辑权限返回
	//弹出div，用来输入信息，保存数据
	//2.有查看权限返回
	//弹出div，用来查看相关信息，不允许编辑
	//3。无权限返回
	curTd=obj;
	$('#Finterid').val(tdid);
	$('#FEntryID').val(tdidd);
	$('#Keyword').val(reptype);
	jQuery.get("FlwAjaxFunction.asp", { "key": reptype, "FItemid": tdid, "FEntryID": tdidd},
	   function(data){
		 if(data.indexOf("###")>-1){
			 var arryreply=data.split("###");
			 $('#Replyer').val(arryreply[0]);
			$('#ReplyDate').val(arryreply[1]);
			$('#ReplyText').val(arryreply[2]);
			if(arryreply[3]==0){
				$('#Replyer').attr('readonly',true);
				$('#ReplyDate').attr('readonly',true);
				$('#ReplyText').attr('readonly',true);
				$('#submitSaveEdit').attr('disabled',true);
			}
		 }
		 $("#ReplyDiv").show('slow');
		 var width=590;//$("#ReplyDiv").css('width');会带有px
		 var height=180;//$("#ReplyDiv").css('height');
		var Position = getPosition();
		var leftadd = (Position.left +(Position.width-width)/2)+ "px";
		var topadd = (Position.top +(Position.height-height)/2)+ "px";
		 $("#ReplyDiv").animate({top:topadd,left:leftadd,opacity:1,marginLeft:"0.6in",fontSize:"3em"});
	 });
}
//6s双击表单对应列
//obj 当前元素
//reptype 提交类型
//tdidd 当前行数据分录内码
function S6ClickTd(obj,reptype,tdidd){
	//判断权限
	//1.有编辑权限返回
	//弹出div，用来输入信息，保存数据
	//2.有查看权限返回
	//弹出div，用来查看相关信息，不允许编辑
	//3。无权限返回
	curTd=obj;
	$('#FEntryID').val(tdidd);
	$('#Keyword').val(reptype);
	jQuery.get("FlwAjaxFunction.asp", { "key": reptype, "FItemid": tdidd},
	   function(data){
		 if(data.indexOf("###")>-1){
			 var arryreply=data.split("###");
			 $('#Replyer').val(arryreply[0]);
			$('#ReplyDate').val(arryreply[1]);
			$('#ReplyText').val(arryreply[2]);
			if(arryreply[3]==0){
				$('#Replyer').attr('readonly',true);
				$('#ReplyDate').attr('readonly',true);
				$('#ReplyText').attr('readonly',true);
				$('#submitSaveEdit').attr('disabled',true);
			}
		 }
		 if(reptype=="T8reply")$('#submitSaveEdit').val("改善结果确认");
		 else if(reptype=="T9reply")$('#submitSaveEdit').val("结案");
		 $("#ReplyDiv").show('slow');
		 var width=590;//$("#ReplyDiv").css('width');会带有px
		 var height=180;//$("#ReplyDiv").css('height');
		var Position = getPosition();
		var leftadd = (Position.left +(Position.width-width)/2)+ "px";
		var topadd = (Position.top +(Position.height-height)/2)+ "px";
		 $("#ReplyDiv").animate({top:topadd,left:leftadd,opacity:1,marginLeft:"0.6in",fontSize:"3em"});
	 });
}
//出货样双击表单对应列
//obj 当前元素
//reptype 提交类型
//tdid 当前行的数据内码
//tdidd 当前行数据分录内码
function SHDClickTd(obj,reptype,tdid,tdidd){
	//判断权限
	//1.有编辑权限返回
	//弹出div，用来输入信息，保存数据
	//2.有查看权限返回
	//弹出div，用来查看相关信息，不允许编辑
	//3。无权限返回
	curTd=obj;
	$('#Finterid').val(tdid);
	$('#FEntryID').val(tdidd);
	$('#Keyword').val(reptype);
	jQuery.get("FlwAjaxFunction.asp", { "key": reptype, "FItemid": tdid, "FEntryID": tdidd},
	   function(data){
		 if(data.indexOf("###")>-1){
			 var arryreply=data.split("###");
			 $('#Replyer').val(arryreply[0]);
			$('#ReplyDate').val(arryreply[1]);
			$('#ReplyText').val(arryreply[2]);
			if(reptype=="SHDSPLreply"){
			  $('#OtherEle').val(arryreply[3]);
			  document.getElementById('OtherEle').style.display="block";
			  document.getElementById('qccp').style.display="none";
//			  $('#OtherEle').show('fast');
//				$('#qccp').hide('fast');
			}
			else if(reptype=="SHDQCreply"){
			  document.getElementById('qccp').style.display="block";
			  document.getElementById('OtherEle').style.display="none";
//				$('#qccp').show('fast');
//				$('#OtherEle').hide('fast');
				if(arryreply[3].indexOf("|01,")>-1)$('#QC01').attr('checked',true);
				if(arryreply[3].indexOf("|02,")>-1)$('#QC02').attr('checked',true);
				if(arryreply[3].indexOf("|03,")>-1)$('#QC03').attr('checked',true);
				if(arryreply[3].indexOf("|04,")>-1)$('#QC04').attr('checked',true);
				if(arryreply[3].indexOf("|05,")>-1)$('#QC05').attr('checked',true);
				if(arryreply[3].indexOf("|06,")>-1)$('#QC06').attr('checked',true);
			}else{
			  document.getElementById('qccp').style.display="none";
			  document.getElementById('OtherEle').style.display="none";
//				$('#OtherEle').hide('fast');
//				$('#qccp').hide('fast');
			}
			if(arryreply[4]==0){
				$('#Replyer').attr('readonly',true);
				$('#ReplyDate').attr('readonly',true);
				$('#ReplyText').attr('readonly',true);
				$('#OtherEle').attr('readonly',true);
				$('#submitSaveEdit').attr('disabled',true);
			}
		 }
		 $("#ReplyDiv").show('slow');
		 var width=590;//$("#ReplyDiv").css('width');会带有px
		 var height=180;//$("#ReplyDiv").css('height');
		var Position = getPosition();
		var leftadd = (Position.left +(Position.width-width)/2)+ "px";
		var topadd = (Position.top +(Position.height-height)/2)+ "px";
		 $("#ReplyDiv").animate({top:topadd,left:leftadd,opacity:1,marginLeft:"0.6in",fontSize:"3em"});
	 });
}