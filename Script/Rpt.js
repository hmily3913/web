var curTd;
function confirmClick(reptype,yorn){
 if($('#FItemid').val()!=''){
	 jQuery.get("AjaxFunction.asp", { "key": reptype, "FItemid": $("#FItemid").val(),"YN":yorn},
	   function(data){
			 if(data.indexOf("###")>-1){
				 if(yorn==1){
				 curTd.parentNode.bgColor="#ff99ff";
				 curTd.innerText="True";
				 }else
				 curTd.innerText="False";
				 $("#ReplyDiv").hide("slow");
				 }
			 else alert(data);
		 });
  }else{
    alert("尚未工作回复，不需要确定！");
		$("#ReplyDiv").hide("slow");
  }
}
function SAClickTd(obj,reptype,tdid){
	//判断权限
	//1.有编辑权限返回
	//弹出div，用来输入信息，保存数据
	//2.有查看权限返回
	//弹出div，用来查看相关信息，不允许编辑
	//3。无权限返回
	curTd=obj;
	$('#FItemid').val(tdid);
	$('#Keyword').val(reptype);
	jQuery.get("AjaxFunction.asp", { "key": reptype, "FItemid": tdid},
	   function(data){
		 if(data.indexOf("###")>-1){
			 var arryreply=data.split("###");
			 $('#Replyer').val(arryreply[0]);
			$('#ReplyDate').val(arryreply[1]);
			$('#ReplyText').val(arryreply[2]);
			if(arryreply[4]!==undefined&&arryreply[4].length>0)$('#ReplyType').val(arryreply[4]);
			if(arryreply[3]==0){
				$('#Replyer').attr('readonly',true);
				$('#ReplyDate').attr('readonly',true);
				$('#ReplyText').attr('readonly',true);
				$('#submitSaveEdit').attr('disabled',true);
			}
		 $("#ReplyDiv").show('slow');
		 var width=590;//$("#ReplyDiv").css('width');
		 var height=180;//$("#ReplyDiv").css('height');
		 }else alert(data);
			var Position = getPosition();
			var leftadd = (document.documentElement.scrollLeft + ((document.documentElement.clientWidth==0?Position.width:document.documentElement.clientWidth) - 590) / 2) + "px";

//			(((xPos-Position.width)>0?(xPos-Position.width):0)+Position.left +(Position.width-width)/2)+ "px";
			var topadd = (document.documentElement.scrollTop + ((document.documentElement.clientHeight==0?Position.height:document.documentElement.clientHeight) - 180) / 2) + "px";
//(((yPos-Position.height)>0?(yPos-Position.height):0)+Position.top +(Position.height-height)/2)+ "px";
			 $("#ReplyDiv").animate({top:topadd,left:leftadd,opacity:1,marginLeft:"0.6in",fontSize:"3em"});
	 });
}

function closead(){
  $("#ReplyDiv").hide("slow");
}
//处理保存回复

function SaveEdit(){
//如果回复内容为空，不提交
 if($('#ReplyText').val()!=''){
  jQuery.get("AjaxFunction.asp", { "key": "update"+$("#Keyword").val(), "FItemid": $("#FItemid").val(),"ReplyText":$('#ReplyText').val(),"ReplyType":$('#ReplyType').val() },
   function(data){
		if(data.indexOf("###")>-1){
			var arryreply=data.split("###");
			if(arryreply[0].length>9)
			  curTd.innerText=arryreply[0].substring(0,8)+"...";
			else
			  curTd.innerText=arryreply[0];
			if(arryreply[1].length>0)get_previousSibling(curTd).innerText=arryreply[1];
			//实时改变提交后背景颜色
			curTd.parentNode.bgColor="#ffff66";
//			if($("#Keyword").val()=="OCreply")curTd.parentNode.bgColor="#ffff66";
//			else if($("#Keyword").val()=="T8reply")curTd.parentNode.bgColor="#ff99ff";
//			else if($("#Keyword").val()=="T9reply")curTd.parentNode.parentNode.removeChild(curTd.parentNode);
		}
		$("#ReplyDiv").hide("slow");
   });
  }else{
    alert("没有回复内容，不需要保存！");
	$("#ReplyDiv").hide("slow");
  }
}


