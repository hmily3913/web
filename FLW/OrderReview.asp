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
<script language="javascript" src="../Script/Admin.js"></script>
<script language="javascript" src="../Script/jquery-1.5.2.min.js"></script>
<link rel="stylesheet" href="../Images/jquery.datepick.css">
<script language="javascript" src="../Script/jquery.datepick.pack.js"></script>
<script language="javascript" src="../Script/jquery.datepick-zh-CN.js"></script>
<link rel="stylesheet" href="../Images/jqi.css">
<script language="javascript" src="../Script/jquery-impromptu.3.1.js"></script>
<script language="javascript">
function closead(){
  $("#ReplyDiv").hide("slow");
}
//处理保存回复
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
	$('#listDiv').load("OrderReviewDetails.asp #listtable",{
	  showType:'DetailsList',
	  Page:arr[0],
		ed:$('#ed').val(),
		sd:$('#sd').val(),
		ped:$('#ped').val(),
		psd:$('#psd').val(),
		os:$('#os').val(),
		fc:$('#fc').val(),
		tdid:$('#tdid').val()
	 },function(response, status, xhr){
	  if (status =="success") {
			if(response.split('###')[3]=="1"){
			$("#listtable table tr td:gt(7)").click(function(event) {
//				alert(event.target.id);
				if(this.className!=undefined&&this.className!=""&&event.target.id==''&&this.className!="DataCol"){
					var obj=this;
					var objvalue=$(obj).text();
					$(obj).html("<input type='text' class='textfield' id='"+obj.className+"' name='"+obj.className+"'>");
					$('#'+obj.className).val(objvalue);
//					if(obj.className=="PMDate"){$('#PMDate').datepick({dateFormat: 'yyyy-mm-dd'});}
					$('#'+obj.className).focus();
					$('#'+obj.className).blur(function(){
						var editvalue=$(this).val();
						$.get("OrderReviewDetails.asp", { showType: "getInfo",detailType:obj.className, InfoID:$(obj).parent().attr("id"),values:editvalue},function(data){
							if(data.length>0)alert(data);
							else{
								$(obj).text(editvalue);
								$('#'+obj.className).remove();
							}
						});
					});
				}
			});
			}
	    pageNavigation('pageN',arr,response.split('###')[1],response.split('###')[2],'showDiv');
		$('#loading-one').empty().append('页面载入完毕.').parent().fadeOut('slow');
	  }	
    })
	//产生分页导航栏
}
function showSearch(){
	var txt='';
	txt+='日期：从<input name="SDate" type="text" class="textfield" id="SDate" style="width:30%;" value="" maxlength="100"  onfocus="$(this).datepick()">到<input name="EDate" type="text" class="textfield" id="EDate" style="width:30%" value="" maxlength="100"  onfocus="$(this).datepick()"><br/>';
	txt+='订单状态：<select id="OrderStat" name="OrderStat" class="textfield" style="width:30%"><option value="">全部</option><option value="0">未转销售订单</option><option value="1">已转销售订单</option></select><br/>';
	txt+='通知单号：<input type="text" id="TZId" name="TZId" class="textfield" style="width:30%"><br/>';
	txt+='评审日期：从<input name="PSDate" type="text" class="textfield" id="PSDate" style="width:30%;" value="" maxlength="100"  onfocus="$(this).datepick()">到<input name="PEDate" type="text" class="textfield" id="PEDate" style="width:30%" value="" maxlength="100"  onfocus="$(this).datepick()"><br/>';
	txt+='分厂：<input type="text" id="fc" name="fc" class="textfield" style="width:30%"><br/>';
	$.prompt(txt,{
		buttons: { 导出: '0', 查看: '1' },
		submit:function(v,m,f){ 
			if(v==1){
				$('#sd').val(f.SDate);
				$('#ed').val(f.EDate);
				$('#os').val(f.OrderStat);
				$('#tdid').val(f.TZId);
				$('#psd').val(f.PSDate);
				$('#ped').val(f.PEDate);
				$('#fc').val(f.fc);
				pageN(1);
				$.prompt.close();
			}else{
				window.open("OrderReviewDetails.asp?print_tag=1&showType=Export&sd="+f.SDate+"&ed="+f.EDate+"&os="+f.OrderStat+"&tdid="+f.TZId+"&psd="+f.PSDate+"&ped="+f.PEDate+"&fc="+encodeURI(f.fc),"Print","","false");
				$.prompt.close();
			}
				return false; 
		 }
	 });
}
function output(){
  window.open("OrderReviewDetails.asp?print_tag=1&showType=PingbiList&pbmonth="+$('#pbmonth').val(),"Print","","false");
}
</script>
</HEAD>
<BODY>
<%
'if Instr(session("AdminPurviewFLW"),"|105,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
%>
<div id="loading" style="position:fixed !important;position:absolute;top:0;left:0;height:100%; width:100%; z-index:999; background:#000 url(../images/load.gif) no-repeat center center; opacity:0.6; filter:alpha(opacity=60);font-size:14px;line-height:20px;overflow:hidden;">
	<p id="loading-one" style="color:#fff;position:absolute; top:50%; left:50%; margin:20px 0 0 -50px; padding:3px 10px;">页面载入中...</p>
</div>
<div align="center" style="top:0;margin:0 auto; ">
<font color="#FF0000"><strong>订单评审工作平台</strong></font>
<p align="left" style="margin-top:0; margin-bottom:0; ">
<font style="background-color:#ff99ff">已经评审</font>&nbsp;
<input type="hidden" id="fc" />
<input type="hidden" id="psd" />
<input type="hidden" id="ped" />
<input type="hidden" id="sd" />
<input type="hidden" id="ed" />
<input type="hidden" id="os" />
<input type="hidden" id="tdid" />
<input type="button" name="seachbutton" id="seachbutton" onClick="showSearch()" value="查找" style='HEIGHT: 18px;WIDTH: 40px;'>
</p>
<div id="ReplyDiv" style="width:100%;height:100%;top:0;left:0;display:none;background-color:#888888;position:absolute;marginTop:-75px;marginLeft:-150px;overflow-y: hidden; overflow-x: hidden;">
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