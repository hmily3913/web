
//check if the next sibling node is an element node

function get_nextsibling(n)
  {
  var x=n.nextSibling;
  while (x.nodeType!=1)
   {
   x=x.nextSibling;
   }
  return x;
  }
function get_previousSibling(n)
  {
  var x=n.previousSibling;
  while (x.nodeType!=1)
   {
   x=x.previousSibling;
   }
  return x;
  }

//获取当前位置
function getPosition() {
var top    = document.documentElement.scrollTop;
var scrtop =window.screenTop;
var left   = document.documentElement.scrollLeft;
var scrleft =window.screenLeft;
var height = document.documentElement.offsetHeight;
var width = document.documentElement.offsetWidth;
return {top:top,left:left,height:height,width:width,scrtop:scrtop,scrleft:scrleft};
}
//改变管理位置标记--------------------------------------------------------------
function changeAdminFlag(Content){
   var row=parent.parent.headFrame.document.all.Trans.rows[0];
   row.cells[3].innerHTML = Content ;
   return true;
}

//通用选择删除条目（反选-全选）--------------------------------------------------------
function CheckOthers(form)
{
   for (var i=0;i<form.elements.length;i++)
   {
      var e = form.elements[i];
      if (e.checked==false)
      {
	     e.checked = true;
      }
      else
      {
	     e.checked = false;
      }
   }
}

function CheckAll(form)
{
   for (var i=0;i<form.elements.length;i++)
   {
      var e = form.elements[i];
      e.checked = true;
   }
}
//相关条目删除提示------------------------------------------------------------
function ConfirmDel(message)
{
   if (confirm(message))
   {
      document.formDel.submit();
   }
}
//调用在线内容编辑器-----------------------------------------------------------
function OpenDialog(sURL, iWidth, iHeight)
{
   var oDialog = window.open(sURL, "_EditorDialog", "width=" + iWidth.toString() + ",height=" + iHeight.toString() + ",resizable=no,left=0,top=0,scrollbars=no,status=no,titlebar=no,toolbar=no,menubar=no,location=no");
   oDialog.focus();
}
//检验输入字符的有效性（0-9，a-z,-,_）-------------------------------------------
function voidNum(argValue) 
{
   var flag1=false;
   var compStr="1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-";
   var length2=argValue.length;
   for (var iIndex=0;iIndex<length2;iIndex++)
   {
	   var temp1=compStr.indexOf(argValue.charAt(iIndex));
	   if(temp1==-1) 
	   {
	      flag1=false;
			break;							
	   }
	   else
	   { flag1=true; }
   }
   return flag1;
} 
//验证数字的有效性
function checkNum(textId) {
 var num;
 num = textId.value;
 var re = /^(-?\d+)(\.\d+)?$/;   //判断字符串是否为数字 
     //判断正整数 /^[1-9]+[0-9]*]*$/  
     if (!re.test(num))
    {
        alert("请输入数字(例:1.01)");
		textId.value=0;
        textId.focus();
        return false;
     }
	 	return true;
}
//正整数检查
function checkInt(textId){
 var num;
 num = textId.value;
 var re = /^-?\d+$/;   //判断字符串是否为数字 
     //判断正整数 /^[1-9]+[0-9]*]*$/  
     if (!re.test(num))
    {
        alert("请输入整数(例:1)");
		textId.value=0;
        textId.focus();
        return false;
     }
	return true;
}
/**
 * 检查日期格式是否正确
 * 输入:str  字符串
 * 返回:true 或 flase; true表示格式正确
 * 注意：此处不能验证中文日期格式
 * 验证短日期（2007-06-05）
 */
function checkDate(obj){
	var str=obj.value;
	if (obj.value==''){
		return true;
	}else{
		//var value=str.match(/((^((1[8-9]\d{2})|([2-9]\d{3}))(-)(10|12|0?[13578])(-)(3[01]|[12][0-9]|0?[1-9])$)|(^((1[8-9]\d{2})|([2-9]\d{3}))(-)(11|0?[469])(-)(30|[12][0-9]|0?[1-9])$)|(^((1[8-9]\d{2})|([2-9]\d{3}))(-)(0?2)(-)(2[0-8]|1[0-9]|0?[1-9])$)|(^([2468][048]00)(-)(0?2)(-)(29)$)|(^([3579][26]00)(-)(0?2)(-)(29)$)|(^([1][89][0][48])(-)(0?2)(-)(29)$)|(^([2-9][0-9][0][48])(-)(0?2)(-)(29)$)|(^([1][89][2468][048])(-)(0?2)(-)(29)$)|(^([2-9][0-9][2468][048])(-)(0?2)(-)(29)$)|(^([1][89][13579][26])(-)(0?2)(-)(29)$)|(^([2-9][0-9][13579][26])(-)(0?2)(-)(29)$))/);
		var value = str.match(/^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2})$/);
		if (value == null) {
			alert("时间错误，格式为：yyyy-mm-dd,注意闰年。");
			var myDate=new Date();
			obj.value=myDate.getFullYear()+"-"+(myDate.getMonth()+1)+"-"+myDate.getDate();
			obj.focus();
			return false;
		}
		else {
			var date = new Date(value[1], value[3] - 1, value[4]);
			return (date.getFullYear() == value[1] && (date.getMonth() + 1) == value[3] && date.getDate() == value[4]);
		}
	}
}
/**
 * 检查时间格式是否正确
 * 输入:str  字符串
 * 返回:true 或 flase; true表示格式正确
 * 验证时间(10:57:10)
 */
function checkTime(obj){
	var str=obj.value;
    var value = str.match(/^(\d{1,2})(:)?(\d{1,2})\2(\d{1,2})$/)
    if (value == null) {
		alert("时间错误，格式为：10:57:10。");
		var myDate=new Date();
		obj.value=myDate.toLocaleTimeString();
		obj.focus();
        return false;
    }
    else {
        if (value[1] > 24 || value[3] > 60 || value[4] > 60) {
			alert("时间错误，格式为：10:57:10。");
			var myDate=new Date();
			obj.value=myDate.toLocaleTimeString();
			obj.focus();
            return false
        }
        else {
            return true;
        }
    }
}

/**
 * 检查全日期时间格式是否正确
 * 输入:str  字符串
 * 返回:true 或 flase; true表示格式正确
 * (2007-06-05 10:57:10)
 */
function checkFullTime(obj){
	var str=obj.value;
	if (obj.value==''){
		return true;
	}else{
		var value = str.match(/^(?:19|20)[0-9][0-9]-(?:(?:0?[1-9])|(?:1[0-2]))-(?:(?:[0-2]?[1-9])|(?:[1-3][0-1])) (?:(?:[0-2]?[0-3])|(?:[0-1]?[0-9])):[0-5]?[0-9]:[0-5]?[0-9]$/);
		if (value == null) {
			alert("日期时间错误，格式为：(2007-06-05 10:57:10)。");
			var myDate=new Date();
			obj.value=myDate.getFullYear()+"-"+(myDate.getMonth()+1)+"-"+myDate.getDate()+" "+myDate.toLocaleTimeString();
			obj.focus();
			return false;
		}
		else {
			return true;
		}
	}
    //var value = str.match(/^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2}) (\d{1,2}):(\d{1,2}):(\d{1,2})$/);
    
}

//检查用户登录------------------------------------------------------------------------------
function CheckAdminLogin()
{
   var check; 
   if (!voidNum(document.AdminLogin.LoginName.value))
   {
	  alert("请正确输入用户名称（由0-9,a-z,-_任意组合的字符串）。");
      document.AdminLogin.LoginName.focus();
	  return false;
	  exit;
   }    
   if (!voidNum(document.AdminLogin.LoginPassword.value))
   {
	  alert("请输入用户密码。");
	  document.AdminLogin.LoginPassword.focus();
	  return false;
	  exit;
   }
/*   if (!voidNum(document.AdminLogin.VerifyCode.value))
   {
      alert("请正确输入验证码。");
      document.AdminLogin.VerifyCode.focus();
	  return false;
	  exit;
   }*/
   return true;
}

//用户退出登录提示--------------------------------------------------------------------------
function AdminOut()
{
   if (confirm("您真的要退出管理操作吗？"))
   location.replace("CheckAdmin.asp?AdminAction=Out")
}
//跳转到第几页-------------------------------------------------------------------------------
function GoPage(Myself)
{
   window.location.href=Myself+"Page="+document.formDel.SkipPage.value;
}
function GoPagebySeach()
{
   window.location.href="PurviewSet.asp?seachname="+escape(document.formDel.seachname.value)+"&Page=1";
}
function GoPagebySeach_Flw()
{
   window.location.href="FLW_PurviewSet.asp?seachname="+escape(document.formDel.seachname.value)+"&Page=1";
}

























//选择起始日期-----------------------------------------------------------------
var DS_x,DS_y;

function dateSelector()  //构造dateSelector对象，用来实现一个日历形式的日期输入框。
{
  var myDate=new Date();
  this.year=myDate.getFullYear();  //定义year属性，年份，默认值为当前系统年份。
  this.month=myDate.getMonth()+1;  //定义month属性，月份，默认值为当前系统月份。
  this.date=myDate.getDate();  //定义date属性，日，默认值为当前系统的日。
  this.inputName='';  //定义inputName属性，即输入框的name，默认值为空。注意：在同一页中出现多个日期输入框，不能有重复的name！
  this.display=display;  //定义display方法，用来显示日期输入框。
}

function display()  //定义dateSelector的display方法，它将实现一个日历形式的日期选择框。
{
  var week=new Array('日','一','二','三','四','五','六');

  document.write("<style type=text/css>");
  document.write("  .ds_font td,span  { font: normal 12px 宋体; color: #000000; }");
  document.write("  .ds_border  { border: 1px solid #000000; cursor: hand; background-color: #DDDDDD }");
  document.write("  .ds_border2  { border: 1px solid #000000; cursor: hand; background-color: #DDDDDD }");
  document.write("</style>");

  document.write("<input style='width:72px;text-align:left;' class='textfield' id='DS_"+this.inputName+"' name='"+this.inputName+"' value='"+this.year+"-"+this.month+"-"+this.date+"' title=双击可进行编缉 ondblclick='this.readOnly=false;this.focus()' onblur='this.readOnly=true' readonly>");
  document.write("<button style='width:60px;height:18px;font-size:12px;margin:1px;border:1px solid #A4B3C8;background-color:#DFE7EF;' type=button onclick=get_nextsibling(this).style.display='block' onfocus=this.blur()>选择日期</button>");

  document.write("<div style='position:absolute;display:none;text-align:center;width:0px;height:0px;overflow:visible' onselectstart='return false;'>");
  document.write("  <div style='position:absolute;left:-60px;top:20px;width:142px;height:165px;background-color:#F6F6F6;border:1px solid #245B7D;' class=ds_font>");
  document.write("    <table cellpadding=0 cellspacing=1 width=140 height=20 bgcolor=#CEDAE7 onmousedown='DS_x=event.x-parentNode.style.pixelLeft;DS_y=event.y-parentNode.style.pixelTop;setCapture();' onmouseup='releaseCapture();' onmousemove='dsMove(this.parentNode)' style='cursor:move;'>");
  document.write("      <tr align=center>");
  document.write("        <td width=12% onmouseover=this.className='ds_border' onmouseout=this.className='' onclick=subYear(this) title='减小年份'>&lt;&lt;</td>");
  document.write("        <td width=12% onmouseover=this.className='ds_border' onmouseout=this.className='' onclick=subMonth(this) title='减小月份'>&lt;</td>");
  document.write("        <td width=52%><b>"+this.year+"</b><b>年</b><b>"+this.month+"</b><b>月</b></td>");
  document.write("        <td width=12% onmouseover=this.className='ds_border' onmouseout=this.className='' onclick=addMonth(this) title='增加月份'>&gt;</td>");
  document.write("        <td width=12% onmouseover=this.className='ds_border' onmouseout=this.className='' onclick=addYear(this) title='增加年份'>&gt;&gt;</td>");
  document.write("      </tr>");
  document.write("    </table>");

  document.write("    <table cellpadding=0 cellspacing=0 width=140 height=20 onmousedown='DS_x=event.x-parentNode.style.pixelLeft;DS_y=event.y-parentNode.style.pixelTop;setCapture();' onmouseup='releaseCapture();' onmousemove='dsMove(this.parentNode)' style='cursor:move;'>");
  document.write("      <tr align=center>");
  for(i=0;i<7;i++)
	document.write("      <td>"+week[i]+"</td>");
  document.write("      </tr>");
  document.write("    </table>");

  document.write("    <table cellpadding=0 cellspacing=2 width=140 bgcolor=#EEEEEE>");
  for(i=0;i<6;i++)
  {
    document.write("    <tr align=center>");
	for(j=0;j<7;j++)
      document.write("    <td width=10% height=16 onmouseover=if(this.innerText!=''&&this.className!='ds_border2')this.className='ds_border' onmouseout=if(this.className!='ds_border2')this.className='' onclick=getValue(this,document.all('DS_"+this.inputName+"'))></td>");
    document.write("    </tr>");
  }
  document.write("    </table>");

  document.write("    <span style=cursor:hand onclick=this.parentNode.parentNode.style.display='none'>【关闭】</span>");
  document.write("  </div>");
  document.write("</div>");

  dateShow(get_nextsibling(get_nextsibling(document.all("DS_"+this.inputName))).childNodes[0].childNodes[2],this.year,this.month)
}

function subYear(obj)  //减小年份
{
  var myObj=obj.parentNode.parentNode.parentNode.cells[2].childNodes;
  myObj[0].innerHTML=eval(myObj[0].innerHTML)-1;
  dateShow(get_nextsibling(get_nextsibling(obj.parentNode.parentNode.parentNode)),eval(myObj[0].innerHTML),eval(myObj[2].innerHTML))
}

function addYear(obj)  //增加年份
{
  var myObj=obj.parentNode.parentNode.parentNode.cells[2].childNodes;
  myObj[0].innerHTML=eval(myObj[0].innerHTML)+1;
  dateShow(get_nextsibling(get_nextsibling(obj.parentNode.parentNode.parentNode)),eval(myObj[0].innerHTML),eval(myObj[2].innerHTML))
}

function subMonth(obj)  //减小月份
{
  var myObj=obj.parentNode.parentNode.parentNode.cells[2].childNodes;
  var month=eval(myObj[2].innerHTML)-1;
  if(month==0)
  {
    month=12;
    subYear(obj);
  }
  myObj[2].innerHTML=month;
  dateShow(get_nextsibling(get_nextsibling(obj.parentNode.parentNode.parentNode)),eval(myObj[0].innerHTML),eval(myObj[2].innerHTML))
}

function addMonth(obj)  //增加月份
{
  var myObj=obj.parentNode.parentNode.parentNode.cells[2].childNodes;
  var month=eval(myObj[2].innerHTML)+1;
  if(month==13)
  {
    month=1;
    addYear(obj);
  }
  myObj[2].innerHTML=month;
  dateShow(get_nextsibling(get_nextsibling(obj.parentNode.parentNode.parentNode)),eval(myObj[0].innerHTML),eval(myObj[2].innerHTML))
}

function dateShow(obj,year,month)  //显示各月份的日
{
  var myDate=new Date(year,month-1,1);
  var today=new Date();
  var day=myDate.getDay();
  var selectDate=get_previousSibling(get_previousSibling(obj.parentNode.parentNode)).value.split('-');
  var length;
  switch(month)
  {
    case 1:
    case 3:
    case 5:
    case 7:
    case 8:
    case 10:
    case 12:
      length=31;
      break;
    case 4:
    case 6:
    case 9:
    case 11:
      length=30;
      break;
    case 2:
      if((year%4==0)&&(year%100!=0)||(year%400==0))
        length=29;
      else
        length=28;
  }
  for(i=0;i<obj.cells.length;i++)
  {
    obj.cells[i].innerHTML='';
    obj.cells[i].style.color='';
    obj.cells[i].className='';
  }
  for(i=0;i<length;i++)
  {
    obj.cells[i+day].innerHTML=(i+1);
    if(year==today.getFullYear()&&(month-1)==today.getMonth()&&(i+1)==today.getDate())
      obj.cells[i+day].style.color='red';
    if(year==eval(selectDate[0])&&month==eval(selectDate[1])&&(i+1)==eval(selectDate[2]))
      obj.cells[i+day].className='ds_border2';
  }
}

function getValue(obj,inputObj)  //把选择的日期传给输入框
{
  var myObj=get_nextsibling(get_nextsibling(inputObj)).childNodes[0].childNodes[0].cells[2].childNodes;
  if(obj.innerHTML)
    inputObj.value=myObj[0].innerHTML+"-"+myObj[2].innerHTML+"-"+obj.innerHTML;
  get_nextsibling(get_nextsibling(inputObj)).style.display='none';
  for(i=0;i<obj.parentNode.parentNode.parentNode.cells.length;i++)
    obj.parentNode.parentNode.parentNode.cells[i].className='';
  obj.className='ds_border2'
}

function dsMove(obj)  //实现层的拖移
{
  if(event.button==1)
  {
    var X=obj.clientLeft;
    var Y=obj.clientTop;
    obj.style.pixelLeft=X+(event.x-DS_x);
    obj.style.pixelTop=Y+(event.y-DS_y);
  }
}

/**
 * 分页导航条
 * 09/01/17
 * @author lym6520@qq.com 
 * @verson v2.0
 * @param {} fnName			翻页时执行的函数名(传入的第一个参数必须是“当前页码”）)
 * @param {} fnNameParams		fnName函数的参数，数组形式（比如：var arr = new Array(); arr[0] = 1;arr[1] = "hello"）
 * @param {} pagetotal			总页码
 * @param {} totalItem			总记录数
 * @param {} showID			页面显示分页导航条的div  ID
 */
function pageNavigation(fnName, fnNameParams, pagetotal, totalItem, showID) {   
    var fnParam = new Array();
    //如果这样 fnParam = fnNameParams;两个都指向同一引用
    for(var i = 0 ; i < fnNameParams.length; i++)
        fnParam[i] = fnNameParams[i];
           
     var pageIndex = parseInt(fnNameParams[0]);//当前页
       
    // 无记录  
    if (pagetotal == 0) {   
        $('#' + showID).empty();//清空翻页导航条   
        return;   
    }   
    // 分页   
    var front = pageIndex - 4;// 前面一截   
    var back = pageIndex + 4;// 后面一截   
    
    $('#' + showID).empty();//清空翻页导航条   
       
    // 页码链接   
    // 首页, 上一页   
    if (pageIndex == 1) {   
        $('#' + showID).append("首页 上一页 ");   
    } else {
        fnParam[0] = 1 ;
        var fn = fnName + "(" + fnParam + ")"; //组装执行的函数  
		var str = "<a href = 'javascript:" + fn + "'>首页</a> ";//创建连接
		$('#' + showID).append(str);
		
		fnParam[0] = pageIndex - 1 ;
	    var fn = fnName + "(" + fnParam + ")"; //组装执行函数         		
		var str = "<a href = 'javascript:" + fn + "'>上一页</a> ";//创建连接
		$('#' + showID).append(str);	         
    }   
  
    if (pagetotal == 1) {   
        $('#' + showID).append("1 ");   
    }   
    // 如果当前页是5,前面一截就是1234,后面一截就是6789   
    if (pagetotal > 1) {   
        var tempBack = pagetotal;   
        var tempFront = 1;   
        if (back < pagetotal)   
            tempBack = back;   
        if (front > 1)   

            tempFront = front;   
        for (var i = tempFront; i <= tempBack; i++) {   
            if (pageIndex == i) {   
                var str = " " + i + " ";   
                $('#' + showID).append(str);   
            } else {   
                fnParam[0] = i;
                var fn = fnName + "(" + fnParam + ")"; //组装执行的函数   
                var str = "<a href = 'javascript:" + fn + "'>[" + i + "]</a>";//创建连接   
                $('#' + showID).append(str);   
            }   
        }   
    }   
  
    // 下一页, 尾页   
    if (pageIndex == pagetotal) {   
        $('#' + showID).append("下一页 尾页 ");   
    } else {   
        fnParam[0] = pageIndex + 1 ;
        var fn = fnName + "(" + fnParam + ")"; //组装执行的函数   
        var str = " <a href = 'javascript:" + fn + "'>下一页</a> ";//创建连接   
        $('#' + showID).append(str);           
           
        fnParam[0] = pagetotal ;
        var fn = fnName + "(" + fnParam +  ")"; //组装执行的函数   
        var str = "<a href = 'javascript:" + fn + "'> 尾页 </a> ";//创建连接   
        $('#' + showID).append(str);           
    }   
       
    // 红色字体显示当前页   
    var str = "<font color = 'red'>" + pageIndex +"</font>";       
    $('#' + showID).append(str);   
       
    // 斜线"/"   
    $('#' + showID).append("/");   
       
    // 蓝色字体显示总页数   
    var str = "<font color = 'blue'>" + pagetotal +"</font>  ";      
    $('#' + showID).append(str);  
    var str = "跳到：第&nbsp;<input id='SkipPage' name='SkipPage' onKeyDown='if(event.keyCode==13)event.returnValue=false' style='HEIGHT: 18px;WIDTH: 30px;'  type='text' class='textfield' value='"+pageIndex+"'>&nbsp;页";
    $('#' + showID).append(str);  
	var fn = fnName + "(get_previousSibling(this).value)"; //组装执行的函数  
	var str="<input style='HEIGHT: 18px;WIDTH: 20px;' name='submitSkip' type='button' class='button' onClick='"+fn+"' value='GO'>";
    $('#' + showID).append(str);  
    var str = "&nbsp;共&nbsp;<font color = 'blue'>"+totalItem+"</font>&nbsp;条记录";      
    $('#' + showID).append(str);  
    //跳转到指定页
}

  Date.prototype.dateDiff = function(interval,objDate){
    //若參數不足或 objDate 不是日期物件則回傳 undefined
    if(arguments.length<2||objDate.constructor!=Date) return undefined;
    switch (interval) {
      //計算秒差
      case "s":return parseInt((objDate-this)/1000);
      //計算分差
      case "n":return parseInt((objDate-this)/60000);
      //計算時差
      case "h":return parseInt((objDate-this)/3600000);
      //計算日差
      case "d":return parseInt((objDate-this)/86400000);
      //計算週差
      case "w":return parseInt((objDate-this)/(86400000*7));
      //計算月差
      case "m":return (objDate.getMonth()+1)+((objDate.getFullYear()-this.getFullYear())*12)-(this.getMonth()+1);
      //計算年差
      case "y":return objDate.getFullYear()-this.getFullYear();
      //輸入有誤
      default:return undefined;
    }
  }
function showdate(n)
{
	var uom = new Date();
	uom.setDate(uom.getDate()+n);
	uom = uom.getFullYear() + "-" + (uom.getMonth()+1) + "-" + uom.getDate();
	return uom;
}
