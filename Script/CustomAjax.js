//自定义ajaxjs
//创建xmlHttp实例
var xmlHttp;
function createXMLHttpRequest()
{
	if(window.ActiveXObject)
	{
		xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
	}
	else if(window.XMLHttpRequest)
	{
		xmlHttp = new XMLHttpRequest();
	}
}
//1.获取职员编号，姓名，id
var editInput;
function getEmp(obj){
	editInput=obj;
	if (obj.value != ''){
		createXMLHttpRequest();
		var url="AjaxFunction.asp?key=emp&FItemid="+encodeURI(obj.value);
		xmlHttp.open("get", url , true) ;
		xmlHttp.onreadystatechange=getBackEmp;
		xmlHttp.send(null) ;
	}
}
function getBackEmp(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		if(xmlHttp.responseText.indexOf("###")>-1){
			var arryEmp=xmlHttp.responseText.split("###");
			editInput.value=arryEmp[1];
			editInput.previousSibling.previousSibling.value=arryEmp[0];
			editInput.parentNode.nextSibling.childNodes[0].value=arryEmp[2];
		}else{
			alert("员工编号不存在");
			editInput.value="";
		}
//		alert(xmlHttp.responseText)
	}
}
//2.获取部门信息
function getDepartment(obj){
	editInput=obj;
	if (obj.value != ''){
		createXMLHttpRequest();
		var url="AjaxFunction.asp?key=department&Fitemid="+encodeURI(obj.value);
		xmlHttp.open("get", url , true) ;
		xmlHttp.onreadystatechange=getBackDepartment;
		xmlHttp.send(null) ;
	}
}
function getBackDepartment(){
	var departname=editInput.id;
	var departid=editInput.parentNode.firstChild.id;
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		if(xmlHttp.responseText.indexOf("###")>-1){
			var wordforout;
			var arrDepartMon=xmlHttp.responseText.split("@@@");
			if(arrDepartMon.length-1==1){
				var arryDepartDetail=arrDepartMon[0].split("###");
//				editInput.value=arryDepartDetail[1];
//				editInput.parentNode.firstChild.value=arryDepartDetail[0];

				wordforout="<input name='"+departid+"' type='hidden' id='"+departid+"' value='"+arryDepartDetail[0]+"'>";
				wordforout+="<input name='"+departname+"' type='text' class='textfield' id='"+departname+"' style='WIDTH: 100;' value='"+arryDepartDetail[1]+"' maxlength='100' onBlur='return getDepartment(this)'>";
				editInput.parentNode.innerHTML=wordforout;
//					wordforout="<input name='FBase1' type="hidden" id="FBase1" value="<%=FBase1%>">
//		<input name="FBase1name" type="text" class="textfield" id="FBase1name" style="WIDTH: 140;" value="<%=FBase1name%>" maxlength="100" onBlur="return getDepartment(this)" <%if CheckFlag=1 then response.Write("readonly") end if%>>
			}else{
				wordforout="<select name='"+departid+"' id='"+departid+"' onChange='changeText(this)' style='WIDTH: 70'>";
				for(var n=0;n<arrDepartMon.length-1;n++){
					var arryDepartDetail=arrDepartMon[n].split("###");
					wordforout=wordforout+"<option value='"+arryDepartDetail[0]+"'>"+arryDepartDetail[1]+"</option>";
				}
				wordforout=wordforout+"</select><input type='text' name='"+departname+"' id='"+departname+"' style='WIDTH: 70;font-size:12px' onChange='getDepartment(this)'>";
				editInput.parentNode.innerHTML=wordforout;
			}
		}else{
			alert("部门名称或编号不存在！");
			editInput.value="";
		}
	}
}
function changeText(obj){
	obj.nextSibling.value=obj.options[obj.selectedIndex].text;
}
//3.addrow()
function AddRow(){
//	var tableObject=new Object();
//	tableObject=document.getElementById("editDetails");
	var tbdetail=document.getElementById("TbDetails");
	var CloneNodeTr=tbdetail.rows[1].cloneNode(true);
//	var len = tableObject.rows.length;
	CloneNodeTr.style.display="block";
	tbdetail.appendChild(CloneNodeTr);
}
function DeleteRow(obj,tp){
	if(obj.parentNode.childNodes[0].firstChild.value!=''){
		if (!confirm("确定要删除吗？")) { 
			return false; 
		}else{
			editInput=obj;
				createXMLHttpRequest();
				var url="AjaxFunction.asp?key="+tp+"deleteDetail&FItemid="+obj.parentNode.childNodes[0].firstChild.value;
				xmlHttp.open("get", url , true) ;
				xmlHttp.onreadystatechange=BackDeleteRow;
				xmlHttp.send(null) ;
		}
	}else
		obj.parentNode.parentNode.removeChild(obj.parentNode);
}
function BackDeleteRow(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		if(xmlHttp.responseText.indexOf("###")>-1){
			editInput.parentNode.parentNode.removeChild(editInput.parentNode);
		}else{
			alert("删除异常，请联系管理员！");
		}
	}
}
//删除宿舍人员信息
function deleteEW(v,obj,tb){
	if(v!=''){
		if (!confirm("确定要删除吗？")) { 
			return false; 
		}else{
			editInput=obj;
				createXMLHttpRequest();
				var url;
				if(tb=="DP")
				  url="AjaxFunction.asp?key=deleteDP&FItemid="+v;
				else if(tb=="EW")
				  url="AjaxFunction.asp?key=deleteEW&FItemid="+v;
				else if(tb=="SC")
				  url="AjaxFunction.asp?key=deleteSC&FItemid="+v;
				xmlHttp.open("get", url , true) ;
				xmlHttp.onreadystatechange=BackdeleteEW;
				xmlHttp.send(null) ;
		}
	}else
		alert("删除异常，请重新登录！");
}
function BackdeleteEW(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		if(xmlHttp.responseText.indexOf("###")>-1){
			editInput.parentNode.parentNode.parentNode.removeChild(editInput.parentNode.parentNode);
		}else{
			alert("删除异常，请联系管理员！");
		}
	}
}
//检查宿舍号
function checkDorm(obj){
	if(obj.value!=''){
		editInput=obj;
		createXMLHttpRequest();
		var url="AjaxFunction.asp?key=checkDorm&FItemid="+obj.value;
		xmlHttp.open("get", url , true) ;
		xmlHttp.onreadystatechange=BackcheckDorm;
		xmlHttp.send(null) ;
	}
}
function BackcheckDorm(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		if(xmlHttp.responseText.indexOf("###")>-1){
			alert("宿舍号不允许重复，请检查！");
			editInput.focus();
		}else{
			return true;
		}
	}
}
//获取上月水电信息
function getEW(obj){
	editInput=obj;
	if (obj.value != ''){
		createXMLHttpRequest();
		var url="AjaxFunction.asp?key=EW&FItemid="+encodeURI(obj.value);
		xmlHttp.open("get", url , true) ;
		xmlHttp.onreadystatechange=getBackEW;
		xmlHttp.send(null) ;
	}
}
function getBackEW(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		if(xmlHttp.responseText.indexOf("###")>-1){
			var arryEmp=xmlHttp.responseText.split("###");
			editInput.parentNode.nextSibling.childNodes[0].value=arryEmp[1];
			editInput.parentNode.nextSibling.nextSibling.childNodes[0].value=arryEmp[2];
			editInput.parentNode.nextSibling.nextSibling.nextSibling.childNodes[0].value=arryEmp[3];
		}else{
			alert("宿舍编号不存在");
			editInput.value="";
		}
	}
}

//获取职员编号，姓名
function getEmpName(obj){
	editInput=obj;
	if (obj.value != ''){
		createXMLHttpRequest();
		var url="AjaxFunction.asp?key=empname&FItemid="+encodeURI(obj.value);
		xmlHttp.open("get", url , true) ;
		xmlHttp.onreadystatechange=getBackEmpName;
		xmlHttp.send(null) ;
	}
}
function getBackEmpName(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		if(xmlHttp.responseText.indexOf("###")>-1){
			var arryEmp=xmlHttp.responseText.split("###");
			editInput.value=arryEmp[1];
			editInput.previousSibling.previousSibling.value=arryEmp[0];
		}else{
			alert("员工不存在");
			editInput.value="";
		}
	}
}
//获得司机信息
function getDriver(obj){
	editInput=obj;
	if (obj.value != ''){
		createXMLHttpRequest();
		var url="AjaxFunction.asp?key=getDriver&FItemid="+encodeURI(obj.value);
		xmlHttp.open("get", url , true) ;
		xmlHttp.onreadystatechange=getBackDriver;
		xmlHttp.send(null) ;
	}else{
		alert("车辆不能为空！");
	}
}
function getBackDriver(){
	if (xmlHttp.readyState==4 || xmlHttp.readyState=="complete")
	{
		if(xmlHttp.responseText.indexOf("###")>-1){
			var arryEmp=xmlHttp.responseText.split("###");
			document.getElementById("DriverName").value=arryEmp[1];
			document.getElementById("Driver").value=arryEmp[0];
			document.getElementById("Startemil").value=arryEmp[2];
		}else{
//			alert("员工不存在");
//			editInput.value="";
		}
	}
}
