var QueryArea=[];
//实例
//QueryArea=[{name:'扣款日期',value:'a.Fdate',type:3},
//{name:'物料名称',value:'f.fName'},
//{name:'处置结果',value:'d.FDefectHandlingID',Content:[{CValue:'1036',title:'拒收'},{CValue:'1077',title:'让步接收'}]},
//{name:'扣款金额',value:'b.famount',type:1}
//];
function AdvSearch(){
	var tr=document.createElement('tr');
	var td = document.createElement('td');
	tr.setAttribute("height","20px");
	td.setAttribute("width","100%");
	td.setAttribute("colSpan","5");
	tr.className='tablemenu';//setAttribute("className","tablemenu");
	tr.setAttribute("class", 'tablemenu');
	td.innerHTML='高级查询';
	$(tr).append(td);
	$('#QueryTable').append(tr);
	
	tr=document.createElement('tr');
	tr.setAttribute("height","20px");
	td = document.createElement('td');
	tr.setAttribute("className","toolbar");
	td = document.createElement('td');
	td.setAttribute("width","20%");
	td.setAttribute("align","center");
	td.innerHTML='字段';
	$(tr).append(td);
	td = document.createElement('td');
	td.setAttribute("width","20%");
	td.setAttribute("align","center");
	td.innerHTML='操作符';
	$(tr).append(td);
	td = document.createElement('td');
	td.setAttribute("width","20%");
	td.setAttribute("align","center");
	td.innerHTML='字段值';
	$(tr).append(td);
	td = document.createElement('td');
	td.setAttribute("width","20%");
	td.setAttribute("align","center");
	td.innerHTML='关系';
	$(tr).append(td);
	td = document.createElement('td');
	td.setAttribute("width","20%");
	$(tr).append(td);
	$('#QueryTable').append(tr);
	
	tr=document.createElement('tr');
	tr.setAttribute("height","20px");
	td=document.createElement('td');
	td.setAttribute("align","center");
	td.innerHTML='<input type="hidden" name="QueryType" id="QueryType" /><input type="hidden" name="AllQuery" id="AllQuery" /><select id="QueryField" name="QueryField" onchange="return ChangeField()" class="adsearchsel"></select>';
 	$(tr).append(td);
	td = document.createElement('td');
	td.setAttribute("align","center");
	td.innerHTML='<select id="QueryOpration" name="QueryOpration" class="adsearchsel"></select>';
	$(tr).append(td);
	td = document.createElement('td');
	td.setAttribute("align","center");
	td.id='td_QueryContent';
	$(tr).append(td);
	td = document.createElement('td');
	td.setAttribute("align","center");
	td.innerHTML='<select id="QueryRelation" name="QueryRelation" class="adsearchsel"> <option value="AND">并且</option> <option value="OR">或者</option> </select>';
	$(tr).append(td);
	td = document.createElement('td');
	td.innerHTML='<input type="button" value="添加" onclick="return AddQuerys()" class="adsearch"/>';
	$(tr).append(td);
	$('#QueryTable').append(tr);
	
	tr=document.createElement('tr');
	td = document.createElement('td');
	td.id='AllQueryStr';
	td.innerHTML="显示查询条件";
	td.setAttribute("colSpan","4");
	td.setAttribute("align","center");
	td.style.cssText="border:1px solid red";
	$(tr).append(td);
	td = document.createElement('td');
	td.innerHTML='<input type="button" value="清空" onclick="return ClearQuerys()" class="adsearch"/>&nbsp;<input type="button" value="查询"  onclick="return doSearch()" class="adsearch"/>';
	$(tr).append(td);
	$('#QueryTable').append(tr);
	td=null;
	tr=null;
	
	var optionStr="<option>请选择字段</option>";
	for(var n=0;n<QueryArea.length;n++){
		optionStr+="<option value='"+QueryArea[n].value+"'>"+QueryArea[n].name+"</option>";
	}
	$('#QueryField').html(optionStr);
}
function ChangeField(){
	var FType=2;
	for(var n=0;n<QueryArea.length;n++){
		if($('#QueryField').val()==QueryArea[n].value){
			if(QueryArea[n].type!==undefined)FType=QueryArea[n].type;
			var validFun='';
			if(FType==1)validFun=' onchange="return checkNum(this)" ';
			if(FType==3)validFun=' onchange="return checkDate(this)" ';
			if(QueryArea[n].Content&&QueryArea[n].Content.length>0){
				var optionStrContent='<select id="QueryContent" name="QueryContent">';
				for(var m=0;m<QueryArea[n].Content.length;m++){
					optionStrContent+='<option value="'+QueryArea[n].Content[m].CValue+'">'+QueryArea[n].Content[m].title+'</option>';
				}
				optionStrContent+='</select>';
				$('#td_QueryContent').html(optionStrContent);
			}else{
				var optionStrContent='<input id="QueryContent" name="QueryContent" type="text" class="textfield"'+validFun+'>';
				$('#td_QueryContent').html(optionStrContent);
				if(FType==3)$('#QueryContent').datepick({dateFormat: 'yyyy-mm-dd'});
			}
		}
	}
	$('#QueryType').val(FType);
	var optionStr="";
	if(FType==1){//数值
		optionStr+="<option value='='>等于</option>";
		optionStr+="<option value='>'>大于</option>";
		optionStr+="<option value='>='>大于等于</option>";
		optionStr+="<option value='<'>小于</option>";
		optionStr+="<option value='<='>小于等于</option>";
	}else if(FType==3){//日期
		optionStr+="<option value='=0'>等于</option>";
		optionStr+="<option value='>0'>大于</option>";
		optionStr+="<option value='>=0'>大于等于</option>";
		optionStr+="<option value='<0'>小于</option>";
		optionStr+="<option value='<=0'>小于等于</option>";
		
	}else{//文本
		optionStr+="<option value='='>等于</option>";
		optionStr+="<option value='like'>包含</option>";
	}
	$('#QueryOpration').html(optionStr);
}
function AddQuerys(){
	if($('#QueryField').val()==''){alert('请先选择条件，再添加！');return false;}
	var thisQuery="",thisQueryStr="";
	if($('#QueryType').val()=="1"){
		thisQuery=" "+$('#QueryField').val()+" "+$('#QueryOpration').val()+" "+$('#QueryContent').val();
		thisQueryStr=" "+$('#QueryField option:selected').text()+" "+$('#QueryOpration').val()+" "+$('#QueryContent').val();
	}
	else if($('#QueryType').val()=="2"){
		if($('#QueryOpration').val()=="="){
			thisQuery=" "+$('#QueryField').val()+" "+$('#QueryOpration').val()+" '"+$('#QueryContent').val()+"'";
			thisQueryStr=" "+$('#QueryField option:selected').text()+" "+$('#QueryOpration').val()+" '"+$('#QueryContent').val()+"'";
		}
		else{
			thisQuery=" "+$('#QueryField').val()+" "+$('#QueryOpration').val()+" '%"+$('#QueryContent').val()+"%'";
			thisQueryStr=" "+$('#QueryField option:selected').text()+" "+$('#QueryOpration').val()+" '%"+$('#QueryContent').val()+"%'";
		}
	}else if($('#QueryType').val()=="3"){
		thisQuery=" datediff(d,'"+$('#QueryContent').val()+"',"+$('#QueryField').val()+") "+$('#QueryOpration').val();
		thisQueryStr=" "+$('#QueryField option:selected').text()+" "+$('#QueryOpration option:selected').text()+" '"+$('#QueryContent').val()+"'";
	}

	if($('#AllQuery').val()==''){
		$('#AllQueryStr').text(thisQueryStr);
		$('#AllQuery').val(thisQuery);
	}else{
		var AllQuery=" ( "+$('#AllQuery').val()+" "+$('#QueryRelation').val()+" "+thisQuery+" ) ";
		var AllQueryStr=" ( "+$('#AllQueryStr').text()+" "+$('#QueryRelation').val()+" "+thisQueryStr+" ) ";
		$('#AllQueryStr').text(AllQueryStr);
		$('#AllQuery').val(AllQuery);
	}
}
function ClearQuerys(){
	$('#AllQueryStr').text('');
	$('#AllQuery').val('');
}