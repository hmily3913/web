<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|210,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="DetailsList" then 
  dim page'页码
      page=clng(request("page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=request("rp")
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Bill_AllPersonalCtrl "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where 1=1 "
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = Request.Form("qtype")
		if Request.Form("qtype")="CtrlDate" then
			datawhere = datawhere&" and datediff(d,"&searchcols&",'"&searchterm&"')=0 "
		else
			datawhere = datawhere&" and "&searchcols&" like '%"&searchterm&"%' "
		end if
	End if
	datawhere = datawhere&Session("AllMessage25")
	Session("AllMessage25")=""
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "SerialNum" 
	Else
	sortname = Request.Form("sortname")
	End If
	Dim sortorder
	if Request.Form("sortorder") = "" then
	sortorder = "desc"
	Else
	sortorder = Request.Form("sortorder")
	End If
      taxis=" order by "&sortname&" "&sortorder
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
  idCount=rs("idCount")
  if(idcount>0) then'如果记录总数=0,则不处理
    if(idcount mod pages=0)then'如果记录总数除以每页条数有余数,则=记录总数/每页条数+1
	  pagec=int(idcount/pages)'获取总页数
   	else
      pagec=int(idcount/pages)+1'获取总页数
    end if
  end if
	'获取本页需要用到的id============================================
    '读取所有记录的id数值,因为只有id所以速度很快
    sql="select SerialNum from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("SerialNum")
	  else
	    sqlid=sqlid &","&rs("SerialNum")
	  end if
	  rs.movenext
    next
%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    sql="select * from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格'
	dim ys:ys="#f7f7f7"
	if rs("CheckFlag")>0 then ys="#ffff66"
%>
{"id":"<%=rs("SerialNum")%>","ys":"<%=ys%>","cell":[
<%
	  for i=0 to rs.fields.count-2
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).value&""",")
		else
		response.write (""""&JsonStr(rs.fields(i).value)&""",")
		end if
	  next
	    if IsNull(rs.fields(i).value) then
		response.write (""""&rs.fields(i).value&"""]}")
		else
		response.write (""""&JsonStr(rs.fields(i).value)&"""]}")
		end if

	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
  end if
  rs.close
  set rs=nothing
response.Write"]}"
'-----------------------------------------------------------'
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
  	if  Instr(session("AdminPurviewFLW"),"|211.1,")>0 then
			for   i=1   to   Request.form("SubDepart").count 
				if Request.form("SubDepart")(i)<>"" and Request.form("ShouldTo")(i)<>"" then
			    set rs = server.createobject("adodb.recordset")
					sql="select * from Bill_AllPersonalCtrl"
					rs.open sql,connzxpt,1,3
					rs.addnew
					rs("CtrlDate")=Request("CtrlDate")
					rs("MainDepart")=Request("MainDepart")
					rs("Biller")=UserName
					rs("BillDate")=now()
					rs("SubDepart")=Request.form("SubDepart")(i)
					rs("ShouldTo")=Request.form("ShouldTo")(i)
					rs("NewTo")=Request.form("NewTo")(i)
					rs("LeftTo")=Request.form("LeftTo")(i)
					rs("DesertTo")=Request.form("DesertTo")(i)
					rs("LeaveTo")=Request.form("LeaveTo")(i)
					rs("GetinTo")=Request.form("GetinTo")(i)
					rs("OutTo")=Request.form("OutTo")(i)
					rs("RealTo")=Request.form("RealTo")(i)
					rs("WorkFlag")=Request.form("WorkFlag")(i)
					rs("Remark")=Request.form("Remark")(i)
					rs("NewToMan")=Request.form("NewToMan")(i)
					rs("LeftToMan")=Request.form("LeftToMan")(i)
					rs("DesertToMan")=Request.form("DesertToMan")(i)
					rs("LeaveToMan")=Request.form("LeaveToMan")(i)
					rs("GetinToMan")=Request.form("GetinToMan")(i)
					rs("OutToMan")=Request.form("OutToMan")(i)
					rs.update
					rs.close
					set rs=nothing 
				end if
			next
			response.write "###"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit" then
  	if  Instr(session("AdminPurviewFLW"),"|211.1,")>0 then
			for   i=1   to   Request.form("SubDepart").count 
				if Request.form("SerialNum")(i)<>"" and Request.form("ShouldTo")(i)<>"" then
			    set rs = server.createobject("adodb.recordset")
					sql="select * from Bill_AllPersonalCtrl where SerialNum="&Request.form("SerialNum")(i)
					rs.open sql,connzxpt,1,3
					if rs("checkFlag")=1 then
						response.write "已审核，不允许编辑！"
						response.end
					end if
					rs("CtrlDate")=Request("CtrlDate")
					rs("MainDepart")=Request("MainDepart")
					rs("Biller")=UserName
					rs("BillDate")=now()
					rs("SubDepart")=Request.form("SubDepart")(i)
					rs("ShouldTo")=Request.form("ShouldTo")(i)
					rs("NewTo")=Request.form("NewTo")(i)
					rs("LeftTo")=Request.form("LeftTo")(i)
					rs("DesertTo")=Request.form("DesertTo")(i)
					rs("LeaveTo")=Request.form("LeaveTo")(i)
					rs("GetinTo")=Request.form("GetinTo")(i)
					rs("OutTo")=Request.form("OutTo")(i)
					rs("RealTo")=Request.form("RealTo")(i)
					rs("WorkFlag")=Request.form("WorkFlag")(i)
					rs("Remark")=Request.form("Remark")(i)
					rs("NewToMan")=Request.form("NewToMan")(i)
					rs("LeftToMan")=Request.form("LeftToMan")(i)
					rs("DesertToMan")=Request.form("DesertToMan")(i)
					rs("LeaveToMan")=Request.form("LeaveToMan")(i)
					rs("GetinToMan")=Request.form("GetinToMan")(i)
					rs("OutToMan")=Request.form("OutToMan")(i)
					rs.update
					rs.close
					set rs=nothing 
				end if
			next
			response.write "###"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Delete" then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from Bill_AllPersonalCtrl where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,1
		while (not rs.eof)
			if Instr(session("AdminPurviewFLW"),"|211.1,")=0 then
				response.write ("你没有权限进行此操作！")
				response.end
			end if
			if rs("CheckFlag")>0 then
				response.write ("当前状态不允许删除！")
				response.end
			end if
			rs.movenext
		wend
		rs.close
		set rs=nothing 
		connzxpt.Execute("Delete from Bill_AllPersonalCtrl where SerialNum in ("&SerialNum&")")
		response.write "###"
	elseif detailType="Check" then
		set rs = server.createobject("adodb.recordset")
			SerialNum=request("SerialNum")
		sql="select * from Bill_AllPersonalCtrl where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if Instr(session("AdminPurviewFLW"),"|211.2,")=0 then
				response.write ("你没有权限进行当前操作！")
				response.end
			end if
			if rs("CheckFlag")=0 then
				rs("Checker")=AdminName
				rs("CheckDate")=now()
				rs("CheckFlag")=1
				'如果是变更单，审核时同时更新旧版本为不生效状态
				if rs("ShouldTo")<>"" and rs("ShouldTo")<>0 and rs("WorkFlag")="是" then
					connzxpt.Execute("update parametersys_PersonalCtrl set NeedNum="&rs("ShouldTo")&" where MainDepart='"&rs("MainDepart")&"' and SubDepart='"&rs("SubDepart")&"'")
				end if
			end if
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write("###")
  elseif detailType="UnCheck" then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from Bill_AllPersonalCtrl where SerialNum in ("&SerialNum&")"
		rs.open sql,connzxpt,1,3
		while (not rs.eof)
			if Instr(session("AdminPurviewFLW"),"|211.2,")=0 then
				response.write ("你没有权限进行当前操作！")
				response.end
			end if
			if rs("CheckFlag")=1 then
				rs("Checker")=AdminName
				rs("CheckDate")=now()
				rs("CheckFlag")=0
			end if		
			rs.movenext
		wend
		rs.update
		rs.close
		set rs=nothing 
		response.write("###")
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="Register" or detailType="RegisterName" then
    InfoID=request("InfoID")
	if InfoID="" then InfoID=UserName
	sql="select 员工代号,姓名 from [N-基本资料单头] where 员工代号='"&InfoID&"' or 姓名='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write(rs("员工代号")&"###"&rs("姓名"))
	end if
	rs.close
	set rs=nothing 
  elseif detailType="ShowPeson" then
    InfoID=request("InfoID")
		sql="select 员工代号,姓名 from [N-基本资料单头] where 员工代号 in ('"&replace(InfoID,";","','")&"')"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
    do until rs.eof'填充数据到表格'
			response.write rs("姓名")&"###"&rs("员工代号")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ";"
		End If
    loop
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from Bill_AllPersonalCtrl where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
    if rs.bof and rs.eof then
        response.write ("对应单据不存在，请检查！")
        response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":[{"
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
					response.write (""""&rs.fields(i).name & """:"""&replace(replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r"),chr(34),"\""")&""",")
				end if
		  next
	    if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}]}")
			else
				response.write (""""&rs.fields(i).name & """:"""&replace(replace(replace(rs.fields(i).value,chr(10),"\n"),chr(13),"\r"),chr(34),"\""")&"""}]}")
			end if
		end if
		rs.close
		set rs=nothing 
  elseif detailType="MainDepart" then
		sql="select distinct MainDepart from parametersys_PersonalCtrl "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.write "["
    do until rs.eof'填充数据到表格'
			Response.Write(""""&rs("MainDepart")&"""")
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		response.write "]"
		rs.close
		set rs=nothing 
  elseif detailType="SubDepart" then
		InfoID=request("InfoID")
		sql=""
		if request("UseLast")="1" then
			sql="select a.NeedNum,b.*,a.NeedNum-b.DesertTo-b.LeaveTo as ThisRealTo from parametersys_PersonalCtrl a,Bill_AllPersonalCtrl b, "&_
			"(select Max(CtrlDate) as Maxcdt from Bill_AllPersonalCtrl where MainDepart='"&InfoID&"' and CheckFlag=1 and WorkFlag='是') c "&_
			"where a.MainDepart='"&InfoID&"' and a.MainDepart=b.MainDepart "&_
			"and a.SubDepart=b.SubDepart and datediff(d,b.CtrlDate,c.Maxcdt)=0"
		else
			sql="select *,0 as DesertTo,0 as LeaveTo,NeedNum as ThisRealTo,'' as Remark,'' as DesertToMan,'' as LeaveToMan from parametersys_PersonalCtrl where MainDepart='"&InfoID&"'"
		end if
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.write "["
    do until rs.eof'填充数据到表格'
		%>
{"subs":["<%=rs("SubDepart")%>","<%=rs("NeedNum")%>","0","0","<%=rs("DesertTo")%>","<%=rs("LeaveTo")%>","0","0","<%=rs("ThisRealTo")%>","是","<%=JsonStr(rs("Remark"))%>","<%=rs("DesertToMan")%>","<%=rs("LeaveToMan")%>"]}
		<%
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		response.write "]"
		rs.close
		set rs=nothing 
  elseif detailType="Peron" then
		InfoID=request("InfoID")
		sql="select * from [N-基本资料单头] where 员工代号='"&InfoID&"' or 姓名='"&InfoID&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
    if rs.eof and rs.bof then
			response.write "对应员工不存在，请检查！"
			response.end
		end if
		response.write rs("员工代号")&"###"&rs("姓名")
		rs.close
		set rs=nothing 
  elseif detailType="SumAll" then
		Dim FDate
		FDate=request("FMonth")&"#"&request("FYear")
		Dim   MyArray(300,32)
		Dim tempStr1:tempStr1="临时字符串"
		dim tempType(7)
		tempType(0)="应到"
		tempType(1)="新进"
		tempType(2)="离职"
		tempType(3)="旷工"
		tempType(4)="请假"
		tempType(5)="调入"
		tempType(6)="调出"
		tempType(7)="实到"
		Dim RowsNum:RowsNum=-1
		sql="select 1 as ord1,1 as ord2,sum(shouldto) as a0,sum(NewTo) as a1,sum(LeftTo) as a2,sum(DesertTo) as a3,  "&_
"sum(LeaveTo) as a4,sum(GetinTo) as a5,sum(OutTo) as a6,sum(RealTo) as a7,  "&_
"a.maindepart,DATEPART(d,ctrldate) as 日期,b.PersonalType  "&_
"from Bill_AllPersonalCtrl a,parametersys_PersonalCtrl b  "&_
"where a.maindepart=b.maindepart and a.SubDepart=b.SubDepart and CtrlDate>='"&Split(getDateRangebyMonth(FDate),"###")(0)&"' and CtrlDate<='"&Split(getDateRangebyMonth(FDate),"###")(1)&"' and CheckFlag=1 "&_
"group by a.maindepart,PersonalType,ctrldate  "&_
"union all "&_
"select 1 as ord1,2 as ord2,sum(shouldto) as a0,sum(NewTo) as a1,sum(LeftTo) as a2,sum(DesertTo) as a3,  "&_
"sum(LeaveTo) as a4,sum(GetinTo) as a5,sum(OutTo) as a6,sum(RealTo) as a7,  "&_
"a.maindepart,DATEPART(d,ctrldate) as 日期,'汇总' as PersonalType  "&_
"from Bill_AllPersonalCtrl a "&_
"where CtrlDate>='"&Split(getDateRangebyMonth(FDate),"###")(0)&"' and CtrlDate<='"&Split(getDateRangebyMonth(FDate),"###")(1)&"' and CheckFlag=1  "&_
"group by a.maindepart,ctrldate  "&_
"union all "&_
"select 2 as ord1,2 as ord2,sum(shouldto) as a0,sum(NewTo) as a1,sum(LeftTo) as a2,sum(DesertTo) as a3,  "&_
"sum(LeaveTo) as a4,sum(GetinTo) as a5,sum(OutTo) as a6,sum(RealTo) as a7,  "&_
"'公司' as maindepart,DATEPART(d,ctrldate) as 日期,'总和' as PersonalType  "&_
"from Bill_AllPersonalCtrl a "&_
"where CtrlDate>='"&Split(getDateRangebyMonth(FDate),"###")(0)&"' and CtrlDate<='"&Split(getDateRangebyMonth(FDate),"###")(1)&"' and CheckFlag=1  "&_
"group by ctrldate  "&_
"order by ord1,maindepart,ord2,PersonalType,日期 asc"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
    while(not rs.eof)
			if tempStr1<>(rs("maindepart")&rs("PersonalType")) then
				tempStr1=(rs("maindepart")&rs("PersonalType"))
				RowsNum=RowsNum+1
			end if
			dim n
			for n=0 to 7
				MyArray((n+RowsNum*8),0)=rs("maindepart")&rs("PersonalType")
				MyArray((n+RowsNum*8),1)=tempType(n)
				MyArray((n+RowsNum*8),(rs("日期")+1))=rs(("a"&n))
			next
	    rs.movenext
    wend
		rs.close
		set rs=nothing 
'		formatnumber(d1,2)'
		%>
<div id="listtable" style="width:100%; height:100%; overflow:scroll">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td class="tablemenu" colspan="35" height="20" width="100%" id="formove2"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:$('#listDiv').hide()" >&nbsp;<strong><%=request("FYear")%>年<%=request("FMonth")%>月各部门人员出勤情况</strong></font></td>
  </tr>
  <tr class="TitleRow">
    <td nowrap bgcolor="#8DB5E9"><strong>序号</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>人员类别</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>考勤状况</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>1日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>2日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>3日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>4日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>5日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>6日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>7日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>8日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>9日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>10日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>11日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>12日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>13日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>14日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>15日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>16日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>17日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>18日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>19日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>20日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>21日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>22日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>23日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>24日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>25日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>26日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>27日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>28日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>29日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>30日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>31日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>总计</strong></td>
	</tr>
<%
	dim o:o=1
	for n=0 to UBound(MyArray,1)
		if MyArray(n,1)<>"" then
			dim bgcolor,strRealData,totalNum
			totalNum=0
			if InStr(n/8, ".")>0 then
				strRealData=Left(n/8, InStr(n/8, ".") - 1)
			else
				strRealData=n/8
			End if
			if Cint(strRealData) mod 3 =0 then
				bgcolor="#FFFF00"
			elseif Cint(strRealData) mod 3 =1 then
				bgcolor="#EBF2F9"
			elseif Cint(strRealData) mod 3 =2 then
				bgcolor="#CCFFFF"
			end if
			Response.Write "<tr bgcolor='"&bgcolor&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
			Response.Write "<td nowrap >"&o&"</td>" & vbCrLf
			o=o+1
			for m=0 to UBound(MyArray,2)
				Response.Write "<td nowrap >"&MyArray(n,m)&"</td>" & vbCrLf
				if m>1 then totalNum=totalNum+MyArray(n,m)
			next
			Response.Write "<td nowrap >"&totalNum&"</td>" & vbCrLf
			Response.Write "</tr>" & vbCrLf
		end if
	next
%>
</table>
</div>
		<%
  elseif detailType="M1" or detailType="M2" or detailType="M3" or detailType="LQ" or detailType="YJ" or detailType="HSH" then
		dim fenchanstr
		if detailType="M1" then
			fenchanstr=" and a.maindepart='一分厂'"
		elseif detailType="M2" then
			fenchanstr=" and a.maindepart='二分厂' "
		elseif detailType="M3" then
			fenchanstr=" and a.maindepart='三分厂' "
		elseif detailType="LQ" then
			fenchanstr=" and a.maindepart='远华公司' "
		elseif detailType="YJ" then
			fenchanstr=" and a.maindepart='眼镜布绳' "
		elseif detailType="HSH" then
			fenchanstr=" and a.maindepart='花生盒' "
		end if
		FDate=request("FMonth")&"#"&request("FYear")
		tempStr1="临时字符串"
		Redim MyArray(120,32)
		Redim tempType(7)
		tempType(0)="应到"
		tempType(1)="新进"
		tempType(2)="离职"
		tempType(3)="旷工"
		tempType(4)="请假"
		tempType(5)="调入"
		tempType(6)="调出"
		tempType(7)="实到"
		RowsNum=-1
		sql="select 1 as ord1,1 as ord2,shouldto as a0,NewTo as a1,LeftTo as a2,DesertTo as a3,  "&_
"LeaveTo as a4,GetinTo as a5,OutTo as a6,RealTo as a7,   "&_
"a.maindepart,DATEPART(d,ctrldate) as 日期,a.SubDepart as PersonalType  "&_
"from Bill_AllPersonalCtrl a  "&_
"where CtrlDate>='"&Split(getDateRangebyMonth(FDate),"###")(0)&"' and CtrlDate<='"&Split(getDateRangebyMonth(FDate),"###")(1)&"' and CheckFlag=1 and a.SubDepart<>'管理组' "&fenchanstr&_
"union all "&_
"select 1 as ord1,2 as ord2,sum(shouldto) as a0,sum(NewTo) as a1,sum(LeftTo) as a2,sum(DesertTo) as a3,  "&_
"sum(LeaveTo) as a4,sum(GetinTo) as a5,sum(OutTo) as a6,sum(RealTo) as a7,  "&_
"a.maindepart,DATEPART(d,ctrldate) as 日期,b.PersonalType  "&_
"from Bill_AllPersonalCtrl a,parametersys_PersonalCtrl b  "&_
"where a.maindepart=b.maindepart and a.SubDepart=b.SubDepart and CtrlDate>='"&Split(getDateRangebyMonth(FDate),"###")(0)&"' and CtrlDate<='"&Split(getDateRangebyMonth(FDate),"###")(1)&"' and CheckFlag=1 "&fenchanstr&_
"group by a.maindepart,PersonalType,ctrldate  "&_
"union all "&_
"select 1 as ord1,2 as ord2,sum(shouldto) as a0,sum(NewTo) as a1,sum(LeftTo) as a2,sum(DesertTo) as a3,  "&_
"sum(LeaveTo) as a4,sum(GetinTo) as a5,sum(OutTo) as a6,sum(RealTo) as a7,  "&_
"a.maindepart,DATEPART(d,ctrldate) as 日期,'汇总' as PersonalType  "&_
"from Bill_AllPersonalCtrl a "&_
"where CtrlDate>='"&Split(getDateRangebyMonth(FDate),"###")(0)&"' and CtrlDate<='"&Split(getDateRangebyMonth(FDate),"###")(1)&"' and CheckFlag=1  "&fenchanstr&_
"group by a.maindepart,ctrldate  "&_
"order by ord1,maindepart,ord2,PersonalType,日期 asc"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
    while(not rs.eof)
			if tempStr1<>(rs("maindepart")&rs("PersonalType")) then
				tempStr1=(rs("maindepart")&rs("PersonalType"))
				RowsNum=RowsNum+1
			end if
			for n=0 to 7
				MyArray((n+RowsNum*8),0)=rs("maindepart")&rs("PersonalType")
				MyArray((n+RowsNum*8),1)=tempType(n)
				MyArray((n+RowsNum*8),(rs("日期")+1))=rs(("a"&n))
			next
	    rs.movenext
    wend
		rs.close
		set rs=nothing 
'		formatnumber(d1,2)'
		%>
<div id="listtable" style="width:100%; height:100%; overflow:scroll">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td class="tablemenu" colspan="35" height="20" width="100%" id="formove2"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:$('#listDiv').hide()" >&nbsp;<strong><%=request("FYear")%>年<%=request("FMonth")%>月各部门人员出勤情况</strong></font></td>
  </tr>
  <tr class="TitleRow">
    <td nowrap bgcolor="#8DB5E9"><strong>序号</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>人员类别</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>考勤状况</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>1日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>2日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>3日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>4日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>5日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>6日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>7日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>8日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>9日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>10日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>11日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>12日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>13日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>14日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>15日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>16日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>17日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>18日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>19日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>20日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>21日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>22日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>23日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>24日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>25日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>26日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>27日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>28日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>29日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>30日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>31日</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>总计</strong></td>
	</tr>
<%
	o=1
	for n=0 to UBound(MyArray,1)
		if MyArray(n,1)<>"" then
			totalNum=0
			if InStr(n/8, ".")>0 then
				strRealData=Left(n/8, InStr(n/8, ".") - 1)
			else
				strRealData=n/8
			End if
			if Cint(strRealData) mod 3 =0 then
				bgcolor="#FFFF00"
			elseif Cint(strRealData) mod 3 =1 then
				bgcolor="#EBF2F9"
			elseif Cint(strRealData) mod 3 =2 then
				bgcolor="#CCFFFF"
			end if
			Response.Write "<tr bgcolor='"&bgcolor&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
			Response.Write "<td nowrap >"&o&"</td>" & vbCrLf
			o=o+1
			for m=0 to UBound(MyArray,2)
				Response.Write "<td nowrap >"&MyArray(n,m)&"</td>" & vbCrLf
				if m>1 then totalNum=totalNum+MyArray(n,m)
			next
			Response.Write "<td nowrap >"&totalNum&"</td>" & vbCrLf
			Response.Write "</tr>" & vbCrLf
		end if
	next
%>
</table>
</div>
		<%
  elseif detailType="SumAllDetails" then
		FDate=request("FMonth")&"#"&request("FYear")
		sql="select  "&_
"case  "&_
"when a.NewToMan like '%'+b.fnumber+'%' then '新进' "&_
"when a.LeftToMan like '%'+b.fnumber+'%' then '离职' "&_
"when a.DesertToMan like '%'+b.fnumber+'%' then '旷工' "&_
"when a.LeaveToMan like '%'+b.fnumber+'%' then '请假' "&_
"when a.GetinToMan like '%'+b.fnumber+'%' then '调入' "&_
"when a.OutToMan like '%'+b.fnumber+'%' then '调出' "&_
"end as 类型, "&_
"b.fnumber,b.FName,CtrlDate,a.MainDepart,a.SubDepart "&_
"from "&AllOPENROWSET&" zxpt.dbo.bill_AllPersonalCtrl) as a,  "&_
"t_emp as b  "&_
" where ((a.LeaveTo>0 and a.LeaveToMan like '%'+b.fnumber+'%')  "&_
"or (a.NewTo>0 and a.NewToMan like '%'+b.fnumber+'%')  "&_
"or (a.LeftTo>0 and a.LeftToMan like '%'+b.fnumber+'%')  "&_
"or (a.DesertTo>0 and a.DesertToMan like '%'+b.fnumber+'%')  "&_
"or (a.GetinTo>0 and a.GetinToMan like '%'+b.fnumber+'%')  "&_
"or (a.OutTo>0 and a.OutToMan like '%'+b.fnumber+'%')  "&_
") and checkFlag=1 and "&_
"CtrlDate>='"&Split(getDateRangebyMonth(FDate),"###")(0)&"' and CtrlDate<='"&Split(getDateRangebyMonth(FDate),"###")(1)&"' "&_
"order by CtrlDate,MainDepart,SubDepart,类型"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		o=1
		%>
<div id="listtable" style="width:700px; height:450; overflow:scroll">
<table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td class="tablemenu" colspan="7" height="20" width="100%" id="formove2"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="javascript:$('#listDiv').hide()" >&nbsp;<strong><%=request("FYear")%>年<%=request("FMonth")%>月各部门人员出勤情况明细</strong></font></td>
  </tr>
  <tr >
    <td nowrap bgcolor="#8DB5E9"><strong>序号</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>主部门</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>子部门</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>类型</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>员工编号</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>员工姓名</strong></td>
	</tr>
<%
		while (not rs.eof)
			Response.Write "<tr bgcolor='#EBF2F9' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
			Response.Write "<td nowrap >"&o&"</td>" & vbCrLf
			o=o+1
			Response.Write "<td nowrap >"&rs("CtrlDate")&"</td>" & vbCrLf
			Response.Write "<td nowrap >"&rs("MainDepart")&"</td>" & vbCrLf
			Response.Write "<td nowrap >"&rs("SubDepart")&"</td>" & vbCrLf
			Response.Write "<td nowrap >"&rs("类型")&"</td>" & vbCrLf
			Response.Write "<td nowrap >"&rs("fnumber")&"</td>" & vbCrLf
			Response.Write "<td nowrap >"&rs("FName")&"</td>" & vbCrLf
			Response.Write "</tr>" & vbCrLf
			rs.movenext
		wend
		rs.close
		set rs=nothing 
  end if
end if
 %>
