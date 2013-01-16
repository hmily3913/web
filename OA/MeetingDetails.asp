<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|501,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
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
      datafrom=" oa_meeting "
  dim datawhere'数据条件
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
	if searchcols = "id" then
	if isnumeric(searchterm) then
		datawhere = " WHERE " & searchcols & " = " & searchterm & ""
	else
		datawhere = " WHERE " & searchcols & " = 56465453143613645641564643156136135136561345643654"
	End if
	Else
		datawhere = " WHERE " & searchcols & " LIKE '%" & searchterm & "%'"
	End if
	End if
'	datawhere=datawhere&Session("AllMessage16")&Session("AllMessage17")&Session("AllMessage18")&Session("AllMessage19")
'	Session("AllMessage16")=""
'	Session("AllMessage17")=""
'	Session("AllMessage18")=""
'	Session("AllMessage19")=""
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
  dim i'用于循环的整数
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
  '获取本页需要用到的id结束============================================
'-----------------------------------------------------------
'-----------------------------------------------------------
%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    sql="select SerialNum,subject,meeting_type,begin_time,end_time,place,Biller,BillDate,Checker,CheckDate,CheckFlag,Dealer,DealDate,Remark from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格
	dim CheckState:CheckState=""
	if rs("CheckFlag")="1" then
	  CheckState="已通知"
	elseif rs("CheckFlag")="2" then
	  CheckState="已开会"
	elseif rs("CheckFlag")="3" then
	  CheckState="已处理"
	else
	  CheckState="待审核"
	end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("subject")%>","<%=rs("meeting_type")%>","<%=rs("begin_time")%>","<%=rs("starthour")%>","<%=rs("end_time")%>","<%=rs("endhour")%>","<%=rs("place")%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("Checker")%>","<%=rs("CheckDate")%>","<%=CheckState%>","<%=rs("Dealer")%>","<%=rs("CheckDate")%>","<%=JsonStr(rs("Remark"))%>"]}
<%		
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
  end if
  rs.close
  set rs=nothing
response.Write"]}"
'-----------------------------------------------------------
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
'  	if  Instr(session("AdminPurviewFLW"),"|501.1,")>0 then
			set rs = server.createobject("adodb.recordset")
			sql="select * from oa_meeting"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("subject")=Request("subject")
			rs("meeting_type")=Request("meeting_type")
			rs("begin_time")=Request("begin_time")
			rs("end_time")=Request("end_time")
			rs("starthour")=Request("starthour")
			rs("endhour")=Request("endhour")
			rs("place")=Request("place")
			rs("schedule")=Request("schedule")
			rs("attachment1")=Request("attachment1")
			rs("Remark")=Request("Remark")
			rs.update
			rs.close
			set rs=nothing 
			response.write "###"
'		else
'			response.write ("你没有权限进行此操作！")
'			response.end
'		end if
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from oa_meeting where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
'		if Instr(session("AdminPurviewFLW"),"|501.1,")=0 then
'			response.write ("你没有权限进行此操作！")
'			response.end
'			end if
		if rs("CheckFlag")>0 then
			response.write ("当前状态不允许编辑，请检查！")
			response.end
		end if
		rs("Biller")=AdminName
		rs("BillerID")=UserName
		rs("BillDate")=now()
		rs("subject")=Request("subject")
		rs("meeting_type")=Request("meeting_type")
		rs("begin_time")=Request("begin_time")
		rs("end_time")=Request("end_time")
			rs("starthour")=Request("starthour")
			rs("endhour")=Request("endhour")
		rs("place")=Request("place")
		rs("schedule")=Request("schedule")
		rs("attachment1")=Request("attachment1")
		rs("Remark")=Request("Remark")
		response.write "###"
    rs.update
		rs.close
		set rs=nothing 
  elseif detailType="Delete" then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from oa_meeting where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
'		if Instr(session("AdminPurviewFLW"),"|501.1,")=0 then
'			response.write ("你没有权限进行此操作！")
'			response.end
'		end if
		if rs("BillerID")<>UserName then
			response.write ("只能删除本人自己添加的数据！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			response.write ("已经审核不允许删除！")
			response.end
		end if
		rs.close
		set rs=nothing 
		connzxpt.Execute("Delete from oa_meeting where SerialNum="&SerialNum)
		response.write "###"
  elseif detailType="Recorder" then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from oa_meeting where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")<>1 then
			response.write ("此单据已经在实施，不需要审核！")
			response.end
		end if
		rs("Recorder")=AdminName
		rs("RecorderID")=UserName
		rs("RecordDate")=now()
		rs("CheckFlag")=2
		rs("CONTENT")=request("CONTENT")
		rs("attachment2")=request("attachment2")
		rs.update
		rs.close
		set rs=nothing 
		response.write "###"
  elseif detailType="inform" then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from oa_meeting where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			response.write ("此单据状态不允许此操作！")
			response.end
		end if
		rs("Checker")=AdminName
		rs("CheckerID")=UserName
		rs("CheckDate")=now()
		rs("CheckFlag")=1
		dim nnnn
		dim strAttaName
		if rs("attachment1")="" then
		strAttaName=""
		else
		strAttaName=Server.MapPath(rs("attachment1"))
		end if
		for nnnn=0 to UBound(split(request("Emps"),";"))-1
			connzxpt.Execute("Insert into oa_meeting_details (meeting_ID,RecieverID,SysMess,Email,TypeFlag) values ("&SerialNum&",'"&split(request("Emps"),";")(nnnn)&"',"&request("sysMessFlag")&","&request("EmailFlag")&",1)")
			if request("sysMessFlag")=1 then 
				connzxpt.Execute("Insert into smmsys_Message (incept,title,content,inceptuserid,sendtime,sender,senderuserid) values ('"&split(request("Empnames"),";")(nnnn)&"','会议通知','"&rs("subject")&"','"&split(request("Emps"),";")(nnnn)&"','"&now()&"','"&AdminName&"','"&UserName&"')")
			end if
		next
		if request("EmailFlag")=1 then SendMail request("Emails"),"会议通知:"&rs("subject"),SerialNum&"<br>"&"会议时间："&rs("begin_time")&" "&rs("starthour")&"到"&rs("end_time")&" "&rs("endhour")&"<br>"&"会议地点："&rs("place"),rs("schedule")&"("&rs("Biller")&")",strAttaName
		rs.update
		rs.close
		set rs=nothing 
		response.write "会议通知发送成功！"
  elseif detailType="dealwith" then
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from oa_meeting where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")<>2 then
			response.write ("此单据状态不允许此操作！")
			response.end
		end if
		rs("Dealer")=AdminName
		rs("DealerID")=UserName
		rs("DealDate")=now()
		rs("CheckFlag")=3
		if rs("attachment2")="" then
		strAttaName=""
		else
		strAttaName=Server.MapPath(rs("attachment2"))
		end if
		for nnnn=0 to UBound(split(request("Emps"),";"))-1
			connzxpt.Execute("Insert into oa_meeting_details (meeting_ID,RecieverID,SysMess,Email,TypeFlag) values ("&SerialNum&",'"&split(request("Emps"),";")(nnnn)&"',"&request("sysMessFlag")&","&request("EmailFlag")&",2)")
			if request("sysMessFlag")=1 then 
				connzxpt.Execute("Insert into smmsys_Message (incept,title,content,inceptuserid,sendtime,sender,senderuserid) values ('"&split(request("Empnames"),";")(nnnn)&"','会议通知','"&rs("subject")&"','"&split(request("Emps"),";")(nnnn)&"','"&now()&"','"&AdminName&"','"&UserName&"')")
			end if
		next
		if request("EmailFlag")=1 then SendMail request("Emails"),"会议纪要:"&rs("subject"),SerialNum&"<br>"&"会议时间："&rs("begin_time")&" "&rs("starthour")&"到"&rs("end_time")&" "&rs("endhour")&"<br>"&"会议地点："&rs("place"),rs("schedule")&rs("CONTENT")&"("&rs("Biller")&")",strAttaName
		rs.update
		rs.close
		set rs=nothing 
		response.write "会议内容发送成功！"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from oa_meeting where SerialNum="&InfoID
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
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
			end if
			next
				if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}]}")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}]}")
			end if
		end if
		rs.close
		set rs=nothing 
	elseif detailType="mails" then
    set rs = server.createobject("adodb.recordset")
    sql="select a.员工代号 as pk,a.姓名 as name,c.部门代号 as pspk,b.FEmail as val,case when d.Serialnum is null then 'false' else 'true' end as val1 from [N-基本资料单头] a inner join AIS20081217153921.dbo.t_Base_Emp b on a.员工代号=b.FNumber and b.FEmail is not null and FEmail<>'' and a.离职否='在职' inner join [G-部门资料表] c on a.部门别=c.部门代号 left join zxpt.dbo.oa_meeting_details d on d.meeting_ID="&request("SNum")&" and d.RecieverID=a.员工代号 and d.Email=1 and d.TypeFlag="&request("Type")&" union all select distinct c.部门代号 as pk,c.部门名称 as name,'0' as pspk,'' as val,'' as val1 from [N-基本资料单头] a,AIS20081217153921.dbo.t_Base_Emp b,[G-部门资料表] c  where a.员工代号=b.FNumber and b.FEmail is not null and FEmail<>'' and a.离职否='在职' and a.部门别=c.部门代号"
    rs.open sql,conn,1,1
		response.Write("[")
		do until rs.eof
	%>
	{"SerialNum": "<%=rs("pk")%>", "name": "<%=rs("name")%>", "PSNum": "<%=rs("pspk")%>","Email":"<%=rs("val")%>"
	<%
		if rs("val1")<>"" then response.Write(",""checked"":"&rs("val1"))
		if rs("pspk")="0" then
	%>
	, isParent:true
	<%
		end if
		response.Write("}")
			rs.movenext
			If Not rs.eof Then
			  Response.Write ","
			End If
		loop
		Response.Write "]"
		rs.close
		set rs=nothing 
  end if
end if
 %>
