<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurviewFLW"),"|109,")=0 then 
  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
  response.end
end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName,Depart
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
Depart=session("Depart")
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
      datafrom=" Flw_PrenatalTest "
  dim datawhere'数据条件
    datawhere=" where 1=1 "
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
		if Request.Form("qtype")="BillDate" then
			datawhere = datawhere&" and datediff(d,"&searchcols&",'"&searchterm&"')=0 "
		else
			datawhere = datawhere&" and "&searchcols&" like '%"&searchterm&"%' "
		end if
	End if
	datawhere=datawhere&Session("AllMessage64")&Session("AllMessage65")&Session("AllMessage66")&Session("AllMessage67")&Session("AllMessage68")
	session.contents.remove "AllMessage64"
	session.contents.remove "AllMessage65"
	session.contents.remove "AllMessage66"
	session.contents.remove "AllMessage67"
	session.contents.remove "AllMessage68"
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
    sql="select * from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
    do until rs.eof'填充数据到表格
		dim checkstat:checkstat="未审核"
		if rs("CheckFlag")=1 then
			checkstat="工程确定"
		elseif rs("CheckFlag")=2 then
			checkstat="生管确定"
		elseif rs("CheckFlag")=3 then
			checkstat="技术员确定"
		elseif rs("CheckFlag")=4 then
			checkstat="分厂确定"
		elseif rs("CheckFlag")=5 then
			checkstat="品保确定"
		end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("BillDate")%>","<%=rs("OrderID")%>","<%=rs("ProductID")%>","<%=rs("ProductName")%>","<%=rs("CustomLv")%>","<%=rs("Employee")%>","<%=rs("QualityLv")%>","<%=rs("PreMeetDate")%>","<%=rs("PostMeetDate")%>","<%=rs("TestType")%>","<%=checkstat%>","<%=rs("Biller")%>","<%=rs("Checker")%>","<%=rs("Determine")%>"]}
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
  	if  Instr(session("AdminPurviewFLW"),"|412.1,")>0 then
			set rs = server.createobject("adodb.recordset")
			sql="select * from Flw_PrenatalTest"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("OrderID")=Request("OrderID")
			rs("CustomLv")=Request("CustomLv")
			rs("QualityLv")=Request("QualityLv")
			rs("Employee")=Request("Employee")
			rs("TestType")=Request("TestType")
			rs("ProductID")=Request("ProductID")
			rs("ProductName")=Request("ProductName")
			rs("TestContact")=Request("TestContact")
			if Request("IsProcess") then rs("IsProcess")=Request("IsProcess")
			if Request("IsMold") then rs("IsMold")=Request("IsMold")
			if Request("IsCarve") then rs("IsCarve")=Request("IsCarve")
			if Request("IsFilm") then rs("IsFilm")=Request("IsFilm")
			if Request("IsCopper") then rs("IsCopper")=Request("IsCopper")
			if Request("IsDraw") then rs("IsDraw")=Request("IsDraw")
			rs("EnEndDate")=Request("EnEndDate")
			rs.update
			rs.close
			set rs=nothing 
			response.write "保存成功！"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Flw_PrenatalTest where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if rs("CheckFlag")=0 then
			if Instr(session("AdminPurviewFLW"),"|109.1,")=0 then
				response.write ("你没有权限进行此操作！")
				response.end
			end if
			if rs("BillerID")<>UserName then
				response.write ("只能编辑自己添加的数据！")
				response.end
			end if
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("OrderID")=Request("OrderID")
			rs("CustomLv")=Request("CustomLv")
			rs("QualityLv")=Request("QualityLv")
			rs("Employee")=Request("Employee")
			rs("TestType")=Request("TestType")
			rs("ProductID")=Request("ProductID")
			rs("ProductName")=Request("ProductName")
			rs("TestContact")=Request("TestContact")
			if Request("IsProcess") then rs("IsProcess")=Request("IsProcess")
			if Request("IsMold") then rs("IsMold")=Request("IsMold")
			if Request("IsCarve") then rs("IsCarve")=Request("IsCarve")
			if Request("IsFilm") then rs("IsFilm")=Request("IsFilm")
			if Request("IsCopper") then rs("IsCopper")=Request("IsCopper")
			if Request("IsDraw") then rs("IsDraw")=Request("IsDraw")
			rs("EnEndDate")=Request("EnEndDate")
    	rs.update
			rs.close
			set rs=nothing 
			response.write "修改成功！"
		elseif rs("CheckFlag")=2 then
			if Instr(session("AdminPurviewFLW"),"|109.4,")=0 then
				response.write ("你没有权限进行此操作！")
				response.end
			end if
			for i=2 to Request.form("SerialNumD").count
				if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNumD")(i)<>"" then
					connzxpt.Execute("Delete from Flw_PrenatalTestDetails where SerialNumD="&Request.Form("SerialNumD")(i))
				elseif Request.Form("SerialNumD")(i)<>"" then
					connzxpt.Execute("update Flw_PrenatalTestDetails set Problem='"&Request.form("Problem")(i)&"',Standard='"&Request.form("Standard")(i)&"',Program='"&Request.form("Program")(i)&"',Responser='"&Request.form("Responser")(i)&"',FinishDate='"&Request.form("FinishDate")(i)&"' where SerialNumD="&Request.Form("SerialNumD")(i))
				elseif Request.form("Problem")(i)<>"" then
						connzxpt.Execute("insert into Flw_PrenatalTestDetails (SNum,Problem,Standard,Program,Responser,FinishDate) values ('"&SerialNum&"','"&Request.form("Problem")(i)&"','"&Request.form("Standard")(i)&"','"&Request.form("Program")(i)&"','"&Request.form("Responser")(i)&"','"&Request.form("FinishDate")(i)&"')")
				end if
			next
			response.write "保存成功！"
		else
			response.Write("当前状态不允许编辑！")
			response.End()
		end if
  elseif detailType="Delete" then
		if Instr(session("AdminPurviewFLW"),"|109.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		sql="select * from Flw_PrenatalTest where CheckFlag=0 and SerialNum="&SerialNum
		if rs.eof and rs.bof then
			rs.close
			set rs=nothing
			response.Write("单据不允许删除！删除失败！")
			response.End()
		else
			rs.close
			set rs=nothing
			connzxpt.Execute("Delete from Flw_PrenatalTest where SerialNum="&SerialNum)
			connzxpt.Execute("Delete from Flw_PrenatalTestDetails where SerialNum="&SerialNum)
			response.write "删除成功！"
		end if
  elseif detailType="Check" then
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from Flw_PrenatalTest where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs("CheckFlag")=0 and Instr(session("AdminPurviewFLW"),"|109.2,")>0 then
			rs("CheckFlag")=1
			rs("Checker")=AdminName
			rs("CheckerID")=UserName
			rs("CheckDate")=now()
			rs("EnEndDate")=Request.form("EnEndDate")
			rs.update
			rs.close
			set rs=nothing
			response.Write("确定成功")
		elseif rs("CheckFlag")=1 and Instr(session("AdminPurviewFLW"),"|109.3,")>0 then
			if Request.form("MtrProvidDate")="" or Request.form("PreMeetDate")="" or Request.form("PostMeetDate")="" then
				response.Write("物料交期、产前会日期、产后总结会日期不能为空！")
				response.End()
			end if
			rs("CheckFlag")=2
			rs("PMer")=AdminName
			rs("PMerID")=UserName
			rs("PMDate")=now()
			rs("MtrProvidDate")=Request.form("MtrProvidDate")
			rs("PreMeetDate")=Request.form("PreMeetDate")
			rs("PostMeetDate")=Request.form("PostMeetDate")
			rs.update
			rs.close
			set rs=nothing
			response.Write("确定成功")
		elseif rs("CheckFlag")=2 and Instr(session("AdminPurviewFLW"),"|109.4,")>0 then
			rs("CheckFlag")=3
			rs("Techer")=AdminName
			rs("TecherID")=UserName
			rs("TechDate")=now()
			rs.update
			rs.close
			set rs=nothing
			response.Write("确定成功")
		elseif rs("CheckFlag")=3 and Instr(session("AdminPurviewFLW"),"|109.5,")>0 then
			rs("CheckFlag")=4
			rs("Brancher")=AdminName
			rs("BrancherID")=UserName
			rs("BranchDate")=now()
			rs.update
			rs.close
			set rs=nothing
			response.Write("确定成功")
		elseif rs("CheckFlag")=4 and Instr(session("AdminPurviewFLW"),"|109.6,")>0 then
			if Request.form("IsTestStand") then rs("IsTestStand")=Request.form("IsTestStand")
			if Request.form("IsTestViscos") then rs("IsTestViscos")=Request.form("IsTestViscos")
			if Request.form("IsTestPull") then rs("IsTestPull")=Request.form("IsTestPull")
			if Request.form("IsTestDown") then rs("IsTestDown")=Request.form("IsTestDown")
			if Request.form("IsTestPack") then rs("IsTestPack")=Request.form("IsTestPack")
			rs("TestViscos")=Request.form("TestViscos")
			rs("TestPull")=Request.form("TestPull")
			rs("TestDown")=Request.form("TestDown")
			rs("TestPack")=Request.form("TestPack")
			rs("Determine")=Request.form("Determine")
			rs("CheckFlag")=5
			rs("Qcer")=AdminName
			rs("QcerID")=UserName
			rs("QcDate")=now()
			rs.update
			rs.close
			set rs=nothing
			response.Write("确定成功")
		else
			response.Write("当前单据不允许审核或者你没有进行此操作的权限！")
			response.End()
		end if
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="ProductID" then
    InfoID=request("InfoID")
		sql="select FNumber,FName from t_icitemcore where FNumber ='"&InfoID&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
				response.write ("产品编号不存在！")
				response.end
		else
			response.write(rs("FNumber")&"###"&rs("FName"))
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from Flw_PrenatalTest where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			response.write ("对应单据不存在，请检查！")
			response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":[{"
			for i=0 to rs.fields.count-1
			if IsNull(rs.fields(i).value) then
				if Left(rs.fields(i).name,2)="Is" then
					response.write (""""&rs.fields(i).name & """:"&LCase(cstr(rs.fields(i).value))&",")
				else
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				end if
			else
				if Left(rs.fields(i).name,2)="Is" then
					response.write (""""&rs.fields(i).name & """:"&LCase(cstr(rs.fields(i).value))&",")
				else
					response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
			end if
			next
			response.write ("""Entrys"":[")
			sql="select * from Flw_PrenatalTestDetails where SNum="&InfoID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connzxpt,1,1
			do until rs.eof
				response.write "{"
				for i=0 to rs.fields.count-1
				if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
				next
				response.write ("""bg"":""#EBF2F9""}")
				rs.movenext
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write "]}]}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="OrderID" then
    InfoID=request("InfoID")
		sql="select a.Fbillno,b.F_103,b.F_104,b.Fnumber,c.FName from t_DHTZD a,t_Organization b,t_emp c where a.Fbillno='"&InfoID&"' and a.FBase7=b.Fitemid and a.FBase3=c.fitemid "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
			response.write ("对应订单不存在！")
			response.end
		else
			response.write rs("Fbillno")&"###"&rs("F_103")&"###"&rs("F_104")&"###"&rs("Fname")
		end if
		rs.close
		set rs=nothing 
	end if
end if
 %>
