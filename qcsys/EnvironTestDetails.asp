<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurview"),"|412,")=0 then 
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
      datafrom=" qcsys_EnvironTest "
  dim datawhere'数据条件
    datawhere=" where 1=1 "
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
		if Request.Form("qtype")="RegDate" then
			datawhere = datawhere&" and datediff(d,"&searchcols&",'"&searchterm&"')=0 "
		else
			datawhere = datawhere&" and "&searchcols&" like '%"&searchterm&"%' "
		end if
	End if
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
		dim TestFee
		if Instr(session("AdminPurview"),"|412.1,")=0 then
			TestFee="*****"
		else
			TestFee=rs("TestFee")
		end if
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("TestYear")%>","<%=rs("RegDate")%>","<%=rs("OrderId")%>","<%=rs("Employer")%>","<%=rs("ProductID")%>","<%=rs("ProductName")%>","<%=rs("TestItem")%>","<%=TestFee%>","<%=rs("Feiyong")%>","<%=JsonStr(rs("TestAgency"))%>","<%=JsonStr(rs("DetectResult"))%>","<%=JsonStr(rs("Remark"))%>","<%=rs("Biller")%>"]}
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
  	if  Instr(session("AdminPurview"),"|412.1,")>0 then
			set rs = server.createobject("adodb.recordset")
			sql="select * from qcsys_EnvironTest"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("TestYear")=Request("TestYear")
			rs("RegDate")=Request("RegDate")
			rs("OrderId")=Request("OrderId")
			rs("Employer")=Request("Employer")
			rs("CustomID")=Request("CustomID")
			rs("ProductID")=Request("ProductID")
			rs("ProductName")=Request("ProductName")
			rs("TestItem")=Request("TestItem")
			rs("TestFee")=Request("TestFee")
			rs("TestAgency")=Request("TestAgency")
			rs("SupplyID")=Request("SupplyIDMain")
			rs("BomID")=Request("BomID")
			rs("Feiyong")=Request("Feiyong")
			rs("DetectResult")=Request("DetectResult")
			rs("DetectReport")=Request("DetectReport")
			rs("Remark")=Request("RemarkMain")
			rs.update
			set rs=connzxpt.Execute("select top 1 SerialNum from qcsys_EnvironTest order by serialnum desc")
			if rs("SerialNum")="" then
			SerialNum=1
			else
			SerialNum=rs("SerialNum")
			end if
			for   i=2   to   Request.form("SerialNumD").count
				if (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") and Request.form("MaterialID")(i)<>"" then
					connzxpt.Execute("insert into qcsys_EnvironTestDetails (SNum,MaterialID,MaterialName,SupplyID,Procurement,Remark) values ('"&SerialNum&"','"&Request.form("MaterialID")(i)&"','"&Request.form("MaterialName")(i)&"','"&Request.form("SupplyID")(i)&"','"&Request.form("Procurement")(i)&"','"&Request.form("Remark")(i)&"')")
				end if
			next
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
		sql="select * from qcsys_EnvironTest where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurview"),"|412.1,")=0 then
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
			rs("TestYear")=Request("TestYear")
			rs("Employer")=Request("Employer")
			rs("RegDate")=Request("RegDate")
			rs("OrderId")=Request("OrderId")
			rs("CustomID")=Request("CustomID")
			rs("ProductID")=Request("ProductID")
			rs("ProductName")=Request("ProductName")
			rs("TestItem")=Request("TestItem")
			rs("TestFee")=Request("TestFee")
			rs("TestAgency")=Request("TestAgency")
			rs("SupplyID")=Request("SupplyIDMain")
			rs("Feiyong")=Request("Feiyong")
			rs("BomID")=Request("BomID")
			rs("DetectResult")=Request("DetectResult")
			rs("DetectReport")=Request("DetectReport")
			rs("Remark")=Request("RemarkMain")
    rs.update
		for   i=2   to   Request.form("SerialNumD").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNumD")(i)<>"" then
				connzxpt.Execute("Delete from qcsys_EnvironTestDetails where SerialNumD="&Request.Form("SerialNumD")(i))
			elseif Request.Form("SerialNumD")(i)<>"" then
				connzxpt.Execute("update qcsys_EnvironTestDetails set MaterialID='"&Request.form("MaterialID")(i)&"',MaterialName='"&Request.form("MaterialName")(i)&"',SupplyID='"&Request.form("SupplyID")(i)&"',Procurement='"&Request.form("Procurement")(i)&"',Remark='"&Request.form("Remark")(i)&"' where SerialNumD="&Request.Form("SerialNumD")(i))
			elseif Request.form("MaterialID")(i)<>"" then
					connzxpt.Execute("insert into qcsys_EnvironTestDetails (SNum,MaterialID,MaterialName,SupplyID,Procurement,Remark) values ('"&SerialNum&"','"&Request.form("MaterialID")(i)&"','"&Request.form("MaterialName")(i)&"','"&Request.form("SupplyID")(i)&"','"&Request.form("Procurement")(i)&"','"&Request.form("Remark")(i)&"')")
			end if
		next
		rs.close
		set rs=nothing 
		response.write "修改成功！"
  elseif detailType="Delete" then
		if Instr(session("AdminPurview"),"|412.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
		connzxpt.Execute("Delete from qcsys_EnvironTest where SerialNum in ("&SerialNum&")")
		connzxpt.Execute("Delete from qcsys_EnvironTestDetails where SNum in ("&SerialNum&")")
		response.write "删除成功！"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="ProductID" or detailType="MaterialID" then
    InfoID=request("InfoID")
		sql="select FNumber,FName from t_icitemcore where FNumber ='"&InfoID&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
				response.write ("员工编号不存在！")
				response.end
		else
			response.write(rs("FNumber")&"###"&rs("FName"))
		end if
		rs.close
		set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from qcsys_EnvironTest where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			response.write ("对应单据不存在，请检查！")
			response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":[{"
			for i=0 to rs.fields.count-1
			if IsNull(rs.fields(i).value) then
				if rs.fields(i).name="Remark" then
					response.write ("""RemarkMain"":"""&rs.fields(i).value&""",")
				elseif rs.fields(i).name="SupplyID" then
					response.write ("""SupplyIDMain"":"""&rs.fields(i).value&""",")
				else
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				end if
			else
				if rs.fields(i).name="Remark" then
					response.write ("""RemarkMain"":"""&JsonStr(rs.fields(i).value)&""",")
				elseif rs.fields(i).name="SupplyID" then
					response.write ("""SupplyIDMain"":"""&rs.fields(i).value&""",")
				else
					response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
			end if
			next
			response.write ("""Entrys"":[")
			sql="select * from qcsys_EnvironTestDetails where SNum="&InfoID
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
  elseif detailType="OrderId" then
    InfoID=request("InfoID")
		sql="select a.Fbillno,b.Fnumber,c.Fname from Seorder a,t_Organization b,t_emp c where a.Fbillno='"&InfoID&"' and a.FCustID=b.Fitemid and a.fempid=c.fitemid "
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
			response.write ("对应订单不存在！")
			response.end
		else
			response.write rs("Fbillno")&"###"&rs("Fnumber")&"###"&rs("Fname")
		end if
		rs.close
		set rs=nothing 
  elseif detailType="showBom" then
    InfoID=request("InfoID")
		sql="select distinct f.Fnumber,f.Fname "
		sql=sql&" from Seorder a,Seorderentry b,t_ICitemcore c, "
		sql=sql&" icbom d,iccustbomchild e,t_ICitemcore f  "
		sql=sql&" where a.FInterID=b.FInterID and b.FItemid=c.Fitemid  "
		sql=sql&" and d.FinterID=e.FinterID and e.Fitemid=f.fItemid "
		sql=sql&" and b.FBomInterID=d.FinterID and a.Fbillno='"&InfoID&"' and c.FNumber='"&request("ProductID")&"'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
			response.write ("对应BOM不存在！")
			response.end
		else'"&rs("FBomNumber")&"
			response.write "{""Info"":""###"",""fieldValue"":"""","
			response.write ("""Entrys"":[")
			do until rs.eof
				response.write "{""MaterialID"":"""&rs("Fnumber")&""",""MaterialName"":"""&rs("Fname")&"""}"
				rs.movenext
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write "]}"
		end if
		rs.close
		set rs=nothing 
	end if
end if
 %>
