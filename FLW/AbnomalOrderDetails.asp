<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
if Instr(session("AdminPurviewFLW"),"|107,")=0 then 
  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
  response.end
end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName,Depart
Depart=session("Depart")
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
      datafrom=" z_AbnomalOrder a,t_Organization b,t_ICItem c,HM_Employees d,t_measureUnit e "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where a.Custom=b.fitemid and a.Product=c.fitemid and a.FUnit=e.fitemid and a.Saler=d.fitemid and e.funitgroupid=1480 and left(c.FNumber,1)=1 "
	if request("AllQuery")<>"" then
		datawhere=datawhere&" and "&request("AllQuery")
	end if
	datawhere=datawhere&Session("AllMessage31")&Session("AllMessage32")&Session("AllMessage33")&Session("AllMessage34")&Session("AllMessage35")&Session("AllMessage36")&Session("AllMessage37")&Session("AllMessage38")&Session("AllMessage39")&Session("AllMessage40")
	session.contents.remove "AllMessage31"
	session.contents.remove "AllMessage32"
	session.contents.remove "AllMessage33"
	session.contents.remove "AllMessage34"
	session.contents.remove "AllMessage35"
	session.contents.remove "AllMessage36"
	session.contents.remove "AllMessage37"
	session.contents.remove "AllMessage38"
	session.contents.remove "AllMessage39"
	session.contents.remove "AllMessage40"
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "a.FID" 
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
  rs.open sql,connk3,0,1
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
    sql="select a.FID from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = pages '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("FID")
	  else
	    sqlid=sqlid &","&rs("FID")
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
    sql="select a.*,b.FName as CustomName,c.Fname as ProductName,d.Name as SalerName,e.Fname as FUnitName "&_
		" from "& datafrom &" "&_
		" where a.Custom=b.fitemid and a.Product=c.fitemid and a.FUnit=e.fitemid and a.Saler=d.fitemid "&_
		" and e.funitgroupid=1480 and left(c.FNumber,1)=1 and a.FID in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    do until rs.eof'填充数据到表格
		dim stat1,stat2,stat3,stat4,stat5
		stat0="否"
		stat1="否"
		stat2="否"
		stat3="否"
		stat4="否"
		stat5="否"
		if rs("CheckFlag")=1 then stat0="是"
		if rs("PCFlag")=1 then stat1="是"
		if rs("PUFlag")=1 then stat2="是"
		if rs("SAFlag")=1 then stat3="是"
		if rs("ENFlag")=1 then stat4="是"
		if rs("VPFlag")=1 then stat5="是"
%>		
		{"id":"<%=rs("FID")%>",
		"cell":["<%=rs("FBillno")%>","<%=rs("FTime")%>","<%=rs("OrderIDs")%>","<%=JsonStr(rs("CustomName"))%>","<%=JsonStr(rs("ProductName"))%>","<%=rs("SalerName")%>","<%=rs("FUnitName")%>","<%=rs("FQty")%>","<%=rs("UseDate")%>","<%=JsonStr(rs("Remark"))%>","<%=rs("FID")%>","<%=stat0%>","<%=stat1%>","<%=stat5%>"]}
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
  	if  Instr(session("AdminPurviewFLW"),"|107.1,")>0 then
		dim rs2,FIDone
		set rs = server.createobject("adodb.recordset")
		sql="select * from z_AbnomalOrder"
		rs.open sql,connk3,1,3
		rs.addnew
		rs("FClassTypeId")="257709051"
		set rs2=connk3.Execute("select max(Fid) as FID,max(FbillNo) as FbillNo from z_AbnomalOrder")
		if rs2("FID")<>"" then
		rs("FID")=Cint(rs2("FID"))+1
		FIDone=Cint(rs2("FID"))+1
		rs("FbillNo")=rs2("FbillNo")+1
		else
		rs("FID")="1000"
		FIDone=1000
		rs("FbillNo")="10000000"
		end if
		rs2.close
		set rs2=nothing
		if Request.Form("FTime")<>"" then rs("FTime")=Request.Form("FTime")
		if Request.Form("UseDate")<>"" then rs("UseDate")=Request.Form("UseDate")
		rs("OrderIDs")=Request.Form("OrderIDs")
		rs("Custom")=Request.Form("Custom")
		rs("Product")=Request.Form("Product")
		rs("Saler")=Request.Form("Saler")
		rs("FUnit")=Request.Form("FUnit")
		rs("FQty")=Request.Form("FQty")
		rs("Remark")=Request.Form("Remark")
		rs("Biller")=AdminName
		rs("BillDate")=now()
		rs.update
		rs.close
		set rs=nothing 
		for   i=1   to   Request.form("FEntryID").count
			if Request.form("MaterialName")(i)<>"" then
				connk3.Execute("insert into z_AbnomalOrderEntry (FID,FIndex,MaterialId,MaterialName,MaterialType,UseQty,MiniQty,Price) values ("&FIDone&","&i&",'"&Request.form("MaterialId")(i)&"','"&Request.form("MaterialName")(i)&"','"&Request.form("MaterialType")(i)&"','"&Request.form("UseQty")(i)&"','"&Request.form("MiniQty")(i)&"','"&Request.form("Price")(i)&"')")
			end if
		next
		response.write "###"
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
  elseif detailType="Edit"  then
		set rs = server.createobject("adodb.recordset")
    FID=request("FID")
		sql="select * from z_AbnomalOrder where FID="&Request.Form("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.1,")=0 and Instr(session("AdminPurviewFLW"),"|107.2,")=0 and Instr(session("AdminPurviewFLW"),"|107.4,")=0 and Instr(session("AdminPurviewFLW"),"|107.6,")=0 and Instr(session("AdminPurviewFLW"),"|107.8,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
    end if
		if rs("CheckFlag")=0 and Instr(session("AdminPurviewFLW"),"|107.1,")>0 then
			if Request.Form("FTime")<>"" then rs("FTime")=Request.Form("FTime")
			if Request.Form("UseDate")<>"" then rs("UseDate")=Request.Form("UseDate")
			rs("OrderIDs")=Request.Form("OrderIDs")
			rs("Custom")=Request.Form("Custom")
			rs("Product")=Request.Form("Product")
			rs("Saler")=Request.Form("Saler")
			rs("FUnit")=Request.Form("FUnit")
			rs("FQty")=Request.Form("FQty")
			rs("Remark")=Request.Form("Remark")
			rs("Biller")=AdminName
			rs("BillDate")=now()
			rs.update
			rs.close
			set rs=nothing 
			for   i=1   to   Request.form("FEntryID").count
				if Request.form("DeleteFlag")(i)="1" or (Request.form("MaterialName")(i)="" and Request.Form("FEntryID")(i)<>"") then
					connk3.Execute("Delete from z_AbnomalOrderEntry where FEntryID="&Request.Form("FEntryID")(i))
				elseif Request.Form("FEntryID")(i)<>"" then
					connk3.Execute("update z_AbnomalOrderEntry set MaterialId='"&Request.form("MaterialId")(i)&"',MaterialName='"&Request.form("MaterialName")(i)&"',MaterialType='"&Request.form("MaterialType")(i)&"',UseQty='"&Request.form("UseQty")(i)&"',MiniQty='"&Request.form("MiniQty")(i)&"',Price='"&Request.form("Price")(i)&"' where FEntryID="&Request.Form("FEntryID")(i))
				elseif Request.Form("MaterialName")(i)<>"" then
					connk3.Execute("insert into z_AbnomalOrderEntry (FID,FIndex,MaterialId,MaterialName,MaterialType,UseQty,MiniQty,Price) values ("&Request.Form("FID")&","&i&",'"&Request.form("MaterialId")(i)&"','"&Request.form("MaterialName")(i)&"','"&Request.form("MaterialType")(i)&"','"&Request.form("UseQty")(i)&"','"&Request.form("MiniQty")(i)&"','"&Request.form("Price")(i)&"')")
				end if
			next
		elseif rs("CheckFlag")=1 and rs("PCFlag")=0 and Instr(session("AdminPurviewFLW"),"|107.2,")>0 then
			rs("PCOption")=request("PCOption")
			rs("PCer")=AdminName
			rs("PCDate")=now()
		elseif rs("CheckFlag")=1 and rs("PUFlag")=0 and Instr(session("AdminPurviewFLW"),"|107.4,")>0 then
			rs("PUOption")=request("PUOption")
			rs("PUer")=AdminName
			rs("PUDate")=now()
		elseif rs("CheckFlag")=1 and rs("SAFlag")=0 and Instr(session("AdminPurviewFLW"),"|107.6,")>0 then
			rs("SAOption")=request("SAOption")
			rs("SAer")=AdminName
			rs("SADate")=now()
		elseif rs("CheckFlag")=1 and rs("ENFlag")=0 and Instr(session("AdminPurviewFLW"),"|107.8,")>0 then
			rs("ENOption")=request("ENOption")
			rs("ENer")=AdminName
			rs("ENDate")=now()
		else
			response.Write("你没有权限进行此操作或者当前状态不允许此操作！")
    end if
		response.write "###"
  elseif detailType="Delete" then
		set rs = server.createobject("adodb.recordset")
		sql="select * from z_AbnomalOrder where FID="&request("FID")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			response.write ("单据已审核，不允许删除！")
			response.end
		end if
		rs.close
		set rs=nothing 
		connk3.Execute("Delete from z_AbnomalOrder where FID="&request("FID"))
		connk3.Execute("Delete from z_AbnomalOrderEntry where FID="&request("FID"))
		response.write "###"
  elseif detailType="审核" then
		dim direct:direct=request("direct")
		set rs = server.createobject("adodb.recordset")
		sql="select * from z_AbnomalOrder where FID="&request("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		rs("Checker")=AdminName
		rs("CheckDate")=now()
		rs("Remark")=request("operattext")
		if direct="Y" then
		rs("CheckFlag")=1
		elseif direct="N" then
		rs("CheckFlag")=0
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "成功！"
  elseif detailType="生管" then
		direct=request("direct")
		set rs = server.createobject("adodb.recordset")
		sql="select * from z_AbnomalOrder where FID="&request("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.2,")=0 and direct="Z" then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.3,")=0 and (direct="Y" or direct="N") then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("PCFlag")=1 and direct="Z" then
			response.write ("单据当前状态不允许此操作！")
			response.end
		end if
		rs("PCOption")=request("operattext")
		if direct="Y" then
		rs("PCChecker")=AdminName
		rs("PCCheckDate")=now()
		rs("PCFlag")=1
		elseif direct="N" then
		rs("PCChecker")=AdminName
		rs("PCCheckDate")=now()
		rs("PCFlag")=0
		elseif direct="Z" then
		rs("PCer")=AdminName
		rs("PCDate")=now()
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "成功！"
  elseif detailType="采购" then
		direct=request("direct")
		set rs = server.createobject("adodb.recordset")
		sql="select * from z_AbnomalOrder where FID="&request("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.4,")=0 and direct="Z" then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.5,")=0 and (direct="Y" or direct="N") then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("PUFlag")=1 and direct="Z" then
			response.write ("单据当前状态不允许此操作！")
			response.end
		end if
		rs("PUOption")=request("operattext")
		if direct="Y" then
		rs("PUChecker")=AdminName
		rs("PUCheckDate")=now()
		rs("PUFlag")=1
		elseif direct="N" then
		rs("PUChecker")=AdminName
		rs("PUCheckDate")=now()
		rs("PUFlag")=0
		elseif direct="Z" then
		rs("PUer")=AdminName
		rs("PUDate")=now()
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "成功！"
  elseif detailType="营销" then
		direct=request("direct")
		set rs = server.createobject("adodb.recordset")
		sql="select * from z_AbnomalOrder where FID="&request("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.6,")=0 and direct="Z" then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.7,")=0 and (direct="Y" or direct="N") then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("SAFlag")=1 and direct="Z" then
			response.write ("单据当前状态不允许此操作！")
			response.end
		end if
		rs("SAOption")=request("operattext")
		if direct="Y" then
		rs("SAChecker")=AdminName
		rs("SACheckDate")=now()
		rs("SAFlag")=1
		elseif direct="N" then
		rs("SAChecker")=AdminName
		rs("SACheckDate")=now()
		rs("SAFlag")=0
		elseif direct="Z" then
		rs("SAer")=AdminName
		rs("SADate")=now()
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "成功！"
  elseif detailType="工程" then
		direct=request("direct")
		set rs = server.createobject("adodb.recordset")
		sql="select * from z_AbnomalOrder where FID="&request("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.8,")=0 and direct="Z" then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.9,")=0 and (direct="Y" or direct="N") then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("ENFlag")=1 and direct="Z" then
			response.write ("单据当前状态不允许此操作！")
			response.end
		end if
		rs("ENOption")=request("operattext")
		if direct="Y" then
		rs("ENChecker")=AdminName
		rs("ENCheckDate")=now()
		rs("ENFlag")=1
		elseif direct="N" then
		rs("ENChecker")=AdminName
		rs("ENCheckDate")=now()
		rs("ENFlag")=0
		elseif direct="Z" then
		rs("ENer")=AdminName
		rs("ENDate")=now()
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "成功！"
  elseif detailType="副总" then
		direct=request("direct")
		set rs = server.createobject("adodb.recordset")
		sql="select * from z_AbnomalOrder where FID="&request("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.10,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		if rs("VPFlag")=1 and direct="Z" then
			response.write ("单据当前状态不允许此操作！")
			response.end
		end if
		rs("VPer")=AdminName
		rs("VPDate")=now()
		rs("VPOption")=request("operattext")
		if direct="Y" then
		rs("VPFlag")=1
		elseif direct="N" then
		rs("VPFlag")=0
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "成功！"
  elseif detailType="CheckAll" then
		set rs = server.createobject("adodb.recordset")
    FID=request("FID")
		sql="select * from z_AbnomalOrder where FID="&Request.Form("FID")
		rs.open sql,connk3,1,3
		if rs.bof and rs.eof then
			response.write ("数据库读取记录出错！")
			response.end
		end if
		if Instr(session("AdminPurviewFLW"),"|107.1,")>0 and CheckFlag=0 then
			rs("Checker")=AdminName
			rs("CheckDate")=now()
			rs("Remark")=request("Remark")
			rs("CheckFlag")=1
		elseif Instr(session("AdminPurviewFLW"),"|107.3,")>0 and CheckFlag=1 and rs("PCFlag")=0 then
			rs("PCOption")=request("PCOption")
			rs("PCChecker")=AdminName
			rs("PCCheckDate")=now()
			rs("PCFlag")=1
		elseif Instr(session("AdminPurviewFLW"),"|107.5,")>0 and CheckFlag=1 and rs("PUFlag")=0 then
			rs("PUOption")=request("PUOption")
			rs("PUChecker")=AdminName
			rs("PUCheckDate")=now()
			rs("PUFlag")=1
		elseif Instr(session("AdminPurviewFLW"),"|107.7,")>0 and CheckFlag=1 and rs("SAFlag")=0 then
			rs("SAOption")=request("SAOption")
			rs("SAChecker")=AdminName
			rs("SACheckDate")=now()
			rs("SAFlag")=1
		elseif Instr(session("AdminPurviewFLW"),"|107.9,")>0 and CheckFlag=1 and rs("ENFlag")=0 then
			rs("ENOption")=request("ENOption")
			rs("ENChecker")=AdminName
			rs("ENCheckDate")=now()
			rs("ENFlag")=1
		elseif Instr(session("AdminPurviewFLW"),"|107.10,")>0 and CheckFlag=1 and rs("VPFlag")=0 then
			rs("VPer")=AdminName
			rs("VPDate")=now()
			rs("VPOption")=request("VPOption")
		end if
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="FID" then
    InfoID=request("InfoID")
		sql="select a.*,b.FName as CustomName,c.Fname as ProductName,d.Name as SalerName,e.Fname as FUnitName "&_
		" from z_AbnomalOrder a,t_Organization b,t_ICItem c,HM_Employees d,t_measureUnit e "&_
		" where a.Custom=b.fitemid and a.Product=c.fitemid and a.FUnit=e.fitemid and a.Saler=d.fitemid "&_
		" and e.funitgroupid=1480 and left(c.FNumber,1)=1 and a.FID="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
			response.write ("对应单据不存在，请检查！")
			response.end
		else
			response.write "{""Info"":""###"",""fieldValue"":[{"
			for i=0 to rs.fields.count-1
			if IsNull(rs.fields(i).value) then
			response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
			else
			response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
			end if
			next
			response.write ("""Entrys"":[")
			sql="select a.*,b.Stock from z_AbnomalOrderEntry a left join (select a.FNumber,sum(b.FQty) as Stock from t_ICItemCore a,ICInventory b where a.FItemID = b.FItemID group by a.FNumber) b on a.MaterialId=b.FNumber where FID="&InfoID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connk3,1,1
			do until rs.eof
				response.write "{"
				for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
				next
				if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}")
				end if
				rs.movenext
			If Not rs.eof Then
				Response.Write ","
			End If
			loop
			response.write "]}]}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="审核" then
    InfoID=request("InfoID")
		sql="select * from z_AbnomalOrder where FID="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
			if rs.bof and rs.eof then
					response.write ("单据编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""operater"":"""&rs("Biller")&""",""operattime"":"""&rs("Billdate")&""",""Checker"":"""&rs("Checker")&""",""CheckDate"":"""&rs("CheckDate")&""",""operattext"":"""&JsonStr(rs("Remark"))&"""}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="生管" then
    InfoID=request("InfoID")
		sql="select * from z_AbnomalOrder where FID="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
			if rs.bof and rs.eof then
					response.write ("单据编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""operater"":"""&rs("PCer")&""",""operattime"":"""&rs("PCDate")&""",""Checker"":"""&rs("PCChecker")&""",""CheckDate"":"""&rs("PCCheckDate")&""",""operattext"":"""&JsonStr(rs("PCOption"))&"""}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="采购" then
    InfoID=request("InfoID")
		sql="select * from z_AbnomalOrder where FID="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
			if rs.bof and rs.eof then
					response.write ("单据编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""operater"":"""&rs("PUer")&""",""operattime"":"""&rs("PUDate")&""",""Checker"":"""&rs("PUChecker")&""",""CheckDate"":"""&rs("PUCheckDate")&""",""operattext"":"""&JsonStr(rs("PUOption"))&"""}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="营销" then
    InfoID=request("InfoID")
		sql="select * from z_AbnomalOrder where FID="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
			if rs.bof and rs.eof then
					response.write ("单据编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""operater"":"""&rs("SAer")&""",""operattime"":"""&rs("SADate")&""",""Checker"":"""&rs("SAChecker")&""",""CheckDate"":"""&rs("SACheckDate")&""",""operattext"":"""&JsonStr(rs("SAOption"))&"""}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="工程" then
    InfoID=request("InfoID")
		sql="select * from z_AbnomalOrder where FID="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
			if rs.bof and rs.eof then
					response.write ("单据编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""operater"":"""&rs("ENer")&""",""operattime"":"""&rs("ENDate")&""",""Checker"":"""&rs("ENChecker")&""",""CheckDate"":"""&rs("ENCheckDate")&""",""operattext"":"""&JsonStr(rs("ENOption"))&"""}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="副总" then
    InfoID=request("InfoID")
		sql="select * from z_AbnomalOrder where FID="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
			if rs.bof and rs.eof then
					response.write ("单据编号不存在！")
					response.end
		else
			response.write "{""Info"":""###"",""operater"":"""&rs("VPer")&""",""operattime"":"""&rs("VPDate")&""",""Checker"":"""&rs("VPer")&""",""CheckDate"":"""&rs("VPDate")&""",""operattext"":"""&JsonStr(rs("VPOption"))&"""}"
		end if
		rs.close
		set rs=nothing 
  elseif detailType="AutoCompCtm" then
		sql="select top "&request("limit")&" FItemid,FNumber,FName from t_Organization where FNumber like '%"&request("q")&"%' or FName like '%"&request("q")&"%'"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		response.write "["
		do until rs.eof
		Response.Write("{")
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
			next
			if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
			else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}")
			end if
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
  elseif detailType="AutoCompPdt" then
		sql="select top "&request("limit")&"  FItemid,FNumber,FName from t_ICItem where left(FNumber,1)=1 and (FNumber like '%"&request("q")&"%' or FName like '%"&request("q")&"%')"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		response.write "["
		do until rs.eof
		Response.Write("{")
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
			next
			if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
			else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}")
			end if
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
  elseif detailType="AutoCompEmp" then
		sql="select top "&request("limit")&"  FItemid,FNumber,name as FName from HM_Employees where Status=1 and (FNumber like '%"&request("q")&"%' or name like '%"&request("q")&"%')"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		response.write "["
		do until rs.eof
		Response.Write("{")
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
			next
			if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
			else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}")
			end if
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
  elseif detailType="AutoCompUnt" then
		sql="select FItemid,FNumber, FName from t_measureUnit where funitgroupid=1480"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		response.write "["
		do until rs.eof
		Response.Write("{")
			for i=0 to rs.fields.count-2
				if IsNull(rs.fields(i).value) then
					response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&""",")
				else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&""",")
				end if
			next
			if IsNull(rs.fields(i).value) then
				response.write (""""&rs.fields(i).name & """:"""&rs.fields(i).value&"""}")
			else
				response.write (""""&rs.fields(i).name & """:"""&JsonStr(rs.fields(i).value)&"""}")
			end if
	    rs.movenext
		If Not rs.eof Then
		  Response.Write ","
		End If
    loop
		Response.Write("]")
		rs.close
		set rs=nothing 
  elseif detailType="Material" then
    InfoID=request("InfoID")
		sql="select a.FNumber,a.FName,sum(b.FQty) as kc from t_ICItemCore a,ICInventory b where a.FItemID = b.FItemID and a.FNumber='"&InfoID&"' group by a.FNumber,a.FName"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connk3,1,1
		if rs.bof and rs.eof then
				response.write ("物料编号不存在！")
				response.end
		else
			response.write(rs("FNumber")&"###"&rs("FName")&"###"&rs("kc"))
		end if
		rs.close
		set rs=nothing 
  end if
end if
 %>
