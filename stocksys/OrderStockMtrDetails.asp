<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurview"),"|508,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,UserName,AdminName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")

if showType="DetailsList" then 
  dim page'页码
      page=clng(request("page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=request("rp")
  dim pagec'总页数

  dim datafrom'数据表名
  sql="(Select u1.SerialNum,t1.FNumber AS FMaterialNumber,t1.FName as FMaterialName,"
  sql=sql&" u1.FBatchNo,t5.FName as FSPName,t3.FName as FBUUnitName,u1.FQty as FBUQty, "
  sql=sql&" t7.fqty as icmoqty ,t6.Department,t6.Reason,t6.Gaishan,t6.Checker,case when t7.fqty=0 then 100 else round(isnull(u1.FQty/t7.fqty*100,0),2) end as bili"
  sql=sql&" From ((Select SerialNum,FBrNo,FItemID,FBatchNo,FStockID,FQty,FBal,FStockPlaceID, "
  sql=sql&" FKFPeriod,ISNULL(FKFDate,'') FKFDate,ISNULL(FKFDate,'') FMyKFDate,"
  sql=sql&" 500 as FStockTypeID,FQtyLock,FAuxPropID,FSecQty "
  sql=sql&" From ICInventory where FQty<>0 "
  sql=sql&" union all"
  sql=sql&" Select SerialNum,FBrNo,FItemID,FBatchNo,FStockID,FQty,FBal,FStockPlaceID, "
  sql=sql&" FKFPeriod,ISNULL(FKFDate,'') FKFDate,ISNULL(FKFDate,'') FMyKFDate,"
  sql=sql&" FStockTypeID,0 as FQtyLock,FAuxPropID,FSecQty "
  sql=sql&" From POInventory where FQty<>0 )) as u1 "
  sql=sql&" left join t_ICItem t1 on u1.FItemID = t1.FItemID "
  sql=sql&" left join t_Stock t2 on u1.FStockID=t2.FItemID "
  sql=sql&" left join t_MeasureUnit t3 on t1.FUnitID=t3.FMeasureUnitID "
  sql=sql&" left join t_StockPlace t5 on u1.FStockPlaceID=t5.FSPID "
  sql=sql&" left join SeorderEntry t7 on t7.FMTONo=u1.FBatchNo "
  sql=sql&" left join zxpt.dbo.stocksys_OrderStockMtr t6 on u1.SerialNum=t6.SNum "
  sql=sql&" where Round(u1.FQty,t1.FQtyDecimal)>1 "
  sql=sql&" and (t7.fauxstockqty>0 or t7.fauxstockqty is null) "
  sql=sql&" and t1.FDeleted=0"
'	sql=sql&" and u1.FBatchNo<>'1' "
  sql=sql&" AND t2.FItemID=184 "
  sql=sql&" AND t2.FTypeID in (500,20291,20293) "
  sql=sql&" and ( t1.FNumber like '3.06.%' "
  sql=sql&" or t1.FNumber like '3.17.01.%'"
  sql=sql&" or t1.FNumber like '3.17.02.%'"
  sql=sql&" or t1.FNumber like '3.17.03.%'"
  sql=sql&" or t1.FNumber like '3.17.04.%'"
  sql=sql&" or t1.FNumber like '3.17.11.%'"
  sql=sql&" or t1.FNumber like '3.18.03.%'"
  sql=sql&" or t1.FNumber like '3.18.05.%'"
  sql=sql&" or t1.FNumber like '3.18.06.%'"
  sql=sql&" or t1.FNumber like '3.19.01.%'"
  sql=sql&" or t1.FNumber like '3.19.02.%'"
  sql=sql&" or t1.FNumber like '3.19.06.%'"
  sql=sql&" or t1.FNumber like '3.19.07.%'"
  sql=sql&" or t1.FNumber like '3.19.08.%'"
  sql=sql&" or t1.FNumber like '3.19.09.%')"
  sql=sql&" ) aaa  "
      datafrom=sql
  dim datawhere'数据条件
  dim i'用于循环的整数
	Dim searchterm,searchcols
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = request.Form("qtype")
		if Request.Form("qtype")="CtrlDate" then
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
  dim rs,sql'sql语句
  '获取记录总数
  sql="select count(1) as idCount from "& datafrom &" " & datawhere
'	response.Write(sql)
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
    sql="select SerialNum from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
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
    rs.open sql,connk3,1,1
    do until rs.eof'填充数据到表格'
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("FMaterialNumber")%>","<%=JsonStr(rs("FMaterialName"))%>","<%=rs("FBatchNo")%>","<%=rs("FSPName")%>","<%=rs("FBUUnitName")%>","<%=rs("FBUQty")%>","<%=rs("icmoqty")%>","<%=rs("bili")%>","<%=rs("Department")%>","<%=JsonStr(rs("Reason"))%>","<%=JsonStr(rs("Gaishan"))%>","<%=rs("Checker")%>","<%=rs("CheckDate")%>"]}
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
'-----------------------------------------------------------'
elseif showType="DataProcess" then 
	detailType=request("detailType")
	if detailType="STSave" then
		if Instr(session("AdminPurview"),"|508.1,")=0 then
			response.Write("你没有权限进行此操作！")
			response.End()
		end if
		sql="select * from stocksys_OrderStockMtr where SNum="&request("SerialNum")
		set rs = server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,3
		if rs.eof then rs.addnew
		rs("SNum")=request("SerialNum")
		rs("Department")=request("Department")
		rs("Reason")=request("Reason")
		rs("CheckerID")=request("CheckerID")
		rs("Checker")=request("Checker")
		rs("Biller")=AdminName
		rs("BillerID")=UserName
		rs("BillDate")=now()
		rs.update
		rs.close
		set rs=nothing 
		response.write "保存成功！"
	elseif detailType="MNSave" then
		sql="select * from stocksys_OrderStockMtr where SNum="&request("SerialNum")
		set rs = server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,3
		if not rs.eof then
			if rs("CheckerID")=UserName then
				rs("Gaishan")=request("Gaishan")
				rs("CheckDate")=now()
			else
				response.Write("只有责任人才能进行此操作！")
				response.End()
			end if
		else
			response.Write("未判定责任部门责任人，无法回复！")
			response.End()
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "保存成功！"
	end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="CheckerID" then
    InfoID=request("InfoID")
		sql="select 员工代号,姓名 from [N-基本资料单头] where 员工代号 like '%"&InfoID&"%' or 姓名 like '%"&InfoID&"%' "
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
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
  sql="(Select u1.SerialNum,t1.FNumber AS FMaterialNumber,t1.FName as FMaterialName,"
  sql=sql&" u1.FBatchNo,t5.FName as FSPName,t3.FName as FBUUnitName,u1.FQty as FBUQty, "
  sql=sql&" t7.fqty as icmoqty ,t6.Department,t6.Reason,t6.Gaishan,t6.Checker,t6.CheckerID,t6.CheckDate"
  sql=sql&" From ((Select SerialNum,FBrNo,FItemID,FBatchNo,FStockID,FQty,FBal,FStockPlaceID, "
  sql=sql&" FKFPeriod,ISNULL(FKFDate,'') FKFDate,ISNULL(FKFDate,'') FMyKFDate,"
  sql=sql&" 500 as FStockTypeID,FQtyLock,FAuxPropID,FSecQty "
  sql=sql&" From ICInventory where FQty<>0 "
  sql=sql&" union all"
  sql=sql&" Select SerialNum,FBrNo,FItemID,FBatchNo,FStockID,FQty,FBal,FStockPlaceID, "
  sql=sql&" FKFPeriod,ISNULL(FKFDate,'') FKFDate,ISNULL(FKFDate,'') FMyKFDate,"
  sql=sql&" FStockTypeID,0 as FQtyLock,FAuxPropID,FSecQty "
  sql=sql&" From POInventory where FQty<>0 )) as u1 "
  sql=sql&" left join t_ICItem t1 on u1.FItemID = t1.FItemID "
  sql=sql&" left join t_Stock t2 on u1.FStockID=t2.FItemID "
  sql=sql&" left join t_MeasureUnit t3 on t1.FUnitID=t3.FMeasureUnitID "
  sql=sql&" left join t_StockPlace t5 on u1.FStockPlaceID=t5.FSPID "
  sql=sql&" left join SEOrderEntry t7 on t7.FMTONo=u1.FBatchNo "
  sql=sql&" left join zxpt.dbo.stocksys_OrderStockMtr t6 on u1.SerialNum=t6.SNum "
  sql=sql&" where Round(u1.FQty,t1.FQtyDecimal)>1 "
  sql=sql&" and (t7.fauxstockqty>0 or t7.fauxstockqty is null)"
  sql=sql&" and t1.FDeleted=0"
  sql=sql&" AND t2.FItemID=184 "
  sql=sql&" AND t2.FTypeID in (500,20291,20293) "
  sql=sql&" and ( t1.FNumber like '3.06.%' "
  sql=sql&" or t1.FNumber like '3.17.01.%'"
  sql=sql&" or t1.FNumber like '3.17.02.%'"
  sql=sql&" or t1.FNumber like '3.17.03.%'"
  sql=sql&" or t1.FNumber like '3.17.04.%'"
  sql=sql&" or t1.FNumber like '3.17.11.%'"
  sql=sql&" or t1.FNumber like '3.18.03.%'"
  sql=sql&" or t1.FNumber like '3.18.05.%'"
  sql=sql&" or t1.FNumber like '3.18.06.%'"
  sql=sql&" or t1.FNumber like '3.19.01.%'"
  sql=sql&" or t1.FNumber like '3.19.02.%'"
  sql=sql&" or t1.FNumber like '3.19.06.%'"
  sql=sql&" or t1.FNumber like '3.19.07.%'"
  sql=sql&" or t1.FNumber like '3.19.08.%'"
  sql=sql&" or t1.FNumber like '3.19.09.%')"
  sql=sql&" ) aaa  "
		sqlall="select * from "&sql&" where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sqlall,connk3,1,1
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
	end if
end if
 %>
