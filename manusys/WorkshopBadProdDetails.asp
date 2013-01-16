<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurview"),"|409,")=0 then 
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
  dim datawhere'数据条件
  dim rs,sql,idCount,page,temstr,pages
	datawhere=""
	if request("AllQuery")<>"" then
		datawhere=datawhere&" and "&request("AllQuery")
	  sql="select count(1) as idCount  from icmo,t_icitemcore d where 1=1 and icmo.Fitemid=d.Fitemid and "&request("AllQuery")
	else
	  sql="select count(1) as idCount  from icmo a "
	end if
  dim i:i=1
	page=clng(request("page"))
	pages=request("rp")
	temstr="###"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
  idCount=rs("idCount")

%>
{"page":"<%=page%>","total":"<%=idcount%>","rows":[
<%
    sql="select icmo.FinterID,icmo.FCheckdate,icmo.FCloseDate,c.FName as n1,icmo.FMtono, "
    sql=sql&" d.FNumber,d.Fname as n2,icmo.Fbillno,icmo.FauxQty,icmo.FAuxQtyFinish, "
    sql=sql&" icmo.FAuxStockQty,sum(n.FAuxQtyPick) as yfsl,sum(n.FAuxStockQty) as slsl,sum(n.FAuxQtySupply) as blsl,round(sum(e.ReductQty),2) as bfsl  "
    sql=sql&" from icmo inner join t_department c on icmo.FworkShop=c.Fitemid and icmo.FCancellation=0 "
    sql=sql&" inner join t_icitemcore d on icmo.Fitemid=d.Fitemid  "
    sql=sql&" left join PPBom m on icmo.FinterID=m.FIcmoInterID "
    sql=sql&" left join PPBomEntry n on m.FInterID=n.FInterID "
    sql=sql&" left join zxpt.dbo.manusys_WorkshopBadProd e on icmo.FBillNo=e.icmo_id and n.FEntryID=e.itemid where 1=1  "

	if request("AllQuery")<>"" then sql=sql&" and "&request("AllQuery")
    sql=sql&" group by icmo.FinterID,icmo.FCheckdate,icmo.FCloseDate,c.FName,icmo.FMtono,d.FNumber,d.Fname,icmo.Fbillno,icmo.FauxQty,icmo.FAuxQtyFinish,icmo.FAuxStockQty order by icmo.FCloseDate desc"
'		response.Write(sql)
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
		rs.pagesize = pages
		rs.absolutepage = page  
		for i=1 to rs.pagesize
		if rs.eof then exit for  
		if(i=1)then
		else
		  Response.Write ","
		end if
%>		
		{"id":"<%=i%>",
		"cell":["<%=rs("n1")%>","<%=rs("FCheckdate")%>","<%=rs("FCloseDate")%>","<%=rs("FMtono")%>","<%=rs("FNumber")%>","<%=rs("n2")%>","<%=rs("FBillNo")%>","<%=rs("FauxQty")%>","<%=rs("FAuxQtyFinish")%>","<%=rs("FAuxStockQty")%>","<%=rs("yfsl")%>","<%=rs("slsl")%>","<%=rs("blsl")%>","<%=rs("bfsl")%>","<%=rs("FinterID")%>"]}
<%		
	    rs.movenext
    next
  rs.close
  set rs=nothing
	response.Write"]}"

elseif showType="getInfo" then 
	sql="select a.FBillNo,n.FEntryID,c.fnumber,c.fname,d.Fname,n.FAuxQtyScrap,n.FBomInputAuxQty,n.FAuxStockQty,sum(e.Qty) as n1,sum(ReductQty) as n2,case when n.FAuxStockQty=0 then 0 else round(sum(ReductQty)*100/n.FAuxStockQty,2) end as n3,e.DealFlag,e.Dealdate,e.DealerID,e.Dealer"
	sql=sql&" from icmo a inner join "
	sql=sql&" PPBom m on a.FINterID=m.FIcmoInterID and a.Finterid="&request("InfoID")&" inner join "
	sql=sql&" PPBomEntry n on m.FInterID=n.FInterID inner join "
	sql=sql&" t_icitemcore c on n.Fitemid=c.Fitemid left join  "
	sql=sql&" t_measureUnit d on d.FMeasureUnitId=n.FUnitID left join  "
	sql=sql&" zxpt.dbo.manusys_WorkshopBadProd e on a.FBillNo=e.icmo_id and n.FEntryID=e.itemid group by a.FBillNo,n.FEntryID,c.fnumber,c.fname,d.Fname,n.FAuxQtyScrap,n.FBomInputAuxQty,n.FAuxStockQty,e.DealFlag,e.Dealdate,e.DealerID,e.Dealer"
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
	response.Write("{""row"":[")
	do until rs.eof
		response.Write("{""data"":[")
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
	response.Write("]}")
	rs.close
	set rs=nothing
elseif showType="getInfoDetails" then 
	sql="select SerialNum,icmo_id,itemid,MtrName,Unit,Qty,ReductQty,[Percent],Type,available,Section,sourceType,sourceReason,ResponsDepart,Dealsugg,Dealresult,Remark from manusys_WorkshopBadProd where icmo_id='"&request("InfoID")&"' and itemid='"&request("InfoID1")&"'"
'	response.Write(sql)
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connzxpt,0,1
	response.Write("{""row"":[")
	do until rs.eof
		response.Write("{""data"":[")
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
	response.Write("]}")
	rs.close
	set rs=nothing
elseif showType="getDepart" then 
	sql="select distinct a.Fitemid,a.FName from t_department a,icmo b where b.FworkShop=a.Fitemid"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connk3,1,1
	response.Write("[")
	do until rs.eof
		response.Write("{""CValue"":"""&rs("Fitemid")&""",""title"":"""&rs("FName")&"""}")
		rs.movenext
	If Not rs.eof Then
		Response.Write ","
	End If
	loop
	response.Write("]")
	rs.close
	set rs=nothing 
elseif showType="DataProcess" then 
	if request("submitType")="SaveOne" then
		for   i=2   to   Request.form("SerialNum").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("SerialNum")(i)<>"" then
				sql="select DealFlag from manusys_WorkshopBadProd where SerialNum="&Request.Form("SerialNum")(i)
				set rs=server.createobject("adodb.recordset")
				rs.open sql,connzxpt,1,1
				if rs("DealFlag")=1 then
				else
					connzxpt.Execute("Delete from manusys_WorkshopBadProd where SerialNum="&Request.Form("SerialNum")(i))
				end if
				rs.close
				set rs=nothing 
			elseif Request.Form("SerialNum")(i)<>"" then
				sql="select DealFlag from manusys_WorkshopBadProd where SerialNum="&Request.Form("SerialNum")(i)
				set rs=server.createobject("adodb.recordset")
				rs.open sql,connzxpt,1,1
				if rs("DealFlag")=1 then
				else
					connzxpt.Execute("update manusys_WorkshopBadProd set MtrName='"&Request.form("MtrName")(i)&"',Unit='"&Request.form("Unit")(i)&"',Qty='"&Request.form("Qty")(i)&"',ReductQty='"&Request.form("ReductQty")(i)&"',[Percent]='"&Request.form("Percent")(i)&"',Type='"&Request.form("MtrType")(i)&"',available='"&Request.form("available")(i)&"',sourceType='"&Request.form("sourceType")(i)&"',sourceReason='"&Request.form("sourceReason")(i)&"',ResponsDepart='"&Request.form("ResponsDepart")(i)&"',Dealsugg='"&Request.form("Dealsugg")(i)&"',Section='"&Request.form("Section")(i)&"',Remark='"&Request.form("Remark")(i)&"',Billdate='"&now()&"',BillerID='"&UserName&"',Biller='"&AdminName&"' where SerialNum="&Request.Form("SerialNum")(i))
				end if
				rs.close
				set rs=nothing 
			elseif Request.form("Qty")(i)<>"" then
					connzxpt.Execute("insert into manusys_WorkshopBadProd (icmo_id,itemid,MtrName,Unit,Qty,ReductQty,[Percent],Type,available,sourceType,sourceReason,ResponsDepart,Dealsugg,Remark,Billdate,BillerID,Biller,Section) values ('"&Request.form("FICMOID")&"','"&Request.form("Fitemid")&"','"&Request.form("MtrName")(i)&"','"&Request.form("Unit")(i)&"','"&Request.form("Qty")(i)&"','"&Request.form("ReductQty")(i)&"','"&Request.form("Percent")(i)&"','"&Request.form("MtrType")(i)&"','"&Request.form("available")(i)&"','"&Request.form("sourceType")(i)&"','"&Request.form("sourceReason")(i)&"','"&Request.form("ResponsDepart")(i)&"','"&Request.form("Dealsugg")(i)&"','"&Request.form("Remark")(i)&"','"&now()&"','"&UserName&"','"&AdminName&"','"&Request.form("Section")(i)&"')")
			end if
		next
		response.Write("保存操作成功！")
	elseif request("submitType")="DealOne" then
		response.Write(request.Form("MtrType"))
		if Instr(session("AdminPurview"),"|306.2,")=0 then
			response.Write("你没有权限进行此次操作！")
			response.End()
		end if
		connzxpt.Execute("update manusys_WorkshopBadProd set Dealdate='"&now()&"',DealerID='"&UserName&"',Dealer='"&AdminName&"',DealFlag='是' where icmo_id='"&request.Form("FICMOID")&"' and itemid ='"&request.Form("Fitemid")&"'")

		for   i=2   to   Request.form("SerialNum").count
			if Request.Form("SerialNum")(i)<>"" and Request.form("DeleteFlag")(i)="0" then
				connzxpt.Execute("update manusys_WorkshopBadProd set Dealresult='"&request.Form("Dealresult")(i)&"',Remark='"&request.Form("Remark")(i)&"' where SerialNum="&Request.Form("SerialNum")(i))
			end if
		next
		response.Write("处置操作成功！")
	end if
elseif showType="Excel" then 
	sql="select e.serialnum,o.FName as workshop,e.Billdate,a.FMtono,a.FauxQty,p.FName as product,c.Fnumber,c.fname as mtrname,d.fname as unitname,n.FAuxQtyScrap,n.FBomInputAuxQty,n.FAuxStockQty,e.*"
	sql=sql&" from icmo a inner join "
	sql=sql&" PPBom m on a.FINterid=m.FIcmoInterID inner join "
	sql=sql&" PPBomEntry n on m.FInterID=n.FInterID inner join "
	sql=sql&" zxpt.dbo.manusys_WorkshopBadProd e on a.FBillNo=e.icmo_id and n.FEntryID=e.itemid inner join "
	sql=sql&" t_department o on a.Fworkshop=o.Fitemid inner join  "
	sql=sql&" t_icitemcore c on n.Fitemid=c.Fitemid inner join  "
	sql=sql&" t_icitemcore p on a.Fitemid=p.Fitemid left join  "
	sql=sql&" t_measureUnit d on d.FMeasureUnitId=n.FUnitID where 1=1 "
	if request("FCJ")<>"" then sql=sql&" and a.Fworkshop="&request("FCJ")
	if request("SDate")<>"" then sql=sql&" and datediff(d,'"&request("SDate")&"',e.BillDate)>=0 "
	if request("EDate")<>"" then sql=sql&" and datediff(d,'"&request("EDate")&"',e.BillDate)<=0 "
	if request("FGD")<>"" then sql=sql&" and e.section='"&request("FGD")&"'"
	if request("DealF")<>"" then
		if request("DealF")="是" then
			sql=sql&" and e.DealFlag='"&request("DealF")&"' "
		else
			sql=sql&" and e.DealFlag is null "
		end if
	end if
	sql=sql&" order by e.serialnum desc"
	if request("Printtag")=1 then
	response.ContentType("application/vnd.ms-excel")
	response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
	end if
%>
<div id="listtable" style="width:100%; height:420; overflow:scroll">
<form id="allBad" name="allBad">
<table>
<tbody id="TbDetails">
<tr>    <td height="20" width="100%" class="tablemenu" colspan="32"><font color="#15428B"><img src="../images/close.jpg" border="0" align="absmiddle" onClick="$('#listtable').hide().css('z-index','550');$('#QueryTable').show();" >&nbsp;<strong>页面查看明细</strong></font></td>
</tr>
<tr bgcolor="#99BBE8">
<td colspan="32" align="left" ><input type="button" name="dealAll" id="dealAll" value="批量处理" class="button" onclick="return DealAll()"/></td>
</tr>
<tr bgcolor="#99BBE8">
<td nowrap="nowrap">勾选</td>
<td nowrap="nowrap" style="display:none" class="FICMOID">FICMOID</td>
<td nowrap="nowrap" style="display:none" class="Fitemid">ID</td>
<td nowrap="nowrap" >ID</td>
<td nowrap="nowrap" >生产车间</td>
<td nowrap="nowrap" >登记日期</td>
<td nowrap="nowrap" >计划跟踪号</td>
<td nowrap="nowrap" >生产任务数量</td>
<td nowrap="nowrap" >产品型号</td>
<td nowrap="nowrap" >物料编号</td>
<td nowrap="nowrap" >物料名称</td>
<td nowrap="nowrap" >还原单位</td>
<td nowrap="nowrap" >单位用量</td>
<td nowrap="nowrap" >BOM用料</td>
<td nowrap="nowrap" >订单用量</td>
<td nowrap="nowrap" class="MtrName">名称</td>
<td nowrap="nowrap" class="Unit">单位</td>
<td nowrap="nowrap" class="Qty">数量</td>
<td nowrap="nowrap" class="ReductQty">还原数量</td>
<td nowrap="nowrap" class="Percent">产生比例</td>
<td nowrap="nowrap" class="MtrType">物料类别</td>
<td nowrap="nowrap" class="available">是否可利用</td>
<td nowrap="nowrap" class="sourceType">来源项目</td>
<td nowrap="nowrap" class="sourceReason">来源说明</td>
<td nowrap="nowrap" class="ResponsDepart">责任单位</td>
<td nowrap="nowrap" class="ResponsDepart">产生工段</td>
<td nowrap="nowrap" class="Dealsugg">处置建议</td>
<td nowrap="nowrap" class="Dealdate">处置日期</td>
<td nowrap="nowrap" style="display:none" class="DealerID">处置人编号</td>
<td nowrap="nowrap" class="Dealer">处置人</td>
<td nowrap="nowrap" class="DealFlag">处置确认</td>
<td nowrap="nowrap" class="Dealresult">处置结果说明</td>
<td nowrap="nowrap" class="Remark">备注</td>
</tr>
<%
  set rs=server.createobject("adodb.recordset")
  rs.open sql,connk3,0,1
	while (not rs.eof)
%>
<tr bgcolor="#EBF2F9" ondblclick="return showEditDiv(this)">
<td nowrap="nowrap" ><input type="checkbox" name="SerialNum" value="<%=rs("serialnum")%>" /></td>
<td nowrap="nowrap" style="display:none" name="FICMOID"><%=rs("icmo_id")%></td>
<td nowrap="nowrap" style="display:none" name="Fitemid"><%=rs("itemid")%></td>
<td nowrap="nowrap" ><%=rs("serialnum")%></td>
<td nowrap="nowrap" ><%=rs("workshop")%></td>
<td nowrap="nowrap" ><%=rs("Billdate")%></td>
<td nowrap="nowrap" ><%=rs("FMtono")%></td>
<td nowrap="nowrap" name="AllQty"><%=rs("FauxQty")%></td>
<td nowrap="nowrap" ><%=rs("product")%></td>
<td nowrap="nowrap" name="Fnumber"><%=rs("Fnumber")%></td>
<td nowrap="nowrap" name="Fname"><%=rs("mtrname")%></td>
<td nowrap="nowrap" name="FUnit"><%=rs("unitname")%></td>
<td nowrap="nowrap" name="DWuse"><%=rs("FAuxQtyScrap")%></td>
<td nowrap="nowrap" name="BomUse"><%=rs("FBomInputAuxQty")%></td>
<td nowrap="nowrap" ><%=rs("FAuxStockQty")%></td>
<td nowrap="nowrap" name="MtrName"><%=rs("MtrName")%></td>
<td nowrap="nowrap" name="Unit"><%=rs("Unit")%></td>
<td nowrap="nowrap" name="Qty"><%=rs("Qty")%></td>
<td nowrap="nowrap" name="ReductQty"><%=rs("ReductQty")%></td>
<td nowrap="nowrap" name="Percent"><%=rs("Percent")%></td>
<td nowrap="nowrap" name="MtrType"><%=rs("MtrType")%></td>
<td nowrap="nowrap" name="available"><%=rs("available")%></td>
<td nowrap="nowrap" name="sourceType"><%=rs("sourceType")%></td>
<td nowrap="nowrap" name="sourceReason"><%=rs("sourceReason")%></td>
<td nowrap="nowrap" name="ResponsDepart"><%=rs("ResponsDepart")%></td>
<td nowrap="nowrap" name="ResponsDepart"><%=rs("Section")%></td>
<td nowrap="nowrap" name="Dealsugg"><%=rs("Dealsugg")%></td>
<td nowrap="nowrap" name="Dealdate"><%=rs("Dealdate")%></td>
<td nowrap="nowrap" style="display:none" name="DealerID"><%=rs("DealerID")%></td>
<td nowrap="nowrap" name="Dealer"><%=rs("Dealer")%></td>
<td nowrap="nowrap" name="DealFlag"><%=rs("DealFlag")%></td>
<td nowrap="nowrap" name="Dealresult"><%=rs("Dealresult")%></td>
<td nowrap="nowrap" name="Remark"><%=rs("Remark")%></td>
</tr>
<%
		rs.movenext
	wend
	
%>
</tbody>
</table>
</form>
</div>
<%
	rs.close
	set rs=nothing
elseif showType="DealAll" then 
	if Instr(session("AdminPurview"),"|306.2,")=0 then
		response.Write("你没有权限进行此次操作！")
		response.End()
	end if
	SerialNum=request("SerialNum")
  connzxpt.Execute("update manusys_WorkshopBadProd set Dealdate='"&now()&"',DealerID='"&UserName&"',Dealer='"&AdminName&"',DealFlag='是' where SerialNum in ("& SerialNum&")")
	response.Write("处理完成，请刷新查看！")
end if
 %>
