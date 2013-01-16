<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurview"),"|309,")=0 then 
'  response.write "{""page"":""1"",""total"":""1"",""rows"":[{""id"":""0"",""cell"":["
'  response.write ("""<font color='red')>你不具有该管理模块的操作权限，请返回！</font>""]}]}")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag,UserName,AdminName,DepartName
UserName=session("UserName")
AdminName=session("AdminName")
showType=request("showType")
print_tag=request("print_tag")
Depart=session("Depart")
DepartName=session("DepartName")
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
      datafrom=" manusys_Inner a,manusys_InnerProduct b,manusys_InnerProductPrice c "
  dim datawhere'数据条件
  dim i'用于循环的整数
    datawhere=" where b.GongxuID=c.SerialNum and a.SerialNum=b.SNum "
	Dim searchterm,searchcols
	
	if Request.Form("query") <> "" then
	searchterm = Request.Form("query")
	searchcols = Request.Form("qtype")
	if isnumeric(searchterm) then
	datawhere = datawhere&" and " & searchcols & " = " & searchterm & " "
	else
	datawhere = datawhere&" and " & searchcols & " LIKE '%" & searchterm & "%' "
	end if
	End if
		 if request("sd")<>"" then datawhere=datawhere&" and datediff(d,a.XDDate,'"&request("sd")&"')<=0 "
		 if request("ed")<>"" then datawhere=datawhere&" and datediff(d,a.XDDate,'"&request("ed")&"')>=0 "
		 if request("wt")<>"" then datawhere=datawhere&" and a.Weituo='"&request("wt")&"' "
		 if request("jg")<>"" then datawhere=datawhere&" and a.jiagong='"&request("jg")&"' "
		 if request("sh")<>"" then datawhere=datawhere&" and a.CheckFlag='"&request("sh")&"' "
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
	Dim sortname
	if Request.Form("sortname") = "" then
	sortname = "a.SerialNum" 
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
    sql="select b.SerialNum from "& datafrom &" " & datawhere & taxis
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
    sql="select a.*,b.Dingdan,b.Danjia,b.Shuliang,b.Sunhao,b.Zhuanru,b.Jine,b.Fujian,c.Unit,c.Chengpin,c.Gongxu from "& datafrom &" where b.GongxuID=c.SerialNum and b.SNum=a.SerialNum and b.SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,1,1
		dim checktag
    do until rs.eof'填充数据到表格'
		checktag="未审"
		if rs("CheckFlag")="1" then checktag="品保审核"
		if rs("CheckFlag")="2" then checktag="委托审核"
		if rs("CheckFlag")="3" then checktag="加工审核"
		if rs("CheckFlag")="4" then checktag="转入确认"
%>		
		{"id":"<%=rs("SerialNum")%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("XDdate")%>","<%=rs("Weituo")%>","<%=rs("jiagong")%>","<%=rs("Dingdan")%>","<%=JsonStr(rs("Chengpin"))%>","<%=rs("Gongxu")%>","<%=rs("Unit")%>","<%=FormatNumber(rs("Danjia"),4,true)%>","<%=FormatNumber(rs("Shuliang"),2,true)%>","<%=FormatNumber(rs("Sunhao"),2,true)%>","<%=FormatNumber(rs("Zhuanru"),2,true)%>","<%=FormatNumber(rs("Jine"),2)%>","<%=Jsonstr(rs("Fujian"))%>","<%=checktag%>","<%=rs("Biller")%>","<%=rs("BillDate")%>","<%=rs("Checker")%>","<%=rs("Checker1")%>","<%=rs("Checker2")%>"]}
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
  if detailType="AddNew" then
  	if  Instr(session("AdminPurview"),"|309.1,")>0 then
			SerialNum=getBillNo("manusys_Inner",3,date())
			set rs = server.createobject("adodb.recordset")
			sql="select * from manusys_Inner"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("SerialNum")=SerialNum
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("Weituo")=Request("Weituo")
			rs("jiagong")=Request("jiagong")
			rs("XDDate")=Request("XDDate")
			rs("JHDate")=Request("JHDate")
			if request("Shengchanlx") then rs("Shengchanlx")=Request("Shengchanlx")
			if request("YP") then rs("YP")=Request("YP")
			if request("FG") then rs("FG")=Request("FG")
			if request("TZ") then rs("TZ")=Request("TZ")
			if request("CP") then rs("CP")=Request("CP")
			if request("YCL") then
				rs("YCL")=Request("YCL")
				rs("YuanCaiLiao")=Request("YuanCaiLiao")
			end if
			rs("Remark")=Request("Remark")
			rs("Gongyi")=Request("Gongyi")
			rs("Pingzhi")=Request("Pingzhi")
			dim nnnn:nnnn=0
			for   i=2   to   Request.form("DSerialNum").count
				if (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") and Request.form("GongxuID")(i)<>"" then
					connzxpt.Execute("insert into manusys_InnerProduct (SNum,Dingdan,GongxuID,Shuliang,Zhuanru,Danjia,Jine,Remark,Fujian) values ("&SerialNum&",'"&Request.form("Dingdan")(i)&"',"&Request.form("GongxuID")(i)&","&Request.form("Shuliang")(i)&","&Request.form("Zhuanru")(i)&","&Request.form("Danjia")(i)&","&Request.form("Jine")(i)&",'"&Request.form("DRemark")(i)&"','"&Request.form("Fujian")(i)&"')")
					nnnn=nnnn+1
				end if
			next
			if nnnn>0 then
			rs.update
			rs.close
			set rs=nothing 
			response.write "保存成功！"
			else
			response.write ("保存失败，明细不能为空！")
			response.end
			end if
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
		if Instr(session("AdminPurview"),"|309.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from manusys_Inner where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs("BillerID")<>UserName then
			response.write ("只能编辑自己添加的数据！")
			response.end
		end if
		if rs("CheckFlag")>0 then
			response.write ("当前状态不允许编辑，请检查！")
			response.end
		end if
			rs("SerialNum")=SerialNum
			rs("Biller")=AdminName
			rs("BillerID")=UserName
			rs("BillDate")=now()
			rs("Weituo")=Request("Weituo")
			if Request("jiagong")<>rs("jiagong") then
				rs("jiagong")=Request("jiagong")
				connzxpt.Execute("Delete from manusys_InnerProduct where SNum="&SerialNum)
			end if
			rs("XDDate")=Request("XDDate")
			rs("JHDate")=Request("JHDate")
			if request("Shengchanlx") then rs("Shengchanlx")=Request("Shengchanlx")
			if request("YP") then rs("YP")=Request("YP")
			if request("FG") then rs("FG")=Request("FG")
			if request("TZ") then rs("TZ")=Request("TZ")
			if request("CP") then rs("CP")=Request("CP")
			if request("YCL") then
				rs("YCL")=Request("YCL")
				rs("YuanCaiLiao")=Request("YuanCaiLiao")
			end if
			rs("Remark")=Request("Remark")
			rs("Gongyi")=Request("Gongyi")
			rs("Pingzhi")=Request("Pingzhi")
			nnnn=0
		for   i=2   to   Request.form("DSerialNum").count
			if Request.form("DeleteFlag")(i)="1" and Request.Form("DSerialNum")(i)<>"" then
				connzxpt.Execute("Delete from manusys_InnerProduct where SerialNum="&Request.Form("DSerialNum")(i))
			elseif Request.Form("DSerialNum")(i)<>"" then
				connzxpt.Execute("update manusys_InnerProduct set Dingdan='"&Request.form("Dingdan")(i)&"',GongxuID='"&Request.form("GongxuID")(i)&"',Shuliang='"&Request.form("Shuliang")(i)&"',Danjia='"&Request.form("Danjia")(i)&"',Jine='"&Request.form("Jine")(i)&"',Remark='"&Request.form("DRemark")(i)&"',Fujian='"&Request.form("Fujian")(i)&"',Zhuanru='"&Request.form("Zhuanru")(i)&"' where SerialNum="&Request.Form("DSerialNum")(i))
				nnnn=nnnn+1
			elseif (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") and Request.form("GongxuID")(i)<>"" then
					connzxpt.Execute("insert into manusys_InnerProduct (SNum,Dingdan,GongxuID,Shuliang,Zhuanru,Danjia,Jine,Remark,Fujian) values ("&SerialNum&",'"&Request.form("Dingdan")(i)&"',"&Request.form("GongxuID")(i)&","&Request.form("Shuliang")(i)&","&Request.form("Zhuanru")(i)&","&Request.form("Danjia")(i)&","&Request.form("Jine")(i)&",'"&Request.form("DRemark")(i)&"','"&Request.form("Fujian")(i)&"')")
					nnnn=nnnn+1
			end if
		next
			if nnnn>0 then
			rs.update
			rs.close
			set rs=nothing 
			response.write "保存成功！"
			else
			response.write ("保存失败，明细不能为空！")
			response.end
			end if
  elseif detailType="Product"  then
		if Instr(session("AdminPurview"),"|309.1,")=0 and Instr(session("AdminPurview"),"|309.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from manusys_Inner where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs("jiagong")<>DepartName then
			response.write ("只有加工方才能进行生产确认！")
			response.end
		end if
		if rs("CheckFlag")<>3 then
			response.write ("当前状态不允许进行生产确认，请检查！")
			response.end
		end if
		rs("JHDate")=Request("JHDate")
		rs("Confirm")=AdminName
		rs("ConfirmID")=UserName
		rs("ConfirmDate")=now()
		rs.update
		rs.close
		set rs=nothing 
		for   i=2   to   Request.form("DSerialNum").count
			if Request.Form("DSerialNum")(i)<>"" and (Request.form("DeleteFlag")(i)="0" or Request.form("DeleteFlag")(i)="") then
				connzxpt.Execute("update manusys_InnerProduct set Jine="&(cdbl(Request.form("Zhuanru")(i))*cdbl(Request.form("Danjia")(i)))&",Zhuanru="&Request.form("Zhuanru")(i)&",Sunhao="&Request.form("Sunhao")(i)&" where SerialNum="&Request.Form("DSerialNum")(i))
			end if
		next
		connzxpt.Execute("update manusys_Inner set ZongJine=aaa.c from (select sum(Jine) c,SNum from manusys_InnerProduct group by SNum) aaa where aaa.SNum=SerialNum and SerialNum="&SerialNum)
		response.write "生产确认成功！"
  elseif detailType="PBQueren"  then
		if Instr(session("AdminPurview"),"|309.3,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		set rs = server.createobject("adodb.recordset")
		SerialNum=request("SerialNum")
		sql="select * from manusys_Inner where SerialNum="&SerialNum
		rs.open sql,connzxpt,1,3
		if rs("CheckFlag")<>0 then
			response.write ("当前状态不允许进行品保审核，请检查！")
			response.end
		end if
		rs("Pingzhi")=Request("Pingzhi")
		rs("QcChecker")=AdminName
		rs("QcCheckerID")=UserName
		rs("QcCheckDate")=now()
		rs("CheckFlag")=1
		rs.update
		rs.close
		set rs=nothing 
		response.write "品保审核成功！"
  elseif detailType="Delete" then
		if Instr(session("AdminPurview"),"|309.1,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
    SerialNum=request("SerialNum")
		sql="select * from manusys_InnerProduct where CheckFlag>0 and BillerID='"&UserName&"' and SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.eof then
			connzxpt.Execute("Delete from manusys_Inner where SerialNum in ("&SerialNum&")")
			connzxpt.Execute("Delete from manusys_InnerProduct where SNum in ("&SerialNum&")")
			response.write "删除成功！"
		else
			response.Write("已审核不允许删除，或者你要删除不属于自己的单据，删除失败！")
			response.End()
		end if
  elseif detailType="Check"  then
		SerialNum=request("SerialNum")
		if Instr(session("AdminPurview"),"|309.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		sql="select * from manusys_Inner where SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,3
		if request("operattext")=1 then
			while (not rs.eof)
				if rs("Weituo")<>DepartName then
					response.Write("只能审核委托方是本部门的单据！")
					response.End()
				end if
				if rs("CheckFlag")=1 then
					rs("Checker")=AdminName
					rs("CheckerID")=UserName
					rs("CheckDate")=now()
					rs("CheckFlag")=2
				end if		
				rs.movenext
			wend
		elseif request("operattext")=0 then
			while (not rs.eof)
				if rs("Weituo")<>DepartName then
					response.Write("只能审核委托方是本部门的单据！")
					response.End()
				end if
				if rs("CheckFlag")=2 then
					rs("Checker")=null
					rs("CheckerID")=null
					rs("CheckDate")=null
					rs("CheckFlag")=1
				end if		
				rs.movenext
			wend
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "审核成功！"
  elseif detailType="Check1"  then
		SerialNum=request("SerialNum")
		if Instr(session("AdminPurview"),"|309.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		sql="select * from manusys_Inner where SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,3
		if request("operattext")=1 then
			while (not rs.eof)
				if rs("jiagong")<>DepartName then
					response.Write("只能审核加工方是本部门的单据！")
					response.End()
				end if
				if rs("Shengchanlx") and rs("CheckFlag")<3 then
					rs("Checker1")=AdminName
					rs("CheckerID1")=UserName
					rs("CheckDate1")=now()
					rs("CheckFlag")=3
				elseif rs("CheckFlag")=2 then
					rs("Checker1")=AdminName
					rs("CheckerID1")=UserName
					rs("CheckDate1")=now()
					rs("CheckFlag")=3
				end if		
				rs.movenext
			wend
		elseif request("operattext")=0 then
			while (not rs.eof)
				if rs("jiagong")<>DepartName then
					response.Write("只能审核加工方是本部门的单据！")
					response.End()
				end if
				if rs("CheckFlag")=3 and rs("CheckerID1")=UserName then
					rs("Checker1")=null
					rs("CheckerID1")=null
					rs("CheckDate1")=null
					if rs("Shengchanlx") then
						rs("CheckFlag")=0
					else
						rs("CheckFlag")=2
					end if
				end if		
				rs.movenext
			wend
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "审核成功！"
  elseif detailType="Check2"  then
		SerialNum=request("SerialNum")
		if Instr(session("AdminPurview"),"|309.2,")=0 then
			response.write ("你没有权限进行此操作！")
			response.end
		end if
		sql="select * from manusys_Inner where SerialNum in ("&SerialNum&")"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,3
		if request("operattext")=1 then
			while (not rs.eof)
				if rs("Weituo")<>DepartName then
					response.Write("只能确认委托方是本部门的单据！")
					response.End()
				end if
				if rs("CheckFlag")=3 then
					rs("Checker2")=AdminName
					rs("CheckerID2")=UserName
					rs("CheckDate2")=now()
					rs("CheckFlag")=4
				end if		
				rs.movenext
			wend
		elseif request("operattext")=0 then
			while (not rs.eof)
				if rs("Weituo")<>DepartName then
					response.Write("只能确认委托方是本部门的单据！")
					response.End()
				end if
				if rs("CheckFlag")=4 and rs("CheckerID2")=UserName then
					rs("Checker2")=null
					rs("CheckerID2")=null
					rs("CheckDate2")=null
					rs("CheckFlag")=3
				end if		
				rs.movenext
			wend
		end if
		rs.update
		rs.close
		set rs=nothing 
		response.write "确认成功！"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="SerialNum" then
    InfoID=request("InfoID")
		sql="select * from manusys_Inner where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
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
			sql="select a.SerialNum as DSerialNum,a.*,a.Remark as DRemark,b.Chengpin,b.Gongxu,b.Unit from manusys_InnerProduct a,manusys_InnerProductPrice b where a.gongxuID=b.Serialnum and a.SNum="&InfoID
			set rs=server.createobject("adodb.recordset")
			rs.open sql,connzxpt,1,1
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
	elseif detailType="Gongxu" then
	sql="select SerialNum,Chengpin,Gongxu,Unit,Zongji from manusys_InnerProductPrice where CheckFlag=1 and jiagong='"&Request("jiagong")&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
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
	end if
elseif showType="Export" then 
%>
 <table width="100%" border="1" cellpadding="3" cellspacing="1">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><strong>单号</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>下单日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>订单号</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>委托方</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>加工方</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>成品名称</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>工序名称</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>单位</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>单价</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>需求数量</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>损耗数量</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>转入数量</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>金额</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>审核</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>备注</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>登记人</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>登记日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>委托审核</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>审核日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>加工审核</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>审核日期</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>转入确认</strong></td>
    <td nowrap bgcolor="#8DB5E9"><strong>审核日期</strong></td>
  </tr>
 <%
    sql="select a.*,b.Remark as DRemark,b.Dingdan,b.Danjia,b.Shuliang,b.Sunhao,b.Zhuanru,b.Jine,c.Unit,c.Chengpin,c.Gongxu from manusys_Inner a,manusys_InnerProduct b,manusys_InnerProductPrice c where b.GongxuID=c.SerialNum and a.SerialNum=b.SNum "
		 if request("sd")<>"" then sql=sql&" and datediff(d,a.XDDate,'"&request("sd")&"')<=0 "
		 if request("ed")<>"" then sql=sql&" and datediff(d,a.XDDate,'"&request("ed")&"')>=0 "
		 if request("wt")<>"" then sql=sql&" and a.Weituo='"&request("wt")&"' "
		 if request("jg")<>"" then sql=sql&" and a.jiagong='"&request("jg")&"' "
		 if request("sh")<>"" then sql=sql&" and a.CheckFlag='"&request("sh")&"' "
		 sql=sql&" order by a.SerialNum desc "
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    while(not rs.eof)'填充数据到表格
		checktag="未审"
		if rs("CheckFlag")="1" then checktag="品保审核"
		if rs("CheckFlag")="2" then checktag="委托审核"
		if rs("CheckFlag")="3" then checktag="加工审核"
		if rs("CheckFlag")="4" then checktag="转入确认"
	  Response.Write "<tr>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Serialnum")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("XDDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Dingdan")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Weituo")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("jiagong")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Chengpin")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Gongxu")&"</td>"
      Response.Write "<td nowrap>"&rs("Unit")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Danjia")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Shuliang")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Sunhao")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Zhuanru")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Jine")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&checktag&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("DRemark")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Biller")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("BillDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Checker")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("CheckDate")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Checker1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("CheckDate1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("Checker2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("CheckDate2")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
	  rs.close
	  set rs=nothing
  %>
  </table>
<% 
elseif showType="ExportOne" then 
		response.ContentType("application/vnd.ms-word")
		response.AddHeader "Content-disposition", "attachment; filename=erpData.doc"
		sql="select * from manusys_Inner where SerialNum="&request.QueryString("SerialNum")
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		if rs.bof and rs.eof then
			response.write ("对应单据不存在，请检查！")
			response.end
		else
%>
<table width="100%" border="1" cellpadding="3" cellspacing="1">
  <tr>
    <td height="20" width="100%" align="center"><strong>内部加工明细</strong></font></td>
  </tr>
  <tr>
    <td height="20" width="100%">
	  <table width="100%" border="1" cellpadding="0" cellspacing="0" id=editNews>
      <tr>
        <td height="20" align="left" width="10%">单据号：</td>
        <td WIDTH='15%'><%=rs("SerialNum")%></td>
        <td height="20" align="left" width="10%">委托部门：</td>
        <td  WIDTH='15%'><%=rs("Weituo")%></td>
        <td height="20" align="left" width="10%">加工单位：</td>
        <td  WIDTH='15%'><%=rs("jiagong")%></td>
        <td height="20" width="10%" align="left">下单日期：</td>
        <td width="15%"><%=rs("XDDate")%></td>
      </tr>
      <tr>
     	  <td>交货日期：</td>
        <td><%=rs("JHDate")%></td>
        <td><input type="checkbox" name="Shengchanlx" id="Shengchanlx" value="1"  <%if rs("Shengchanlx") then response.Write("checked")%> /><label for="Shengchanlx">子制令</label>
        </td>
      <td>备注：</td>
      <td colspan="4"><%=rs("Remark")%></td>
      </tr>
      <tr>
      <td colspan="8">
      <fieldset><legend>提供物品项目及明细：</legend>
      <input type="checkbox" name="YP" id="YP" value="1"  <%if rs("YP") then response.Write("checked")%> /><label for="YP">提供样品</label>
      <input type="checkbox" name="FG" id="FG" value="1"  <%if rs("FG") then response.Write("checked")%> /><label for="FG">提供附稿</label>
      <input type="checkbox" name="TZ" id="TZ" value="1"  <%if rs("TZ") then response.Write("checked")%> /><label for="TZ">提供图纸</label>
      <input type="checkbox" name="CP" id="CP" value="1"  <%if rs("CP") then response.Write("checked")%> /><label for="CP">提供试做的产品</label>
      <input type="checkbox" name="YCL" id="YCL" value="1" <%if rs("YCL") then response.Write("checked")%> /><label for="YCL">提供原材料</label>
      <br />
      <%=rs("YuanCaiLiao")%>
		  </fieldset>
      </td>
      </tr>
      <tr>
        <td height="20" colspan="8">
		<table width="100%" border="1" id="editDetails" cellpadding="3" cellspacing="1">
		  <tbody id="TbDetails">
		  <tr height="24">
			<td width="6%"><strong>订单号</strong></td>
			<td width="8%"><strong>成品名称</strong></td>
			<td width="10%"><strong>工序名称</strong></td>
			<td width="4%"><strong>单位</strong></td>
			<td width="4%"><strong>单价</strong></td>
			<td width="4%"><strong>需求数量</strong></td>
			<td width="4%"><strong>损耗数量</strong></td>
			<td width="4%"><strong>转入数量</strong></td>
			<td width="6%"><strong>金额</strong></td>
			<td width="8%"><strong>备注</strong></td>
		  </tr>

<%
			sql1="select a.*,b.Chengpin,b.Gongxu,b.Unit from manusys_InnerProduct a,manusys_InnerProductPrice b where a.gongxuID=b.Serialnum and a.SNum="&request.QueryString("SerialNum")
			set rs2=server.createobject("adodb.recordset")
			rs2.open sql1,connzxpt,1,1
			while(not rs2.eof)
%>
		  <tr id="CloneNodeTr" >
			<td><%=rs2("Dingdan")%></td>
			<td><%=rs2("Chengpin")%></td>
			<td><%=rs2("Gongxu")%></td>
			<td><%=rs2("Unit")%></td>
			<td><%=rs2("Danjia")%></td>
			<td><%=rs2("Shuliang")%></td>
			<td><%=rs2("Sunhao")%></td>
			<td><%=rs2("Zhuanru")%></td>
			<td><%=rs2("Jine")%></td>
			<td><%=rs2("Remark")%></td>
		  </tr>

<%			
			rs2.movenext
			wend
		rs2.close
		set rs2=nothing 
%>
		  </tbody>
		</table>
    </td>
    </tr>
      <tr>
      <td colspan="8" style="height:50px;">
      <fieldset><legend>工艺要求：</legend>
		  <%=rs("Gongyi")%></fieldset>
      </td>
      </tr>
      <tr>
      <td colspan="8" style="height:50px;">
      <fieldset><legend>质量要求：</legend>
		  <%=rs("Pingzhi")%></fieldset>
      </td>
      </tr>
      <tr>
		<td colspan="8"><fieldset><legend>审核：</legend>
    委托制单/日期：<%=rs("Biller")%>&nbsp;
    <%=rs("Billdate")%>&nbsp;
    品保审核/日期：<%=rs("QcChecker")%>&nbsp;
    <%=rs("QcCheckDate")%>&nbsp;
    <br />
    委托审核/日期：<%=rs("Checker")%>&nbsp;
    <%=rs("CheckDate")%>&nbsp;
    加工审核/日期：<%=rs("Checker1")%>&nbsp;
    <%=rs("CheckDate1")%>&nbsp;
    <br />
    生产确认/日期：<%=rs("Confirm")%>&nbsp;
    <%=rs("ConfirmDate")%>&nbsp;
    转入确认/日期：<%=rs("Checker2")%>&nbsp;
    <%=rs("CheckDate2")%>&nbsp;
		  </fieldset>
		</td>
      </tr>
    </table>
		</td>
  </tr>
</table>
<%
		end if
		rs.close
		set rs=nothing 
end if
 %>
