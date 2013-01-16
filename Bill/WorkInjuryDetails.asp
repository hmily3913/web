<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<%
'if Instr(session("AdminPurviewFLW"),"|203,")=0 then 
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
      datafrom=" Bill_WorkInjury "
  dim datawhere'数据条件
	datawhere="where 1=1 "
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
	datawhere=datawhere&Session("AllMessage16")&Session("AllMessage17")&Session("AllMessage18")&Session("AllMessage19")&Session("AllMessage63")
	session.contents.remove "AllMessage16"
	session.contents.remove "AllMessage17"
	session.contents.remove "AllMessage18"
	session.contents.remove "AllMessage19"
	session.contents.remove "AllMessage63"
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
	dim ys:ys="#f7f7f7"
	if rs("CheckFlag")>0 then ys="#ffff66"
	dim CheckState:CheckState=""
	if rs("CheckFlag")="1" then
	  CheckState="主管审核"
	elseif rs("CheckFlag")="2" then
	  CheckState="人资确认"
	elseif rs("CheckFlag")="4" then
	  CheckState="财务确认"
	elseif rs("CheckFlag")="5" then
	  CheckState="已审批"
	elseif rs("CheckFlag")="3" then
	  ys="#ff99ff"
	  CheckState="已实施"
	else
	  CheckState="待审核"
	end if
	dim ClearState:ClearState="未结"
	if rs("ClearFlag")="1" then
	  ClearState="已结"
	  ys="#B3CFFC"
	end if
%>		
		{"id":"<%=rs("SerialNum")%>","ys":"<%=ys%>",
		"cell":["<%=rs("SerialNum")%>","<%=rs("RegisterName")%>","<%=rs("RegDate")%>","<%=rs("Departmentname")%>","<%=rs("Scene")%>","<%=rs("InDate")%>","<%=rs("Position")%>","<%=rs("InjuryType")%>","<%=rs("InsuranceFlag")%>","<%=rs("InsuranceId")%>","<%=CheckState%>","<%=replace(replace(rs("Details"),chr(10),""),chr(13),"")%>","<%=replace(replace(rs("Diagnosis"),chr(10),""),chr(13),"")%>","<%=rs("CauseType")%>","<%=replace(replace(rs("CauseAnalys"),chr(10),""),chr(13),"")%>","<%=replace(replace(rs("Measures"),chr(10),""),chr(13),"")%>","<%=rs("ProcessFlow")%>","<%=ClearState%>","<%=rs("ClearDate")%>"]}
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
  	if  Instr(session("AdminPurviewFLW"),"|209.1,")>0 then
			SerialNum=getBillNo("Bill_WorkInjury",3,date())
			set rs = server.createobject("adodb.recordset")
			sql="select * from Bill_WorkInjury"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("SerialNum")=SerialNum
			rs("Biller")=UserName
			rs("BillDate")=now()
			rs("Register")=Request("Register")
			rs("RegisterName")=Request("RegisterName")
			rs("RegDate")=Request("RegDate")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("Scene")=Request("Scene")
			rs("InDate")=Request("InDate")
			rs("Position")=Request("Position")
			rs("InjuryType")=Request("GJ01")&Request("GJ02")&Request("GJ03")&Request("GJ04")&Request("GJ05")&Request("GJ06")&Request("GJ07")&Request("GJ08")&Request("GJ09")&Request("GJ10")
			rs("InsuranceFlag")=Request("InsuranceFlag")
			rs("InsuranceId")=Request("InsuranceId")
			rs("Details")=Request("Details")
			rs("CauseType")=Request("CauseType")
			rs("CauseAnalys")=Request("CauseAnalys")
			rs("Measures")=Request("Measures")
			rs("Ages")=Request("Ages")
			rs("Gender")=Request("Gender")
			rs("MedicalHospital")=Request("MedicalHospital")
			rs("CardId")=Request("CardId")
			rs("Address")=Request("Address")
			rs("qiancheng1")=Request("qiancheng1")
			rs("qiancheng2")=Request("qiancheng2")
			rs("qiancheng3")=Request("qiancheng3")
			rs("RegTime")=Request("RegTime")
			rs.update
			rs.close
			set rs=nothing 
			response.write "###"
		else
			response.write ("你没有权限进行此操作！")
			response.end
		end if
  elseif detailType="Edit"  then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_WorkInjury where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|209.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
    end if
	if rs("Biller")<>UserName and rs("Register")<>UserName and Instr(session("AdminPurviewFLW"),"|209.5,")=0 then
		response.write ("只能编辑自己添加的数据！")
		response.end
	end if
	if rs("CheckFlag")>0 and Instr(session("AdminPurviewFLW"),"|209.5,")=0 then
		response.write ("当前状态不允许编辑，请检查！")
		response.end
	end if
'			rs("Biller")=UserName
			rs("BillDate")=now()
			rs("Register")=Request("Register")
			rs("RegisterName")=Request("RegisterName")
			rs("RegDate")=Request("RegDate")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("Scene")=Request("Scene")
			rs("InDate")=Request("InDate")
			rs("Position")=Request("Position")
			rs("InjuryType")=Request("GJ01")&Request("GJ02")&Request("GJ03")&Request("GJ04")&Request("GJ05")&Request("GJ06")&Request("GJ07")&Request("GJ08")&Request("GJ09")&Request("GJ10")
			rs("InsuranceFlag")=Request("InsuranceFlag")
			rs("InsuranceId")=Request("InsuranceId")
			rs("Details")=Request("Details")
			rs("CauseType")=Request("CauseType")
			rs("CauseAnalys")=Request("CauseAnalys")
			rs("Measures")=Request("Measures")
			rs("Ages")=Request("Ages")
			rs("Gender")=Request("Gender")
			rs("MedicalHospital")=Request("MedicalHospital")
			rs("CardId")=Request("CardId")
			rs("Address")=Request("Address")
			rs("qiancheng1")=Request("qiancheng1")
			rs("qiancheng2")=Request("qiancheng2")
			rs("qiancheng3")=Request("qiancheng3")
			rs("RegTime")=Request("RegTime")
			response.write "###"
    rs.update
	rs.close
	set rs=nothing 
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_WorkInjury where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|209.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
	end if
	if rs("Biller")<>UserName and rs("Register")<>UserName then
		response.write ("只能删除本人自己添加的数据！")
		response.end
	end if
	if rs("CheckFlag")>0 then
		response.write ("已经审核不允许删除！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_WorkInjury where SerialNum="&SerialNum)
	response.write "###"
  elseif detailType="Check" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_WorkInjury where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if rs("CheckFlag")=5 then
		response.write ("此单据已经在审批，不需要审核！")
		response.end
	end if
	if rs("CheckFlag")=0 and Instr(session("AdminPurviewFLW"),"|209.2,")>0 then
    rs("DepartReplyer")=session("AdminName")
    rs("DepartReplyDate")=now()
    rs("CheckFlag")=1
	elseif rs("CheckFlag")=1 and Instr(session("AdminPurviewFLW"),"|209.3,")>0 then
    rs("RelatedReplyer")=session("AdminName")
    rs("RelatedReplyDate")=now()
    rs("CheckFlag")=2
	elseif rs("CheckFlag")=3 and Instr(session("AdminPurviewFLW"),"|209.6,")>0 then
    rs("Cwer")=AdminName
    rs("CwerID")=UserName
    rs("CwDate")=now()
    rs("CheckFlag")=4
	elseif rs("CheckFlag")=4 and Instr(session("AdminPurviewFLW"),"|209.4,")>0 then
    rs("DirectorReplyer")=AdminName
    rs("DirectorReplyDate")=now()
    rs("CheckFlag")=5
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="unCheck" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_WorkInjury where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if rs("CheckFlag")=0 or rs("ClearFlag")=1 then
		response.write ("此单据未审核或已结，不允许反审核！")
		response.end
	end if
	if rs("CheckFlag")=1 and Instr(session("AdminPurviewFLW"),"|209.2,")>0 then
    rs("DepartReplyer")=session("AdminName")
    rs("DepartReplyDate")=now()
    rs("CheckFlag")=0
	elseif rs("CheckFlag")=2 and Instr(session("AdminPurviewFLW"),"|209.3,")>0 then
    rs("RelatedReplyer")=session("AdminName")
    rs("RelatedReplyDate")=now()
    rs("CheckFlag")=1
	elseif rs("CheckFlag")=4 and Instr(session("AdminPurviewFLW"),"|209.6,")>0 then
    rs("Cwer")=AdminName
    rs("CwerID")=UserName
    rs("CwDate")=now()
    rs("CheckFlag")=3
	elseif rs("CheckFlag")=5 and Instr(session("AdminPurviewFLW"),"|209.4,")>0 then
    rs("DirectorReplyer")=session("AdminName")
    rs("DirectorReplyDate")=now()
    rs("CheckFlag")=4
	else
		response.write ("你没有权限进行此操作！")
		response.end
	end if
	rs.update
	rs.close
	set rs=nothing 
	response.write "###"
  elseif detailType="Update"  then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_WorkInjury where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|209.5,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
  end if
'	if rs("CheckFlag")>3 or rs("ClearFlag")=1 then
'		response.write ("此单据不允许修改，请检查！")
'		response.end
'  end if
			rs("Diagnosis")=Request("Diagnosis")
			rs("MedicalFee")=Request("MedicalFee")
			rs("InsuranceSubsidyFee")=Request("InsuranceSubsidyFee")
			rs("CompanySubsidyFee")=Request("CompanySubsidyFee")
			rs("InjuryLeaveFee")=Request("InjuryLeaveFee")
			rs("OtherFee")=Request("OtherFee")
			rs("BorrowFee")=Request("BorrowFee")
			rs("ReimburseFee")=Request("ReimburseFee")
			rs("Balance")=Request("Balance")
			rs("Remark")=Request("Remark")
			rs("ProcessFlow")=Request("ProcessFlow")
			rs("MedicalFee2")=Request("MedicalFee2")
			rs("AppraisalFee")=Request("AppraisalFee")
			rs("InjuryLevel")=Request("InjuryLevel")
			if Request("ClearDate")<>"" then rs("ClearDate")=Request("ClearDate")
			rs("MedicalHospital")=Request("MedicalHospital")
			if Request("ProcessFlow")="财务审核中" and rs("CheckFlag")=2 then
				rs("ImplementReplyer")=session("AdminName")
				rs("ImplementReplyDate")=now()
				rs("CheckFlag")=3
			end if
			response.write "###"
    rs.update
	rs.close
	set rs=nothing 
  elseif detailType="Clear"  then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_WorkInjury where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|209.5,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
    end if
	if rs("CheckFlag")<2 or rs("ClearFlag")=1 then
		response.write ("当前状态不允许结清，请检查！")
		response.end
	end if
	if Request("ClearDate")="" then
		response.write ("结清日期不能为空，请检查！")
		response.end
	end if
			rs("ClearDate")=Request("ClearDate")
			rs("ClearFlag")=1
			response.write "###"
    rs.update
	rs.close
	set rs=nothing 
  elseif detailType="Borrow"  then
	set rs = server.createobject("adodb.recordset")
	dim FEntryID:FEntryID=request("FEntryID")
    SerialNum=request("Snum")
	sql="select * from Bill_WorkInjury where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|209.5,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
    end if
	if rs("ClearFlag")=1 or rs("CheckFlag")>3 then
		response.write ("当前状态不允许借款，请检查！")
		response.end
	end if
	if request("BorrowDate")="" or request("EmpName")="" or request("BorrowMoney")="" then
		response.write ("借款日期，借款人，借款金额不允许为空！")
		response.end
	end if
	if FEntryID="" then
	sql="insert into Bill_WorkInjuryDetails (SNum,EmpDepart,EmpName,BorrowDate,BorrowMoney) values ("&SerialNum&",'"&request("EmpDepart")&"','"&request("EmpName")&"','"&request("BorrowDate")&"',"&request("BorrowMoney")&")"
	connzxpt.Execute(sql)
	sql="Update Bill_WorkInjury set BorrowFee=AllBorrow,Balance=ReimburseFee-AllBorrow from (select sum(BorrowMoney) as AllBorrow,SNum from Bill_WorkInjuryDetails where SNum="&SerialNum&" Group By SNum) aaa where aaa.SNum=SerialNum"
	connzxpt.Execute(sql)
	set rs = server.createobject("adodb.recordset")
	sql="select top 1 FEntryID from Bill_WorkInjuryDetails where SNum="&SerialNum&" order by FEntryID desc"
	rs.open sql,connzxpt,1,1
	FEntryID=rs("FEntryID")
	else
	sql="update Bill_WorkInjuryDetails set SNum="&SerialNum&",EmpDepart='"&request("EmpDepart")&"',EmpName='"&request("EmpName")&"',BorrowDate='"&request("BorrowDate")&"',BorrowMoney="&request("BorrowMoney")&" where FEntryID="&FEntryID
	connzxpt.Execute(sql)
	sql="Update Bill_WorkInjury set BorrowFee=AllBorrow,Balance=ReimburseFee-AllBorrow from (select sum(BorrowMoney) as AllBorrow,SNum from Bill_WorkInjuryDetails where SNum="&SerialNum&" Group By SNum) aaa where aaa.SNum=SerialNum"
	connzxpt.Execute(sql)
	end if
	response.write "###"&FEntryID
	rs.close
	set rs=nothing 
  elseif detailType="DeleteBorrow"  then
	set rs = server.createobject("adodb.recordset")
	FEntryID=request("FEntryID")
    SerialNum=request("Snum")
	sql="select * from Bill_WorkInjury where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|209.5,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
    end if
	if rs("ClearFlag")=1 then
		response.write ("当前状态不允许借款，请检查！")
		response.end
	end if
	sql="Delete from Bill_WorkInjuryDetails where FEntryID="&FEntryID
	connzxpt.Execute(sql)
	sql="Update Bill_WorkInjury set BorrowFee=AllBorrow,Balance=ReimburseFee-AllBorrow from (select sum(BorrowMoney) as AllBorrow,SNum from Bill_WorkInjuryDetails where SNum="&SerialNum&" Group By SNum) aaa where aaa.SNum=SerialNum"
	connzxpt.Execute(sql)
	response.write "###"
	rs.close
	set rs=nothing 
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="Register" then
    InfoID=request("InfoID")
	sql="select a.*,b.部门名称,datediff(yy,a.出生日期,getdate()) as ages from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###"&rs("性别")&"###"&rs("工作岗位")&"###"&rs("到职日")&"###"&rs("保险类型")&"###"&rs("工伤保险号")&"###"&rs("社保号")&"###"&rs("ages")&"###"&rs("身份证号")&"###"&rs("户籍地址"))
	end if
	rs.close
	set rs=nothing 
  elseif detailType="SerialNum" then
    InfoID=request("InfoID")
	sql="select * from Bill_WorkInjury where SerialNum="&InfoID
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
  elseif detailType="ShowBorrow" then
    InfoID=request("InfoID")
	sql="select * from Bill_WorkInjuryDetails where SNum="&InfoID
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,1,1
	response.write "{""Info"":""###"",""Borrows"":["
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
	  response.write "]}"
	rs.close
	set rs=nothing 
  elseif detailType="doc" then
    InfoID=request("InfoID")
		sql="select * from Bill_WorkInjury where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.ContentType("application/vnd.ms-word")
		response.AddHeader "Content-disposition", "attachment; filename=erpData.doc"
	%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />

</HEAD>
<BODY>
&nbsp;
<p align="center">
	<span style="font-size:24px;">工&nbsp;伤&nbsp;事&nbsp;故&nbsp;报&nbsp;告</span>
</p>
<p>
	<span style="font-size:18px;">开发区人事劳动局：</span>
</p>
<p>
	<span style="font-size:18px;">&nbsp;&nbsp;&nbsp;&nbsp;兹有<%= rs("RegisterName") %>，<%= rs("Gender") %>，身份证<%= rs("CardId") %>，汉族，家住<%= rs("Address") %>，现系我公司<%= rs("Departmentname") %><%= rs("Position") %>。</span>
</p>
<p>
	<span style="font-size:18px;">&nbsp;&nbsp;&nbsp;&nbsp;<%= rs("Details") %></span>后速将其送至<%= rs("MedicalHospital") %>进行救治，经诊断为<%= rs("Diagnosis") %>。现伤情已稳定，就此事向上级主管部门提出工伤认定申请，望予受理，不胜感激！</span>
</p>
<p align="right">
	<span style="font-size:18px;">温州市蓝道工业发展有限公司</span>
</p>
<p align="right">
	<span style="font-size:18px;"><%= date() %></span>
</p>
</body>
</html>

	<%
	rs.close
	set rs=nothing 
  elseif detailType="xls" then
    InfoID=request("InfoID")
		sql="select * from Bill_WorkInjury where SerialNum="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		response.ContentType("application/vnd.ms-word")
		response.AddHeader "Content-disposition", "attachment; filename=erpData.doc"
	%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />

</HEAD>
<BODY>
&nbsp;
<p align="center">
	<span style="font-size:24px;">工&nbsp;伤&nbsp;事&nbsp;故&nbsp;报&nbsp;告</span>
</p>
	  <table width="100%" border="1" cellpadding="0" cellspacing="0" id="editNews" style="border:1px; font-size:12px">
      <tr>
        <td height="20px" align="left" width="60px">单据号：</td>
        <td width="80px"><%=rs("SerialNum")%></td>
        <td height="20" align="left" width="60px">伤者工号：</td>
        <td width="80px"><%=rs("Register")%></td>
        <td height="20px" align="left" width="60px">伤者姓名：</td>
        <td width="80px"><%=rs("RegisterName")%></td>
        <td height="20px" align="left" width="60px">性别：</td>
        <td width="80px"><%=rs("Gender")%></td>
      </tr>
      <tr>
        <td height="20" align="left">年龄：</td>
        <td><%=rs("Ages")%></td>
        <td height="20" align="left">部门编号：</td>
        <td><%=rs("Department")%></td>
        <td height="20" align="left">部门名称：</td>
        <td><%=rs("Departmentname")%></td>
        <td height="20" align="left">工种：</td>
        <td><%=rs("Position")%></td>
      </tr>
      <tr>
        <td height="20" align="left">入职时间：</td>
        <td><%=rs("InDate")%></td>
        <td height="20" align="left">发生时间：</td>
        <td><%=rs("RegDate")%></td>
        <td height="20" align="left">发生地点：</td>
        <td><%=rs("Scene")%></td>
        <td height="20" align="left">是否参保：</td>
        <td><%=rs("InsuranceFlag")%></td>
      </tr>
      <tr>
        <td height="20" align="left">保险编号：</td>
        <td><%=rs("InsuranceId")%></td>
        <td height="20" align="left" >事故类型：</td>
        <td colspan="5"><%=rs("InjuryType")%></td>
      </tr>
      <tr>
        <td height="80" align="left">事故发生详细情况：</td>
        <td colspan="7">
	  <%=rs("Details")%>
	  </td>
      </tr>
      <tr>
        <td height="60" align="left">原因分析：</td>
        <td><%=rs("CauseType")%></td>
        <td colspan="6">
	  <%=rs("CauseAnalys")%>
	  </td>
      </tr>
      <tr>
        <td height="60" align="left">改正及预防措施：</td>
        <td colspan="7">
	  <%=rs("Measures")%>
	  </td>
      </tr>
      <tr>
        <td height="20" align="left">责任部门：</td>
        <td><%=rs("DepartReplyer")%></td>
        <td height="20" align="left">审核时间：</td>
        <td><%=rs("DepartReplyDate")%></td>
        <td height="20" align="left">人资审核：</td>
        <td><%=rs("RelatedReplyer")%></td>
        <td height="20" align="left">审核时间：</td>
        <td><%=rs("RelatedReplyDate")%></td>
      </tr>
      <tr>
        <td height="20" align="left">副总审批：</td>
        <td><%=rs("DirectorReplyer")%></td>
        <td height="20" align="left">审批时间：</td>
        <td><%=rs("DirectorReplyDate")%></td>
        <td height="20" align="left">实施：</td>
        <td><%=rs("ImplementReplyer")%></td>
        <td height="20" align="left">实施时间：</td>
        <td><%=rs("ImplementReplyDate")%></td>
      </tr>
</table>
</body>
</html>

	<%
	rs.close
	set rs=nothing 
 	elseif detailType="qiancheng" then
		InfoID=request("InfoID")
		sql="select * from [N-签呈表] where 编号="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
%>
<div id="qianc" style="width:100%;height:100%;">
<div id="formove2" class="tablemenu"><font color="#15428B"><strong>签呈明细查看</strong></font></div>
<div style="background-color:#ddd; border:1px #666 solid;">
<div style="width:25%; float:left; background-color:#ddd;">签呈编号：<%=rs("编号")%></div><div style="width:25%; float:left; background-color:#ddd;">输入人员：<%=rs("员工姓名")%></div><div style="width:25%; float:left; background-color:#ddd;">签呈部门：<%=rs("部门名称")%></div><div style="width:25%; float:left; background-color:#ddd;">签呈日期：<%=rs("签呈日期")%></div>
<div style="width:50%; float:left; background-color:#ddd;">主题：<%=rs("主题")%></div><div style="width:25%; float:left;">性质：<%=rs("性质")%></div><div style="width:25%; float:left; background-color:#ddd;">签呈类别：<%=rs("类别")%></div>
<div style="width:100%; background-color:#ddd;">内容1：<%=rs("签呈内容")%></div>
<div style="width:100%; height:60px; background-color:#ddd;">内容2：<%=rs("签呈内容1")%></div>
</div>
<!-- tabs -->
<ul class="css-tabs">
	<li><a href="#">签呈明细</a></li>
	<li><a href="#">会签人员</a></li>
</ul>
<!-- panes -->
<div class="css-panes">

	<div>
  <table style="border:1px #666 solid;">
  <tr>
  <td>项次</td>
  <td>员工代号</td>
  <td>姓名</td>
  <td>奖惩项目</td>
  <td>奖点</td>
  <td>惩点</td>
  <td>事由</td>
  <td>奖点</td>
  <td>惩点</td>
  <td>工作岗位</td>
  <td>职等</td>
  </tr>
<%
		sql="select * from [N-签呈奖惩明细] where 编号="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		while (not rs.eof)
%>
  <tr>
  <td><%=rs("项次")%></td>
  <td><%=rs("员工代号")%></td>
  <td><%=rs("姓名")%></td>
  <td><%=rs("奖惩项目")%></td>
  <td><%=rs("奖点")%></td>
  <td><%=rs("惩点")%></td>
  <td><%=rs("事由")%></td>
  <td><%=rs("奖点")%></td>
  <td><%=rs("惩点")%></td>
  <td><%=rs("工作岗位")%></td>
  <td><%=rs("职等")%></td>
  </tr>
<%
			rs.movenext
		wend
%>
  </table>
	</div>
	
	<div>
  <table style="border:1px #666 solid;">
  <tr>
  <td>项次</td>
  <td>签合人员</td>
  <td>姓名</td>
  <td>顺序</td>
  <td>意见</td>
  <td>签合时间</td>
  </tr>
<%
		sql="select * from [N-签合表单身] where 编号="&InfoID
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		while (not rs.eof)
%>
  <tr>
  <td><%=rs("项次")%></td>
  <td><%=rs("员工代号")%></td>
  <td><%=rs("员工姓名")%></td>
  <td><%=rs("顺序")%></td>
  <td><%=rs("意见")%></td>
  <td><%=rs("时间")%></td>
  </tr>
<%
			rs.movenext
		wend
%>
  </table>
 	</div>
</div>
<div align="center"><input type="button" value="关闭" onClick="$(this).parent().parent().parent().hide();"/></div>
</div>
<%
		rs.close
		set rs=nothing 
  end if
end if
 %>
