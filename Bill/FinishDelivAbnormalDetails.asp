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
  dim Depart,wherestr,seachword,flag4search
  wherestr=""
  seachword=request("seachword")
  if seachword<>"" then
  start_date=dateadd("d",4,split(getDateRange(seachword,2012),"###")(0))
  end_date=dateadd("d",4,split(getDateRange(seachword,2012),"###")(1))
  wherestr=" and datediff(d,RegDate,'"&start_date&"')<=0 and datediff(d,RegDate,'"&end_date&"')>=0 "
  end if
  flag4search=request("flag4search")
  if flag4search="1" then 
    wherestr=wherestr&" and Replyer is null "
  elseif flag4search="2" then 
    wherestr=wherestr&" and Replyer is not null "
  end if
  dim page'页码
      page=clng(request("page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=request("limit")
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" Bill_FinishDelivAbnormal "
  dim datawhere'数据条件
		 datawhere="where 1=1"&wherestr&Session("AllMessage15")
		 Session("AllMessage15")=""
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by SerialNum desc"
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
	jsonstr="{'idcount':"&idcount&",'data':["
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    dim sql2,rs2,hejikoufen
	hejikoufen=0
	dim formdata(3),bgcolors
    sql="select *,left(AbnormalNote,10) as a1,left(ReplyText,10) as a2 from "& datafrom &" where SerialNum in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connzxpt,0,1
    while(not rs.eof)'填充数据到表格
  	  bgcolors="#EBF2F9"
		if Len(rs("ReplyText"))>0 then
		  bgcolors="#ff99ff"'粉色
		end if
		jsonstr=jsonstr&"{'SerialNum':"&rs("SerialNum")&","
		jsonstr=jsonstr&"'RegDate':'"&rs("RegDate")&"',"
		jsonstr=jsonstr&"'RegisterName':'"&rs("RegisterName")&"',"
		jsonstr=jsonstr&"'Departmentname':'"&rs("Departmentname")&"',"
		jsonstr=jsonstr&"'Register':'"&rs("Register")&"',"
		jsonstr=jsonstr&"'Department':'"&rs("Department")&"',"
		jsonstr=jsonstr&"'SEOutID':'"&rs("SEOutID")&"',"
		jsonstr=jsonstr&"'OrderID':'"&rs("OrderID")&"',"
		jsonstr=jsonstr&"'ProductId':'"&rs("ProductId")&"',"
		jsonstr=jsonstr&"'ProductName':'"&rs("ProductName")&"',"
		jsonstr=jsonstr&"'Model':'"&rs("Model")&"',"
		jsonstr=jsonstr&"'Unit':'"&rs("Unit")&"',"
		jsonstr=jsonstr&"'Quantity':'"&rs("Quantity")&"',"
		jsonstr=jsonstr&"'SendDate':'"&rs("SendDate")&"',"
		jsonstr=jsonstr&"'OutDate':'"&rs("OutDate")&"',"
		jsonstr=jsonstr&"'Replyer':'"&rs("Replyer")&"',"
		jsonstr=jsonstr&"'ReplyDate':'"&rs("ReplyDate")&"',"
		jsonstr=jsonstr&"'CheckFlag':'"&rs("CheckFlag")&"',"
		jsonstr=jsonstr&"'AbnormalNote':'"&replace(replace(rs("AbnormalNote"),chr(10),""),chr(13),"")&"',"
		jsonstr=jsonstr&"'ReplyText':'"&replace(replace(rs("ReplyText"),chr(10),""),chr(13),"")&"'},"
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
	jsonstr=Left(jsonstr,len(jsonstr)-1)&"]}"
  else
	jsonstr=jsonstr&"]}"
  end if
'-----------------------------------------------------------
  rs.close
  set rs=nothing
	response.Write(jsonstr)
elseif showType="AddEditShow" then 
  dim detailType
  detailType=request("detailType")
'数据处理
  dim SerialNum,Register,RegisterName,RegDate,Department,Departmentname,OrderID,Product,ReceivDepartment
  dim ProductType,CustomRanke,CustomLevel,OrderState,OrderQuantity,Agenter,AbnormalType,AbnormalNote
  dim CustomID,CheckFlag,style1,style2,style3,Replyer,ReplyText,ReplyDate
  dim IsLoss,LossAmount,OrderDate,CustomDate,MCReplyDate
  if detailType="AddNew" then
	sql="select a.*,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&UserName&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
	Register=UserName
	RegisterName=AdminName
	RegDate=date()
	Department=rs("部门别")
	Departmentname=rs("部门名称")
	style1="block;"
	style2="none;"
  elseif detailType="Edit" or detailType="Check" then
    SerialNum=request("SerialNum")
	sql="select * from Bill_FinishDelivAbnormal where SerialNum="&SerialNum
	set rs=server.createobject("adodb.recordset")
	rs.open sql,connzxpt,0,1
	Register=rs("Register")
	RegisterName=rs("RegisterName")
	RegDate=rs("RegDate")
	Department=rs("Department")
	Departmentname=rs("Departmentname")
	ReceivDepartment=rs("ReceivDepartment")
	OrderID=rs("OrderID")
	Product=rs("Product")
	ProductType=rs("ProductType")
	CustomRanke=rs("CustomRanke")
	CustomLevel=rs("CustomLevel")
	OrderState=rs("OrderState")
	OrderQuantity=rs("OrderQuantity")
	Agenter=rs("Agenter")
	AbnormalType=rs("AbnormalType")
	AbnormalNote=rs("AbnormalNote")
	CustomID=rs("CustomID")
	IsLoss=rs("IsLoss")
	LossAmount=rs("LossAmount")
	OrderDate=rs("OrderDate")
	CustomDate=rs("CustomDate")
	MCReplyDate=rs("MCReplyDate")
	Replyer=rs("Replyer")
	ReplyText=rs("ReplyText")
	ReplyDate=rs("ReplyDate")
  end if
    rs.close
    set rs=nothing
elseif showType="DataProcess" then 
  detailType=request("detailType")
  if detailType="AddNew" then
  	if  Instr(session("AdminPurviewFLW"),"|208.1,")>0 then
			set rs = server.createobject("adodb.recordset")
			sql="select * from Bill_FinishDelivAbnormal"
			rs.open sql,connzxpt,1,3
			rs.addnew
			rs("Biller")=UserName
			rs("BillDate")=now()
			rs("Register")=Request("Register")
			rs("RegisterName")=Request("RegisterName")
			rs("RegDate")=Request("RegDate")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("SEOutID")=Request("SEOutID")
			rs("OrderID")=Request("OrderID")
			rs("ProductId")=Request("ProductId")
			rs("ProductName")=Request("ProductName")
			rs("Model")=Request("Model")
			rs("Unit")=Request("Unit")
			rs("Quantity")=Request("Quantity")
			rs("SendDate")=Request("SendDate")
			rs("OutDate")=Request("OutDate")
			rs("AbnormalNote")=Request("AbnormalNote")
			rs.update
			rs.close
			set rs=nothing 
			response.write "{'success':true,'MSG':'对应单据添加成功！"&Request("SEOutID")&"'}"
		else
			response.write ("{'success':false,'MSG':'你没有权限进行此操作！'}")
			response.end
		end if
  elseif detailType="Edit"  then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_FinishDelivAbnormal where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("{'success':false,'MSG':'数据库读取记录出错！'}")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|208.1,")=0 then
		response.write ("{'success':false,'MSG':'你没有权限进行此操作！'}")
		response.end
  end if
			rs("Biller")=UserName
			rs("BillDate")=now()
			rs("Register")=Request("Register")
			rs("RegisterName")=Request("RegisterName")
			rs("RegDate")=Request("RegDate")
			rs("Department")=Request("Department")
			rs("Departmentname")=Request("Departmentname")
			rs("SEOutID")=Request("SEOutID")
			rs("OrderID")=Request("OrderID")
			rs("ProductId")=Request("ProductId")
			rs("ProductName")=Request("ProductName")
			rs("Model")=Request("Model")
			rs("Unit")=Request("Unit")
			rs("Quantity")=Request("Quantity")
			rs("SendDate")=Request("SendDate")
			rs("OutDate")=Request("OutDate")
			rs("AbnormalNote")=Request("AbnormalNote")
			response.write "{'success':true,'MSG':'对应单据修改成功！'}"
  rs.update
	rs.close
	set rs=nothing 
  elseif detailType="Delete" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_FinishDelivAbnormal where SerialNum="&SerialNum
	rs.open sql,connzxpt,0,1
	if rs.bof and rs.eof then
		response.write ("数据库读取记录出错！")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|208.1,")=0 then
		response.write ("你没有权限进行此操作！")
		response.end
	end if
	if rs("Biller")<>UserName and rs("Register")<>UserName then
		response.write ("只能删除本人自己添加的数据！")
		response.end
	end if
	if rs("CheckFlag")>0 then
		response.write ("已经有回复不允许删除！")
		response.end
	end if
	rs.close
	set rs=nothing 
	connzxpt.Execute("Delete from Bill_FinishDelivAbnormal where SerialNum="&SerialNum)
	response.write "删除成功！"
  elseif detailType="Check" then
	set rs = server.createobject("adodb.recordset")
    SerialNum=request("SerialNum")
	sql="select * from Bill_FinishDelivAbnormal where SerialNum="&SerialNum
	rs.open sql,connzxpt,1,3
	if rs.bof and rs.eof then
		response.write ("{'success':false,'MSG':'数据库读取记录出错！'}")
		response.end
	end if
	if Instr(session("AdminPurviewFLW"),"|208.2,")=0 then
		response.write ("{'success':false,'MSG':'你没有权限进行此操作！'}")
		response.end
  end if
    rs("Replyer")=session("AdminName")
    rs("Replydate")=now()
    rs("ReplyText")=Request("ReplyText")
    rs("CheckFlag")=1
	rs.update
	rs.close
	set rs=nothing 
	response.write "{'success':true,'MSG':'对应单据回复成功！'}"
  end if
elseif showType="getInfo" then 
  dim InfoID
  detailType=request("detailType")
  if detailType="OrderID" then
    InfoID=request("InfoID")
	sql="select a.FOrderBillNo,c.fname as a1,c.fnumber,c.fmodel,a.fqty,Min(e.FCheckDate) as SendDate,b.fheadselfs0247,f.fname as a2 from SEOutStockEntry a INNER JOIN "&_
	" SEOutStock b ON a.FBrNo = b.FBrNo AND a.FInterID = b.FInterID INNER JOIN "&_
	" t_ICItem c ON a.FItemID = c.FItemID left join "&_
	" ICStockBillEntry d on d.FSourceInterID=a.FInterID and d.FSourceEntryID=a.FEntryID  "&_
	" and d.FSourceTranType=b.FTranType left JOIN  "&_
	" ICStockBill e on e.FInterID = d.FInterID left join "&_
	" t_measureUnit f on a.Funitid=f.fitemid "&_
	" where b.fbillno='"&InfoID&"' "&_
	" group by a.FOrderBillNo,c.fname,c.fnumber,c.fmodel,a.fqty,b.fheadselfs0247,f.fname"
	set rs=server.createobject("adodb.recordset")
	jsonstr="{'idcount':1,'data':["
	rs.open sql,connk3,1,1
    if rs.bof and rs.eof then
	jsonstr=jsonstr&"]}"
	else
	while (not rs.eof)
		jsonstr=jsonstr&"{'OrderID':'"&rs("FOrderBillNo")&"',"
		jsonstr=jsonstr&"'ProductName':'"&rs("a1")&"',"
		jsonstr=jsonstr&"'ProductId':'"&rs("fnumber")&"',"
		jsonstr=jsonstr&"'Model':'"&rs("fmodel")&"',"
		jsonstr=jsonstr&"'Quantity':'"&rs("fqty")&"',"
		jsonstr=jsonstr&"'SendDate':'"&rs("SendDate")&"',"
		jsonstr=jsonstr&"'OutDate':'"&rs("fheadselfs0247")&"',"
		jsonstr=jsonstr&"'Unit':'"&rs("a2")&"'},"
	  rs.movenext
    wend
	  rs2.close
	  set rs2=nothing
	jsonstr=Left(jsonstr,len(jsonstr)-1)&"]}"
	end if
	response.write(jsonstr)
	rs.close
	set rs=nothing 
  elseif detailType="Register" then
    InfoID=request("InfoID")
	sql="select a.员工代号,a.姓名,a.部门别,b.部门名称 from [N-基本资料单头] a,[G-部门资料表] b where a.部门别=b.部门代号 and a.员工代号='"&InfoID&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,1
    if rs.bof and rs.eof then
        response.write ("员工编号不存在！")
        response.end
	else
		response.write(rs("员工代号")&"###"&rs("姓名")&"###"&rs("部门别")&"###"&rs("部门名称")&"###")
	end if
	rs.close
	set rs=nothing 
  end if
end if
 %>
