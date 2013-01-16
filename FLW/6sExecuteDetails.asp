<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<% Option Explicit %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HEAD>
<TITLE>欢迎进入系统后台</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
<META NAME="copyright" CONTENT="Copyright 2011-2012" />
<META NAME="Author" CONTENT="报表系统" />
<META NAME="Keywords" CONTENT="" />
<META NAME="Description" CONTENT="" />
<link rel="stylesheet" href="../Images/CssAdmin.css">

</HEAD>
<!--#include file="../CheckAdmin.asp" -->
<!--#include file="../Include/ConnSiteData.asp" -->
<BODY>
<%

'if Instr(session("AdminPurviewFLW"),"|103,")=0 then 
'  response.write ("<font color='red')>你不具有该管理模块的操作权限，请返回！</font>")
'  response.end
'end if
'========判断是否具有管理权限
dim showType,start_date,end_date,print_tag
showType=request("showType")
print_tag=request("print_tag")
if print_tag=1 then
response.ContentType("application/vnd.ms-excel")
response.AddHeader "Content-disposition", "attachment; filename=erpData.xls"
end if
if showType="DetailsList" then 
%>
 <div id="listtable" style="width:100%; height:100%; overflow:scroll">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检查单号</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检查日期</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>被查部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>被查单位</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检查区域</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>问题项目</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>类别</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>问题内容说明</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>扣分</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>检查人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>责任人</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>频率</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>改善对策</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>改善确认</strong></font></td>
    <td nowrap bgcolor="#8DB5E9"><font color="#FFFFFF"><strong>结案</strong></font></td>
  </tr>
 <%
  dim Depart,wherestr
  wherestr=""
  Depart=request("s6partment")
  start_date=request("start_date")
  end_date=request("end_date")
  if Depart="KD01.0001.0001" or Depart="12" then
	wherestr=" and (FComboBox1='人资部' or FComboBox1='总经办' or FComboBox1='管理部')"
  elseif Depart="KD01.0001.0002" then
	wherestr=" and FComboBox1='工程部'"
  elseif Depart="KD01.0001.0003" then
	wherestr=" and FComboBox1='采购部'"
  elseif Depart="KD01.0005.0004" then
	wherestr=" and FComboBox1='营销部'"
  elseif Depart="KD01.0001.0005" then
	wherestr=" and FComboBox1='生技部'"
  elseif Depart="KD01.0001.0006" then
	wherestr=" and FComboBox1='仓储科'"
  elseif Depart="KD01.0001.0007" then
	wherestr=" and FComboBox1='二分厂'"
  elseif Depart="KD01.0001.0008" then
	wherestr=" and FComboBox1='三分厂'"
  elseif Depart="KD01.0001.0009" then
	wherestr=" and FComboBox1='财务部'"
  elseif Depart="KD01.0001.0010" then
	wherestr=" and FComboBox1='一分厂'"
  elseif Depart="KD01.0001.0011" then
	wherestr=" and FComboBox1='品保部'"
  elseif Depart="KD01.0004.0001" or Depart="KD01.0004.0002" then
	wherestr=" and FComboBox1='娄桥办'"
  elseif Depart="KD01.0001.0017" then
	wherestr=" and FComboBox1='生管部'"
  elseif Len(Depart)=3 then
    wherestr=" and FComboBox1='"&Depart&"'"
  end if
  if start_date<>""  then wherestr=wherestr&" and b.FDate1>='"&start_date&"'"
  if end_date<>""  then	wherestr=wherestr&" and b.FDate1<='"&end_date&"'"

		  
  dim page'页码
      page=clng(request("Page"))
  dim idCount'记录总数
  dim pages'每页条数
      pages=20
  dim pagec'总页数

  dim datafrom'数据表名
      datafrom=" t_6s a,t_5sEntry b "
  dim datawhere'数据条件
		 datawhere="where a.fid=b.fid and a.fuser>0 and ftext9!='Y'"+wherestr&Session("AllMessage1")
		 Session("AllMessage1")=""
  dim sqlid'本页需要用到的id
  dim taxis'排序的语句 asc,desc
      taxis=" order by fentryid desc"
  dim i'用于循环的整数
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
    sql="select fentryid from "& datafrom &" " & datawhere & taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,1,1
    rs.pagesize = 20 '每页显示记录数
	rs.absolutepage = page  
    for i=1 to rs.pagesize
	  if rs.eof then exit for  
	  if(i=1)then
	    sqlid=rs("fentryid")
	  else
	    sqlid=sqlid &","&rs("fentryid")
	  end if
	  rs.movenext
    next
  '获取本页需要用到的id结束============================================
'-----------------------------------------------------------
'-----------------------------------------------------------
  if sqlid<>"" then'如果记录总数=0,则不处理
    '用in刷选本页所语言的数据,仅读取本页所需的数据,所以速度快
    dim sql2,rs2,hejikoufen
	hejikoufen=0
	dim formdata(3),bgcolors
    sql="select *,left(FText7,10) as a1 from "& datafrom &" where a.fid=b.fid and fentryid in("& sqlid &") "&taxis
    set rs=server.createobject("adodb.recordset")
    rs.open sql,connk3,0,1
    while(not rs.eof)'填充数据到表格
	hejikoufen=hejikoufen+rs("FInteger")
  	  bgcolors="#EBF2F9"
		if Len(rs("FText7"))>0 then
		  bgcolors="#ffff66"'黄色
		end if
		if instr(rs("FText8"),"Y")>0 then
		  bgcolors="#ff99ff"'粉色
		end if
	  Response.Write "<tr bgcolor='"&bgcolors&"' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FBillNo")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FDate1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FComboBox1")&"</td>"
      Response.Write "<td nowrap>"&rs("FText4")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText5")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText6")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FComboBox")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FInteger")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText1")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText2")&"</td>" & vbCrLf
      Response.Write "<td nowrap>"&rs("FText10")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return S6ClickTd(this,'T7reply',"&rs("fentryid")&")"">"&rs("a1")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return S6ClickTd(this,'T8reply',"&rs("fentryid")&")"">"&rs("FText8")&"</td>" & vbCrLf
      Response.Write "<td nowrap onDblClick=""return S6ClickTd(this,'T9reply',"&rs("fentryid")&")"">"&rs("FText9")&"</td>" & vbCrLf
      Response.Write "</tr>" & vbCrLf
	  rs.movenext
    wend
    Response.Write "<tr>" & vbCrLf
    Response.Write "<td colspan='8' nowrap  bgcolor='#EBF2F9'>合计扣分&nbsp;</td>" & vbCrLf
    Response.Write "<td colspan='7' nowrap  bgcolor='#EBF2F9'>"
	Response.Write (hejikoufen)
	Response.Write "</td>" & vbCrLf
    Response.Write "</tr>" & vbCrLf
	  rs2.close
	  set rs2=nothing
  else
    response.write "<tr><td height='50' align='center' colspan='15' nowrap  bgcolor='#EBF2F9'>暂无产品信息</td></tr>"
  end if
'-----------------------------------------------------------
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
	response.Write("###"&pagec&"###"&idcount&"###")
elseif showType="PingbiList" then 
dim pbmonth
pbmonth=request("pbmonth")
start_date=split(getDateRangebyMonth(pbmonth&"#"&request("pbyear")),"###")(0)
end_date=split(getDateRangebyMonth(pbmonth&"#"&request("pbyear")),"###")(1)
'数据处理
dim bumen(22),koufen(26),jifen(26),defen(26),fakuan(26),pbdefen(26),perfakuan
perfakuan=5'每份罚款金额
'设置基础分
jifen(0)=6
jifen(1)=6
jifen(2)=12
jifen(3)=12
jifen(4)=6
jifen(5)=16
jifen(6)=12
jifen(7)=12
jifen(8)=6
jifen(9)=11
jifen(10)=12
jifen(11)=12
jifen(12)=15
jifen(13)=12
jifen(14)=11
jifen(15)=6
jifen(16)=24
jifen(17)=5
jifen(18)=8
jifen(19)=5
jifen(20)=18
jifen(21)=8
jifen(22)=4
jifen(23)=48
jifen(24)=52
jifen(25)=41
jifen(26)=27
'设置部门
bumen(0)="一分厂展示柜"
bumen(1)="一分厂眼镜绳布"
bumen(2)="一分厂展示架"
bumen(3)="一分厂有机玻璃"
bumen(4)="一分厂木工房"
bumen(5)="三分厂下料段"
bumen(6)="三分厂针车段"
bumen(7)="三分厂打胶段"
bumen(8)="二分厂冲床"
bumen(9)="二分厂注塑段"
bumen(10)="二分厂包盒段"
bumen(11)="二分厂组装段"
bumen(12)="六分厂下料段"
bumen(13)="六分厂花生盒段"
bumen(14)="工程部"
bumen(15)="总经办"
bumen(16)="仓储科"
bumen(17)="采购部"
bumen(18)="品保部"
bumen(19)="财务部"
bumen(20)="管理部"
bumen(21)="营销部"
bumen(22)="生管部"
'jifen(22)=0
'获取扣分
sql="select isnull(sum(a0),0) as b0, "&_
"isnull(sum(a1),0) as b1, "&_
"isnull(sum(a2),0) as b2, "&_
"isnull(sum(a3),0) as b3, "&_
"isnull(sum(a4),0) as b4, "&_
"isnull(sum(a5),0) as b5, "&_
"isnull(sum(a6),0) as b6, "&_
"isnull(sum(a7),0) as b7, "&_
"isnull(sum(a8),0) as b8, "&_
"isnull(sum(a9),0) as b9, "&_
"isnull(sum(a10),0) as b10, "&_
"isnull(sum(a11),0) as b11, "&_
"isnull(sum(a12),0) as b12, "&_
"isnull(sum(a13),0) as b13, "&_
"isnull(sum(a14),0) as b14, "&_
"isnull(sum(a15),0) as b15, "&_
"isnull(sum(a16),0) as b16, "&_
"isnull(sum(a17),0) as b17, "&_
"isnull(sum(a18),0) as b18, "&_
"isnull(sum(a19),0) as b19, "&_
"isnull(sum(a20),0) as b20, "&_
"isnull(sum(a21),0) as b21, "&_
"isnull(sum(a22),0) as b22, "&_
"isnull(sum(a23),0) as b23, "&_
"isnull(sum(a24),0) as b24, "&_
"isnull(sum(a25),0) as b25, "&_
"isnull(sum(a26),0) as b26 "&_
"from ( "&_
"select  "&_
"case when FComboBox1='一分厂' and Ftext4 like '展示柜%' then b.FInteger else 0 end as a0,  "&_
"case when FComboBox1='一分厂' and Ftext4 like '眼镜绳布%' then b.FInteger else 0 end as a1,  "&_
"case when FComboBox1='一分厂' and Ftext4='展示架' then b.FInteger else 0 end as a2,  "&_
"case when FComboBox1='一分厂' and Ftext4='有机玻璃' then b.FInteger else 0 end as a3,  "&_
"case when FComboBox1='一分厂' and Ftext4='木工房' then b.FInteger else 0 end as a4,  "&_
"case when FComboBox1='三分厂' and Ftext4='下料段' then b.FInteger else 0 end as a5,  "&_
"case when FComboBox1='三分厂' and Ftext4='针车段' then b.FInteger else 0 end as a6,  "&_
"case when FComboBox1='三分厂' and Ftext4='打胶段' then b.FInteger else 0 end as a7,  "&_
"case when FComboBox1='二分厂' and Ftext4='冲床段' then b.FInteger else 0 end as a8,  "&_
"case when FComboBox1='二分厂' and Ftext4='注塑段' then b.FInteger else 0 end as a9,  "&_
"case when FComboBox1='二分厂' and Ftext4='包盒段' then b.FInteger else 0 end as a10,  "&_
"case when FComboBox1='二分厂' and Ftext4='组装段' then b.FInteger else 0 end as a11,  "&_
"case when FComboBox1='六分厂' and Ftext4='六厂下料段' then b.FInteger else 0 end as a12,  "&_
"case when FComboBox1='六分厂' and Ftext4='六厂花生盒段' then b.FInteger else 0 end as a13,  "&_
"case when FComboBox1='工程部' then b.FInteger else 0 end as a14,  "&_
"case when FComboBox1='生技部' or FComboBox1='总经办' then b.FInteger else 0 end as a15,  "&_
"case when FComboBox1='仓储科' then b.FInteger else 0 end as a16,  "&_
"case when FComboBox1='采购部' then b.FInteger else 0 end as a17,  "&_
"case when FComboBox1='品保部' then b.FInteger else 0 end as a18,  "&_
"case when FComboBox1='财务部' then b.FInteger else 0 end as a19,   "&_
"case when FComboBox1='人资部' or FComboBox1='管理部' then b.FInteger else 0 end as a20,   "&_
"case when FComboBox1='营销部' then b.FInteger else 0 end as a21,   "&_
"case when FComboBox1='生管部' then b.FInteger else 0 end as a22,   "&_
"case when FComboBox1='一分厂' then b.FInteger else 0 end as a23,   "&_
"case when FComboBox1='三分厂' then b.FInteger else 0 end as a24,   "&_
"case when FComboBox1='二分厂' then b.FInteger else 0 end as a25,   "&_
"case when FComboBox1='六分厂' then b.FInteger else 0 end as a26   "&_
"from t_6s a,t_5sEntry b   "&_
"where a.fid=b.fid and a.fuser>0 and b.FDate1>='"&start_date&"' and b.FDate1<='"&end_date&"') aaa"
set rs=server.createobject("adodb.recordset")
rs.open sql,connk3,0,1
if not rs.eof then
for i=0 to UBound(koufen)
koufen(i)=rs("b"&i)'扣分
defen(i)=jifen(i)-koufen(i)'得分=基分-扣分
pbdefen(i)=formatnumber(koufen(i)/jifen(i),2)'评比分=扣分/基分
if defen(i)*perfakuan>0 then
fakuan(i)=0
else
fakuan(i)=defen(i)*perfakuan'罚款金额
end if
next
end if
rs.close
set rs=nothing
'排序
dim gdmingci(2,13),gdmingcit(13),zjmingci(2,6),zjmingcit(6),jjmingci(2,5),jjmingcit(5)
dim zjbl,j,n'中间变量
'工段名次开始
for i=0 to UBound(gdmingci,2)
gdmingci(0,i)=pbdefen(i)
gdmingci(1,i)=i
next
for i=0 to UBound(gdmingci,2)
  for j=i to UBound(gdmingci,2)
    if gdmingci(0,i)>gdmingci(0,j) then
	  zjbl=gdmingci(0,i)
	  n=gdmingci(1,i)
	  gdmingci(0,i)=gdmingci(0,j)
	  gdmingci(0,j)=zjbl
	  gdmingci(1,i)=gdmingci(1,j)
	  gdmingci(1,j)=n
	end if
  next
next
for i=0 to UBound(gdmingci,2)
  if gdmingci(0,0)=gdmingci(0,i) then 
    gdmingci(2,i)=1
  elseif gdmingci(0,13)=gdmingci(0,i) then 
    gdmingci(2,i)=-1
  else
    gdmingci(2,i)=0
  end if
next

for i=0 to UBound(gdmingci,2)
  for j=0 to UBound(gdmingci,2)
    if i=gdmingci(1,j) and gdmingci(2,j)=1 then 
	  gdmingcit(i)="<font color='#0000FF'>工段第1名</font>"
	elseif i=gdmingci(1,j) and gdmingci(2,j)=-1 then 
	  gdmingcit(i)="<font color='#0000FF'>工段最后1名</font>"
	elseif i=gdmingci(1,j) then
	  gdmingcit(i)=""&(j+1)
	end if
  next
next
'间接单位名次开始
for i=0 to UBound(jjmingci,2)
jjmingci(1,i)=i
jjmingci(0,i)=pbdefen(17+i)
next
for i=0 to UBound(jjmingci,2)
  for j=i to UBound(jjmingci,2)
    if jjmingci(0,i)>jjmingci(0,j) then
	  zjbl=jjmingci(0,i)
	  n=jjmingci(1,i)
	  jjmingci(0,i)=jjmingci(0,j)
	  jjmingci(0,j)=zjbl
	  jjmingci(1,i)=jjmingci(1,j)
	  jjmingci(1,j)=n
	end if
  next
next

for i=0 to UBound(jjmingci,2)
  if jjmingci(0,0)=jjmingci(0,i) then 
    jjmingci(2,i)=1
  elseif jjmingci(0,5)=jjmingci(0,i) then 
    jjmingci(2,i)=-1
  else
    jjmingci(2,i)=0
  end if
next

for i=0 to UBound(jjmingci,2)
  for j=0 to UBound(jjmingci,2)
    if i=jjmingci(1,j) and jjmingci(2,j)=1 then 
	  jjmingcit(i)="<font color='#0000FF'>间接单位第1名</font>"
	elseif i=jjmingci(1,j) and jjmingci(2,j)=-1 then 
	  jjmingcit(i)="<font color='#0000FF'>间接单位最后1名</font>"
	elseif i=jjmingci(1,j) then
	  jjmingcit(i)=""&(j+1)
	end if
'	if jjmingcit(i)="1" then jjmingcit(i)="<font color='#0000FF'>间接单位第1名</font>"
'	if jjmingcit(i)="6" then jjmingcit(i)="<font color='#0000FF'>间接单位最后1名</font>"
  next
next
'直接单位名次开始
for i=0 to UBound(zjmingci,2)
zjmingci(1,i)=i
next
zjmingci(0,0)=pbdefen(23)
zjmingci(0,1)=pbdefen(24)
zjmingci(0,2)=pbdefen(25)
zjmingci(0,3)=pbdefen(26)
zjmingci(0,4)=pbdefen(14)
zjmingci(0,5)=pbdefen(15)
zjmingci(0,6)=pbdefen(16)
for i=0 to UBound(zjmingci,2)
  for j=i to UBound(zjmingci,2)
    if zjmingci(0,i)>zjmingci(0,j) then
	  zjbl=zjmingci(0,i)
	  n=zjmingci(1,i)
	  zjmingci(0,i)=zjmingci(0,j)
	  zjmingci(0,j)=zjbl
	  zjmingci(1,i)=zjmingci(1,j)
	  zjmingci(1,j)=n
	end if
  next
next
for i=0 to UBound(zjmingci,2)
  if zjmingci(0,0)=zjmingci(0,i) then 
    zjmingci(2,i)=1
  elseif zjmingci(0,6)=zjmingci(0,i) then 
    zjmingci(2,i)=-1
  else
    zjmingci(2,i)=0
  end if
next

for i=0 to UBound(zjmingci,2)
  for j=0 to UBound(zjmingci,2)
    if i=zjmingci(1,j) and zjmingci(2,j)=1 then 
	  zjmingcit(i)="<font color='#0000FF'>直接单位第1名</font>"
	elseif i=zjmingci(1,j) and zjmingci(2,j)=-1 then 
	  zjmingcit(i)="<font color='#0000FF'>直接单位最后1名</font>"
	elseif i=zjmingci(1,j) then
	  zjmingcit(i)=""&(j+1)
	end if
  next
next

%>
 <div id="pingbibiao" style="width:100%; height:100%; overflow:auto; ">
 <table width="100%" border="0" cellpadding="3" cellspacing="1" bgcolor="#99BBE8">
  <tr>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>部门</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>合计</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center" width="40"><font color="#FFFFFF"><strong>扣分合计</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>基分</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center" width="40"><font color="#FFFFFF"><strong>基分合计</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>得分</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center" width="55"><font color="#FFFFFF"><strong>每分罚款金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center" width="40"><font color="#FFFFFF"><strong>罚款金额</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center" width="50"><font color="#FFFFFF"><strong>工段评比得分</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>工段名次</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center" width="50"><font color="#FFFFFF"><strong>部门评比得分</strong></font></td>
    <td nowrap bgcolor="#8DB5E9" align="center"><font color="#FFFFFF"><strong>部门名次</strong></font></td>
  </tr>
 <%
 bgcolors="#EBF2F9"
 dim tempbumen,tempdwmingci,jjbg
 tempbumen="#@!@#@%$@"
for i=0 to UBound(bumen)
if i>16 then
  tempdwmingci=jjmingcit(i-17)
  bgcolors="#FF6901"
else
  tempdwmingci=zjmingcit(i-10)
end if

  Response.Write "<tr bgcolor='"&bgcolors&"' align='right' onMouseOver = ""this.style.backgroundColor = '#FFFFFF'"" onMouseOut = ""this.style.backgroundColor = ''"" style='cursor:hand'>" & vbCrLf
  Response.Write "<td nowrap>"&bumen(i)&"</td>" & vbCrLf
  Response.Write "<td nowrap><a href=""javascript:ShowDetails('"&left(bumen(i),3)&"','"&start_date&"','"&end_date&"')"">"&koufen(i)&"</a></td>" & vbCrLf
  
  if Instr(bumen(i),"一分厂")>0 and tempbumen<>"一分厂" then
  Response.Write "<td nowrap rowspan='5'><a href=""javascript:ShowDetails('"&left(bumen(i),3)&"','"&start_date&"','"&end_date&"')"">"&koufen(23)&"</a></td>"
  Response.Write "<td nowrap>"&jifen(i)&"</td>"
  Response.Write "<td nowrap rowspan='5'>"&jifen(23)&"</td>"
  Response.Write "<td nowrap>"&defen(i)&"</td>"
  Response.Write "<td nowrap>"&perfakuan&"</td>"
  Response.Write "<td nowrap>"&fakuan(i)&"</td>"
  Response.Write "<td nowrap>"&pbdefen(i)&"</td>"
  Response.Write "<td nowrap>"&gdmingcit(i)&"</td>"
  Response.Write "<td nowrap rowspan='5'>"&pbdefen(23)&"</td>"
  Response.Write "<td nowrap rowspan='5'>"&zjmingcit(0)&"</td>"
  tempbumen="一分厂"
  elseif Instr(bumen(i),"三分厂")>0 and tempbumen<>"三分厂" then
  Response.Write "<td nowrap rowspan='3'><a href=""javascript:ShowDetails('"&left(bumen(i),3)&"','"&start_date&"','"&end_date&"')"">"&koufen(24)&"</a></td>"
  Response.Write "<td nowrap>"&jifen(i)&"</td>"
  Response.Write "<td nowrap rowspan='3'>"&jifen(24)&"</td>"
  Response.Write "<td nowrap>"&defen(i)&"</td>"
  Response.Write "<td nowrap>"&perfakuan&"</td>"
  Response.Write "<td nowrap>"&fakuan(i)&"</td>"
  Response.Write "<td nowrap>"&pbdefen(i)&"</td>"
  Response.Write "<td nowrap>"&gdmingcit(i)&"</td>"
  Response.Write "<td nowrap rowspan='3'>"&pbdefen(24)&"</td>"
  Response.Write "<td nowrap rowspan='3'>"&zjmingcit(1)&"</td>"
  tempbumen="三分厂"
  elseif Instr(bumen(i),"二分厂")>0 and tempbumen<>"二分厂" then
  Response.Write "<td nowrap rowspan='4'><a href=""javascript:ShowDetails('"&left(bumen(i),3)&"','"&start_date&"','"&end_date&"')"">"&koufen(25)&"</a></td>"
  Response.Write "<td nowrap>"&jifen(i)&"</td>"
  Response.Write "<td nowrap rowspan='4'>"&jifen(25)&"</td>"
  Response.Write "<td nowrap>"&defen(i)&"</td>"
  Response.Write "<td nowrap>"&perfakuan&"</td>"
  Response.Write "<td nowrap>"&fakuan(i)&"</td>"
  Response.Write "<td nowrap>"&pbdefen(i)&"</td>"
  Response.Write "<td nowrap>"&gdmingcit(i)&"</td>"
  Response.Write "<td nowrap rowspan='4'>"&pbdefen(25)&"</td>"
  Response.Write "<td nowrap rowspan='4'>"&zjmingcit(2)&"</td>"
  tempbumen="二分厂"
  elseif Instr(bumen(i),"六分厂")>0 and tempbumen<>"六分厂" then
  Response.Write "<td nowrap rowspan='2'><a href=""javascript:ShowDetails('"&left(bumen(i),3)&"','"&start_date&"','"&end_date&"')"">"&koufen(26)&"</a></td>"
  Response.Write "<td nowrap>"&jifen(i)&"</td>"
  Response.Write "<td nowrap rowspan='2'>"&jifen(26)&"</td>"
  Response.Write "<td nowrap>"&defen(i)&"</td>"
  Response.Write "<td nowrap>"&perfakuan&"</td>"
  Response.Write "<td nowrap>"&fakuan(i)&"</td>"
  Response.Write "<td nowrap>"&pbdefen(i)&"</td>"
  Response.Write "<td nowrap>"&gdmingcit(i)&"</td>"
  Response.Write "<td nowrap rowspan='2'>"&pbdefen(26)&"</td>"
  Response.Write "<td nowrap rowspan='2'>"&zjmingcit(3)&"</td>"
  tempbumen="六分厂"
  elseif Instr(bumen(i),"分厂")>0 then
  Response.Write "<td nowrap>"&jifen(i)&"</td>"
  Response.Write "<td nowrap>"&defen(i)&"</td>"
  Response.Write "<td nowrap>"&perfakuan&"</td>"
  Response.Write "<td nowrap>"&fakuan(i)&"</td>"
  Response.Write "<td nowrap>"&pbdefen(i)&"</td>"
  Response.Write "<td nowrap>"&gdmingcit(i)&"</td>"
  else
  Response.Write "<td nowrap><a href=""javascript:ShowDetails('"&left(bumen(i),3)&"','"&start_date&"','"&end_date&"')"">"&koufen(i)&"</a></td>" & vbCrLf
  Response.Write "<td nowrap>"&jifen(i)&"</td>"
  Response.Write "<td nowrap>"&jifen(i)&"</td>" & vbCrLf
  Response.Write "<td nowrap>"&defen(i)&"</td>"
  Response.Write "<td nowrap>"&perfakuan&"</td>"
  Response.Write "<td nowrap>"&fakuan(i)&"</td>"
  Response.Write "<td nowrap>"&pbdefen(i)&"</td>"
  Response.Write "<td nowrap></td>"
  Response.Write "<td nowrap>"&pbdefen(i)&"</td>"
  Response.Write "<td nowrap>"&tempdwmingci&"</td>"
  end if
  
  Response.Write "</tr>" & vbCrLf
next
  rs.close
  set rs=nothing
  %>
  </table>
  </div>
<% 
end if
 %>
</body>
</html>
