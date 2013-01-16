<!--#include file="Include/ConnSiteData.asp" -->
<%
		response.Write("[")
		sql="select * from smmsys_Tree where PermissType='RPT' order by SerialNum asc"
		set rs=server.createobject("adodb.recordset")
		rs.open sql,connzxpt,1,1
		do until rs.eof
			dim showflag:showflag=false
			if Instr(session("AdminPurview"),rs("Permission"))>0 or rs("Permission")="" then
				showflag=true
	%>
	{"SerialNum": "<%=rs("SerialNum")%>", "TreeUrl": "<%=rs("TreeUrl")%>", "name": "<%=rs("TreeName")%>", "PSNum": "<%=rs("PSNum")%>"
	<%
	if instr(PermissionID,rs("PermissionID"))>0 then response.Write(",checked : true")
	
		set rs2=server.createobject("adodb.recordset")
		sql2="select count(1) as idcount from smmsys_Tree where PermissType='RPT' and PSNum="&rs("SerialNum")
		rs2.open sql2,connzxpt,1,1
		if rs2("idcount")>0 then
	%>
	, isParent:true
	<%
		end if
		response.Write("}")
		end if
			rs.movenext
			If Not rs.eof and showflag Then
			  Response.Write ","
			End If
		loop
		Response.Write "]"
		rs2.close
		set rs2=nothing 
%>