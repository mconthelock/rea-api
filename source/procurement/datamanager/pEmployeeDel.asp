<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,Page,gPage,MaxPage,keywords,sDivcode,sDepcode,sSeccode,ob,sPoscode,status
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
	keywords = UCase(TRIM( Request( "SearchUser" ) ))
	sDivcode=Request("ListDivision")
	sDepcode=Request("ListDepartment")
	sSeccode=Request("ListSection")
	sposcode=Request("ListPosition")
	ob=Request("ListOrderby")
	  sSQL = "DELETE FROM aauthorize WHERE sEmpno='" & Request.QueryString("sEmpno") &"'"
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  ' Create a recordset and open it for adding records

	  sSQL = "DELETE FROM aAuthorize WHERE sEmpno='" & Request.QueryString("sEmpno") &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL = "DELETE FROM aEmployee WHERE sEmpno='" & Request.QueryString("sEmpno") &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  Set Rs = Nothing
		sTatus="del"
	  Response.Redirect("pEmployee.asp?sTatus=" & status)

%>


        