<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,RS2,RS3,sCatalog,pos1,pos2,sn
	  ' Create a recordset and open it for adding records
		sSQL="Select sCatID from pproduct where sCatID='" & Request.QueryString("sCatID") &"'"
	  Set Rs = Server.CreateObject("ADODB.RecordSet")	
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if RS.EOF then
	  sSQL = "SELECT * FROM pCatalog WHERE sCatID='" & Request.QueryString("sCatID") &"'"
	  Set Rs3 = Server.CreateObject("ADODB.RecordSet")	
	  Rs3.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  
	  'jeab pos1 = Instr(Rs3("sCatalog"), "(")
	  'jeab pos2 = Instr(Rs3("sCatalog"), ")")
	  'jeab sn = Mid(Rs3("sCatalog"),pos1,(pos2-pos1)+1)
	   sCatalog = Replace(Replace(Rs3("SCATACC"),"(",""),")","")
	  Rs3.Close
	  sSQL =  "SELECT * FROM EBUDGET.EBUDGETCNTLMST WHERE sCatalog='" & sCatalog &"'"
	  Rs3.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  if Not Rs3.EOF then
		  Rs3("CSTATUS") = "0"
		  Rs3("USER_UPDATE") = session("sEmpno")
		  Rs3("DATE_UPDATE") = Now()
		  Rs3.Update	
	  end if
	  Rs3.Close
	 ' sSQL =  "DELETE FROM EBUDGET.EBUDGETCNTLMST WHERE sCatalog='" & sCatalog &"'"
	  'Rs3.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  sSQL = "DELETE FROM pCatalog WHERE sCatID='" & Request.QueryString("sCatID") &"'"
	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")	
	  Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	 ' Rs2.Close
	 ' Rs3.Close
	  Set Rs2 = Nothing
	  Set Rs3 = Nothing
		if Request("status")="1" then
		  Response.Redirect "pCatalog.asp?sCattype=" & Request("sCattype")
		else			
		  Response.Redirect "pCatalogall.asp"
		end if
	else
		response.write "<br><center><font color=red size=+2>Can not Delete this Sub-Group in Expense Type use in Product !!!</font></br></center>"
	end if
	  Set Rs = Nothing


%>
