<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/Program.asp" -->
<%
 Dim  Rs,sSQL, v, f
 Set Rs = Server.CreateObject("ADODB.RecordSet")

 v = request.querystring("v")
 f = request.querystring("focus")
 
 sSQL = "SELECT * FROM PDIVISION WHERE SDIVCODE = '"&v&"'"
 Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText 
 IF NOT Rs.EOF  Then
	Rs.Close
	sSQL = "SELECT * FROM (SELECT SDIVCODE AS SUPPLYCODE , SDIV AS SUPPLY FROM PDIVISION WHERE SDIVCODE = '"&v&"'" _
			&" union " _
			&"SELECT SDEPCODE AS SUPPLYCODE , SDEPT AS SUPPLY FROM PDEPARTMENT WHERE SDEPT not like '%Cancel%' and SDEPCODE <> '00' and  SUBSTR(SDEPCODE,1,2) = SUBSTR('"&v&"',1,2) "_
			&" union " _
			&"SELECT SSECCODE AS SUPPLYCODE , SSEC AS SUPPLY FROM PSECTION WHERE SSEC not like '%Cancel%' and SSECCODE <> '00' and  SUBSTR(SSECCODE,1,2) = SUBSTR('"&v&"',1,2)) ORDER BY SUPPLY "
	 Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	 Response.Write "<select name='vorgsupply' >"
	      Response.Write "<option selected value= ''>---------</option>"& vbCrLf
	 Do While NOT Rs.EOF
		if Rs("SUPPLYCODE") = f then
			Response.Write "<option selected value= '"&Rs("SUPPLYCODE")& "'>"& Rs("SUPPLY") &"</option>"& vbCrLf
			
		else
			Response.Write "<option value= '"&Rs("SUPPLYCODE")& "'>"& Rs("SUPPLY") &"</option>"& vbCrLf
			
		end if
		Rs.MoveNext
	Loop
	Response.Write "</select>"
	Rs.Close
' sSQL = "SELECT  * FROM pexpenstype  ORDER BY sexptype "
' Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
 ELSE
	Rs.Close
	sSQL = "SELECT * FROM PDEPARTMENT WHERE SDEPCODE = '"&v&"'"
	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs.EOF  Then
		Rs.Close
		sSQL = "SELECT * FROM (SELECT SDEPCODE AS SUPPLYCODE , SDEPT AS SUPPLY FROM PDEPARTMENT WHERE SDEPCODE = '"&v&"' " _
			  &" union " _
			  &"SELECT SSECCODE AS SUPPLYCODE , SSEC AS SUPPLY FROM PSECTION WHERE SSEC not like '%Cancel%' and SSECCODE <> '00' and  SUBSTR(SSECCODE,1,4) = SUBSTR('"&v&"',1,4)) ORDER BY SUPPLY "
		Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
		Response.Write "<select name='vorgsupply' >"
		 Response.Write "<option selected value= ''>---------</option>"& vbCrLf
		Do While NOT Rs.EOF
				if Rs("SUPPLYCODE") = f then
			Response.Write "<option selected value= '"&Rs("SUPPLYCODE")& "'>"& Rs("SUPPLY") &"</option>"& vbCrLf
			
		else
			Response.Write "<option value= '"&Rs("SUPPLYCODE")& "'>"& Rs("SUPPLY") &"</option>"& vbCrLf
			
		end if
				Rs.MoveNext
		Loop
		Response.Write "</select>"
		Rs.Close
	ELSE
		Rs.Close
		sSQL = "SELECT * FROM (SELECT SSECCODE AS SUPPLYCODE , SSEC AS SUPPLY FROM PSECTION WHERE SSECCODE = '"&v&"' " _
			  &" union " _
			  &"SELECT SDEPCODE AS SUPPLYCODE , SDEPT AS SUPPLY FROM PDEPARTMENT WHERE SDEPT not like '%Cancel%' and SDEPCODE <> '00' and  SUBSTR(SDEPCODE,1,2) = '06' "_
			  &" union " _
			  &"SELECT SSECCODE AS SUPPLYCODE , SSEC AS SUPPLY FROM PSECTION WHERE SSEC not like '%Cancel%' and SSECCODE <> '00' and  SUBSTR(SSECCODE,1,2) = '06') ORDER BY SUPPLY "

		Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
		
		Response.Write "<select name='vorgsupply' >"
		 Response.Write "<option selected value= ''>---------</option>"& vbCrLf
		Do While NOT Rs.EOF
				if Rs("SUPPLYCODE") = f then
			Response.Write "<option selected value= '"&Rs("SUPPLYCODE")& "'>"& Rs("SUPPLY") &"</option>"& vbCrLf
			
		else
			Response.Write "<option value= '"&Rs("SUPPLYCODE")& "'>"& Rs("SUPPLY") &"</option>"& vbCrLf
			
		end if
				Rs.MoveNext
		Loop
		Response.Write "</select>"
		Rs.Close
	
	END IF
 END IF
 
 
 'response.write(Request.qu)
%>