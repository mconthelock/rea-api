<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: worakrit yodthonglert
	' Version: 1.01
%>
<%
	Dim sExptype, sCatid, sCatType
	Dim i
	Dim  Rs,sSQL
	Set Rs = Server.CreateObject("ADODB.RecordSet")
%>
<HTML>
<HEAD>
<TITLE>Summary By Total</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		

  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top" width="1%" height="2"><img src="../images/ccb.gif" width="10" height="10"></td>
      <td width="99%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="8" height="8"></td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="99%" height="30" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="+1" color="#666666"><b><b><b><b><b><b><b><font color="#BF00BF">Report 
        Manager</font></b></b></b></b></b></b></b></font></td>
    </tr>
    <tr> 
      <td width="1%"><img src="../images/spacer.gif" width="1" height="1"></td>
      <td width="99%" align="center" valign="top" height="0"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="98%" height="2" align="left" valign="top" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../images/y01.jpg" width="12" height="12"></b></font></td>
          </tr>
          <tr> 
            <td width="98%" align="left" valign="top" height="0"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr align="left" valign="top" bgcolor="#FFD553"> 
                  <td colspan="2" align="center">&nbsp;</td>
                </tr>
                <tr align="center" valign="top"> 
                  <td colspan="2" bgcolor="#FFD553"> 
                    
                  <table border=0 width=437 align="Center" cellspacing="1" cellpadding="0">
                    <tr bgcolor="#CC00CC"> 
                      <td colspan="2" height="20"> 
                        <div align="center"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Matching Report</font></font></b></div>
                      </td>
                    </tr>
                    <form Action="finRecReport.asp" method="post">
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" nowrap> 
                          <div align="right">Expense Type &nbsp;</div>
                        </td>
                        <td width="69%" height="10"> 
							<%
							sSQL = "SELECT  * FROM pexpenstype  ORDER BY sexptype "
						 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
							'response.write Rs.EoF
							%>
							                          <select name="sexptype" size=1>
                            <option value="0" checked>All</option>

							<%
							Do While Not RS.EOF	
							%> 
                            <option value="<%=RS("sexptype")%>"><%=RS("sexptype")%>-<%=RS("sexpdesc")%></option>
                            <%
								RS.MoveNext
							Loop
							Rs.Close
							%>
							
                            <%' List_Pexptype(sexptype) %> 
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" nowrap> 
                          <div align="right">Group in Exp: &nbsp;</div>
                        </td>
                        <td width="69%" height="10"> 
                          <select name="scattype" size=1>
                            <option value="0" checked>All</option>
                            <% Show_Pcattype(sCatType) %> 
                          </select>
                        </td>
                      </tr>
					<tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" nowrap> 
                          <div align="right">Sub-Group in Exp: &nbsp;</div>
                        </td>
                        <td width="69%" height="10"> 
                          <select name="scatid" size=1>
                            <option value="0" checked>All</option>
                            <% Show_Pcatalog(sCatid) %> 
                          </select>
                        </td>
                      </tr>					  					  
                      <tr bgcolor="#EEEEEE"> 
                      	<td width="31%" height="10"> 
                          <div align="right">PR No. &nbsp;</div>
                        </td>
                        <td width="69%" height="10"> 
                          <input type="text" name="txtPrNo" size="20" maxlength="12">
                        </td>	
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                      	<td width="31%" height="10"> 
                          <div align="right">PO No. &nbsp;</div>
                        </td>
                        <td width="69%" height="10"> 
                          <input type="text" name="txtPoNo" size="20" maxlength="12">
                        </td>	
                      </tr>					  
                      <tr bgcolor="#EEEEEE"> 
                      	<td width="31%" height="10"> 
                          <div align="right">INV. No. &nbsp;</div>
                        </td>
                        <td width="69%" height="10"> 
                          <input type="text" name="txtInvNo" size="20" maxlength="12">
                        </td>	
                      </tr>					  		      
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" bgcolor="#EEEEEE" align="right">Division 
                          :</td>
                        <td width="69%" height="10"> 
                          <select name="ListDivision">
                            <option value="0" selected>&lt;-- Search List By Division--&gt;</option>
                            <%List_Division()%> 
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" bgcolor="#EEEEEE" align="right">Department 
                          :</td>
                        <td width="69%" height="10"> 
                          <select name="ListDepartment">
                            <option value="0" selected>&lt;-- Search List By Department--&gt;</option>
                            <%List_Department()%> 
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" bgcolor="#EEEEEE" align="right">Section 
                          : </td>
                        <td width="69%" height="10"> 
                          <select name="ListSection">
                            <option value="0" selected>&lt;-- Search List By Section--&gt;</option>
                            <%List_Section()%> 
                          </select>
                        </td>
                      </tr>
					   <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" bgcolor="#EEEEEE" align="right">Date Type
                          : </td>
                        <td width="69%" height="10"> 
                          <select name="datetype">
                            <option value="0" >PR Date</option>
							<option value="1" >PO Date</option>
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10"> 
                          <div align="right">From&nbsp;</div>
                        </td>
                        <td width="69%" height="10"> 
                          <select name="From_day">
                            <%Create_daylist(day(date-7)) %> 
                          </select>
                          <select name="From_month">
                            <% Create_MonthList(month(date-7))%> 
                          </select>
                          <select name="From_Year">
                            <% Create_YearList(Year(date))%> 
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10"> 
                          <div align="right">To &nbsp;</div>
                        </td>
                        <td width="69%" height="10"> 
                          <select name="To_Day">
                            <%Create_daylist(day(date)) %> 
                          </select>
                          <select name="To_Month">
                            <% Create_MonthList(month(date))%> 
                          </select>
                          <select name="To_Year">
                            <% Create_YearList(Year(date))%> 
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10"> 
                          <div align="right"><img src="../images/dot.gif"></div>
                        </td>
                        <td width="69%" height="10"> 
                          <input type="submit" name="Submit1" value="Submit">
                          <input type="reset" name="Reset1" value="Reset">
                        </td>
                      </tr>
                    </Form>
                  </table>
                  </td>
                </tr>
                <tr align="left" valign="top"> 
                  <td width="98%" bgcolor="#FFD553">&nbsp;</td>
                  <td width="2%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
    </tr>
  </table>
  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr align="center"> 
      <td width="1%"> <img src="../report/images/spacer.gif" width="1" height="1"><br>
      </td>
      <td width="99%"> 
         <!--#include file="../inc/copyRight.inc"-->
		 </td>
    </tr>
  </table>

</BODY>
</HTML>
<%
Sub Create_Yearlist(searchyear)
Dim syear
	For i=2000 to searchyear
		if i=searchyear then
			Response.Write "<option value= '"& i & "' selected>"& i &"</option>"& vbCrLf
		else
			Response.Write "<option value= '"& i & "'>"& i &"</option>"& vbCrLf
		End if
	Next
End Sub
%>
<%
Sub List_Position()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from Pposition Order By sposcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
		Response.Write "<option value= '" & Rs2("sposcode") & "'" & cSelect &">"& Rs2("sposition") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Division()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from PDivision where sdiv not like '%Cancel%' Order By sDivcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
		Response.Write "<option value= '" & Rs2("sDivcode") & "'" & cSelect &">" & Rs2("sDivcode") & "-"& Rs2("sDiv") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Department()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from PDepartment where sdept not like '%Cancel%' Order By sDepcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
		Response.Write "<option value= '" & Rs2("sDepcode") & "'" & cSelect &">" & Rs2("sDepcode") & "-"& Rs2("sDept") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Section()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from PSection where ssec not like '%Cancel%' Order By sSeccode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
		Response.Write "<option value= '" & Rs2("sSeccode") & "'" & cSelect &">" & Rs2("sSeccode") & "-"& Rs2("sSec") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub

Sub Show_Pcatalog(scatid)
Dim sSQL2,RS2,cSelect
	sSQL2= "Select * from pcatalog where cstatus='1' Order By sCatID"
	'response.write sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("scatid") = scatid Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("scatid")& "'" & cSelect &">"&Rs2("scatid")&"-" & Rs2("scatalog") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	set Rs2=Nothing
End Sub

Sub Show_Pcattype(scattype)
Dim sSQL2,RS2,cSelect
	sSQL2= "select * from pcattype where cstatus='1' Order By sCatType"
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("scattype") = scattype Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("scattype")& "'" & cSelect &">"&Rs2("scattype")&"-" & Rs2("scatdesc") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	set Rs2=Nothing
End Sub

%>