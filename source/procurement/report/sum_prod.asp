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

	Dim i
   Dim  Rs,sSQL,sSQL2,RS2
	Dim sScriptPath,sexptype,sPoscode,sVendcode,sVenddesc
	sVendcode=Request("sVendercode")
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM pexpenstype  ORDER BY sexptype "
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

	Dim rsOra, sqlOra
	set rsOra = createObject("adodb.recordset")
%>
<HTML>
<HEAD>
<TITLE>P/O ITEM QUERY FORM</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
<SCRIPT LANGUAGE="JavaScript">
<!--
function Form_Load(){
	document.frmName.SPRODCODE.focus();
}

function OpenNewWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
-->
</SCRIPT>
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
              <tr align="center" valign="top"> 
                <td colspan="2" bgcolor="#FFD553"> 
                  <table border=0 width=100% align="Center" cellspacing="1" cellpadding="0">
                    <tr bgcolor="#CC00CC"> 
                      <td colspan="2" height="20"> 
                        <div align="center"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">P/O 
                          ITEM REPORT QUERY</font></font></b></div>
                      </td>
                    </tr>
                    <Form Action="sum_prodRpt.asp">
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" align="right" nowrap><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Date 
                          range:&nbsp;</font></b></td>
                        <td width="69%" height="10" nowrap> From 
                          <select name="selPOSDate">
                            <%Create_daylist(day(date-7)) %>
                          </select>
                          <select name="selPOSMonth">
                            <% Create_MonthList(month(date-7))%>
                          </select>
                          <select name="selPOSYear">
                            <% Create_YearList(Year(date))%>
                          </select>
                          to 
                          <select name="selPOEDate">
                            <%Create_daylist(day(date)) %>
                          </select>
                          <select name="selPOEMonth">
                            <% Create_MonthList(month(date))%>
                          </select>
                          <select name="selPOEYear">
                            <% Create_YearList(Year(date))%>
                          </select>
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" align="right" nowrap><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Product 
                          code:&nbsp;</font></b></td>
                        <td width="69%" height="10"> 
                          <input type="text" name="txtProdCode" value="">
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td width="31%" height="10" align="right" nowrap><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Product 
                          name:&nbsp;</font></b></td>
                        <td width="69%" height="10"> 
                          <input type="text" name="txtProdName" value="">
                        </td>
                      </tr>
                      <tr bgcolor="#EEEEEE"> 
                        <td colspan="2" height="10" align="center"> 
                          <div align="right"><img src="../images/dot.gif"></div>
                          <br>
                          <input type="submit" name="Submit1" value="Submit">
                          <input type="reset" name="Reset1" value="Reset">
                          <input type="button" value="Back" onClick="javascript: window.location.href = 'report.asp'">
                          <br>
                          &nbsp; </td>
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
        <!--#include file="../inc/copyRight.inc"--></td>
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
		sSQL2= "Select * from PDivision Order By sDivcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
		Response.Write "<option value= '" & Rs2("sDivcode") & "'" & cSelect &">"& Rs2("sDiv") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Department()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from PDepartment Order By sDepcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
		Response.Write "<option value= '" & Rs2("sDepcode") & "'" & cSelect &">"& Rs2("sDept") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Section()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from PSection Order By sSeccode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
		Response.Write "<option value= '" & Rs2("sSeccode") & "'" & cSelect &">"& Rs2("sSec") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub
%>