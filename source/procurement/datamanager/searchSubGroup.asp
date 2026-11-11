<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	' Author: Ekkapan Kongchoo

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sCattype,sCatdesc,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sedesc,nprice,simage,dadddate,sempno
	Dim img_sprodcode	,img_sexptype,img_svendcode,img_nprice,sbread,sExptype
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year ,sErrImage
	Dim yearA
	Dim sn , chksn , pos1 , pos2
	
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sCattype=Request.QueryString("sCattype")
	sExptype=Request("sExptype")
	'sCatdesc=Request.QueryString("sCatdesc")

'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
'	response.write sCattype
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
' Create a recordset and open the given event. If the Record is not found
' set the mode to ADD
	sSQL="Select * from pCattype where sCattype='" & sCattype & "'"
'	response.write sSQL
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	sCatdesc=RS("sCatdesc")
	RS.Close

'Rem by Dermpan
'  sSQL = "SELECT * FROM PCatalog  WHERE sCattype='" & sCattype  & "'"

'if inStr("11,12,13,14,15,16,17,18,19,1A,1B,1C,1D,1E,1F,1G,1H,1I,1J,1K" _
'& ",21,22,23,24,25,26,27,28,29.1,29.2,29A,31,32,41,42,43,44,45,46", sCattype) > 0 then
'#####Test Comment#########
'if inStr("11,12,13,14,15,16,17,18,19,1A,1B,1C,1D,1E,1F,1G,1H,1I,1J,1K" _
'& ",21,22,23,24,25,26,27,28,29.1,29.2,29A" _
'& ",31,32,33,41,42,43,44,45,46" _
'& "47,48,42.2,45.1,45.2,46.1,46.2,46.3,47.1," _
'& "47.2,47.3,45.3,4.8,4.9,4.5,4.9A,4.9B", sCattype) > 0 then
'#####End Test Comment#######

if inStr("'11',12,13,14,15,16,17,18,19,1A,1B,1C,1D,1E,1F,1G,1H,1I,1J,1K" _
& ",21,22,23,24,25,26,27,28,29.1,29.2,29.3,29.4,29A" _
& ",31,32,33,41,42,43,44,45,46" _
& "47,48,'42.2',45.1,45.2,46.1,46.2,46.3,47.1," _
& "47.2,47.3,45.3,4.8,4.9,4.5,4.9A,4.9B", sCattype) > 0 then
'response.write sCattype
'  sSQL = "SELECT * FROM PCatalog  WHERE sCatId like '" & sCattype & right(year(date), 1) & "%'"
	'Modify by Dermpan, 7-JAN-03, [14]
	
	yearA = right(year(date), 1)
'	yearA = month(date)
	'Account start at 1 APR to 31 MAR
	if month(yearA) > 3 then
		yearA = right(year(date)+1, 1)
	end if

'28/06/2006 remove comment by Pik
'  sSQL = "SELECT * FROM PCatalog  WHERE sCatId like '" & sCattype & yearA & "%' or sCatId like '" & sCattype & yearA + 1 & "%'"
'  sSQL = "SELECT * FROM PCatalog  WHERE sCatId like '" & sCattype & right(year(date) - 1, 1) & "%' or sCatId like '" & sCattype & right(year(date), 1) & "%' or sCatId like '" & sCattype & right(year(date) + 1, 1) & "%' order by scatId"
  sSQL = "SELECT pCatalog.*, case when INSTR(pCatalog.SCATALOG,pCatalog.SCATACC)  = 0 then '('||SCATACC||')'||SCATALOG else  SCATALOG end MSCATALOG FROM PCatalog  WHERE cstatus = '1' and sCatalog like '%" & year(date) & "%'  and sCattype = '" & sCattype & "' order by scatId"
  if IsNumeric(sCattype) then
''  sSQL = "SELECT * FROM PCatalog  WHERE sCatId like '" & int(sCattype) & right(year(date), 1) & "%' or sCatId like '" & int(sCattype) & right(year(date) + 1, 1) & "%' order by scatId"
'Begin Year
'	sSQL = "SELECT * FROM PCatalog  WHERE sCatId like '" & int(sCattype) & right(year(date)-1, 1) & "%' or sCatId like '" & int(sCattype) & right(year(date), 1) & "%' order by scatId"
  sSQL = "SELECT pCatalog.*, case when INSTR(pCatalog.SCATALOG,pCatalog.SCATACC)  = 0 then '('||SCATACC||')'||SCATALOG else  SCATALOG end MSCATALOG FROM PCatalog  WHERE cstatus = '1' and sCatalog like '%" & year(date) & "%'  and sCattype = '" & sCattype & "' order by scatId"
  else
  sSQL = "SELECT pCatalog.*, case when INSTR(pCatalog.SCATALOG,pCatalog.SCATACC)  = 0 then '('||SCATACC||')'||SCATALOG else  SCATALOG end MSCATALOG FROM PCatalog  WHERE cstatus = '1' and sCatId like '" & sCattype & right(year(date), 1) & "%' or sCatId like '" & sCattype & right(year(date) + 1, 1) & "%'"  
'  sSQL = "SELECT * FROM PCatalog  WHERE sCatalog like '%" & year(date) & "%'  and sCattype = '" & sCattype & "' order by scatId"
'' sSQL = "SELECT * FROM PCatalog  WHERE sCatId like '" & sCattype & right(year(date)-1, 1) & "%' or sCatId like '" & sCattype & right(year(date) , 1) & "%' order by scatId"  
  end if 
 
'  sSQL = "SELECT * FROM PCatalog  WHERE sCatId like '" & sCattype & right(year(date)-1, 1) & "%' or sCatId like '" & sCattype & right(year(date) + 1, 1) & "%'"
  sSQL = "SELECT pCatalog.*, case when INSTR(pCatalog.SCATALOG,pCatalog.SCATACC)  = 0 then '('||SCATACC||')'||SCATALOG else  SCATALOG end MSCATALOG FROM PCatalog  WHERE cstatus = '1' and sCatId like '" & sCattype & right(year(date), 1) & "%' or sCatId like '" & sCattype & right(year(date) + 1, 1) & "%'"  

else
  sSQL = "SELECT pCatalog.*, case when INSTR(pCatalog.SCATALOG,pCatalog.SCATACC)  = 0 then '('||SCATACC||')'||SCATALOG else  SCATALOG end MSCATALOG FROM PCatalog  WHERE  cstatus = '1' and sCattype = '" & sCattype  & "' order by scatId"
end if
'sSQL = "SELECT * FROM PCatalog  WHERE sCattype = '" & sCattype  & "' order by scatId"		''''''
'response.write sSQL

  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText

%>

 <html> <head> 
<title>Show Product :: Computer</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
<Script Langauge="JavaScript">
<!---
function CloseWindow_onClick() { //v2.0
  window.close();
}
//-->
</Script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2">&nbsp;</td>
    <td width="63%" height="2" align="center" valign="top"><img src="../security/images/spacer.gif" width="1" height="1"></td>
    <td width="36%" height="2" align="center"><img src="../security/images/spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td width="1%" height="2"><img src="../security/images/spacer.gif" width="1" height="1"></td>
    <td colspan="2" height="2" align="center" valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="../datamanager/images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../security/images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top"><b><img src="../images/plus.gif" width="9" height="9"> 
            DATA MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" height="181">&nbsp;</td>
          <td width="96%" align="center" valign="top" height="181"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  Product </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2" height="76"> 
                  <table width="98%" border="0" cellspacing="1" cellpadding="2">
                    <tr> 
                      <td colspan="3" bgcolor="#FF9999"><%=sCatdesc%></td>
                    </tr>
<%
				If Not RS.EOF then
					Do While Not Rs.EOF
							'pos1 = Instr(Rs("SCATALOG"), "(")
							'pos2 = Instr(Rs("SCATALOG"), ")")
							sn = Rs("SCATACC")
							'IF pos1 > 0 THEN
							'	sn = Mid(Rs("SCATALOG"),pos1,((pos2-pos1)+1))
							'	sn = Replace(sn,"(","")
							'	sn = Replace(sn,")","")
							'END IF
							chksn = 1
							'if Left(Trim(Rs("SCATALOG")),2) <> "20" then
								chk_owner()
							'end if
							'response.write chksn&"<br/>"
							
					If chksn = 1 then
%>
                    <tr> 
		                 
                      <td width="75%" > <%If Not RS.EOF then%> 
								<%=RS("MSCATALOG")%> 
                        <%End if%> </td>
 	                    <td width="25%">
						<a href="pproductedt.asp?sExptype=<%=sExptype%>&sCattype=<%=sCattype%>&sCatID=<%=RS("sCatID")%>&status=select&svendcode=<%=Request("sVendcode")%>" target="num">Select</a><%RS.MoveNext%>
						</td>
                    </tr>
<%
					else
						RS.MoveNext
					end if
					Loop
				End if
%>
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;</td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="../datamanager/datamanager.asp">Data Manager</a>
			| <A HREF="Javascript:window.close()"> Close </A></font></p>
          </td>
          <td width="2%" height="181">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom" height="2"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom" height="2"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom" height="2"><img src="../security/images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
           <td> <br>
            <!--#include file="../inc/copyRight.inc"-->
		   </td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
  </tr>
</table>
<br>
<br>
</body>
</html>
<%

Sub chk_owner()
		Dim sSQL2,RS2,RS3  ,cSelect
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Set Rs3 = Server.CreateObject("ADODB.RecordSet")
		sSQL2= "Select * from ebudget.EBUDGETCNTLMST Where UPPER(SCATALOG) = UPPER('"&sn&"') and CSTATUS = '1'"
		'response.write sSQL2&"<br/>"
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
			IF  Rs2.EOF Then
				chksn=1
			ELSE
				sSQL2= "Select * from ebudget.EBUDGETCNTLMST Where UPPER(SCATALOG) = UPPER('"&sn&"') and CSTATUS = '1' and VORGGROUP IN (SELECT VORGGROUP FROM EBUDGET.EBUDGET_ORGANIZE WHERE SEMPNO IN ('"&Session("SEMPNO")&"'))"
				Rs3.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
				IF NOT Rs2.EOF Then
					chksn=1
				ELSE
					chksn=0
				END IF
			END IF
End Sub

'Sub chk_owner()
'Dim sSQL2,RS2,RS3  ,cSelect
'		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
'		Set Rs3 = Server.CreateObject("ADODB.RecordSet")
		'Set Rs4 = Server.CreateObject("ADODB.RecordSet")
		'sSQL2= "Select * from ebudget.EBUDGETCNTLMST Where UPPER(SCATALOG) = UPPER('"&sn&"') and CSTATUS = '0'"
	
		'Rs4.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
		'IF NOT Rs4.EOF Then
		'	chksn=1
		'ELSE
			
'			sSQL2= "Select * from ebudget.EBUDGETCNTLMST Where UPPER(SCATALOG) = UPPER('"&sn&"') and CSTATUS = '1' and VORGGROUP IN (SELECT VORGGROUP FROM EBUDGET.EBUDGET_ORGANIZE WHERE SEMPNO IN ('"&Session("SEMPNO")&"'))"
			'response.write sSQL2
'			Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
			'response.write chksn&"<br/>"
'			IF NOT Rs2.EOF Then
'				chksn=1
'			ELSE
'				sSQL2 = "Select * from ebudget.EBUDGETCNTLMST WHERE UPPER(SCATALOG) = UPPER('"&sn&"')"
'				Rs3.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
'				IF Rs3.EOF Then
'					chksn=1
'				ELSE
'					chksn=0
'				END IF
'			END IF
	   'END IF
'End Sub
%>