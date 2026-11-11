<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sexptype,sCatalog,sCattype,sCatID , sCatacc , cyear , sn , des , temp , temp1 , vorgnoname , vorgnocode , supply , supplycode, sError
	Dim img_sexptype, img_CatID, img_Cattype , img_sn , img_vorg
	Dim cStatus

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	sCatID=Request("sCatID")
	sCattype=Request("sCAttype")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")

	If IsEmpty(Request.Form("EVENT")) Then
		If  sCatID <> "" Then
			sMode = "EDIT"
			Get_Catalog(sCatID)
		Else
			sMode = "ADD"
			cStatus = 1
		End If	
		Show_Form()
	Else
	   
		Update_Catalog(Request.Form("EVENT"))
	    Update_EBUDGET(Request.Form("EVENT"))
		if Request("status")="1" then
			Response.Redirect "pCatalog.asp?sCattype=" & Request("sCattype")
		else
		    Response.Redirect "pCatalogall.asp?sCatID="&sCatID
		end if
		
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Catalog(sCatID)
	Dim sSQL, Rs , Rs2, pos1 , pos2 , tempsup 
  ' Create a recordset and open the given event. If the Record is not found
	'set the mode to ADD
   'sSQL = "SELECT * FROM PCatalog WHERE sCatID='" & sCatID & "' AND (SCATACC IS NOT NULL OR "
	sSQL = " SELECT * FROM PCATALOG A JOIN PCATTYPE B ON A.SCATTYPE = B.SCATTYPE "_
		&" WHERE sCatID='" & sCatID & "' "_
		&" AND (SCATACC IS NOT NULL OR SEXPTYPE > '40') "_
		&" ORDER BY SEXPTYPE, A.SCATTYPE"

  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sCatID = Rs("sCatID")
		sCattype = Rs("sCattype")
		cStatus = Rs("cStatus")	
		if Rs("sCatacc") <> "" then 
			sn = Rs("sCatacc")
			pos1 = Instr(Rs("sCatalog"), Rs("sCatacc"))			
			if pos1 > 0 then
				temp = Split(Rs("sCatalog"),sn)
				des = mid(Rtrim(Ltrim(temp(1))), 2, len(Rs("sCatalog")))
			else
				des= Rs("sCatalog")
			end if		
		else	
			pos1 = Instr(Rs("sCatalog"), "(")
			pos2 = Instr(Rs("sCatalog"), ")")
			sn = Mid(Rs("sCatalog"),pos1,(pos2-pos1)+1)			
			temp = Split(Rs("sCatalog"),sn)	
			cyear = Ltrim(Rtrim(temp(0)))
			des = Rtrim(Ltrim(temp(1)))
			sn = Replace(sn,"(","")
			sn = Replace(sn,")","")			
		end if
		
		sSQL = "SELECT VORGNO , VORGSUPPLY FROM EBUDGET.EBUDGETCNTLMST WHERE SCATALOG = '"&sn&"'"
		Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
		IF NOT Rs2.EOF then
		    vorgnocode = Rs2("VORGNO")
			supplycode = Rs2("VORGSUPPLY")
		    tempsup = Rs2("VORGSUPPLY")
			sSQL = "SELECT VORGNAME FROM OWNERBUDGET WHERE VORGNO = '"&Rs2("VORGNO")&"'"
			Rs2.Close
			Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			IF NOT Rs2.EOF then
				vorgnoname = Rs2("VORGNAME")
			END IF
			Rs2.Close
			sSQL = "SELECT SDIV FROM PDIVISION WHERE SDIVCODE = '"&tempsup&"'"
			Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			IF NOT Rs2.EOF then
				supply = Rs2("SDIV")
				Rs2.Close
			ELSE
				Rs2.Close
				sSQL = "SELECT SDEPT FROM PDEPARTMENT WHERE SDEPCODE = '"&tempsup&"'"
				Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
				IF NOT Rs2.EOF then
					supply = Rs2("SDEPT")
					Rs2.Close
				ELSE
					Rs2.Close
					sSQL = "SELECT SSEC FROM PSECTION WHERE SSECCODE = '"&tempsup&"'"
					Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
					IF NOT Rs2.EOF then
						supply = Rs2("SSEC")
					END IF
					Rs2.Close
				END IF
			END IF
		ELSE
			Rs2.Close
		END IF
		
		
	Else
		sMode = "ADD"
	End If
End Sub	
%>	

<%
' This routines presents the form to the user. Depending on the mode, the
' form fields are filled with data or blank.
Sub Show_Form()
	' Set the title of the page depending on if an error occurred during
	' form submission or not
%>
<html>
<head>
<title>Edit CATALOG </title>
<link rel="stylesheet" href="../styles.css">
<script>
function showsupply(supply, focus){
   //alert("xxx"+focus);
        var xmlhttp = new XMLHttpRequest();
        xmlhttp.onreadystatechange = function() {
            if (this.readyState == 4 && this.status == 200) {
                document.getElementById("dvsupply").innerHTML = this.responseText;
            }
        };
        xmlhttp.open("GET", "getsupply.asp?v=" + supply + '&focus='+focus, true);
        xmlhttp.send();
}

</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <% If sMode="EDIT" then %> onload="showsupply('<%=vorgnocode%>', '<%=supplycode%>');" <% end if %> >
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
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"><b><img src="../images/plus.gif" width="9" height="9"> 
            DATA MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="95%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  Sub-Group in Expense Type  : Sub-Group in Expense Type Description of Products</font></b></font></td>
                <td width="5%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>" onsubmit="return checkSave()">
                  <table width="98%" border="0" cellspacing="1" cellpadding="2">
					<tr> 
						<td width="20%" align="Right"><b style="color: #cc0000">Sub-Group in Expense Type ID* :</b></td>
                        <td width="80%">
							<% If sMode="EDIT" then
								Response.write sCatID
							Else
								Response.write "<input type='text' name='sCatID' id='sCatID' size='8' maxlength='6' value='"& sCatID &"'>"
							End if 
							%>
						</td>
					</tr>
					<!--<tr> 
						<td align="Right"><b></b></td>
                        <td width="20%"><font color="#CC3333">
							Sub-Group in Expense Type ID(KEY) Data (Character) : Not more than 6 Characters ,Example: 010101. <b> (Not dupplicate)</b></font>
						</td>
					</tr>-->
                    <tr> 
						<td align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Catalog Type : </font></b></td>
                        <td > 
							<select name="sCattype" id="sCattype">
								<% CAll Create_CATTList(sCattype)%> 
							</select>
                        </td>
					</tr>
                    <tr> 
                        <td align="Right"><b></b></td>
                        <td><font color="#CC3333">Sub-Group in Expense Type Data. Select from List.</font></td>
                    </tr>
                    <tr> 
                        <td align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">SN: </font></b></td>
                        <td>
							<input type="text" name="sn" id = "sn" value="<%=sn%>" size="10" maxlength="15">
						</td>
					</tr>
					<tr> 
                        <td align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Description: </font></b></td>	
						<td><input type="text" name="des" value="<%=des%>" size="50" maxlength="120"></td>
					</tr>
					<tr> 
                        <td align="Right">&nbsp;</td>
                        <td><font color="#CC3333">Sub-Group in Expense Type Description(Character) 
                          : Not more than 120 Characters.</font>
						</td>
					</tr>
					<tr>
						<td align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Owner budget. : </font></b></td>
						<td>
							<% If sMode="EDIT" then %>
								<select name="vorgno" id="vorgno" onchange="showsupply(this.value);">
									<% CALL create_ownerbudget(vorgnocode)%>
								</select>
							<% else %>
								<select name="vorgno" id="vorgno" onchange="showsupply(this.value);">
									<option value="">--------</option>
									<% CALL create_ownerbudget("")%>
								</select>
							<% end if %>
						</td>
					</tr>
					<tr>
						<td align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Support to : </font></b></td>
						<td>
							<div id="dvsupply">
								<select name="vorgsupply" >
									<option value="">--------</option>
								</select>
							</div>
						</td>
					</tr>
                    <tr> 
						<td align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Visible : </font></b></td>
						<td>
							<input type="radio" name="radStatus" value="1"<% if cStatus = 1 then response.write " checked" end if %>>Show
							<input type="radio" name="radStatus" value="2"<% if cStatus = 2 then response.write " checked" end if %>>Hide
                        </td>
					</tr>					  
                    <tr> 
						<td width="19%"  align="Right">&nbsp;</td>
                        <td width="81%"><font color="#CC0000"><b>*</b></font>Required Field.</td>
					</tr>
				</table>
                <br>
                
				<input type=HIDDEN name=EVENT value="<%=sMode%>">
                <input type=HIDDEN name=status value="<%=Request("status")%>">
                <%If sMode<> "ADD" and bError=false  then%> 
					<input type=HIDDEN name=sCatID value="<%=sCatID%>">
                <%End if%> 
				
				<input type="submit" name="Submit" value="Save">
                <input type="submit" name="Submit2" value="Reset">
                <br>
			</form>
           </td>
		</tr>
        <tr align="center"> 
			<td colspan="2">&nbsp; </td>
        </tr>
	</table>
	
	
	
    <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
    | <a href="../datamanager/datamanager.asp">Data Manager</a></font></p>
	</td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../security/images/y04.jpg" width="12" height="12"></td>
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
<script type="text/javascript">
	
	function checkSave()
	{	
		if(document.getElementById("sCatID").value == "")
		{
			alert("Please input Sub-Group in Expense Type ID!");
			return false;
		}
		if(document.getElementById("sCattype").value == "")
		{
			alert("Please input Catalog Type!");
			return false;
		}
		if(document.getElementById("sn").value == "")
		{
			alert("Please input SN!");
			return false;
		}
	
		
		return true;
	}
	
		
		
		
</script>
</body>
</html>

<% End Sub %>

<%
Sub Update_EBUDGET(sUpdateMode)
	Dim sErrImage
	Dim Rs, sSQL
	Dim vorgno , scatalog , cstatus , vorggroup , vorgsupply
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	vorgno = Request.Form("vorgno")
	vorgsupply = Request.Form("vorgsupply")
	scatalog = Request.Form("sn")
	cstatus = "1" 

		sSQL = "SELECT * FROM PSECTION WHERE SSECCODE = '"&vorgno&"'"
		Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		If Not Rs.EOF Then
			vorggroup = "051401"
			'vorgno = vorggroup
			Rs.Close
		ELSE
			Rs.Close
			vorggroup = vorgno
		end if
		
		If sUpdateMode = "EDIT" Then
			  sSQL = "SELECT * FROM  EBUDGET.EBUDGETCNTLMST WHERE SCATALOG='" & scatalog &"'"
			  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			  If Rs.EOF then
				Rs.Close
				Rs.Open "EBUDGET.EBUDGETCNTLMST", conn, adOpenKeySet, adLockPessimistic, adCmdTable
				Rs.AddNew
			  end if
		Else
			  Rs.Open "EBUDGET.EBUDGETCNTLMST", conn, adOpenKeySet, adLockPessimistic, adCmdTable
			  Rs.AddNew
		End If
		Rs("VORGNO") = vorgno
		Rs("SCATALOG") = sCattype
		Rs("sCatalog") = scatalog
		Rs("CSTATUS") = cstatus
		Rs("VORGGROUP")	= vorggroup
		Rs("VORGSUPPLY") = vorgsupply
		Rs("USER_UPDATE") = session("sEmpno")
		Rs("DATE_UPDATE") = Now()
		Rs.Update	
		Rs.Close
		Set Rs = Nothing
	
    
End Sub
	
Sub Update_Catalog(sUpdateMode)
	Dim sErrImage, Rs, Rs2, sSQL, sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")	
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	'Get the form data into variables
	 
	' Check to see if Start date is a valid date
		sCatID = Request.Form("sCatID")
		sCattype = Request.Form("sCattype")
		sCatalog = Request.Form("des")
	'	sCatalog = Request.Form("cyear")&" ("&Request.Form("sn")&") "&Request.Form("des")
		sCatacc = Request.Form("sn")
		cStatus = Request.Form("radStatus")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty
	If (Trim(sCatID) = "") Then
		bError = True
		img_CatID = sErrImage
	End If
	If (Trim(sCattype) = "") Then
		bError = True
		img_Cattype = sErrImage
	else
		sSQL = "select * From amec.PCATTYPE where SCATTYPE = '"&sCattype&"' and SEXPTYPE in ('10','20','30','40')"
		Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		If Not Rs2.EOF then
				If Request.Form("vorgno") = "" then
						response.write "<h1 style='color:red'>Please input Owner budget</h1>"
						response.end	
				End If
		End If
		Rs2.Close
	End If
	If (Trim(sCatacc) = "") Then
		bError = True
		img_sn = sErrImage
	End If

	
	If bError = False Then
		' Create a recordset
		Set Rs = Server.CreateObject("ADODB.RecordSet")		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
			sSQL = "SELECT * FROM  PCatalog WHERE sCatID='" & sCatID &"'"
			Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Rs("SCATACC") <> sCatacc then
				Set Rs2 = Server.CreateObject("ADODB.RecordSet")
				sSQL2 = "SELECT * FROM  PCatalog WHERE scatacc ='" & sCatacc &"'"
				Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
				if Not Rs2.EOF  Then
				    if Rs2("SCATID") <> Rs("SCATID") then
						response.write "<h1 style='color:red'>" & Rs("SCATACC") & "is already exist</h1>"
						response.write "<div>"& Rs2("sCatID")& " - " & Rs2("ScatAcc") &" - "& Rs2("scatalog") & "</div>"
						response.end
					end if
				end if
				Rs2.Close
			end if
		Else
			'Check SCATACC => Can't dupplicate
			sSQL = "SELECT * FROM  PCatalog WHERE sCatID='" & sCatID &"' or sCatAcc = '" & sCatacc & "'"
			Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			If Not Rs.EOF then
				response.write "<h1 style='color:red'>Dupplicate Sub-Group  ID or Sub-Group Code</h1>"
				response.write "<div>"& Rs("sCatID")&" - "& Rs("scatalog") & "</div>"
				response.end
			end if
			Rs.Close
			
			Rs.Open "pCatalog", conn, adOpenKeySet, adLockPessimistic, adCmdTable
			Rs.AddNew
		End If
		' Response.Write		 sSQL
		' Populate the RecordSet fields	
		' dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("sCatID") = sCatID
		Rs("sCattype") = sCattype
		Rs("sCatalog") = sCatalog
		Rs("SCATACC") = sCatacc		
		Rs("cStatus") = cStatus
		Rs.Update	
		Rs.Close
		Set Rs = Nothing
	else
		Show_Form()
	End If
End Sub	
%>