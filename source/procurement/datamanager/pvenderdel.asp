<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	'/**********************************************************************************/
	'/*COPYRIGHT (C) 2015 Mitsubishi Elevator Asia	                                   */
	'/* PROGRAM NO Program outline function explanation (PUR_001P_M003 Data Manager)   */
	'/* Revision No  1                                                                 */
	'/* DEV by  (Kanittha - WSD) 	                                                   */
	'/* Release on 14 Sep 2015 	                                                   */ 
	'/* Reason = “Mark wording "(DO NOT USE)" infront of vendor name of delete vendor  */
    '/* Add', 'Update', 'Delete', System alert email to PUR PIC, SEM, DDEM, DIM”       */  
	'/**********************************************************************************/
%>
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	
<%
	Dim sSQL, Rs, status, msg
	Dim svendcode, sename
 ' Create a recordset and open it for adding records
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  sSQL = "select count(svendcode) as nr from ppritem where svendcode = '"&Request.QueryString("svendcode")&"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  if Rs("nr") = "0" then
		  Rs.close
		  sSQL = "select count(svendcode) as nr from pproduct where svendcode = '"&Request.QueryString("svendcode")&"'"
	      Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		  if Rs("nr") = "0" then
			  Rs.close
			  sSQL = "select svendcode , sename from pvender WHERE svendcode='" & Request.QueryString("svendcode") &"'"
			  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			  svendcode = Rs("svendcode") 
			  sename = Rs("sename")
			  Rs.close
			  sSQL = "DELETE FROM pvender WHERE svendcode='" & Request.QueryString("svendcode") &"'"
			  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			  Set Rs = Nothing
			  KeepLog "D","Delete Vendor", svendcode ,sename
				'msg = "Vendor Code "&svendcode&" is Enable by "&Session("sName")
				msg = " To: PUR(PIC), PUR SEM, PUR DDEM, PUR DEM, PUR DIM" _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  &"    Input date: " & Now _
					  & chr(13) & chr(10) _
					  & "    Vendor status are transferred  from Workflow systerm .  Please see as below." _
					  & chr(13) & chr(10) _
					  &"    - 'Delete' , Vendor code : "&svendcode&" , Vendor name : "&sename&" , Input by "&Session("sName") _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  &"  Best Regards," _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  &"  Auto Send mail System. "
				'msg = "Vendor Code "&svendcode&" is Enable by "&Session("sName")			
			  'sendMail MAIL_SERVER, SYSTEM_NAME, SYSTEM_ADMIN_MAIL, "kanittha@MitsubishiElevatorAsia.co.th", "Vendor Management", msg  
			  sendMail MAIL_SERVER, SYSTEM_NAME, SYSTEM_ADMIN_MAIL, "PURALERTMAIL@MitsubishiElevatorAsia.co.th", "Vendor Management", msg  
			  status="del"
			  Response.Redirect "pvender.asp?status=" &status
			  
		   else
				response.write "Can not Delete. Should you use Edit Function"
				response.end
		   end if
	  
	  else
		  response.write "Can not Delete. Should you use Edit Function"
		  response.end
	  end if
	  

	  
Sub KeepLog(status,comment,vendcode,vendname)
	Dim Rs2 , sSQL , id
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL = "select (count(logid)) as id from logvendor"
	Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs2.EOF Then
		If Rs2("id") = "0" Then
			id =  1
		ELSE
			Rs2.close
			sSQL = "select max(logid) as id from logvendor"
			Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			id = cint(Rs2("id")) + 1
		END IF
		'response.write id
		Rs2.close
		Rs2.Open "logvendor", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		Rs2.AddNew
		Rs2("logid") = id
		Rs2("logdate") = Now
		Rs2("USERID") = Session("sEmpno")
		Rs2("USERNAME") = Session("sName")
		Rs2("STATUS") = status
		Rs2("VCOMMENT") = comment
		Rs2("SVENDCODE") = vendcode
		Rs2("SENAME") = vendname	
		Rs2.Update	
	END IF
	Rs2.Close
	'Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	'response.write sSQL
	'if Not Rs.EOF Then
	'	response.write "IFFFFFFFFFFFFF"
	'END IF
	'if Rs2.EOF then
	'	response.write "IFffffffffffffff"
	'else
	'	response.writ "ELSEeeeeeeeeeee"
	'end if
End Sub
%>
