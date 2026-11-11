<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<!-- #include file="ppo.inc" -->
<%

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim sprno,sexptype,dprdate,srequester,skey,cprtype,nstep,capprove
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	Dim Item,Prno,CashReq,sEmpno,chksEmpno,sDelivery
	sExptype=Request("sExptype")
	CashReq=Request("CashReq")
	sDelivery=Request("sDelivery")
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
'response.write "[" & request("dDelivery") & "/" & request("mDelivery") & "/" & request("yDelivery") & "]"
'response.end
	if request("sPrno")="" then
		dDate=Now
		dDay=(Day(dDate))
			If dDay < 10 then dDay="0" & Day(dDate)
					dMonth=Month(dDate)
			If dMonth < 10 then dMonth="0" & Month(dDate)
					dYear=Mid(Year(dDate),2,2)
					dHour=Hour(dDate)
			If dHour < 10 then dHour="0" & Hour(dDate)
					dMinute=Minute(dDate)
			If dMinute < 10 then dMinute="0" & Minute(dDate)
					dSecond=Second(dDate)
			If dSecond < 10 then dSecond="0" & Second(dDate)
			   sPrno=dYear & dMonth & dDay & dHour & dMinute & dSecond
		sKey=session("sEmpno")
	else

		sEmpno=Request("srequester")
		chkReq()

		
	if chksEmpno=1 then
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		'########## CHECK FOR GRANTED USER ##########
		'By Dermpan on 13-JUN-03 working report job on 13-JUN-03 job 2 
		Rs.Open "select sSecCode, sDepCode, sDivCode from aEmployee where sEmpNo = '" & replace(request("sRequester"), "'", "''") & "'", conn, adOpenKeySet, adLockPessimistic, adCmdText

		'if not (isGrantedUrgentUser(rs("sSecCode")) or isGrantedUrgentUser(rs("sDepCode")) or isGrantedUrgentUser(rs("sDivCode")) or isGrantedUrgentUserID(request("sRequester"))) and request("cPRType") = "1" then
		if not isGrantedExp() and request("cPRType") = "1" then
			response.write "Can not create urgent P/R for this expense type!"
			response.write "<br><a href=""javascript: window.history.back()"">Back</a>"
			response.end
		end if

		rs.close
		'########## END CHECK FOR GRANTED USER ##########

		  Rs.Open "ppr", conn, adOpenKeySet, adLockPessimistic, adCmdTable

		  Rs.AddNew
 	    sPrno = CreatePRNO

		Rs("sprno") = sPrno
		Rs("sexptype") =Mid(Request("sexptype"),1,2)
		Rs("dprdate") = date
		RS("sDelivery")=sDelivery
		
		if request("dDelivery") <> "" and request("mDelivery") <> "" and request("yDelivery") <> "" then
		'######		04/08/2009	change format date from mm/dd/yyyy to dd/mm/yyyy	#####
		'	rs("dDelivery") = request("mDelivery") & "/" & request("dDelivery") & "/" & request("yDelivery")
		'######		05/01/2010	change to old format	#####	
			rs("dDelivery") = request("dDelivery") & "/" & request("mDelivery") & "/" & request("yDelivery")
		end if
		
		Rs("srequester") =Ucase(Request("srequester"))
		Rs("skey") =Ucase(Request("skey"))
		RS("cApprove")="9"
		RS("cpo")="0"
		RS("sCashReq")=CashReq

		if Request("cPrtype")<> "1" then 
			cPrtype="0"
		else
			cPrtype=Request("cPrtype")
		end if

		Rs("cprtype") =cprtype
		Rs("nstep") =99
	'	Rs("capprove") =capprove

	  Rs.Update	
		
	  Rs.Close	
	  response.redirect "showpr.asp?sprno=" & sprno  

	else


	End if
end if
%>

<html>
<head>
<title>ExpensStep</title>
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
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
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="../images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">...</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top"><b><img src="../images/plus.gif" width="9" height="9"> 
            WORK FLOW MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" valign="top" > 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  PURCHASE REQUISITION</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="26" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form name="frmMain" method="post" action="<%=sScriptPath%>">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Request 
                          Date :</b>&nbsp;</font><%response.write Show_date(date)%><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          </font></td>
                        <td width="63%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Expense 
                          Type </b></font>&nbsp; 
                          <SELECT NAME="sexptype" SIZE=1>
                            <% List_Pexptype(sexptype) %> 
                          </SELECT>
                        </td>
                      </tr>
                      <tr> 
                        <td colspan="3" height="2" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Key 
                          By</b> </font>&nbsp;<%=sKey%>(<%=Session("sName")%>) </td>
                        <td width="63%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b></b> </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          </font> </td>
                      </tr>
                      <tr> 
                        <td colspan="3" height="2" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>
                          </b> </font>&nbsp; </td>
                        <td width="63%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>*Request 
                          By</b> 
                          <input type="text" name="sRequester" size="12" maxlength="12">
                          <br>
                          Requester's Employee No.(Can't be blank) </font></td>
                      </tr>
                      <tr> 
                        <td colspan="3" height="49" valign="top">&nbsp;</td>
                        <td width="63%" height="49"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Delivery Date</b> 
                          <input type="text" name="sdelivery" size="21" maxlength="20">
                          </font></td>
                      </tr>
                      <tr> 
                        <td colspan="3" height="49" valign="top">&nbsp;</td>
                        <td width="63%" height="49"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Delivery Expected Date</b></font>
	<select name="dDelivery"><%SCM_DayList(0)%></select>&nbsp<select name="mDelivery"><%SCM_MonthList(0)%></select>&nbsp<select name="yDelivery"><%Create_YearList(0)%></select>
	</td>
                      </tr>
                      <tr> 
                        <td colspan="3" height="20" valign="top">&nbsp;</td>
                        <td width="63%" height="20"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Cash 
                          Requirement No.</b> 
                          <input type="text" name="CashReq" size="6" maxlength="8"></font></td>
                      </tr>
                      <tr align="left"> 
                        <td colspan="3">&nbsp; </td>
                        <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Cash 
                          Requirement No. :Not more than 8 Characters </font></td>
                      </tr>
					  <%
					  'Add condition by Dermpan on 13-JUN-03 Job in working report is 13-JUN-03 item 2
					 'if isGrantedUrgentUser(session("sSecCode")) or isGrantedUrgentUser(session("sDepCode")) or isGrantedUrgentUser(session("sDivCode")) or isGrantedUrgentUserID(session("sEmpNo")) then
					  if isGrantedExp() then
					  %>
	<!--
                      <tr> 
                        <td colspan="3">&nbsp;</td>
                        <td width="63%"> 
                          <input type="checkbox" name="cPrtype" value="1">
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          <b>Urgent</b></font> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Requisition</font></b></td>
                      </tr>
	-->
					  <%end if%>
                      <tr align="left"> 
                        <td colspan="4">&nbsp; </td>
                      </tr>
                      <INPUT TYPE="Hidden" NAME="sPrno" value="<%=sPrno%>">
                      <INPUT TYPE="Hidden" NAME="sKey" value="<%=sKey%>">
                      <tr> 
                        <td colspan="4" align="center"> 
                          <INPUT TYPE="submit" NAME="Submit" VALUE="Submit">
                          <INPUT TYPE="reset" NAME="Reset" VALUE="Reset">
                        </td>
                      </tr>
                    </table>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2"> </td>
              </tr>
            </table>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">*Urgent</font> 
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Requisition</font> 
            : Use for very special case. Approval steps will be shortened.<br>
            <font color="#CC3333">*</font> Press Submit to add Product Items.<br>
          </td>
          <td width="2%" >&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553">
          <td width="2%" align="left" valign="bottom">&nbsp;</td>
          <td width="96%" bgcolor="#FFD553" valign="bottom" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()"><br>
            back</a> | <a href="workflow.asp">WorkFlow Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font> 
          </td>
          <td width="2%" align="right" valign="bottom">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td> <br>
            <!--#include file="../inc/copyRight.inc"--></td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
  </tr>
</table>
</body>
</html>
<%
'##########
'# Function isGrantExp()
'# Description: Create by Dermpan on 26-MAY-03.
'# This function check for granted expense type which user can specify urgent requisition.
'# Return:
'# 1.true if user was granted.
'# 2.false if user wasn't granted.
'##########
function isGrantedExp()
	if trim(sExpType) = "" then
		isGrantedExp = true
		exit function
	end if

	'if inStr("10,20,30,40,60,B0,C0", sExpType) <> 0 or left(trim(sExpType), 1) = "5" then
	if trim(sExpType) <> "90" then
		isGrantedExp = false
	else
		isGrantedExp = true
	end if
end function

'##########
'# Function isGrantUrgentUserID()
'# Description: Create by Dermpan on 04-JUL-03 Job in working report is 04-JUL-03.
'# This function check for granted user who can specific urgent requisition.
'# Return:
'# 1.true if user was granted.
'# 2.false if user wasn't granted.
'##########
function isGrantedUrgentUserID(userID)
	Dim userList
	
	userList = "97070,92151,02066,03064,02041,96321,94098,92043"
	
	if trim(userID) = "" then
		isGrantedUrgentUserID = false
		exit function
	end if
	
	if inStr(userList, lCase(userID)) > 0 then
		isGrantedUrgentUserID = true
	else
		isGrantedUrgentUserID = false
	end if
end function

'##########
'# Function isGrantUrgentUser()
'# Description: Create by Dermpan on 13-JUN-03 Job in working report is 13-JUN-03 item 2.
'# This function check for granted user who can specific urgent requisition.
'# Return:
'# 1.true if user was granted.
'# 2.false if user wasn't granted.
'##########
function isGrantedUrgentUser(user_section)
	Dim section
	
	section = "40402,50404,50405,50403,50502,50506"
	
	if left(user_section, 2) = "60" or inStr(section, user_section) > 0 then
		isGrantedUrgentUser = true
	else
		isGrantedUrgentUser = false
	end if
end function

Sub chkReq()
	Dim Rs2,sSQL2
    Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select * from aEmployee where sEmpno='" & Ucase(sEmpno) & "'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS2.EOF then
			chksEmpno=1
		End if
	Rs2.Close
	Set Rs2=nothing
End Sub
%>