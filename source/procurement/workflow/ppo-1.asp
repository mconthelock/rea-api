<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../Connas400.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<!-- #include file="ppo.inc" -->

<%' Version: 1.01	%>	
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim spono,svendcode,dpodate,srequester,skey,sRefno,sShipto,sShipby,dDelivery,sDelivery,sByer,sAcceptby,sPayterm,sTradeterm
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond,sCurcode,sRef,cstatus,stermcode
	Dim Item,Pono,stname,sEaddress,sCurrency,sEname,dAccept_day,dAccept_month,dAccept_year,dAcceptdate_day,dAcceptdate_month,dAcceptdate_year
spono=Request("sPono")

sSQL="Select * from ppo,pvender,ptermcode where ppo.svendcode=pvender.svendcode" &_
" and pvender.stermcode=ptermcode.stermcode(+)"&_
" and ppo.sPono='" & sPono & "'" 
'response.write sSQL
Set Rs=Server.CreateObject("ADODB.Recordset")
Rs.MaxRecords=1
Rs.Open sSQL,conn,adOpenStatic,adLockReadOnly,adCmdText
If Not Rs.EOF  then
	sPono=Rs("sPono")
	dPodate=Rs("dPodate")
	sRefno=Rs("sRefno")
	svendcode = rs("sVendCode")
	sRequester=Rs("sRequester")
	sKey=Rs("sKey")
	sRef=Rs("sRef")
	sTname=Rs("sTname")
	sEname=Rs("sEname")
	sEaddress=Rs("sEaddress")
	sShipto=Rs("sShipto")
	sShipby=Rs("sShipby")
	sPayterm=Rs("sPayterm")
	sTradeterm=Rs("sTradeterm")
	sDelivery=Rs("sDelivery")

	if rs("dDelivery") <> "" then
		dDelivery = rs("dDelivery")
	end if

	sByer=Rs("sByer")
	cStatus=Rs("cstatus")
	stermcode=RS("stermdesc")
	spayterm=stermcode
	
	if Trim(sByer)="0" then
		sByer=Session("sEmpno")
	End if

	sAcceptby=Rs("sAcceptby")
	sCurrency=RS("sCurrency")
End if
%>
<html>
<head>
<title>PURCHASE ORDER</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
<script language="JavaScript" src="ppo.js"></script>
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
          <td width="2%" height="113">&nbsp;</td>
          <td width="96%" align="center" valign="top" height="113"> 
		 <form name="frmMain" method="post" action="updatepo.asp"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td width="50%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"> 
                  PURCHASE ORDER</font></b></font></td>
                <td width="50%" bgcolor="#FF9999" align="center"> 
                  <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">DATE 
                    : <%=show_date(dpodate)%></font></b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"></font></div>
                </td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2">
                    <table width="93%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td colspan="3" rowspan="3" vlign="Top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Attention</b> 
                          : </font> 
                          <p>&nbsp;</p>
                        </td>
                        <td width="45%" vlign="Top" height="36"><%=sEname%></td>
                        <td width="19%" height="36"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Purchase 
                          Order No. </b></font></td>
                        <td width="27%" height="36"><%=sPono%></td>
                      </tr>
                      <tr> 
                        <td valign="top" vlign="Top" width="45%" height="76" rowspan="2"><%=sEaddress%></td>
                        <td width="19%" height="37"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Reference 
                          No. </b></font></td>
                        <td width="27%" height="37"> 
                         <%=sRefno%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" height="40"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Quotation 
                          No. </b></font></td>
                        <td width="27%" height="40"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b> 
                          <input type="text" name="sAcceptby" maxlength="60" size="13" value="<%=sAcceptby%>">
                          </b></font></td>
                      </tr>
                      <tr> 
                        <td colspan="6"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b> 
                          </b> </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Ship 
                          To: </b></font><font color="#FFCCCC"> 
                          <input type="text" name="sShipto" value="<%=sShipto%>" size="45" maxlength="90">
                          </font> </td>
                      </tr>
                      <tr> 
                        <td colspan="6"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Ship 
                          By:</b> 
                          <input type="text" name="sShipby" value="<%=sShipby%>" size="45" maxlength="90">
                          </font></td>
                      </tr>
                      <tr> 
                        <td colspan="4" rowspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Terms 
                          Of Payment:</b> 
                          <input type="text" name="sPayterm" value="<%=sPayterm%>" size="30" maxlength="45">
                          </font></td>
                        <td colspan="2" height="48"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Currency: 
                          </b></font> 
                          <select name="sCurcode">
                            <%   List_Currency(sCurrency) %> 
                          </select>
                        </td>
                      </tr>
                      <tr> 
                        <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Delivery:</b> 
						  <select name="dDelivery">
							<%SCM_daylist(day(dDelivery))%> 
						  </select>
						  <select name="mDelivery">
							<%SCM_MonthList(month(dDelivery))%> 
						  </select>
						  <select name="yDelivery">
							<%Create_YearList(Year(dDelivery))%> 
						  </select>
						  <%
						  if not isNull(sDelivery) then
						  	if trim(sDelivery) <> "" and isNull(dDelivery) then
								response.write sDelivery
							else
								response.write "<br>(" & sDelivery & ")"
							end if
						  end if
						  %>
                          <input type="hidden" name="sDelivery" value="<%=sDelivery%>" size="15" maxlength="30">
                          </font></td>
                      </tr>
                      <tr> 
                        <td colspan="4" rowspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Terms 
                          Of Trade:</b> 
                          <input type="text" name="sTradeterm" value="<%=sTradeterm%>" size="30" maxlength="45">
                          </font></td>
                        <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Buyer :</b> <%if check_user(sByer) <> "" then response.write "(" & sByer & ") " & getEmpName(sByer) else response.write sByer end if%>
                          <input type="hidden" name="sByer" value="<%=sByer%>" size="12" maxlength="12">
                          </font></td>
                      </tr>
                      <tr> 
                        <td colspan="2">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="6"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b> 
                          </b></font></td>
                      </tr>
                      <!--          <tr> 
                      <td colspan="4"><font color="#FFCCCC">.</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Date 
                        : </b></font> 
                        <select name="dAcceptdate_day">
                          <%    if daccept_day="" then daccept_day=day(date)
														Create_daylist(daccept_day) %> 
                        </select>
                        <select name="dAcceptdate_month">
                          <% if dAccept_month="" then dAccept_month=month(date)
																Create_MonthList(dAccept_month)%> 
                        </select>
                        <% if dAccept_year="" then dAccept_year=year(date) %> 
                        <input type="text" name="dAcceptdate_year" size="4" maxlength="4" value="<%=dAccept_year%>">
                      </td>
                      <td width="21%"><font color="#FFCCCC">...</font> </td>
                      <td width="14%"><font color="#FFCCCC">...</font> </td>
                    </tr> --> 
                    </table>
                </td>
              </tr>
			  <%'if not isSCMVender(svendcode, conn) then%>
              <tr align="center"> 
                <td colspan="2">&nbsp;<% Show_Item()%></td>
              </tr>
			  <%'end if%>
            </table>
            <p> 
              <input type="hidden" name="spono" value="<%=sPono%>">
              <input type="hidden" name="sRefno" value="<%=sRefno%>">
				<%if not alreadySentSCM(sPoNo) or (alreadySentSCM(sPoNo) and not isInSCMTray(sPoNo)) then%>
				  <!--<input type="submit" name="Submit" value="Submit">
	              <input type="reset" name="Submit2" value="Reset">-->
			  <%end if%>

				<%'if isSCMVender(svendcode, conn) and (alreadySentSCM(sPoNo) and not isInSCMTray(sPoNo)) then%>
					<!input type="hidden" name="Print" value="3">
		            <!input type="submit" value="Move to P/O Printed">
				<%'end if%>
			  <br>
			  </form>
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()"><font size="1">Back</font></a> 
              <font size="1">|</font> <a href="workflow.asp"><font size="1">Workflow 
              Manager</font></a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
              </font>
			  </p>
          </td>
          <td width="2%" height="113">&nbsp;</td>
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
<br>
<br>
</body>
</html>
<%
	Sub Show_Item()
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,sexptype,Color
Color="#F8D7D1"
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM ppoitem  where spono='" & sPono &_
	"' Order By sProdcode"
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
'response.write RS2("MaxItem")	
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr bgcolor="#FFD553"> 

          <td width="100%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">

              <tr bgcolor="#DB4B37"> 
                <td  width="15%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Code</font></b></font></td>
                <td width="45%" cellspacing="2" cellpadding="2" width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Name</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Quantity</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">U/M</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">UPrice</font></b></font></td>
                <td   width="10%" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Amount</font></b></font></td>

                <td   cellspacing="2" cellpadding="2" width="9%" bgcolor="#8B2718" align="right" valign="top"><img src="../images/y05.jpg" width="12" height="12"></td>
              </tr>
<% 
Item=1
Do While Not RS.EOF 
'Modify by Dermpan on 14-JAN-03
'Add formatNumber to rs("nAmount")
%>
              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sProdcode")%>&nbsp;</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sprodname")%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nQuality"))%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUnitcode")%></td>
                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nPrice"), 3)%></td>
                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nAmount"), 3)%></td>
                <td  cellspacing="2" cellpadding="2" width="8%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
<%if not alreadySentSCM(sPoNo) and not isInSCMTray(sPoNo) then%><a href='ppoitemedt.asp?spono=<%=spono%>&sProdcode=<%=RS("sProdcode")%>'><img src="../images/iconedit2.gif" width="30" height="26" alt="EDIT" border="0"></a><%end if%></td>
              </tr>
<%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
'Item=Cint(RS("nItem"))+1
 RS.MoveNext
Loop
Rs.Close
	sSQL = "SELECT sCurrency,Sum(nAmount) sumamount FROM ppoitem,pCurrency  where sPono='" & sPono & "'" &_
	" and ppoitem.sCurcode=pCurrency.sCurcode" &_
	" Group By sCurrency" &_
	" Order By sCurrency"
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<tr>
<td colspan="9" align="center">
<%	IF Not Rs.EOF then %>
	<table width="50%">
		<tr><td colspan="2" align="center" bgcolor="teal"><Font color="#FFFFFF"><B>Total</B></Font></td></tr>
		<tr><td  width="50%" align="right" bgcolor="#F1B3A7"><Font color="#1111FF"><b><%=Formatnumber(RS("sumamount"), 3)%></b></font></td><td bgcolor="#F1B3A7"><%=RS("sCurrency")%></td></tr>
	</table>
<%	End if%>
</td>
</tr>
<tr>
	<td colspan="7" align="center">
		<table border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td bgcolor="0000FF">
				<table align="center" width="100%" border="0" cellpadding="2" cellspacing="1">
					<%if alreadySentSCM(sPoNo) and not isInSCMTray(sPoNo) then%>
					<tr>
						<td bgcolor="FFFFFF" nowrap colspan="2" align="center"><font color="#FF0000"><b>This P/O is already on Supply Chain System!</b></font></td>
					</tr>
					<%end if%>

					<%if updatedBySCM(sPoNo) then%>
					<tr>
						<td bgcolor="FFFFFF" nowrap colspan="2" align="center"><font color="#FF0000"><b>This P/O was updated by Supply Chain System on <%=getImpDate(sPoNo)%>!</b></font></td>
					</tr>
					<%end if%>
					<%if isSCMVender(svendcode, conn) and (not alreadySentSCM(sPoNo) or isInSCMTray(sPoNo)) and not updatedBySCM(sPoNo) and timeIsOK() then%>
					<tr>
						<td bgcolor="99AAEE" nowrap>Send to vender through<br>Supply Chain System</td>
						<td bgcolor="FFFFCC" acolspan="6" align="center">
										<input type="hidden" name="scm" value="">
										<%if not alreadySentSCM(sPoNo) then%>
												<input type="button" name="" value="Send to SCM" onClick="javascript: sendToSCM()">
											<%'end if%>
											<%elseif isInSCMTray(sPoNo) then%>
												<input type="hidden" name="Print" value="1">
												<input type="button" name="" value="Remove from SCM" onClick="javascript: removeFromSCM()">
										<%end if%>
						</td>
					</tr>
					<%end if%>
				
				<tr>
				<td bgcolor="99AAEE">Send to vender by paper</td>
				<td bgcolor="FFFFCC" nowrap acolspan="6" align="center">
				<%if not alreadySentSCM(sPoNo) or (alreadySentSCM(sPoNo) and not isInSCMTray(sPoNo)) then%>
					For Print : <input type="radio" name="Print" value="1" <%if cStatus="1" then response.write "checked"%>>Yes
					<input type="radio" name="Print" value="3"  <%if cStatus="3" then response.write "checked"%>> No 
					<%if not alreadySentSCM(sPoNo) then%>
					<br><input type="radio" name="Print" value="9" <%if cStatus="9" then response.write "checked"%>> Cancel P/O
					<%end if%>
					<br>
					 <input type="submit" name="Submit" value="Submit">
					 <input type="reset" name="Submit2" value="Reset">
				<%end if%>
				</td>
				<!--<td   width="8%" bgcolor="#FFD553" align="right" height="8"></td>-->
				</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
            </table>

          </td>

        </tr>
 
      </table>

<%
End Sub
%>