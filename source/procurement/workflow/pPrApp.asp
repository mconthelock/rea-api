<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim sprno,sexptype,dprdate,srequester,skey,cprtype,nstep,capprove,sexpdesc,sCashreq
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year,chk3
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond,sNote,sEmpno
	Dim Item,Prno,sDelivery,sStepdesc,cApprov
	Dim dDelivery
			sPrno=Request("sPrno")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
		  sSQL = "SELECT * FROM  ppr,pexpenstype WHERE ppr.sexptype=pexpenstype.sexptype" &_
			" and sprno='" & sprno &"'"
'			response.write sSQL
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
Do while Not RS.EOF
	sPrno=Rs("sPrno")
	sExpdesc=RS("sExpdesc")
	sExptype = RS("sExptype")
	dPrdate=RS("dPrdate")
	sDelivery=Rs("sDelivery")
	dDelivery = rs("dDelivery")
	sRequester=RS("sRequester")
	sKey=RS("sKey")
	cprtype=Rs("cprtype")
	sCashreq=Rs("sCashReq")
RS.MoveNext
Loop
Session("sExptype") = sExptype
			%>
<html>
<head>
<title>PURCHASE REQUISITION Approve</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
<Script Langauge="JavaScript">
<!---
function OpenNewWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="../images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">...</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top" > 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  PURCHASE REQUISITION</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                    <table width="93%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="3" width="50%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>No.</b> 
                          </font>&nbsp;<%=sprno%></td>
                        
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Expense 
                          Type</b> </font>&nbsp;<%=sExpdesc%>
                        	</td>
                      </tr>
                      <tr> 
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Date :</b>&nbsp;</font><%response.write Show_date(dprdate)%><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          </font></td>
                        <td width="50%" colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Delivery Date :</b> </font><%=sDelivery%>&nbsp;<%if dDelivery <> "" then response.write "<br><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""2"" color=""#CC3333""><b>Expected date : </b></font>" & dDelivery end if%></td>
 
                      </tr>
                      <tr> 
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Request 
                          By :</b> </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          </font>
                        	(<%=sRequester%>)
							<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sRequester %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
							 <%Show_Empno(sRequester)%> </a><%Show_Empdesc()%></td>
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Cash Requirement No. : </b></font><%response.write sCashReq%></td>
              
                      </tr>
                      <tr> 
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">
						<b>Key  By :</b></font> (<%=skey%>) 
						<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sKey %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
						<%Show_Empno(sKey)%></a></td>
                        <td width="50%" colspan="3"><%if cprtype="1" then %><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b><blink>Urgent Requisition</blink></b> </font><%End if%>&nbsp;</td>
                      </tr>
					  		<%
					Show_Investment()
					%>
                    <tr> 
                      <td width="15%"><font color="#CC3333"><b>Quotation.:</b></font></td>
                      <td colspan="2"><%Show_file()%>&nbsp;</td>
					  <td align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Purchase
                          :</b></font> <%Show_Purchase()%> (<%=sEmpno%>) <%Show_Empno(sEmpno)%>
						</td>
                    </tr>

							<tr><td colspan="5" align="right"><A HREF="ppritemdescedt.asp?sPrno=<%=sPrno%>&status=ADD">Add Quotation</a>
							<%Session("sprno")=sprno%>
							</td></tr>
					<tr>
					<!-- 06/23/2023 Edit By Golf 15199 -->
						<td width="15%" colspan="4"><font color="#CC3333"><b>Bidding No. :</b>&nbsp;&nbsp;</font>
							<%
							sSQL = "SELECT * FROM AMEC.PPRBIDDING WHERE SPRNO = '" & sPrno & "'"
							Set Rs = Server.CreateObject("ADODB.RecordSet")
							Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
							If Not Rs.EOF Then
								Do While Not Rs.EOF
									Dim biddingnoRaw, arrBidding, i
									Dim oneBidding, cyear2, nrunno
									Dim parts, yy, runPart

									biddingnoRaw = Trim("" & Rs.Fields("biddingno"))
									If biddingnoRaw <> "" Then
										arrBidding = Split(biddingnoRaw, "|")
										For i = 0 To UBound(arrBidding)
											oneBidding = Trim(arrBidding(i))
											If oneBidding <> "" Then
												parts = Split(oneBidding, "-")
												If UBound(parts) >= 2 Then
													yy = Right(Trim(parts(1)), 2)
													cyear2 = "20" & yy
													runPart = Trim(parts(2))
													nrunno = CLng(runPart)

													%>
													<a href="JavaScript:OpenNewWindow('/form/gpform/bidding/index.asp?sr=1&amp;no=11&amp;orgNo=030101&amp;y=23&amp;y2=<%=cyear2%>&amp;runNo=<%=nrunno%>&amp;m=3&amp;bp=%2Fform%2Fworkflow%2FmineList%2Easp&amp;menu=1','ShowBiddingForm','width=950,height=700 scrollbars=yes')"><%=oneBidding%></a><br>
													<%
												End If
											End If
										Next
									End If
									Rs.MoveNext
								Loop
							End If
							
						%>
						</td>
					   <td>
					  </td>
					</tr>

                    </table>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;<% Show_Item()%></td>
              </tr>
            </table>
  					<br>
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="../workflow/workflow.asp">Workflow Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font>
          </td>
          <td width="2%" >&nbsp;</td>
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
            <hr noshade size="1">
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#666666"><b><font size="1">MITSUBISHI ELEVATOR ASIA CO.,LTD.</font><br>
            </b><font size="1">700/86 Amata Nakorn Industrial Estate,Moo 6,Bangna-Trad Rd., <br>
T.Don Hua Roh,Muang District,Chonburi 20000 <br>
Tel:(038) 213-170-80 Fax: (038) 213-184-6</font></font></td>
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
%>
<%
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,sexptype,Color,Counter
Counter = 0
Color="#F8D7D1"
	sExptype = Session("sexptype")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

	sSQL = "SELECT  ppritem.*,PPRODUCT.HAZARDstatus FROM ppritem LEFT JOIN PPRODUCT ON PPRODUCT.SPRODCODE = PPRITEM.SPRODCODE where ppritem.sPrno='" & sPrno & "' and PPRODUCT.SEXPTYPE='"&sExptype&"'" &_
	 " Order By nItem"

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

	'RS.PageSize = 5
'response.write RS2("MaxItem")	
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
%>
                  <form method="post" action="../workflow/Approve.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="3">
              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Item</font></b></font>
					</td>
                <td  width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Product Name</font></b></font>
					</td>
				<td  width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Hazard Substance</font></b></font>
					</td>
                <td  width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Vendor</font></b></font>
					</td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Quantity</font></b></font>
					</td>
                <td  width="3%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">U/M</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">UPrice</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Amount</font></b></font>
					</td>
                <td width="3%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Cur</font></b></font>
					</td>
                <td width="3%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="1">Received Budget(ß.)</font></b></font>
					</td>
                <td width="3%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="1">Used Budget(ß.)</font></b></font>
					</td>
					
                <td width="8%" bgcolor="#8B2718">&nbsp;
					</td>
                <td width="19%" bgcolor="#8B2718" align="center">
						<b><font color="#FFFFFF" size="2">Approve</font></b>
					</td>
                <td width="19%" bgcolor="#8B2718" align="center">
						<b><font color="#FFFFFF" size="2">Note</font></b>
					</td>

              </tr>
<% 
Dim budget
Do While Not RS.EOF
Counter = Counter + 1
if cint(RS("cStatus")) <> 0 Then
'End if
budget = split(Show_Budget(RS("sProdCode"), dPrdate,sExptype),",")

%>

              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2" valign="top"><%=RS("nItem")%>&nbsp;
						</td>
                <td  bgcolor="<%=Color%>" height="2" valign="top">
<A HREF="JavaScript:OpenNewWindow('../workflow/ppritemdesc.asp?sprno=<%=RS("sprno")%>&nitem=<%=Rs("nItem")%>&sExptype=<%=sExptype%>','ShowPr','width=600,height=500')"><%=RS("sProdname")%></a>
						</td>
				<td  bgcolor="<%=Color%>" height="2" valign="top">
				<%if RS("HAZARDstatus") = "1" Then%>
					Yes
				<%elseif RS("HAZARDstatus") = "0" then %>
					No
				<%else%>
				<%end if%>
				</td>
                <td  bgcolor="<%=Color%>" height="2" valign="top"><%=RS("sVendcode")%></td>
                <td  bgcolor="<%=Color%>" height="2" valign="top"><%=FormatNumber(RS("nQuality"),3)%></td>
                <td  bgcolor="<%=Color%>" height="2" valign="top"><%=RS("sUnitcode")%></td>
                <td  bgcolor="<%=Color%>" height="2" valign="top"><%=FormatNumber(RS("nPrice"), 4)%></td>
                <td  bgcolor="<%=Color%>" height="2" valign="top"><%=FormatNumber(RS("nAmount"),4)%></td>
                <td  bgcolor="<%=Color%>" height="2" valign="top"><%=RS("sCurcode")%><img src="../images/dot.gif"></td>
                <td  bgcolor="<%=Color%>" height="2" valign="top"><%=FormatNumber(budget(0),0)%></td>				
                <td  bgcolor="<%=Color%>" height="2" valign="top">
				<%	
					if cCur(budget(0)) < cCur(budget(1)) then
						response.write "<font color='#FF0000'>"
					else
						response.write "<font color='#000000'>"		
					end if
					response.write FormatNumber(budget(1),0)
					response.write "<font>"
				%>
				</td>								
                <td width="8%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
						<a href="../workflow/ppritemedtedt.asp?nitem=<%=RS("nitem")%>&sprno=<%=RS("sprno")%>&sexptype=<%=sExptype%>">
						<img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a>
						</td>
                <td width="9%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
						<input type="radio" name='cStatus_<%=RS("nItem")%>' value="1" checked> A <br>
						<input type="radio" name='cStatus_<%=RS("nItem")%>' value="0"> R 
					</td>
                <td width="9%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
	                    <input type="text" name="sNote_<%=RS("nItem")%>" size="12" maxlength="60" value="<%=sNote%>">					</td>
              </tr>
<%

Item = cInt(RS("nItem"))+1
	if Color="#F8D7D1" then
			Color="#F1B3A7"
	else
			Color="#F8D7D1"
	end if
End if
RS.MoveNext
Loop

Rs.Close
	sSQL = "SELECT sCurrency,Sum(nAmount) sumamount FROM ppritem,pCurrency  where sprno='" & sPrno & "'" &_
	" and ppritem.sCurcode=pCurrency.sCurcode" &_
	" and cstatus<>'0'" &_
	" Group By sCurrency" &_
	" Order By sCurrency"
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<tr>
<td colspan="13" align="center">
<%	IF Not Rs.EOF then %>
	<table width="50%">
		<tr><td colspan="4" align="center" bgcolor="teal"><Font color="#FFFFFF"><B>Total</B></Font></td></tr>
		<%	Do While Not RS.EOF %>
		<tr><td  width="50%" align="right" bgcolor="#F1B3A7"><Font color="#1111FF"><b><%=Formatnumber(RS("sumamount"), 3)%></b></font></td><td bgcolor="#F1B3A7"><%=RS("sCurrency")%></td></tr>
		<%	RS.MoveNext
			Loop
		%>
	</table>
<%	End if%>
</td>
</tr>
            </table>
<DIV Align="Center">
							<INPUT type="hidden" Value='<%=sPrno%>' name="sPrno" >
							<INPUT type="hidden" Value='<%=item%>' name="nItem" >
							<INPUT type="hidden" Value='<%=Counter%>' name="Count_Item" >
                     <input type="submit" name="Submit" value="Enter">
</DIV>
          </td>
        </tr>
       </table>
<br>
            <table border="0" width="100%" cellspacing="1">
              <tr bgcolor="#8B2718"> 
                <td width="3%" height="16" >&nbsp;</td>
                <td width="27%" height="16"><b><font color="Red">Step No.</font></b></td>
                <td width="39%" height="16"><font color="Red"><b>Approve By</b></font></td>
                <td width="18%" height="16" align="center"><font color="Red"><b>Date</b></font></td>
                <td width="13%" height="16" align="center"><font color="Red"><b>Time</b></font></td>
              </tr>
 
				<%Showdesc()%>			
            </table>
</Form>
<%
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
'Response.write("<BR>nItem = "&Item)
'Response.write("<BR>Counter = "&Counter)
%>
<%
End Sub
%>
<%
Sub chkprapp()
	Dim RS2,sSQL2,RS3
	chk3=0
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select *  from pprapp where  sprno='" & sprno & "'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF  Rs2.EOF then
			chk3="1"
		End if

End Sub

Sub show_file()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from ppritemdesc Where sprno='" & sprno & "'"

		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
%><a href="../workflow/ppritemdescimages/<%=RS2("sfile")%>" target="_Blank">
<%					Response.Write(Rs2("sfile") & "</a>&nbsp;&nbsp;&nbsp;")
					Response.write(RS2("sdesc") & "&nbsp;&nbsp;&nbsp;&nbsp;")
		 %>
			<a Href="delfile.asp?sprno=<%=sprno%>&nitem=<%=RS2("nitem")%>&sExptype=<%=Request("sExptype")%>&sitemno=<%=RS2("sitemno")%>&status=DEL">Del</a>
<%		
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub Show_Empdesc()
	Dim RS2,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select *  from vAemployee where  sEmpno='" & sRequester & "'"

	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF  Not Rs2.EOF then
			response.write "<font size=1> " & RS2("sSeccode") & "/"
			response.write RS2("sDepcode") & "</font>"
		End if

End Sub

Sub Show_Purchase()
	Dim RS2,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select sEmpno  from pprapp where  sprno='" & sPrno & "'" &_
	" and nstep='30'"

	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF  Not Rs2.EOF then
			sEmpno=RS2("sEmpno")
		End if

End Sub

Sub Showdesc()
Dim RS2,sSQL2,RS3
Dim sEmpno,dappdate,tapptime,nPrStep,sEmpnoapp,sComment
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select nStep from ppr where sprno='" & sprno &"'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
	Do While Not Rs2.EOF
		nPrStep=Rs2("nStep")
	Rs2.MoveNext
	Loop
	Rs2.Close
'	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
		  sSQL2 = "SELECT * FROM  pprapp,pFlowstep WHERE  pprapp.sstep=pFlowstep.sstep and sprno='" & sprno &"'" &_
			" Order By nstep"
'			response.write sSQL2
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
Do while Not RS2.EOF
	sEmpno=Rs2("sEmpno")
	sEmpnoapp=RS2("sEmpnoapp")
	nstep=RS2("nstep")
	sStepdesc=RS2("sStepdesc")
	dappdate=RS2("dappdate")
	tapptime=RS2("tapptime")
	cApprov=Rs2("cApprov")
	sComment=RS2("sComment")
if nstep<>"" and nprstep<>"" then
If Cint(nStep)=Cint(nPrStep) and capprov<>"1" then
%>
<tr><td><img src='../images/istepappi.gif'></td><td><%=sStepdesc%></td>
<td><Font Color=red>(<%=sEmpno%>)<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpno %>','SearchProduct','width=500','height=250')"><%Show_Empno(sEmpno)%></A><%if sEmpnoapp<>"" then%> /(<%=sEmpnoapp%>)<%End if%>
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=RS2("sEmpnoapp") %>','showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sEmpnoapp)%></A>
</font></td><td align="Center"><%=dappdate%></td><td align="Center"><%= tapptime%></td></tr>
<%
Elseif cApprov="1" or cApprov="5" then %>
<tr><td><img src='../images/istepapp.Gif'></td><td><%=sStepdesc%></td><td><Font Color=red>
<%if sEmpNo <> sEmpNoApp and not isNull(sEmpNoApp) then%>
<%if sEmpNo = sComment then response.write "->" end if%>
(<%=sEmpNo%>)
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpNo %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sEmpNo)%></A>/
<%end if%>

<%if sEmpNoApp = sComment then response.write "->" end if%>
<%if sEmpNo <> sEmpNoApp and not isNull(sEmpNoApp) then%>
(<%=sEmpNoApp%>)
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpNoApp%>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sEmpNoApp)%></A>
<%else%>
(<%=sComment%>)
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sComment%>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sComment)%></A>
<%end if%>
</Blink>

</font></td><td align="Center"><%=dappdate%></td><td align="Center"><%= tapptime%></td></tr>
<%
Else %>
<tr><td><img src='../images/istepappw.Gif'></td><td><%=sStepdesc%></td><td><Font Color=red>(<%=sEmpno%>)<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpno %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sEmpno)%></A></Blink><%If sEmpnoapp <> "" then %> / (<%=sEmpnoapp%>)<%End if%>
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpnoapp %>','SearchProduct','width=500','height=250')">
<%Show_Empno(sEmpnoapp)%></A>
</font></td><td align="Center"><%=dappdate%></td><td align="Center"><%= tapptime%></td></tr>
<%
End if
End if
RS2.MoveNext
Loop
End Sub
Sub Show_Emp(sEmpno)
    Dim  Rs,sSQL,sStatus,sScriptPath,smode
	Dim scurcode,scurrency
	Dim img_sEmpno, img_sName, img_sSurname, img_sTName,  img_sTSurname, img_sLogin ,img_sPassword
	Dim img_sDivision, img_sDepartment, img_sSection, img_sPosition
	Dim sDivision, sDepartment, sSection, sPosition
	Dim sDivcode, sDepcode, sSeccode, sPoscode
	Dim sDivc, sDepc, sSecc, sPO, sName, sSurname, sTName, sTSurname
	Dim sLogin, sPassword, sPassword1
	Dim bError, sbread
	Dim sErrImage
	bError = False
	sErrImage = "<IMG SRC='../images/icerror.gif'>"
sScriptPath = Request.ServerVariables("SCRIPT_NAME")


  sSQL = "SELECT * FROM vaEmployee WHERE sEmpno ='" & sEmpno & "'"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sEmpno = Rs("sEmpno")
		sDivcode = Rs("sDivcode")
		sDivision=RS("sDivision")
		sDepcode = Rs("sDepcode")
		sDepartment=RS("sDepartment")
		sSeccode = Rs("sSeccode")
		sSection=Rs("sSection")
		sPoscode = Rs("sPoscode")
		sPosition=RS("sPosition")
		sName = Rs("sName")
		sSurname = Rs("sSurname")
		sTName = Rs("sTName")
		sTSurname = Rs("sTSurname")
		sLogin = Rs("sLogin")
		sPassword = RS("sPassword")
Response.write sName
	End If
End Sub

Sub Show_Investment()

    Dim Rs2, sSQL2, eBudgetNo, arr, yy, running
    Set Rs2 = Server.CreateObject("ADODB.RecordSet")

    sSQL2 = "SELECT EBUDGETNO FROM PPRBIDDING WHERE sprno='" & sprno & "'"
    Rs2.Open sSQL2, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    IF Not Rs2.EOF  Then
	  IF Not IsNull(Rs2("EBUDGETNO")) And Trim(Rs2("EBUDGETNO") & "") <> "" Then
        eBudgetNo = Rs2("EBUDGETNO") & ""
        arr = Split(eBudgetNo, "-")

        yy = Mid(arr(1), 4, 2)
        running = CLng(arr(2))
%>
	<tr>
            <td colspan="3">
                <font color="#CC3333"><b>Investment Document</b></font>&nbsp;&nbsp;<a href="javascript:OpenNewWindow(
                'https://amecweb.mitsubishielevatorasia.co.th/webflow/ieform/eBudget/reqtion/index/<%=Session("sEmpno")%>/11/051401/24/20<%=yy%>/<%=running%>',
                'eBudget',
                'width=950,height=700,scrollbars=yes'
                )">
                <%=Rs2("EBUDGETNO") %>
                </a>
            </td>
        </tr>
<%
	  End If
    End If

    Rs2.Close
    Set Rs2 = Nothing

End Sub

'02/02/2010 not show represent
'Sub Showdesc()
'Dim RS2,sSQL2,RS3
'Dim sEmpno,dappdate,tapptime,nPrStep,sEmpnoapp,sComment
'	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
'	sSQL2="Select nStep from ppr where   sprno='" & sprno &"'"
'	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
'	Do While Not Rs2.EOF
'		nPrStep=Rs2("nStep")
'	Rs2.MoveNext
'	Loop
'	Rs2.Close

'		  sSQL2 = "SELECT * FROM  pprapp,pFlowstep WHERE  pprapp.sstep=pFlowstep.sstep and sprno='" & sprno &"'" &_
'			" Order By nstep"
'		  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
'Do while Not RS2.EOF
'	sEmpno=Rs2("sEmpno")
'	sEmpnoapp=RS2("sEmpnoapp")
'	nstep=RS2("nstep")
'	sStepdesc=RS2("sStepdesc")
'	dappdate=RS2("dappdate")
'	tapptime=RS2("tapptime")
'	cApprov=Rs2("cApprov")
'	sComment=RS2("sComment")

'if nstep<>"" and nprstep<>"" then
'If Cint(nStep)=Cint(nPrStep) and capprov<>"1" then
%>
<!--
<tr>
<td><img src='../images/istepappi.gif'></td><td><%=sStepdesc%></td>
<td><Font Color=red>(<%=sEmpno%>)<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpno %>','SearchProduct','width=500','height=250')">
<%'Show_Empno(sEmpno)%></A><%'if sEmpnoapp<>"" then%> /(<%'=sEmpnoapp%>)<%'End if%>
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%'=RS2("sEmpnoapp") %>','showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%'Show_Empno(sEmpnoapp)%></A>
</font></td><td align="Center"><%'=dappdate%></td><td align="Center"><%'= tapptime%></td></tr>
<%
'Elseif cApprov="1" or cApprov="5" then %>
<tr><td><img src='../images/istepapp.Gif'></td><td><%'=sStepdesc%></td><td><Font Color=red>
(<%'=sComment%>)
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%'=scomment %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%'Show_Empno(sComment)%></A></Blink>

</font></td><td align="Center"><%'=dappdate%></td><td align="Center"><%'= tapptime%></td></tr>
<%
'Else %>
<tr><td><img src='../images/istepappw.Gif'></td><td><%'=sStepdesc%></td><td><Font Color=red>(<%'=sEmpno%>)<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpno %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%'Show_Empno(sEmpno)%></A></Blink><%'If sEmpnoapp <> "" then %> / (<%'=sEmpnoapp%>)<%'End if%>
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%'=sEmpnoapp %>','SearchProduct','width=500','height=250')">
<%'Show_Empno(sEmpnoapp)%></A>
</font></td><td align="Center"><%'=dappdate%></td><td align="Center"><%'= tapptime%></td></tr>
-->
<%
'End if
'End if
'RS2.MoveNext
'Loop
'End Sub
%>
