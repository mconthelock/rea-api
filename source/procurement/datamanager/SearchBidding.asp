<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	'/*****************************************************************/
	'/*COPYRIGHT (C) 2015 Mitsubishi Elevator Asia	*/
	'/* PROGRAM NO PUR_001P_M003 */
	'/* Revision No 1 */
	'/* DEV by  (Kanittha - WSD) 	*/
	'/* Release on  15-09-14	*/ 
	'/* Reason = ìDon't show product that Mark wording "(DO NOT USE)" infront of vendor nameî*/ 
	'/*****************************************************************/
%>	


<%
	'########## DERMPAN KHOOMRAKSA ##########
	'I change formatNumber(xxx, y) to formatNumber(xxx, 3) on 11-NOV-2002
	'########## END DERMPAN KHOOMRAKSA ##########
	
	' Author: Ekkapan Kongchoo
	
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color,sExptype,keywords,sCattype,sCatalog,sCatID
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Dim sMode,sPrno,nItem,sExptype2,status,sItemno
	Dim txt_biddingno, txt_req_no, txt_req_name
	Color="#F8D7D1"
	'suserid = Request("uid")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	'sExptype=Request("ListProduct")
	status = Request("status")
	sExptype=Request("sExptype")
	scattype=Request("ListCattype")
	scatalog=Request("ListCatalog")
	sCatID=Request("sCatID")
	MaxPage=Request("MaxPage")
	sItemno=Request("sItemno")
	'sSearchProduct=Request("SearchProduct")
	keywords = Ucase(TRIM( Request( "SearchProduct" )))
	
	txt_biddingno = Ucase(TRIM( Request( "txt_bid" )))
	txt_req_no = TRIM( Request( "txt_reqno" ))
	txt_req_name = Ucase(TRIM( Request( "txt_reqname" )))
	sPrno = Request("sPrno")
	nItem = Request("nItem")
	sExptype2=Request("sExptype")
	If IsEmpty(Request("EVENT")) Then
		If  sExptype <> "" Then
			sMode = "Search"
		Else
			sMode = ""
		End If	
		SearchProduct()
		
	Else
			Get_Product(sExptype)
			ShowProduct(sPrno)
		
			
	End If	

Function Get_Product(sExptype)
'response.write sExptype

if txt_biddingno = "" and txt_req_no = "" and txt_req_name = "" then
	sSQL = "SELECT * FROM WEBFORM.VIEW_BIDDING ORDER BY CYEAR2, NRUNNO"
else
	sSQL = "SELECT * FROM WEBFORM.VIEW_BIDDING WHERE BIDDINGNO IS NOT NULL "
	if txt_biddingno <> "" then
		sSQL = sSQL & " AND BIDDINGNO LIKE '%" & txt_biddingno &"%'" 
	elseif txt_req_no <> "" then
		sSQL = sSQL & " AND SEMPNO = '"& txt_req_no &"'"
	elseif txt_req_name <> "" then
		sSQL = sSQL & " AND UPPER(SNAME) LIKE '%"& txt_req_name &"%'"
	end if
	sSQL = sSQL & "  ORDER BY CYEAR2, NRUNNO"
end if

if Request("EVENT")="Show" then
	'sSQL="Select pproduct.* from pproduct where scatID='" & sCatID & "'"
	sSQL = "select * from ( "
	sSQL = sSQL&"Select pproduct.* from pproduct,pvender where pproduct.svendcode = pvender.svendcode and pvender.sename not like '(DO NOT USE)%' and scatID='" & sCatID & "'"
	sSQL = sSQL&" union Select pproduct.* from pproduct where svendcode in NULL and scatID='" & sCatID & "'"
	sSQL = sSQL & ")"
	
	sSQL = "SELECT * FROM WEBFORM.VIEW_BIDDING ORDER BY CYEAR2, NRUNNO"
	
End if
	'response.write sSQL
'	response.end
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
	page=Request("page")
	gPage=Request("PageNo")
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	TotalRows=Rs.RecordCount	
	RS.PageSize = 15
	PageSize=RS.PageSize
	TotalPage=Rs.PageCount
		if gPage="" then	
				PageNo=1
		else
				PageNo=gpage
		end if

	Select Case page
	Case "Prev"
			If gPage=1 then 
				PageNo=1
			else
				PageNo=gPage-1
			end if
		Case "Next"
			 If Cint(gPage)>=Cint(TotalPage) then
					PageNo=TotalPage
			else
					PageNo=gPage+1
			end if
End Select
'			Response.write ("sPrno = ")
'			Response.write sPrno
'	Response.write sSQL
End Function
%>

<html>
<head>
<title>Search Selection no</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
<Script Langauge="JavaScript">
<!---
function CloseWindow_onClick() { //v2.0
  window.close();
}
//-->
</Script>
</head>

<body bgcolor="#FFFFFF" leftmargin="3" topmargin="3" marginwidth="3" marginheight="3">
<% Function SearchProduct() %> 
<FORM METHOD="post" ACTION="<%=sScriptPath%>">
  <DIV ALIGN="CENTER"> <BR>
    <TABLE WIDTH="98%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
      <TR> 
        <TD ALIGN="left" VALIGN="top" WIDTH="1%" HEIGHT="2">&nbsp;</TD>
        <TD WIDTH="63%" HEIGHT="2" ALIGN="center" VALIGN="top"><IMG SRC="../security/images/spacer.gif" WIDTH="1" HEIGHT="1"></TD>
        <TD WIDTH="36%" HEIGHT="2" ALIGN="center"><IMG SRC="../security/images/spacer.gif" WIDTH="1" HEIGHT="1"></TD>
      </TR>
      <TR> 
        <TD WIDTH="1%" HEIGHT="2"><IMG SRC="../security/images/spacer.gif" WIDTH="1" HEIGHT="1"></TD>
        <TD COLSPAN="2" HEIGHT="2" ALIGN="center" VALIGN="top"> <FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="3" COLOR="#666666"></FONT> 
          <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
            <TR BGCOLOR="#FFD553"> 
              <TD ALIGN="left" VALIGN="top" WIDTH="2%" HEIGHT="2"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="3" COLOR="#666666"><IMG SRC="images/y01.jpg" WIDTH="12" HEIGHT="12"></FONT></TD>
              <TD WIDTH="96%" HEIGHT="2" BGCOLOR="#FFD553"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="3" COLOR="#666666"><B><FONT COLOR="#FFCC66">..</FONT></B></FONT></TD>
              <TD ALIGN="right" VALIGN="top" WIDTH="2%" HEIGHT="2"><IMG SRC="../security/images/y02.jpg" WIDTH="12" HEIGHT="12"></TD>
            </TR>
            <TR BGCOLOR="#FFD553"> 
              <TD WIDTH="2%">&nbsp;</TD>
              <TD WIDTH="96%" ALIGN="left" VALIGN="top" BGCOLOR="#FFD553"> 
                <DIV ALIGN="CENTER"> 
                  <TABLE WIDTH="300" BORDER="0" CELLSPACING="1" CELLPADDING="0" HEIGHT="41">
                    <TR> 
                      <TD COLSPAN="2"> 
                    <% 'if sPrno = "" then %>
					Selection No. 
                    <input type="text" name="txt_bid" size="50" maxlength="20" placeholder="Ex.GP-BID23-0000001"><br>
					Request No.  
                    <input type="text" name="txt_reqno" size="50" maxlength="5" placeholder="Ex.15199"><br>
					Request Name
                    <input type="text" name="txt_reqname" id="txt_reqname" size="50" maxlength="50"  placeholder="Ex.SOMCHAI"><br><br>
					<input type="submit" name="Submit22" value="Search !!">
					<% 'end if %> 
                        <INPUT TYPE="hidden" NAME="sPrno" VALUE='<%=sPrno%>'>
                        <INPUT TYPE="hidden" NAME="nItem" VALUE='<%=nItem%>'>
                        <INPUT TYPE="hidden" NAME="sItemno" VALUE='<%=sItemno%>'>
						<INPUT TYPE="hidden" NAME="sExptype" VALUE='<%=sExptype2%>'>
                        <INPUT TYPE="hidden" NAME="EVENT" VALUE="Search">
                        <INPUT TYPE="hidden" NAME="status" VALUE="<%=Request("status")%>">
                        <INPUT TYPE="hidden" NAME="smode" VALUE="<%=Request("smode")%>">
                        <INPUT TYPE="hidden" NAME="approve" VALUE="<%=Request("approve")%>">
						

                      </TD>
                    </TR>
						<TR>
                      <TD COLSPAN="2" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37"></font> 
                        <font color="#CC3333"><b><a href="../datamanager/svpBidding.asp?sexptype=<%=sexptype%>&sprno=<%=sprno%>&nitem=<%=nitem%>&status=<%=status%>&sItemno=<%=sItemno%>&smode=<%=Request("smode")%>">View 
                        Selection no</a></b></font></TD>
                    </TR>
                  </TABLE>
                  <FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2"><A HREF="javascript:history.back()">back</A> 
                  | <A HREF="Javascript:window.close()"> Close </A></FONT><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="1"><BR>
                  </FONT></DIV>
              </TD>
              <TD WIDTH="2%">&nbsp;<img src="../images/dot.gif"></TD>
            </TR>
            <TR BGCOLOR="#FFD553"> 
              <TD WIDTH="2%" ALIGN="left" VALIGN="bottom"><IMG SRC="../security/images/y03.jpg" WIDTH="12" HEIGHT="12"></TD>
              <TD WIDTH="96%" BGCOLOR="#FFD553" VALIGN="middle" ALIGN="center"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="1"> 
                </FONT><img src="../images/dot.gif"></TD>
              <TD WIDTH="2%" ALIGN="right" VALIGN="bottom"><IMG SRC="../security/images/y04.jpg" WIDTH="12" HEIGHT="12"></TD>
            </TR>
          </TABLE>
        </TD>
      </TR>
    </TABLE>
  </DIV>
</FORM>
<BR>
<% End Function%>

<% Function ShowProduct(sPrno) %>
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
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../security/images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td colspan="5" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Selection</font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>

				  <tr bgcolor="#1E90FF" align="center" valign="top">
					<td width="3%" >No.</td>
					<td width="17%" >Selection No</td>
					<td width="25%" >Requester</td>
					<td width="25%" >Sec / Dept/ Div</td>
					<td width="30%" >Project</td>
				 </tr>
              <%If  RS.EOF then 
				Response.Write  " <tr bgcolor='#F1B3A7'> " & vbCrLf
		        Response.Write  " <td colspan='5' valign='top'>‰¡Ë¡’√“¬°“√°“√∑’Ë§ÈπÀ“"
				Response.Write  "</td></tr> "& vbCrLf
		else
				RS.Absolutepage=PageNo
				RowCount=1
				Do While Not RS.EOF and RowCount<=PageSize %> 
                    <tr bgcolor="<%=Color%>"> 
                        <td  bgcolor="<%=Color%>" height="10">
					       <DIV ALIGN="RIGHT"><%=((PageNo-1)*RS.PageSize)+RowCount%>.&nbsp;</DIV>
						</td>
						<td  bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;
							<a href=../workflow/showprdesc.asp?biddingno=<%=RS("biddingno")%>&sprno=<%=Request("sPrno")%>&nitem=<%=Request("nItem")%>&status=<%=Request("status")%>&sExptype=<%=sExptype%> Target="num"><%=RS("biddingno")%></a>
						</td>
						<td  bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;<%=RS("sempno") & "_" & RS("sname") %></td>
						<td  bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;<%=RS("ssec") & "/" & RS("sdept") & "/" & RS("sdiv") %></td>
						<td  bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;<%=RS("project")%></td>
					</tr>
              <%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
			RowCount=RowCount+1
		 	RS.MoveNext
			Loop
	end if
	Rs.close
	set Rs=nothing
%> 
            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>

<%if TotalPage > 1 then%>
			<form action= "<%=sScriptPath%>" method="Post">
			<tr><td colspan="4" align="center" bgcolor="#FFD553">
			      <table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr><td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
						<td align="left" bgcolor="#FFD553" width="45%">
<%
							if PageNo=1 then
											Response.write "  "
							else
%>
											<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&ListProduct=<%=sExptype%>&SearchProduct=<%=keywords%>&page=Prev&sExptype=<%=sExptype2%>&sPrno=<%=sPrno%>&EVENT=Search&nitem=<%=nitem%>&status=<%=status%>&sItemno=<%=Request("sItemno")%>&smode=<%=Request("smode")%>&approve=<%=Request("approve")%>">Previous Page</a>
<%					End if
%>
							</td>
						<td align="right" bgcolor="#FFD553" width="45%">
<%
						If PageNo=TotalPage Or PageNo=MaxPage Then
							Response.write " "
						else
%>
									<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&ListProduct=<%=sExptype%>&SearchProduct=<%=keywords%>&page=Next&sExptype=<%=sExptype2%>&sPrno=<%=sPrno%>&EVENT=Search&nitem=<%=nitem%>&status=<%=status%>&sItemno=<%=Request("sItemno")%>&smode=<%=Request("smode")%>&approve=<%=Request("approve")%>">Next Page</a>
						<%End if%>
						</td>
						<td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td></tr>
<tr>
							<td align="center" bgcolor="#FFD553" colspan="4">
<%
						For i=1 to TotalPage
							If Cint(PageNo)=Cint(i) then
%>
								<Font color="red"><%=i%></Font>
							<%Else%>
							<A Href="<%=sScriptPath%>?PageNo=<%=i%>&ListProduct=<%=sExptype%>&SearchProduct=<%=keywords%>&MaxPage=<%=TotalPage%>&sExptype=<%=sExptype2%>&sPrno=<%=sPrno%>&EVENT=Search&nitem=<%=nitem%>&status=<%=status%>&sItemno=<%=Request("sItemno")%>&smode=<%=Request("smode")%>&approve=<%=Request("approve")%>"><%=i%></a>
<%				
							End if	
							Next
 %>
						</td>
</tr>
			</table>
<%End if%>

			</td></tr>
			</form>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom" height="13"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center" height="13">		<BR>
					<FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2">
						<A HREF="javascript:history.back()">back</A> 
			            | <A HREF="Javascript:window.close()"> Close </A>
					</FONT>
				<BR></td>
          <td width="2%" align="right" valign="bottom" height="13"><img src="../security/images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td> <br>
            <hr noshade size="1">
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#666666"><b><font size="1">MITSUBISHI 
            ELEVATOR ASIA CO.,LTD.</font><br>
            </b><font size="1">700/86 Amata Nakorn Industrial Estate,Moo 6,Bangna-Trad Rd.,<br> 
T.Don Hua Roh,Muang District,Chonburi 20000 <br>
Tel:(038) 213-170-80 Fax: (038) 213-184-6</font></font></td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
  </tr>
</table>
<% End Function%><br>
<br>
<div align="right"></div>
</body>
</html>
