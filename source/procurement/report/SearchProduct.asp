<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->

<!-- #include file="../List.asp" -->

<%
	' Author: Ekkapan Kongchoo
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color,sExptype,keywords,sCattype,sCatalog,sCatID
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Dim sMode,sPrno,nItem,sExptype2,status,sItemno
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

	sSQL = "SELECT  * FROM pproduct Where  sTprodname LIKE '%" &  keywords & "%' " &_
		" OR sEprodname Like '%" & keywords & "%' " &_
		" OR sprodcode Like '%" & keywords & "%' " &_
 		" ORDER BY sprodcode"

if Request("EVENT")="Show" then
	sSQL="Select * from pproduct where scatID='" & sCatID & "' ORDER BY sprodcode"
End if

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
End Function
%>

<html>
<head>
<title>Search Products</title>
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
                  <TABLE WIDTH="500" BORDER="0" CELLSPACING="1" CELLPADDING="0" HEIGHT="41">
                    <TR> 
                      <TD COLSPAN="2" align="center"> 
                    Search For 
                    <input type="text" name="SearchProduct" size="20" maxlength="20">
                    <input type="submit" name="Submit22" value="Go !!">
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
                        <font color="#CC3333"><b><a href="../datamanager/svpproduct.asp?sexptype=<%=sexptype%>&sprno=<%=sprno%>&nitem=<%=nitem%>&status=<%=status%>&sItemno=<%=sItemno%>&smode=<%=Request("smode")%>">View 
                        Product</a></b></font></TD>
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
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product 
                        :: <% Show_Pexptype(sExptype)%></font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718" HEIGHT="2"> 
                  <DIV ALIGN="CENTER"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></DIV>
                </td>
                <td width="15%" bgcolor="#8B2718" HEIGHT="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">&nbsp;&nbsp;Product 
                  Code</font></b></font></td>
                <td width="49%" bgcolor="#8B2718" HEIGHT="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">&nbsp;&nbsp;Product 
                  Name</font></b></font></td>
                <td align="center" width="12%" bgcolor="#8B2718" HEIGHT="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">UPrice</font></b></font></td>
                <td width="10%" bgcolor="#8B2718" HEIGHT="2"><FONT COLOR="#993300" FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="1">...</FONT></td>
              </tr>
              <%If  RS.EOF then 
       		Response.Write  " <tr bgcolor='#F1B3A7'> " & vbCrLf
		       Response.Write  " <td colspan='5' valign='top'>ไม่มีรายการการที่ค้นหา"
	       	Response.Write  "</td></tr> "& vbCrLf
		else
				RS.Absolutepage=PageNo
				RowCount=1
				Do While Not RS.EOF and RowCount<=PageSize %> 
              <tr bgcolor="<%=Color%>"> 
                <td width="5%" bgcolor="<%=Color%>" height="10">
					<DIV ALIGN="RIGHT">
						<%=((PageNo-1)*RS.PageSize)+RowCount%>.&nbsp;
                  </DIV>
                </td>
                <td width="15%" bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;<%=RS("sprodcode")%>&nbsp;</td>
                <td width="49%" bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;
					<a href=sProductReport.asp?nItem=<%=nItem%>&sExptype=<%=Rs("sExptype")%>&sPrno=<%=sPrno%>&sProdcode=<%=RS("sProdcode")%>&nprice=<%=RS("nprice")%>&status=change&sItemno=<%=Request("sItemno")%>&smode=<%=Request("smode")%>&approve=<%=Request("approve")%> Target="num"><%=RS("seprodname")%></a>
				</td>
                <td  align="right" width="12%" bgcolor="<%=Color%>" height="10"><%=FormatNumber(RS("nPrice"), 3)%>&nbsp;</td>
                <td width="10%" bgcolor="<%=Color%>" align="center" height="10" valign="top"><A HREF="ProductDetail.asp?sprodcode=<%=RS("sProdcode")%>&sexptype=<%=RS("sExptype")%>">Detail</a></td>
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
          <!--#include file="../inc/copyRight.inc"--></td>
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
