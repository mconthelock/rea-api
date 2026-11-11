<% Option Explicit %>
<%
           ' Author : Ekkapan Kongchoo
%>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	
   Dim  Rs,RS2,sSQL,sSQL2
	Dim sScriptPath,svendcode,Color,sExptype,keywords,vendcode,sProdcode
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Dim sMode
	Dim sPrno,nItem,sExptype2
		Color="#F8D7D1"
		sScriptPath = Request.ServerVariables("SCRIPT_NAME")
		Set Rs = Server.CreateObject("ADODB.RecordSet")       	
		sExptype=Request("ListVender")
		sProdcode=Request("sProdcode")
		keywords = Ucase(TRIM( Request( "SearchVender" ) ))
		sPrno = Request.QueryString("sPrno")
		nItem = Request.QueryString("nItem")
'		sVendcode = Request.QueryString("sVendcode")
		sExptype2 = Request("sExptype")
		sExptype = "0"
		sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	If IsEmpty(Request("EVENT")) Then
		If sExptype = "" Then
			sMode = "Search"
		Else
			sMode = ""
		End If	
		SearchVender()
	Else
			Get_Vender(keywords)
			sPrno = Request("sPrno")
			nItem = Request("nItem")

'Response.write ("sPrno = "&sPrno)
'Response.write ("nItem = "&nItem)
			ShowVender(sPrno)
	End If	

Function Get_Vender(keywords)
	'	sSQL = "SELECT distinct  svendcode  FROM  pProduct"
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
if  keywords="" then 
		sSQL = "SELECT distinct  svendcode  FROM  pVender" &_
 		" ORDER BY svendcode"
else 
	sSQL = "Select * from pVender where stname LIKE '%" &  keywords & "%' " &_
		" OR sename Like '%" & keywords & "%' "
End if
		page=Request("page")
		gPage=Request("PageNo")
'response.write gpage
		MaxPage=Request("MaxPage")
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
'			sPrno = Request("sPrno")
'			nItem = Request("nItem")
'			Response.write ("sPrno = "&sPrno)
'			Response.write ("nItem = "&nItem)
'			Response.write " sProdcode = "&sProdcode
'			Response.write " sVendcode = "&sVendcode
'			Response.write " sExptype = "&sExptype2
'			Response.write "<BR>"
'			Response.write ("sExptype2 = "&sExptype2)

End Function
%>

<html>
<head>
<title>Search Vender</title>
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

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<% Function SearchVender() %> 
<FORM METHOD="post" ACTION="<%=sScriptPath%>">
  <DIV ALIGN="CENTER">
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
              <TD WIDTH="96%" HEIGHT="2" BGCOLOR="#FFD553">&nbsp;</TD>
              <TD ALIGN="right" VALIGN="top" WIDTH="2%" HEIGHT="2"><IMG SRC="../security/images/y02.jpg" WIDTH="12" HEIGHT="12"></TD>
            </TR>
            <TR BGCOLOR="#FFD553"> 
              <TD WIDTH="2%">&nbsp;</TD>
              <TD WIDTH="96%" ALIGN="left" VALIGN="top" BGCOLOR="#FFD553">
                <DIV ALIGN="CENTER">
                <TABLE WIDTH="250" BORDER="0" CELLSPACING="1" CELLPADDING="0" HEIGHT="41">
                  <TR> 
                    <TD COLSPAN="2"> 
                      <DIV ALIGN="RIGHT"> 
<!--
                        <SELECT NAME="ListVender">
                          <OPTION VALUE="0" selected>&lt;-- Search List --&gt;</OPTION>
                          <%
							sSQL = "SELECT  * FROM pexpenstype  ORDER BY sexptype "
 							Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
							Do While Not RS.EOF	%> 
                          <OPTION VALUE="<%=RS("sexptype")%>"><%=RS("sexpdesc")%></OPTION>
                          <%
							RS.MoveNext
							Loop
							RS.Close
							Set RS=Nothing
							%> 
                        </SELECT>
-->
                      </DIV>
                    </TD>
                  </TR>
                  <TR> 
                    <TD> 
                      <DIV ALIGN="RIGHT"> 
                        <INPUT TYPE="text" NAME="SearchVender" SIZE="20" MAXLENGTH="20">
                      </DIV>
                    </TD>
                    <TD> 
                      <INPUT TYPE="submit" NAME="Submit2" VALUE="Go !!">
                      <INPUT TYPE="hidden" NAME="sProdcode" VALUE='<%=Request("sProdcode")%>'>
                      <INPUT TYPE="hidden" NAME="sExptype" VALUE='<%=sExptype2%>'>
                      <INPUT TYPE="hidden" NAME="nItem" VALUE='<%=nItem%>'>
                      <INPUT TYPE="hidden" NAME="EVENT" VALUE="Search">
                    </TD>
                  </TR>
                </TABLE>
                  <BR>
                  <FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2"><A HREF="javascript:history.back()">back</A> 
                  | <A HREF="Javascript:window.close()"> Close </A></FONT><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="1"><BR>
                  </FONT></DIV>
              </TD>
              <TD WIDTH="2%">&nbsp;</TD>
            </TR>
            <TR BGCOLOR="#FFD553"> 
              <TD WIDTH="2%" ALIGN="left" VALIGN="bottom"><IMG SRC="../security/images/y03.jpg" WIDTH="12" HEIGHT="12"></TD>
              <TD WIDTH="96%" BGCOLOR="#FFD553" VALIGN="middle" ALIGN="center"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="1"> &nbsp;&nbsp;
                </FONT></TD>
              <TD WIDTH="2%" ALIGN="right" VALIGN="bottom"><IMG SRC="../security/images/y04.jpg" WIDTH="12" HEIGHT="12"></TD>
            </TR>
          </TABLE>
			</TD>
      </TR>
    </TABLE>
    <BR>
  </DIV>
</FORM>
<BR>
<% End Function%> 
<% Function ShowVender(sPrno) %> 
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
                <td colspan="4" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vender
                        :: <% ' Show_Pexptype(sExptype)%>  </font><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="3" COLOR="#666666"><B><FONT COLOR="#FFFFFF" SIZE="2"><%'=sExptype2%></FONT></B></FONT></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718" HEIGHT="5"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="27%" bgcolor="#8B2718" HEIGHT="5"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vendor 
                  Code</font></b></font></td>
                <td width="52%" bgcolor="#8B2718" HEIGHT="5"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vendor 
                  Name</font></b></font></td>
                <td BGCOLOR="#993300" WIDTH="16%" HEIGHT="5"> </td>
              </tr>
              <%If  RS.EOF then 
       		Response.Write  " <tr bgcolor='#F1B3A7'> " & vbCrLf
		       Response.Write  " <td colspan='3' valign='top'>ไม่มีรายการการที่ค้นหา"
	       	Response.Write  "</td></tr> "& vbCrLf
		else
		RS.Absolutepage=PageNo
				RowCount=1
				Do While Not RS.EOF and RowCount<=PageSize 
				Vendcode=RS("svendcode")
				if keywords="" then 
					sSQL2="Select * from pvender where svendcode='" & Vendcode & "'" &_
					"Order By svendcode"
				else
					sSQL2 = "SELECT  *   FROM pvender Where svendcode='" & Vendcode & "'" &_
					" AND (sTname LIKE '%" &  keywords & "%' " &_
					"OR sEname Like '%" & keywords & "%' )" &_
			 		" ORDER BY svendcode"
					End if

				Set Rs2 = Server.CreateObject("ADODB.RecordSet")
				Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
				If Not RS2.EOF then
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="5%" bgcolor="<%=Color%>" height="10"> 
                  <DIV ALIGN="Right"><%=((PageNo-1)*RS.PageSize)+RowCount%>.&nbsp;</DIV>
                </td>
                <td width="27%" bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;<%=RS2("svendcode")%>&nbsp;</td>
                <td width="52%" bgcolor="<%=Color%>" height="10"> 
                  <a href=../Report/sPTaxReport.asp?sVendercode=<%=RS2("sVendcode")%>&sProdcode=<%=sProdcode%>&sExptype=<%=sExptype2%>&Mode=Search' Target="num"><%=RS2("sEname")%></a>	
                </td>
                <td WIDTH="16%" HEIGHT="10">
                  <DIV ALIGN="CENTER"><a href="VenderDetail.asp?svendcode=<%= Rs2("svendcode")%>">Show 
                    Detail</a></DIV>
                </td>
              </tr>
              <%
				End if
				if Color="#F8D7D1" then
							Color="#F1B3A7"
				else
							Color="#F8D7D1"
				End if
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

<%
if TotalPage > 1 then
%>
			<form action= "<%=sScriptPath%>" method="Post">
			<tr>
            <td colspan="4" align="center" bgcolor="#FFD553" HEIGHT="59"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif" WIDTH="8" HEIGHT="8"></td>
						<td align="left" bgcolor="#FFD553" width="45%">
<%
			if PageNo=1 then
					response.write " "
			Else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&ListVender=<%=sExptype%>&SearchVender=<%=keywords%>&page=Prev&sProdcode=<%=sProdcode%>&sExptype2=<%=sExptype2%>&Mode=Search&EVENT=Search">Previous Page</a>
<%
			End if
%>
</td>

<td align="right" bgcolor="#FFD553" width="45%">
<%	
		If PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
		Else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&ListVender=<%=sExptype%>&SearchVender=<%=keywords%>&page=Next&sProdcode=<%=sProdcode%>&sExptype2=<%=sExptype2%>&Mode=Search&EVENT=Search">Next Page</a>
<%End If%>
</td>
	<td align="center" bgcolor="#FFD553" width="2%">
		<img src="images/dot.gif">
	</td>
</tr>
<tr>
<td align="center" bgcolor="#FFD553" colspan="4">
<%
	For i=1 to TotalPage
		If Cint(PageNo)=Cint(i) then
%>
<Font color="Red"><%=i%></Font>
<%Else%>
<A Href="<%=sScriptPath%>?PageNo=<%=i%>&ListVender=<%=sExptype%>&SearchVender=<%=keywords%>&MaxPage=<%=TotalPage%>&sProdcode=<%=sProdcode%>&sExptype2=<%=sExptype2%>&Mode=Search&EVENT=Search"><%=i%></a>
<%
		End if
	Next 
%></td>
</tr>
</table>
<%End if%>

			</td></tr>
			</form>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center"><br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
            | <a href="Javascript:window.close()"> Close </a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
            <br>
            </font></td>
          <td width="2%" align="right" valign="bottom"><img src="../security/images/y04.jpg" width="12" height="12"></td>
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
<% End Function%>
</body>
</html>

