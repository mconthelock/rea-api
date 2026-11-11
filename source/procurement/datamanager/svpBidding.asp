<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	'########## DERMPAN KHOOMRAKSA ##########
	'I change formatNumber(xxx, y) to formatNumber(xxx, 3) on 11-NOV-2002
	'########## END DERMPAN KHOOMRAKSA ##########
	
	' Author: Ekkapan Kongchoo

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,RS2,status,RS3
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sedesc,nprice,simage,dadddate,sempno
	Dim img_sprodcode	,img_sexptype,img_svendcode,img_nprice,sbread,sItemno
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year ,sErrImage,chkGroup
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage,color,sCattype,sCatalog
	Dim sCatID,keywords,sprno,nitem,sExptype2,chk
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sprodcode=Request.QueryString("sprodcode")
	sexptype=Request.QueryString("sexptype")
	svendcode=Request.QueryString("svendcode")
'	scorpcode = Session("gbcorpcode")
	status = Request("status")	
	sitemno = Request("sItemno")
	
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD

 'sSQL = "SELECT * FROM PExpenstype where sexptype='" & sexptype & "'" &_
	'" Order By sExptype"
	
sSQL ="SELECT * FROM WEBFORM.VIEW_BIDDING ORDER BY CYEAR2, NRUNNO"

   Set Rs = Server.CreateObject("ADODB.RecordSet")
   Set Rs2 = Server.CreateObject("ADODB.RecordSet")
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
  
  'if Not Rs.EOF then
	'Do While Not RS.EOF 
		'Response.write rs.fields("PROJECT")
		'rs.movenext
	'loop
  'end if

  'Response.write rs.fields("PROJECT")
  'Response.end


%>

 <html> <head> 
<title>Show Selection </title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
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
                  Selection No </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2" height="76"> 
                  <table width="98%" border="0" cellspacing="1" cellpadding="2">
				  <tr bgcolor="#1E90FF" align="center" valign="top">
					<td width="5%" >No.</td>
					<td width="20%" >Selection No</td>
					<td width="30%" >Project</td>
					<td width="10%" >Requester No</td>
					<td width="35%" >Requester Name</td>
				  </tr>
<%		
					sSQL ="SELECT * FROM WEBFORM.VIEW_BIDDING ORDER BY CYEAR2, NRUNNO"
					dim j
					j =1
					 Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
						IF Not RS2.EOF then
							Do While Not Rs2.EOF
%>
		  	            <tr>
							<td>
								<% response.write j%>
							</td>
  		                    <td>
								<a href="../workflow/showprdesc.asp?biddingno=<%=RS2.fields("biddingno")%>&sprno=<%=Request("sPrno")%>&nitem=<%=Request("nItem")%>&status=<%=Request("status")%>&sExptype=<%=sExptype%>"  Target="num" ><%=RS2.fields("biddingno")%></a>
							</td>
  		                    <td>
								<% response.write RS2.fields("project")%>
							</td>
							<td>
								<% response.write RS2.fields("sempno")%>
							</td>
  		                    <td>
								<% response.write RS2.fields("sname")%>
							</td>
  		                  </tr>
						  
<%							rs2.movenext
							j = j+1
							 Loop
							chk=1
						else
							sHow_Product()
							chk=0
						End if

						RS2.Close
			if chk=1 then
%>
				</table>
<%end if%>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;</td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
<!--              | <a href="workflow.asp">Workflow Manager</a> -->
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
Sub Show_Product()
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
'	Dim sMode,sPrno,nItem,sExptype2,status,sItemno
	Color="#F8D7D1"
	'suserid = Request("uid")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file 

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs3 = Server.CreateObject("ADODB.RecordSet")       	
	'sExptype=Request("ListProduct")
	status = Request("status")
	sExptype=Request("sExptype")
	scattype=Request("ListCattype")
	scatalog=Request("ListCatalog")
	sCatID=Request("sCatID")
	MaxPage=Request("MaxPage")
	sItemno=Request("sItemno")
	'sSearchProduct=Request("SearchProduct")
	keywords = TRIM( Request( "SearchProduct" ))
	sPrno = Request("sPrno")
	nItem = Request("nItem")
	sExptype2=Request("sExptype")
'	If IsEmpty(Request("EVENT")) Then
'		If  sExptype <> "" Then
'			sMode = "Search"
'		Else
'			sMode = ""
'		End If	
'		SearchProduct()
'	Else
'			Get_Product()
			ShowProduct()
			
'	End If	
End Sub
Sub ShowProduct()
if keywords="" then
	sSQL = "SELECT  * FROM pproduct Where sExptype='" & sExptype & "'" &_
		" and (sTprodname LIKE '%" &  keywords & "%' " &_
		" OR sEprodname Like '%" & keywords & "%' " &_
		" OR sprodcode Like '%" & keywords & "%' " &_
 		") ORDER BY sprodcode"
		sSQL ="SELECT * FROM WEBFORM.VIEW_BIDDING ORDER BY CYEAR2, NRUNNO"
else
	sSQL = "SELECT  * FROM pproduct Where sTprodname LIKE '%" &  keywords & "%' " &_
		" OR sEprodname Like '%" & keywords & "%' " &_
		" OR sprodcode Like '%" & keywords & "%' " &_
 		" ORDER BY sprodcode"
		sSQL ="SELECT * FROM WEBFORM.VIEW_BIDDING ORDER BY CYEAR2, NRUNNO"
end if

if Request("EVENT")="Show" then
	sSQL="Select * from pproduct where sExptype='" & sExptype & "'"
End if

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
	page=Request("page")
	gPage=Request("PageNo")
 	Rs3.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	TotalRows=Rs3.RecordCount	
	RS3.PageSize = 10
	PageSize=RS3.PageSize
	TotalPage=Rs3.PageCount
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

%>

      <table width="100%" border="0" cellspacing="0" cellpadding="0">

              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718" HEIGHT="2"> 
                  <DIV ALIGN="CENTER"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></DIV>
                </td>
                <td width="15%" bgcolor="#8B2718" HEIGHT="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">&nbsp;&nbsp;Selection No</font></b></font></td>
                <td width="49%" bgcolor="#8B2718" HEIGHT="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">&nbsp;&nbsp;request No</font></b></font></td>
                <td align="center" width="12%" bgcolor="#8B2718" HEIGHT="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Request Name</font></b></font></td>
              </tr>
              <%If  RS3.EOF then 
       		Response.Write  " <tr bgcolor='#F1B3A7'> " & vbCrLf
		       Response.Write  " <td colspan='5' valign='top'>ไม่มีรายการการที่ค้นหา"
	       	Response.Write  "</td></tr> "& vbCrLf
		else
				RS3.Absolutepage=PageNo
				RowCount=1
				Do While Not RS3.EOF and RowCount<=PageSize %> 
              <tr bgcolor="<%=Color%>"> 
                <td width="5%" bgcolor="<%=Color%>" height="10">
					<DIV ALIGN="RIGHT">
						<%=((PageNo-1)*RS3.PageSize)+RowCount%>.&nbsp;
                  </DIV>
                </td>
                <td width="15%" bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;<%=RS3("sprodcode")%>&nbsp;</td>
                <td width="49%" bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;
				<%if Request("status")="change" then %> 
					<a href=../workflow/showpr.asp?nItem=<%=nItem%>&sExptype=<%=Rs3("sExptype")%>&sPrno=<%=sPrno%>&sProdcode=<%=RS3("sProdcode")%>&nprice=<%=RS3("nprice")%>&status=change&sItemno=<%=Request("sItemno")%>&smode=<%=Request("smode")%>&approve=<%=Request("approve")%> Target="txt_bid"><%=RS3("seprodname")%></a>
				<%else%>
					<a href=../workflow/ppritemedt.asp?nItem=<%=nItem%>&sExptype=<%=Rs3("sExptype")%>&sPrno=<%=sPrno%>&sProdcode=<%=RS3("sProdcode")%>&nprice=<%=RS3("nprice")%>&status=change&sItemno=<%=Request("sItemno")%>&smode=<%=Request("smode")%>&approve=<%=Request("approve")%> Target="num"><%=RS3("seprodname")%></a>
				<%end if%>
				</td>

              </tr>
              <%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
			RowCount=RowCount+1
		 	RS3.MoveNext
			Loop
	end if
	Rs3.close
	set Rs3=nothing
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
<%End if
End Sub

%>
