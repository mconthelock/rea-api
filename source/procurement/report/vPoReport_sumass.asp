<% Option Explicit %>
<%
           ' Author : Ekkapan Kongchoo
%>
<HTML>
<HEAD>
<TITLE>Summary By Total</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	Dim RS,sSQL
	Dim sExptype,From_Day,From_Month,From_Year,To_Day,To_Month,To_Year,FromDate,ToDate,scatid,scattype
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage,color,sScriptPath
	'Dim sDiv,sDep,sSec,sOrg,sCashReq,sUrgent,sOrg1
	Dim sumbht, sumyen, sumusd, sumgbp
	Dim sSQL2,RS2, whereClause
	Dim cntLine
	cntLine = 1
	sumbht = 0
	sumyen = 0
	sumusd = 0
	sumgbp = 0
	Set RS2=Server.CreateObject("ADODB.Recordset")

	Color="#FFCCCC"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	sExptype=Request("sExptype")
	'sDiv=Request("ListDivision")
	'sDep=Request("ListDepartment")
	'sSec=Request("ListSection")
    scatid=Request("scatid")
	scattype=Request("scattype")
	'datetype=Request("datetype")
	'response.write "[" & scatid & "][" & scattype & "]"
	'if sDiv <> "0" then 
	'	sOrg=sOrg & " and sDivcode='" & sDiv & "' "
	'end if
	'if sDep <> "0" then
	'	sOrg=sOrg & " and sDepcode='" & sDep & "' "
	'end if
	'if  sSec<> "0" then
	'	sOrg=sOrg & " and sSeccode='" &  sSec & "' "	
	'end if
	
	'if sExptype <> "0" then
	'	whereClause = " and ppr.sexptype = '" & sExptype & "'"
	'end if
	'if request("txtPrNo") <> "" then
	'	whereClause = " and ppr.sprNo like '%" & request("txtPrNo") & "%'"
	'end if
	'if request("txtPoNo") <> "" then
	'	whereClause = " and ppo.spoNo like '%" & request("txtPoNo") & "%'"
	'end if
    if sExptype <> "0" then
		whereClause = " and pproduct.sexptype = '" & sExptype & "'"
	end if
	'if request("txtPrNo") <> "" then
	'	whereClause = whereClause & " and ppr.sprNo like '%" & request("txtPrNo") & "%'"
	'end if
	'if request("txtPoNo") <> "" then
	'	whereClause = whereClause & " and ppo.spoNo like '%" & request("txtPoNo") & "%'"
	'end if
	if scatid <> "0" then
		whereClause = whereClause & " and pproduct.scatid = '" & scatid & "'"
	end if
	if scattype <> "0" then
		whereClause = whereClause & " and pproduct.scattype = '" & scattype & "'"
	end if	
	
	'response.write "[" & whereClause & "]"
	
	From_Day=Request("From_Day")
	From_month=Request("From_Month")
	From_Year=Request("From_Year")
''		FromDate=From_Month & "/" & From_Day & "/" & From_Year
	FromDate=From_Day & "/" & From_Month & "/" & From_Year
	To_Day=Request("To_Day")
	To_month=Request("To_Month")
	To_Year=Request("To_Year")
	
    ToDate=To_Day & "/" & To_Month & "/" & To_Year
	
	
	sSQL = "select pproduct.sexptype , pproduct.scattype , pproduct.scatid , pcatalog.scatalog , sum(ppoitem.namount) as amt ," _
         	& " ppoitem.scurcode from ppo , ppoitem , pproduct , pexpenstype , pcattype , pcatalog where ppo.spono = ppoitem.spono" _
			& " and ppoitem.sprodcode = pproduct.sprodcode and pproduct.sexptype = pexpenstype.sexptype" _
			& " and ppo.dpodate between to_date('" & FromDate & "','dd/mm/yyyy') " _
			& " and to_date('" & ToDate & "','dd/mm/yyyy') and " _
            & "	pproduct.scattype = pcattype.scattype and pcattype.sexptype = pproduct.sexptype and pcatalog.scatid = pproduct.scatid" _
			& " and pcatalog.scattype = pproduct.scattype "& whereClause &"group by  pproduct.sexptype , pproduct.scattype" _
			& ", pproduct.scatid , pcatalog.scatalog , ppoitem.scurcode  "
'	response.write sSQL
'	response.end
	Set RS=Server.CreateObject("ADODB.Recordset")

	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")

	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	TotalRows=Rs.RecordCount	
	RS.PageSize = 30
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
			 If Cint(gPage)=Cint(TotalPage) then
					PageNo=TotalPage
			else
					PageNo=gPage+1
			end if
	End Select
%>
<script language="javascript">
	function expExc_onclick() {
//		alert("<%=FromDate%>");
		window.location.href="finRecExcell.asp?sSQL=<%=server.urlencode(sSQL)%>"
	}
</script>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
<tr> 
  <td align="left" valign="top" width="1%" height="2"><img src="../images/ccb.gif" width="10" height="10"></td>
  <td width="99%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="8" height="8"></td>
</tr>
<tr> 
  <td width="1%">&nbsp;</td>
  <td width="99%" height="30" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="+1" color="#666666"><b><font color="#BF00BF">Report 
    Manager</font></b></font></td>
</tr>
<tr> 
  <td width="1%"><img src="../images/spacer.gif" width="1" height="1"></td>
  <td width="99%" align="center" valign="top" height="0"> 
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td width="98%" height="2" align="left" valign="top" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../images/y01.jpg" width="12" height="12"></b></font></td>
      </tr>
      <tr> 
        <td width="98%" align="left" valign="top" height="0"> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr align="left" valign="top" bgcolor="#FFD553"> 
              <td colspan="2" align="center">&nbsp;</td>
            </tr>
            <tr align="left" valign="top" bgcolor="#FFD553"> 
              <td colspan="2" align="center"> 
                <table border=0 width="100%" align="Center" cellspacing="1" cellpadding="0" height="18">
                  <tr bgcolor="#CC00CC"> 
                    <td colspan="2" height="20"> 
                      <div align="center"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Summary Report by Asset and Non Asset</font></font></b></div>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <td colspan="2" bgcolor="silver"><font color="blue"><b>Report From</b> <%=Show_date(fromdate)%> - <%=Show_date(todate)%></font></td>
            </tr>
            <tr>
              <td colspan="2" bgcolor="gray">	
		<font color="white">
		<b>Report in :</b> 
	
		</font>						
              </td>
            </tr>
            <tr>
		<td bgcolor="silver" colspan="2">	
			<font color="blue">
			<b>Expense Type :</b> <%Show_pExptype(sExptype)%>
		</font>
		</td>
            </tr>
            <tr align="center" valign="top"> 
                  <td colspan="2" bgcolor="#003333"> 
                    <table border=0 width="100%" align="Center" cellspacing="1" cellpadding="0" height="18">
					  <tr bgcolor="#CCCCFF">
						<td colspan="4"></td>
						<td colspan="3" align="center"> Thai Baht (BHT)</td>
						<td colspan="3" align="center"> Japanese Yen</td>
						<td colspan="3" align="center"> American Dollar (USD)</td>
						<td colspan="3" align="center"> Great British Pound (GBP)</td>
					  </tr>
                      <tr bgcolor="#CCCCFF">
						<td height="20" align="center">Expense Type</td>        
						<td height="20" align="center">Group in Exp.</td>
						<td height="20" align="center">Sub Group in Exp.</td>
						<td height="20" align="center">Catalog</td>
						<td height="20" align="center">PO Amount Total</td>
						<td height="20" align="center">Asset Amount</td>
						<td height="20"  align="center">Non Asset Amount</td>
						<td height="20" align="center">PO Amount Total</td>
						<td height="20" align="center">Asset Amount</td>
						<td height="20"  align="center">Non Asset Amount</td>
						<td height="20" align="center">PO Amount Total</td>
						<td height="20" align="center">Asset Amount</td>
						<td height="20"  align="center">Non Asset Amount</td>
						<td height="20" align="center">PO Amount Total</td>
						<td height="20" align="center">Asset Amount</td>
						<td height="20"  align="center">Non Asset Amount</td>					
                      </tr>	
			<%
				If Not Rs.EOF then
					RS.Absolutepage=PageNo
					RowCount=1
				Do While Not Rs.EOF and RowCount<=PageSize
					
			%>
                      <tr bgcolor="#FFCCCC"> 
		        		<td bgcolor="<%=Color%>" height="2"><%=Rs("sexptype")%>&nbsp;</td>
                        <td height="20" bgcolor="<%=Color%>"><%=Rs("scattype")%></td>
                        <td height="20" bgcolor="<%=Color%>"><%=RS("scatid")%></td>
                        <td height="20" bgcolor="<%=Color%>"><%=RS("scatalog")%></td>
                        <td height="20" bgcolor="<%=Color%>" align="right"><% If RS("scurcode") = "Bt." then Response.write FormatNumber(RS("amt"), 3) end if %></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><% If RS("scurcode") = "Yen" then Response.write FormatNumber(RS("amt"), 3) end if %></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><% If RS("scurcode") = "US$" then Response.write FormatNumber(RS("amt"), 3) end if %></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><% If RS("scurcode") = "GBP" then Response.write FormatNumber(RS("amt"), 3) end if %></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
               </tr>
	    <%
				if Color="#FFCCCC" then
					Color="#99CCFF"
				else
					Color="#FFCCCC"
				end if
			RowCount=RowCount+1
			RS.MoveNext
			Loop 
		End if
		if PageNo = maxPage or TotalPage = 1 then
			sSQL = "select sum(ppoitem.namount) as amt ," _
         	& " ppoitem.scurcode from ppo , ppoitem , pproduct , pexpenstype , pcattype , pcatalog where ppo.spono = ppoitem.spono" _
			& " and ppoitem.sprodcode = pproduct.sprodcode and pproduct.sexptype = pexpenstype.sexptype" _
			& " and ppo.dpodate between to_date('" & FromDate & "','dd/mm/yyyy')" _
			& " and to_date('" & ToDate & "','dd/mm/yyyy') and" _
            & "	pproduct.scattype = pcattype.scattype and pcattype.sexptype = pproduct.sexptype and pcatalog.scatid = pproduct.scatid" _
			& " and pcatalog.scattype = pproduct.scattype "& whereClause &" group by ppoitem.scurcode"
			
		   Rs2.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
			Do While Not Rs2.EOF
				if Rs2("scurcode") = "Bt." then
					sumbht = Rs2("amt")
				end if
				if Rs2("scurcode") = "US$" then
					sumusd = Rs2("amt")
				end if
				if Rs2("scurcode") = "Yen" then
					sumyen = Rs2("amt")
				end if
				if Rs2("scurcode") = "GBP" then
					sumgbp = Rs2("amt")
				end if
			
			RS2.MoveNext
			Loop 
	%>
	
			<tr bgcolor="#CCCCFF"><td colspan="4" align="center">Total</td><td align="right"><% If sumbht <> "0" then Response.write FormatNumber(sumbht, 3) end if %></td><td></td><td></td><td align="right"><% If sumyen <> "0" then Response.write FormatNumber(sumyen, 3) end if %></td><td></td><td></td><td align="right"><% If sumusd <> "0" then Response.write FormatNumber(sumusd, 3) end if %></td><td></td><td></td><td align="right"><% If sumgbp <> "0" then Response.write FormatNumber(sumgbp, 3) end if %></td><td></td><td></td></tr>
	<%
		end if
	%>
                   
					</table>
                  </td>
                </tr>
        <%if TotalPage > 1 then%> 
        <form action= "<%=sScriptPath%>" method="Post">
          <tr> 
            <td colspan="2" align="center" bgcolor="#FFD553"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../images/dot.gif"></td>
                  <td align="left" bgcolor="#FFD553" width="45%"> <%
			If PageNo=1 then
				response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&From_day=<%=From_Day%>&From_month=<%=From_Month%>&From_Year=<%=From_Year%>&to_day=<%=To_Day%>&to_Month=<%=To_month%>&to_Year=<%=To_year%>&sExptype=<%=sExptype%>&scatid=<%=scatid%>&scattype=<%=scattype%>">Previous 
                    Page</a> <%End if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> 
			<% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&From_day=<%=From_Day%>&From_month=<%=From_Month%>&From_Year=<%=From_Year%>&to_day=<%=To_Day%>&to_Month=<%=To_month%>&to_Year=<%=To_year%>&sExptype=<%=sExptype%>&scatid=<%=scatid%>&scattype=<%=scattype%>">Next Page</a> <%End if%> 
                  </td>
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../images/dot.gif"></td>
                </tr>
                <tr> 
                  <td align="center" bgcolor="#FFD553" colspan="4"> <%
	For i=1 to TotalPage
		If Cint(PageNo)=Cint(i) Then
%> <Font color="Red"><%=i%></font> 
                    <%else%> 
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>&From_day=<%=From_Day%>&From_month=<%=From_Month%>&From_Year=<%=From_Year%>&to_day=<%=To_Day%>&to_Month=<%=To_month%>&to_Year=<%=To_year%>&sExptype=<%=sExptype%>&scatid=<%=scatid%>&scattype=<%=scattype%>"><%=i%></a> 
                    <%
	End if
	Next %></td>
                </tr>
              </table>
              </td>
          </tr>
        </form>
          <%End if%> 				    
			<tr align="left" valign="top"> 
			  <td width="98%" bgcolor="#FFD553">&nbsp;</td>
			  <td width="2%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/dot.gif"></td>
			</tr>		  
			<tr align="center" valign="top"> 
			  <td width="98%" bgcolor="#FFD553">&nbsp;<%'SumOrgTaxAmount%><input type="button" name="btnExport" value="Export Excel" onClick="expExc_onclick()"></td>
			  <td width="2%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/dot.gif"></td>
			</tr>
			<tr align="left" valign="top"> 
			  <td width="98%" bgcolor="#FFD553">&nbsp;</td>
			  <td width="2%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
			</tr>
          </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr align="center"> 
      <td width="1%"> <img src="../report/images/spacer.gif" width="1" height="1"><br>
      </td>
      <td width="99%"> 
        <!--#include file="../inc/copyRight.inc"--></td>
    </tr>
  </table>	
</BODY>
</HTML>
