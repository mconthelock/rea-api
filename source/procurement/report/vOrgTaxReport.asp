<% Option Explicit %>
<%
           ' Author : Ekkapan Kongchoo
%>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%

%>
<%
	Dim RS,sSQL
	Dim sExptype,From_Day,From_Month,From_Year,To_Day,To_Month,To_Year,FromDate,ToDate
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage,color,sScriptPath
	Dim sDiv,sDep,sSec,sOrg,sCashReq,sUrgent,sOrg1
	Color="#FFCCCC"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
		sExptype=Request("sExptype")
		sCashReq=Request("sCashreq")
		sUrgent=Request("sUrgent")
		sDiv=Request("ListDivision")
		sDep=Request("ListDepartment")
		sSec=Request("ListSection")

	if sDiv="0" and sDep="0" and sSec= "0" then 
		sOrg1="0"
		sOrg=""
	elseif sDiv<>"0" and sDep<>"0" and sSec<> "0" then
		sOrg=" and sDivcode='" & sDiv & "' "
		sOrg=sOrg & " and sDepcode='" & sDep & "' "
		sOrg=sOrg & " and sSeccode='" & sSec & "' "
	elseif sDiv<>"0" and sDep<>"0" and sSec= "0" then 
		sOrg=" and sDivcode='" & sDiv & "' "
		sOrg=sOrg & " and sDepcode='" & sDep & "' "
	elseif sDiv<>"0" and sDep="0" and sSec= "0" then 
		sOrg=" and sDivcode='" & sDiv & "' "
	else
		sOrg=" and sDivcode='" & sDiv & "' "
		sOrg=sOrg & " and sDepcode='" & sDep & "' "
		sOrg=sOrg & " and sSeccode='" & sSec & "' "
	end if
			
		From_Day=Request("From_Day")
		From_month=Request("From_Month")
		From_Year=Request("From_Year")
''		FromDate=From_Month & "/" & From_Day & "/" & From_Year
		FromDate=From_Day & "/" & From_Month & "/" & From_Year
		To_Day=Request("To_Day")
		To_month=Request("To_Month")
		To_Year=Request("To_Year")
''		ToDate=To_Month & "/" & To_Day & "/" & To_Year
		ToDate=To_Day & "/" & To_Month & "/" & To_Year
		Set RS=Server.CreateObject("ADODB.Recordset")
if sExptype="0" then
	if sCashreq<>"1" and sUrgent <> "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq<>"1" and sUrgent<> "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq<>"1" and sUrgent= "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq="1" and sUrgent<> "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq="1" and sUrgent= "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq<>"1" and sUrgent= "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq="1" and sUrgent<> "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq="1" and sUrgent= "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" and sCashReq<>' '" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	end if
else
	if sCashreq<>"1" and sUrgent <> "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq<>"1" and sUrgent<> "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq<>"1" and sUrgent= "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq="1" and sUrgent<> "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq="1" and sUrgent= "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" and cprtype='1'" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq<>"1" and sUrgent= "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq="1" and sUrgent<> "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	elseif sCashReq="1" and sUrgent= "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,sTaxno,ddate,sVendTax,sRecno,spono From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" and sExptype='" & sExptype & "'" &_
		" and sCashReq<>' '" &_
		" Group By sCurrency,sTaxno,ddate,sVendTax,sRecno,spono" &_
		" order By ddate,sTaxno,scurrency,sVendTax"
	end if
End if

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
<HTML>
<HEAD>
<TITLE>Summary By Total</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		
  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top" width="1%" height="2"><img src="../images/ccb.gif" width="10" height="10"></td>
      <td width="99%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="8" height="8"></td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="99%" height="30" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="+1" color="#666666"><b><b><b><b><b><b><b><font color="#BF00BF">Report 
        Manager</font></b></b></b></b></b></b></b></font></td>
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
                          
                        <div align="center"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Summary 
                          Report Overview By Organization(Tax Invoice)</font></font></b></div>
                        </td>
                      </tr>
                    </table>
					</td>
                </tr>

				<tr>
					<td colspan="2" bgcolor="silver">	
						<font color="blue">
						<b>Tax Report From</b> <%=Show_date(fromdate)%> - <%=Show_date(todate)%>
					</td>
					</td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="gray">	
						<font color="white">
						<b>Tax Report in :</b> 
							<%ShowDivision(sDiv)%> 
						<%if sDep<> "0" then response.write "-->"%>
							<%ShowDepartment(sDep)%>
						<%if sSec<>"0" then response.write "-->"%>
							<%ShowSection(sSec)%>
					</td>
					</td>
				</tr>
				<tr>
					<td bgcolor="silver" colspan="2">	
						<font color="blue">
						<b>Expense Type :</b> <%Show_pExptype(sExptype)%>
					</td>
					</td>
				</tr>
				<tr>
					<td bgcolor="gray" colspan="2">	
						<font color="white"><b>Status :</b>
						<%	If sCashReq="1" then 
								Response.write "Cash Requirement" 
							else
								Response.write "No Cash Requirement"
							End if
							If sUrgent="1" then 
								Response.write " : Urgent Requisition" 
							else
								Response.write " : No Urgent Requisition"
							End if%>
					</td>
					</td>
				</tr>

                <tr align="center" valign="top"> 
                  <td colspan="2" bgcolor="#FFD553"> 
                    <table border=0 width="100%" align="Center" cellspacing="1" cellpadding="0" height="18">
                      <tr bgcolor="#CCCCFF">
                        <td height="20" align="center" width="5%"> No.</td>
                        
                      <td height="20" align="center" width="19%">Tax No.</td>
                      <td height="20" align="center" width="15%">Receive No.</td>
                        <td height="20" align="center" width="10%">Date</td>
                        <td height="20" width="15%" align="center">Amount</td>
                        <td height="20" width="24%" align="center">Currency</td>
                      </tr>
					<%
					If Not Rs.EOF then
						RS.Absolutepage=PageNo
						RowCount=1
						Do While Not Rs.EOF and RowCount<=PageSize%>
                      <tr bgcolor="#FFCCCC"> 
               			<td width="5%" bgcolor="<%=Color%>" height="2"><a href="pTaxdesc.asp?sTaxno=<%=RS("sTaxno")%>&EVENT=Show"><%=((PageNo-1)*RS.PageSize)+RowCount%></a>&nbsp;</td>
                        <td width="19%" height="20" bgcolor="<%=Color%>"><a href="pTaxdesc.asp?sTaxno=<%=RS("sTaxno")%>&EVENT=Show"><%=Rs("sVendTax")%></a></td>
                        <td height="20" width="15%" bgcolor="<%=Color%>"><%Show_Rec(RS("sRecno"))%></td>
                        <td height="20" width="10%" bgcolor="<%=Color%>"><%=RS("ddate")%></td>
                        <td width="15%" height="20" bgcolor="<%=Color%>" align="right"><%=FormatNumber(RS("SumAmount"), 4)%></td>
                        <td width="24%" height="20" bgcolor="<%=Color%>">&nbsp;<%=RS("sCurrency")%></td>
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
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&From_day=<%=From_Day%>&From_month=<%=From_Month%>&From_Year=<%=From_Year%>&to_day=<%=To_Day%>&to_Month=<%=To_month%>&to_Year=<%=To_year%>&sExptype=<%=sExptype%>&ListDivision=<%=sDiv%>&ListDepartment=<%=sDep%>&ListSection=<%=sSec%>&sCashReq=<%=sCashReq%>&sUrgent=<%=sUrgent%>">Previous 
                    Page</a> <%End if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> 
			<% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&From_day=<%=From_Day%>&From_month=<%=From_Month%>&From_Year=<%=From_Year%>&to_day=<%=To_Day%>&to_Month=<%=To_month%>&to_Year=<%=To_year%>&sExptype=<%=sExptype%>&ListDivision=<%=sDiv%>&ListDepartment=<%=sDep%>&ListSection=<%=sSec%>&sCashReq=<%=sCashReq%>&sUrgent=<%=sUrgent%>">Next Page</a> <%End if%> 
                  </td>
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../images/dot.gif"></td>
                </tr>
                <tr> 
                  <td align="center" bgcolor="#FFD553" colspan="4"> <%
	For i=1 to TotalPage
		If Cint(PageNo)=Cint(i) Then
%> <Font color="Red"><%=i%></font> 
                    <%else%> 
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>&From_day=<%=From_Day%>&From_month=<%=From_Month%>&From_Year=<%=From_Year%>&to_day=<%=To_Day%>&to_Month=<%=To_month%>&to_Year=<%=To_year%>&sExptype=<%=sExptype%>&ListDivision=<%=sDiv%>&ListDepartment=<%=sDep%>&ListSection=<%=sSec%>&sCashReq=<%=sCashReq%>&sUrgent=<%=sUrgent%>"><%=i%></a> 
                    <%
	End if
	Next %></td>
                </tr>
              </table>
              <%End if%> 
            </td>
          </tr>
        </form>
                <tr align="center" valign="top"> 
                  <td width="98%" bgcolor="#FFD553">&nbsp;<%SumOrgTaxAmount%></td>
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
        <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
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
<%
Sub SumOrgTaxAmount()
Dim sSQL2,RS2
		Set RS2=Server.CreateObject("ADODB.Recordset")
if sExptype="0" then
	if sCashreq<>"1" and sUrgent <> "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq<>"1" and sUrgent<> "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq<>"1" and sUrgent= "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq="1" and sUrgent<> "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq="1" and sUrgent= "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq<>"1" and sUrgent= "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq="1" and sUrgent<> "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq="1" and sUrgent= "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" and sCashReq<>' '" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	end if
else
	if sCashreq<>"1" and sUrgent <> "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq<>"1" and sUrgent<> "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq<>"1" and sUrgent= "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq="1" and sUrgent<> "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq="1" and sUrgent= "1" and sOrg<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" and cprtype='1'" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq<>"1" and sUrgent= "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq="1" and sUrgent<> "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and sCashReq<>' '" &_
		" and sExptype='" & sExptype & "'" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	elseif sCashReq="1" and sUrgent= "1" and sOrg1="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vpTax where ddate >=" & To_Date(FromDate) &_
		" and ddate <=" & To_Date(ToDate) & sOrg &_
		" and cprtype='1'" &_
		" and sExptype='" & sExptype & "'" &_
		" and sCashReq<>' '" &_
		" Group By sCurrency" &_
		" order By sCurrency"
	end if
End if
		Rs2.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
		if Not Rs2.EOF then
%>
		<table bgcolor="silver" width="60%" border="0"> 
            <tr align="center" valign="top"> 
               <td bgcolor="teal"  colspan="2"><b><font color="red">Total</font></b></td>
             </tr>
<%			Do While Not RS2.EOF
%>
			<tr>
				<td align="right" width="50%"><font color="blue"><%=FormatNumber(RS2("SumAmount"), 4)%></font></td>
				<td>&nbsp;<%=RS2("sCurrency")%></td>
			</tr>
<%			RS2.MoveNext
			Loop
%>
		</table>
<%
		End if
		Rs2.Close
	
End Sub

Function Show_Rec(sRecno)
	Dim RS2,sSQL2
	Set RS2=Server.CreateObject("ADODB.Recordset")
	sSQL2="Select srecby from preceive where srecno='" & sRecno & "'"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	if Not RS2.EOF then
		response.write RS2("sRecby")
	End if
	Set RS2=Nothing
End Function
%>
