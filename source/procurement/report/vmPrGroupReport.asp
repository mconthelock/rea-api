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
	Dim sumAmount,svendcode,scashreq,sCattype,sCatalog,exp,surgent,exp1
	Color="#FFCCCC"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
		sExptype=Request("ListProduct")
		sCattype=Request("ListCattype")
		sCatalog=Request("ListCatalog")
		sVendcode=Request("vender")
		if Trim(Request("Vender"))="" then sVendcode="0"
		sCashreq=Request("scashreq")
		sUrgent=Request("sUrgent")

		From_Day="1"
		From_month=Request("mFrom_Month")
		From_Year=Request("mFrom_Year")
''		FromDate=From_Month & "/" & From_Day & "/" & From_Year
		FromDate=From_Day & "/" & From_Month & "/" & From_Year
		To_month=Request("mTo_Month")
		To_Year=Request("mTo_Year")
		To_Day=MonthDay(To_Month,To_Year)
''		ToDate=To_Month & "/" & To_Day & "/" & To_Year
		ToDate=To_Day & "/" & To_Month & "/" & To_Year
	if sExptype="0" and sCattype="0" and sCatalog="0" then 
		exp1=0
		exp=""
	elseif sExptype<>"0" and sCattype<>"0" and sCatalog<>"0" then 
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		exp=exp & " and sCatID='" & sCatalog & "'"
	elseif sExptype<>"0" and sCattype<>"0" and sCatalog="0" then 
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
	elseif sExptype<>"0" and sCattype="0" and sCatalog="0" then 
		exp=" sExptype='" & sExptype & "'"
	else
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		exp=exp & " and sCatID='" & sCatalog & "'"
	end if

		Set RS=Server.CreateObject("ADODB.Recordset")
if sCashreq<>"1" and sUrgent<> "1" then
	if Exp1="0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp<>"0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where " & exp &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp <>"0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where " & exp &_
		" and sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	end if
elseif sCashReq="1" and sUrgent<>"1" then
	if Exp1="0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp<>"0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where " & exp &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp <>"0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where " & exp &_
		" and sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	end if
elseif sCashReq="1" and sUrgent="1" then
	if Exp1="0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp<>"0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where " & exp &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp <>"0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where " & exp &_
		" and sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	end if
elseif sCashReq<>"1" and sUrgent="1" then
	if Exp1="0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and cprtype='1'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp<>"0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where " & exp &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and cprtype='1'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and cprtype='1'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	elseif Exp <>"0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount,dprdate From vPRGroup where " & exp &_
		" and sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and cprtype='1'" &_
		" Group By sCurrency,dprdate" &_
		" order By dprdate,scurrency"
	end if
End if

'		" and cpo='3'" &_
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
'response.write sSQL
		Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	TotalRows=Rs.RecordCount	
	RS.PageSize = 20
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
                            Report Overview By PR</font></font></b></div>
                        </td>
                      </tr>
                    </table>
					</td>
                </tr>

				<tr>
					<td colspan="2" bgcolor="silver">	
						<font color="blue">
						<b>P/R Report From</b> <%=Show_date(fromdate)%> - <%=Show_date(todate)%>
					</td>
					</td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="gray">	
						<font color="white">
						<b>P/R Report in :</b> 
							<%Show_pExptype(sExptype)%> 
						<%if sCattype<> "0" then response.write "-->"%>
							<%ShowCattype(sCAttype)%>
						<%if sCatalog<>"0" then response.write "-->"%>
							<%ShowCatalog(sCatalog)%>
					</td>
					</td>
				</tr>
				<tr>
					<td bgcolor="silver" colspan="2">	
						<font color="blue">
						<b>Vendor :</b> <%Show_Vender(sVendcode)%>
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
                        <td height="20" align="center" width="20%">Date</td>
                        <td height="20" width="20%" align="center">Amount</td>
                        <td height="20" width="36%" align="center">Currency</td>
                      </tr>
					<%
					If Not Rs.EOF then
						RS.Absolutepage=PageNo
						RowCount=1
						Do While Not Rs.EOF and RowCount<=PageSize%>
                      <tr bgcolor="#FFCCCC"> 
               			<td width="5%" bgcolor="<%=Color%>" height="2"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                        <td height="20" width="28%" bgcolor="<%=Color%>"><%=FormatDatetime(RS("dprdate"))%></td>
                        <td width="24%" height="20" bgcolor="<%=Color%>" align="right"><%=FormatNumber(RS("SumAmount"), 4)%></td>
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
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&From_day=<%=From_Day%>&mFrom_month=<%=From_Month%>&mFrom_Year=<%=From_Year%>&to_day=<%=To_Day%>&mto_Month=<%=To_month%>&mto_Year=<%=To_year%>&ListProduct=<%=sExptype%>&ListCattype=<%=sCattype%>&ListCatalog=<%=sCatalog%>&sCashReq=<%=sCashreq%>&sUrgent=<%=sUrgent%>&vender=<%=sVendcode%>">Previous 
                    Page</a> <%End if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> 
			<% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&From_day=<%=From_Day%>&mFrom_month=<%=From_Month%>&mFrom_Year=<%=From_Year%>&to_day=<%=To_Day%>&mto_Month=<%=To_month%>&mto_Year=<%=To_year%>&ListProduct=<%=sExptype%>&ListCattype=<%=sCattype%>&ListCatalog=<%=sCatalog%>&sCashReq=<%=sCashreq%>&sUrgent=<%=sUrgent%>&vender=<%=sVendcode%>">Next Page</a> <%End if%> 
                  </td>
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../images/dot.gif"></td>
                </tr>
                <tr> 
                  <td align="center" bgcolor="#FFD553" colspan="4"> <%
	For i=1 to TotalPage
		If Cint(PageNo)=Cint(i) Then
%> <Font color="Red"><%=i%></font> 
                    <%else%> 
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>&From_day=<%=From_Day%>&mFrom_month=<%=From_Month%>&mFrom_Year=<%=From_Year%>&to_day=<%=To_Day%>&mto_Month=<%=To_month%>&mto_Year=<%=To_year%>&ListProduct=<%=sExptype%>&ListCattype=<%=sCattype%>&ListCatalog=<%=sCatalog%>&sCashReq=<%=sCashreq%>&sUrgent=<%=sUrgent%>&vender=<%=sVendcode%>"><%=i%></a> 
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
                  <td width="98%" bgcolor="#FFD553">&nbsp;<%SumPrAmount%></td>
                  <td width="2%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/dot.gif"></td>
                </tr>
                <tr align="center" valign="top"> 
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
Sub SumPrAmount()
Dim sSQL2,RS2
		Set RS2=Server.CreateObject("ADODB.Recordset")
if sCashreq<>"1" and sUrgent<> "1" then
	if Exp1="0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp<>"0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where " & exp &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp <>"0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where " & exp &_
		" and sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	end if
elseif sCashReq="1" and sUrgent<>"1" then
	if Exp1="0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp<>"0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where " & exp &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp <>"0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where " & exp &_
		" and sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" Group By sCurrency" &_
		" order By scurrency"
	end if
elseif sCashReq="1" and sUrgent="1" then
	if Exp1="0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp<>"0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where " & exp &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp <>"0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where " & exp &_
		" and sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and sCashreq<>' '" &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	end if
elseif sCashReq<>"1" and sUrgent="1" then
	if Exp1="0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp<>"0" and sVendcode="0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where " & exp &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By scurrency"
	elseif Exp <>"0" and sVendcode<>"0" then
		sSQL="Select sCurrency,Sum(nAmount) as SumAmount From vPRGroup where " & exp &_
		" and sVendcode='" & svendcode & "'" &_
		" and dprdate >=" & To_Date(FromDate) &_
		" and dprdate <=" & To_Date(ToDate) &_
		" and cstatus<>'0'" &_
		" and cprtype='1'" &_
		" Group By sCurrency" &_
		" order By scurrency"
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
				<td align="right" width="50%"><font color="blue"><%=Formatnumber(RS2("SumAmount"), 4)%></font></td>
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
%>
