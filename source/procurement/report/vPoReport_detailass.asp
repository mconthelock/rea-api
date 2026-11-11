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
	Dim sExptype,From_Day,From_Month,From_Year,To_Day,To_Month,To_Year,FromDate,ToDate,scatid,scattype,Product
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage,color,sScriptPath
	'Dim sDiv,sDep,sSec,sOrg,sCashReq,sUrgent,sOrg1
	Dim sSQL2,RS2, whereClause
	Dim cntLine
	cntLine = 1
	Set RS2=Server.CreateObject("ADODB.Recordset")

	Color="#FFCCCC"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	sExptype=Request("sExptype")
	'sDiv=Request("ListDivision")
	'sDep=Request("ListDepartment")
	'sSec=Request("ListSection")
    scatid=Request("scatid")
	scattype=Request("scattype")
	Product=Request("Product") 
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
	'response.write "[" & request("Product") & "]"
	if Product <> "" then
		whereClause = whereClause & " and pproduct.SPRODCODE LIKE '%" & Product & "%'"
	end if
    if sExptype <> "0" then
		whereClause = " and ppr.sexptype = '" & sExptype & "'"
	end if
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
	
	
	sSQL = "select ppr.sprno, ppr.dprdate, ppr.srequester, ppritem.namount pramt , ppoitem.namount poamt, ppoitem.scurcode, " _
         	& "pvender.sename, ppr.sexptype, substr(pcat.scatalog,1,instr(pcat.scatalog,')')) scatalog , pproduct.scatid, " _
			& "pproduct.scattype, ppo.spono, dpodate , pproduct.SEPRODNAME , pproduct.SPRODCODE , ppoitem.nquality , ppoitem.nprice from ppr, ppritem, pvender, pproduct," _
			& "ppo, ppoitem, aemployee a, pcatalog pcat where ppr.sprno = ppritem.sprno and pproduct.sprodcode = ppritem.sprodcode" _
			& " and ppritem.sprodcode = ppoitem.sprodcode and pproduct.sprodcode = ppoitem.sprodcode and ppritem.svendcode = pvender.svendcode" _
			& " and ppritem.svendcode = ppo.svendcode and ppritem.sprno = ppo.srefno and ppo.spono = ppoitem.spono and ppr.sexptype = pproduct.sexptype" _
			& " and ppritem.cstatus <> '0' and a.sempno = ppr.srequester and pproduct.scatid=pcat.scatid " & whereClause _
			& " and ppo.dpodate between to_date('" & FromDate & "','dd/mm/yyyy') " _
			& " and to_date('" & ToDate & "','dd/mm/yyyy') order by ppr.sprno, ppo.spono "
	
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
						<td height="20" align="center">Expense Type</td>        
						<td height="20" align="center">Group in Exp.</td>
						<td height="20" align="center">Sub Group in Exp.</td>
						<td height="20" align="center">Catalog</td>
						<td height="20" align="center">Requester</td>
						<td height="20" align="center">Budget Requisition Control No.</td>
						<td height="20"  align="center">PR No.</td>
						<td height="20" align="center">PR Date</td>
						<td height="20" align="center">PR Amount</td>
						<td height="20"  align="center">PO No.</td>
						<td height="20" align="center">PO Date</td>
						<td height="20" align="center">PO Amount</td>
						<td height="20"  align="center">Product Name</td>
						<td height="20" align="center">Quantity</td>
						<td height="20" align="center">Price</td>
						<td height="20"  align="center">Asset Amount</td>	
						<td height="20"  align="center">Non Asset Amount</td>
						<td height="20"  align="center">Currency</td>
						<td height="20"  align="center">Invoice No.</td>	
						<td height="20"  align="center">Received Date</td>
						<td height="20"  align="center">Amount(Invoice)</td>
						
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
                        <td height="20" bgcolor="<%=Color%>" align="right"><%=RS("srequester")%></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><%=RS("sprno")%></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><%=RS("dprdate")%></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><%=RS("pramt")%></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><%=RS("spono")%></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><%=RS("dpodate")%></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><%=FormatNumber(RS("poamt"), 3)%></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><%=RS("SEPRODNAME")%></td>
						
						<td height="20" bgcolor="<%=Color%>" align="right"><%=RS("nquality")%></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><%=RS("nprice")%></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"></td>
						<td height="20" bgcolor="<%=Color%>" align="right"><%=RS("scurcode")%></td>
				
			<%
			sSQL = "select srecby, to_date(drecdate,'dd/mm/yyyy') as drecdate, namount from preceive , preceiveitem , pproduct where  preceive.srecno = preceiveitem.srecno and preceiveitem.sprodcode = pproduct.sprodcode and preceive.spono = '" & RS("sPoNo") & "' and pproduct.scatid = '" & RS("sCatid") & "' and pproduct.scattype = '"& RS("sCattype") & "' and pproduct.sexptype = '"& RS("sExptype") &"' and preceiveitem.sprodcode = '" & RS("SPRODCODE") &"'" 
			'response.write sSQL
			Rs2.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
			if Not Rs2.EOF then
				Do While Not RS2.EOF
			%>
				<td bgcolor="<%=Color%>">&nbsp;<%=RS2("srecby")%></td>
				<td bgcolor="<%=Color%>">&nbsp;<%=RS2("drecdate")%></td>
				<td bgcolor="<%=Color%>" align="right" width="50%"><font color="blue"><%=FormatNumber(RS2("namount"), 3)%></font></td>		
			<%			
					RS2.MoveNext
					cntLine = cntLine +1 
					if (cntLine > 1) and (Not RS2.EoF) then
			%>
                      </tr>
                      <tr bgcolor="<%=Color%>">
                        <td colspan=18>&nbsp;</td>
			<%	
						end if
					Loop
			Else
				response.write "<td bgcolor=" & Color &">&nbsp;</td>"
				response.write "<td bgcolor=" & Color &">&nbsp;</td>"
				response.write "<td bgcolor=" & Color &">&nbsp;</td>"												
			End if
			Rs2.Close
		%>
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
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&From_day=<%=From_Day%>&From_month=<%=From_Month%>&From_Year=<%=From_Year%>&to_day=<%=To_Day%>&to_Month=<%=To_month%>&to_Year=<%=To_year%>&sExptype=<%=sExptype%>&scatid=<%=scatid%>&scattype=<%=scattype%>&Product=<%=Product%>">Previous 
                    Page</a> <%End if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> 
			<% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&From_day=<%=From_Day%>&From_month=<%=From_Month%>&From_Year=<%=From_Year%>&to_day=<%=To_Day%>&to_Month=<%=To_month%>&to_Year=<%=To_year%>&sExptype=<%=sExptype%>&scatid=<%=scatid%>&scattype=<%=scattype%>&Product=<%=Product%>">Next Page</a> <%End if%> 
                  </td>
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../images/dot.gif"></td>
                </tr>
                <tr> 
                  <td align="center" bgcolor="#FFD553" colspan="4"> <%
	For i=1 to TotalPage
		If Cint(PageNo)=Cint(i) Then
%> <Font color="Red"><%=i%></font> 
                    <%else%> 
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>&From_day=<%=From_Day%>&From_month=<%=From_Month%>&From_Year=<%=From_Year%>&to_day=<%=To_Day%>&to_Month=<%=To_month%>&to_Year=<%=To_year%>&sExptype=<%=sExptype%>&scatid=<%=scatid%>&scattype=<%=scattype%>&Product=<%=Product%>"><%=i%></a> 
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
