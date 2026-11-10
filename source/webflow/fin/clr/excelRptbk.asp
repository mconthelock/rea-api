<%@ Language=VBScript %>
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/connUtil.inc"--> 
<!--#include file="../../inc/formUtil.inc"-->
<!--#include file="../../inc/db.inc"--> 

<%
    Response.Buffer = true
	Response.ContentType ="Application/vnd.excel" 
	Response.AddHeader "Content-Disposition", "attachment; filename=clr.xls" 
'landscape
%>


<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<title>Clear Advance Form Excel</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<META name=ProgId content=Excel.Sheet> 
<META name=Generator content="Microsoft Excel 9"> 
<LINK rel=File-List href="./z.files/filelist.xml"> 
<LINK rel=Edit-Time-Data href="./z.files/editdata.mso"> 
<LINK rel=OLE-Object-Data href="./z.files/oledata.mso"> 
<!--[if gte mso 9]>
<xml> 
<o:DocumentProperties> 
<o:LastAuthor>wonder</o:LastAuthor> 
<o:LastPrinted>2003-05-08T07:02:15Z</o:LastPrinted> 
<o:Created>2003-05-07T10:59:26Z</o:Created> 
<o:LastSaved>2003-05-08T07:02:18Z</o:LastSaved> 
<o:Version>9.6615</o:Version> 
</o:DocumentProperties> 
<o:OfficeDocumentSettings> 
<o:DownloadComponents/> 
<o:LocationOfComponents HRef="file:///D:/program/application/office%25202000/msowc.cab"/> 
</o:OfficeDocumentSettings> 
</xml>
<![endif]--> 
<style>
 .txt {mso-number-format:\@;}
<!--table

{
mso-displayed-decimal-separator:"\.";

mso-displayed-thousand-separator:"\,";}

@page
      {
      mso-page-orientation:Portrait;
      margin:.2in .2in .5in .2in;
      mso-header-margin:.1in;      
      mso-footer-margin:.1in;
      mso-footer-data:"&R&P of &N";
      mso-horizontal-page-align:top;
      mso-vertical-page-align:top;
      }

tr {mso-height-source:auto;}

td {mso-width-source:auto;}

col {mso-width-source:auto;}

.fontHead {font-family: ms sans serif; color: #575757; font-size: 16pt;}
.fontCol {font-family: ms sans serif; color: #575757; font-size: 14pt;; text-align:center;}
.fontCol1 {font-family: ms sans serif; color: #575757; font-size: 12pt;; text-align:left;}
.fontCol2 {font-family: ms sans serif; color: #575757; font-size: 12pt;; text-align:center;}
-->	  
</style>

<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#FF0000" alink="#999999">
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/date.inc"-->
<!--#include file="../../inc/date2.inc"-->
<!--#include file="../inc/flowExcel.inc"-->
<%		
	dim sqlOra
	dim rsOra
	dim cntLine
	
	set rsOra = CreateObject("adodb.recordset")
	cntLine = 1
	sqlOra = "select  form.dreqdate, a.sempno inputer, a.sname inpName, cyadv, nnoadv" _
	& " , b.sempno reqNo, b.sname reqName, sec.ssec, dep.sdept, div.sdiv, pos.sposition " _
	& " from clrForm, form, amec.aemployee a, amec.aemployee b, amec.psection sec, amec.pdepartment dep, amec.pdivision div, amec.pposition pos, cashpayby " _
	& " where b.sseccode = sec.sseccode and b.sdepcode= dep.sdepcode and b.sdivcode = div.sdivcode and b.sposcode = pos.sposcode " _
	& " and clrForm.nfrmno = form.nfrmno  and clrForm.vorgno = form.vorgno " _
	& " and clrForm.cyear = form.cyear and clrForm.cyear2 = form.cyear2 " _
	& " and clrForm.nrunno = form.nrunno and a.sempno = form.vinputer and b.sempno = form.vreqno " _
	& " and form.nfrmno = '" & request.querystring("no") & "'" _
	& " and form.vorgno = '" & request.querystring("orgNo")  & "'" _
	& " and form.cyear = '" & request.querystring("y") & "'" _
	& " and form.cyear2 = '" & request.querystring("y2")  & "'" _
	& " and form.nrunno = '" & request.querystring("runNo")  & "'"	
'	response.write sqlOra
'	response.end
	rsOra.open sqlOra, connOra
			
%>
<table width="100%" border="0">	
  <tr><td>&nbsp;</td></tr> 
  <tr>
	  <td rowspan=2 colspan=2>&nbsp;</td>  
	  <td>
		<table width="100%" border="0">	  
		  <tr>
			<td colspan=5 align=center class="fontHead" nowrap><b>MITSUBISHI ELEVATOR ASIA CO., LTD.</b></td>
		  </tr>		
 		  <tr>
				<td colspan=5 align=center nowrap>&nbsp;</td>
		  </tr>				
<!--		  <tr>
			<td colspan=3 align=center  bgcolor="#CCFFFF" class="fontHead" nowrap><b>Clear Advance</b></td>
		  </tr>-->
		  <tr> 
			<td class="fontCol1">Form No.</td>
			<td colspan=4 class="fontCol1"><%=toFormNumber(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"), request.querystring("y2"), request.querystring("runNo"))%></td>
		  </tr>
		  <tr>   
			<td class="fontCol1">Input By</td> 
			<td colspan=4 class="fontCol1"><%="(" & rsOra("inputer") & ") " & rsOra("inpName") %></td>	
		  </tr>
		  <tr>     
			<td class="fontCol1">Request By</td>
			<td  colspan=4 class="fontCol1"><%="(" & rsOra("reqNo") & ") " & rsOra("reqName") %></td>		
		  </tr>  
		  <tr> 
			<td class="fontCol1">DIV / DEPT / SEC</td>
			<td  colspan=4 class="fontCol1"><%=rsOra("sdiv")%>/<%=rsOra("sdept")%>/<%=rsOra("ssec")%></td>
		  </tr>  	
		  <tr> 
			<td class="fontCol1">Request Date</td>
			<td  colspan=4 class="fontCol1"><%=rsOra("dreqdate")%></td>
		  </tr>  	
		  <tr>   
			<td colspan=5 bgcolor="#CCFFCC" class="fontCol" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;"><b>Cash Advance</b></td>
		  </tr> 
		  <tr align= center bgcolor="#ec858f" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
		  	<td class="fontCol2">No.</td>
			<td class="fontCol2" colspan=3>Description</td>
			<td class="fontCol2">Amount</td>
		  </tr>
		  <%
				sqlOra = "select id, description, to_number(namount) namount from cashAdvList where cyear2 = '" & rsOra("cyadv") & "' and nrunno = '" & rsOra("nnoadv") & "'"
				rsOra.Close
				rsOra.open sqlOra, connOra			
				'response.write sqlOra
			%>
				<%	do while not rsOra.EoF %>
			  <tr align="center" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
  				<td nowrap class="fontCol2"><%=rsOra("id")%></td>
				<td nowrap class="fontCol2" colspan=3><%=rsOra("description")%></td>
				<td nowrap class="fontCol2"><%=formatNumber(rsOra("namount"), 2, true, true, true)%></td>				
			  </tr>				
				<%
							totalAdv = totalAdv + rsOra("namount")
							cntLine = cntLine + 1
							rsOra.MoveNext
						loop
						rsOra.Close
				%>
			<tr bgcolor="#ec858f"  style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
				<td colspan=3 align=center class="fontCol"><strong>Clear Advance Amount</strong></td>
				<td colspan=2 align=center class="fontCol"><strong><%=formatNumber(totalAdv, 2, true, true, true)%></strong></td>
			</tr>
			<tr><td colspan= 5>&nbsp;</td></tr>
			<%						
				sqlOra = "select id, svendinv, to_number(namount) namount " _
				& " from clrList" _
				& " where clrList.nfrmno = '" & request.querystring("no") & "'" _
				& " and clrList.vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and clrList.cyear = '" & request.querystring("y") & "'" _
				& " and clrList.cyear2 = '" & request.querystring("y2")  & "'" _
				& " and clrList.nrunno = '" & request.querystring("runNo")  & "' "
				
				'response.write sqlOra
				rsOra.open sqlOra, connOra						
			  %>
			 <tr>   
				<td colspan=5 bgcolor="#CCFFCC" class="fontCol" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;"><b>Clear Advance</b></td>
			  </tr> 			  
			  <tr bgcolor="#ec858f" align="center" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
				<td nowrap class="fontCol2"><strong>No.</strong></td>			  
				<td nowrap colspan=3 class="fontCol2"><strong>Receipt No</strong></td>
				<td class="fontCol2">Amount</td>
			  </tr>
				<%	do while not rsOra.EoF %>
			  <tr align="center" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
  				<td nowrap class="fontCol2"><%=rsOra("id")%></td>
				<td nowrap colspan=3 class="fontCol2"><%=rsOra("svendinv")%></td>
				<td nowrap class="fontCol2"><%=formatNumber(rsOra("namount"), 2, true, true, true)%></td>				
			  </tr>
				<%
							totalExp = totalExp + rsOra("namount")
							rsOra.MoveNext
						loop
						rsOra.Close
						cntLine = 1
				%>			  
			<tr bgcolor="#ec858f"  style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
				<td colspan=3 align=center class="fontCol"><strong>Total Expense Amount</strong></td>
				<td colspan=2 align=center class="fontCol"><strong><%=formatNumber(totalExp, 2, true, true, true)%></strong></td>
			</tr>
			<tr style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
				<td colspan=3 align =right class="fontCol2"><strong>
				<%
					if totalAdv > totalExp then
						color= "#FF0000"
						response.write "Return cash to company"
					else
						color= "#000000"
						response.write "Addition pay to employee"					
					end if
				%>
				</strong></td>
				<td align=center colspan=2 class="fontCol2">
				<font color=<%=color%>>
				<%=formatNumber(totalExp-totalAdv, 2, true, true, true)%>
				</font>
				</td>
			</tr>	
			<tr><td></td></tr>
		  <tr><td colspan=4>Remark</td></tr>   
		  <tr><td colspan=4>1. Please attached original Receipt slip & Tax invoice</td></tr>   
		  <tr><td colspan=4>2. After process complete please send Receipt slip or Tax Invoice to SEM stamp</td></tr>       
		  <tr><td colspan=4>3. Please attached P/R</td></tr>         																	
		</table>
	  </td>
  </tr>
</table>
<%
	call  showFlowExcel(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),request.querystring("y2"), request.querystring("runNo"), requester, session("uid"),"remark,legend")
%>
</html>
