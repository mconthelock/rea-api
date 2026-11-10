<%@ Language=VBScript %>

<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->

<%
    Response.Buffer = true
	Response.ContentType ="Application/vnd.excel" 
	Response.AddHeader "Content-Disposition", "attachment; filename=Borrow.xls" 
'landscape
%>


<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<title>FIXED ASSET REGISTER FORM Excel</title>
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
	dim rsOra , rsOra2
	dim cntLine
	dim grp , lc , grpdesc , lcdesc
	set rsOra = CreateObject("adodb.recordset")
	set rsOra2 = CreateObject("adodb.recordset")
	
	cntLine = 1
	sqlOra = "select  distinct form.dreqdate, a.sempno inputer, a.sname inpName " _
	& " , b.sempno reqNo, b.sname reqName, sec.ssec, dep.sdept, div.sdiv, pos.sposition, fxa_borrowform.borrowtype , fxa_borrowform.tloccode, fxa_borrowform.borrowto " _
	& " , fxa_borrowform.empcode , fxa_borrowform.duration , fxa_borrowform.cause" _
	& " from fxa_borrowform, form, amec.aemployee a, amec.aemployee b, amec.psection sec, amec.pdepartment dep, amec.pdivision div, amec.pposition pos " _
	& " where b.sseccode = sec.sseccode and b.sdepcode= dep.sdepcode and b.sdivcode = div.sdivcode and b.sposcode = pos.sposcode " _
	& " and fxa_borrowform.nfrmno = form.nfrmno  and fxa_borrowform.vorgno = form.vorgno " _
	& " and fxa_borrowform.cyear = form.cyear and fxa_borrowform.cyear2 = form.cyear2 " _
	& " and fxa_borrowform.nrunno = form.nrunno and a.sempno = form.vinputer and b.sempno = form.vreqno " _
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
	  <td rowspan=2 >&nbsp;</td>  
	  <td>
		<table width="100%" border="0">	  
		  <tr>
			<td colspan=11 align=center class="fontHead" nowrap><b>MITSUBISHI ELEVATOR ASIA CO., LTD.</b></td>
		  </tr>		
 		  <tr>
				<td colspan=11 align=center nowrap>&nbsp;</td>
		  </tr>				
<!--		  <tr>
			<td colspan=3 align=center  bgcolor="#CCFFFF" class="fontHead" nowrap><b>Clear Advance</b></td>
		  </tr>-->
		  <tr> 
			<td class="fontCol1">Form No.</td>
			<td colspan=10 class="fontCol1"><%=toFormNumber(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"), request.querystring("y2"), request.querystring("runNo"))%></td>
		  </tr>
		  <tr>   
			<td class="fontCol1">Input By</td> 
			<td colspan=10 class="fontCol1"><%="(" & rsOra("inputer") & ") " & rsOra("inpName") %></td>	
		  </tr>
		  <tr>     
			<td class="fontCol1">Request By</td>
			<td  colspan=10 class="fontCol1"><%="(" & rsOra("reqNo") & ") " & rsOra("reqName") %></td>		
		  </tr>  
		  <tr> 
			<td class="fontCol1">DIV / DEPT / SEC</td>
			<td  colspan=10 class="fontCol1"><%=rsOra("sdiv")%>/<%=rsOra("sdept")%>/<%=rsOra("ssec")%></td>
		  </tr>  	
		  <tr> 
			<td class="fontCol1">Request Date</td>
			<td  colspan=10 class="fontCol1"><%=rsOra("dreqdate")%></td>
		  </tr> 
		  <tr>
			<td class="fontCol1">CAUSE</td>
			<td  colspan=10 class="fontCol1"><% if rsOra("borrowtype") = "I" then %>BORROW INSIDE COMPANY<% else %> BORROW OUTSIDE COMPANY <% end if %></td>
		  </tr>
		  <% if rsOra("borrowtype") = "O" then %>
		    <tr>
			<td class="fontCol1">BORROW BY :</td>
			<td  colspan=10 class="fontCol1">
			<%
						
							  response.write rsOra("borrowto")	
						  
			%>
			
			</td>
		  </tr>
		  <%
				end if
		  %>
		  
		   <tr>
			<td class="fontCol1">Duration </td>
			<td  colspan=10 class="fontCol1"><%=rsOra("duration")%></td>
		  </tr>
			<tr>
			<td class="fontCol1">Cause Of Borrow </td>
			<td  colspan=10 class="fontCol1"><%=rsOra("cause")%></td>
		  </tr>				
		  
         <%
		 
		 		sqlOra = "select b.grpcode dgrpcode, a.Floccode dloccode, a.assetno dassetno, a.qty, a.remark vremark , b.*, c.locname " _
				& "  from fxa_borrowform a, fxa_faMst b, fxa_locMst c " _
				& "  where a.assetno = b.assetno and a.Floccode = c.loccode " _
				& " and a.nfrmno = '" & request.querystring("no") & "'" _
				& " and a.vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and a.cyear = '" & request.querystring("y") & "'" _
				& " and a.cyear2 = '" & request.querystring("y2") & "'" _
				& " and a.nrunno = '" & request.querystring("runNo") & "'"  _
				& " order by a.assetno"
				rsOra.Close
				rsOra.open sqlOra, connOra	
					i = 1
		 %>
		 
		  <tr align= center bgcolor="#ec858f" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
		  	<td class="fontCol2">ITEM</td>
			<td class="fontCol2" >GROUP</td>
			<td class="fontCol2">ASSET NO.</td>
			<td class="fontCol2">L/C CODE</td>
			<td class="fontCol2">LOCATION</td>
			<td class="fontCol2">DESCRIPTION</td>
			<td class="fontCol2">QTY</td>
			<td class="fontCol2">PURCHASE DATE</td>
			<td class="fontCol2">INITIAL VALUE</td>
			<td class="fontCol2">BOOK VALUE</td>
			<td class="fontCol2">REMARK</td>
		  </tr>
	
				<%	do while not rsOra.EoF %>
			  <tr align="center" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
  				
				<td nowrap class="fontCol2"><%=i%></td>
								<td nowrap class="fontCol2"><%=rsOra("dgrpcode")%></td>
								<td nowrap class="fontCol2"><%=rsOra("dassetno")%></td>
								<td nowrap class="fontCol2"><%=rsOra("dloccode")%></td>
								<td nowrap class="fontCol2"><%=rsOra("locname")%></td>
								<td nowrap class="fontCol2"><%=rsOra("assdesc")%></td>
								<td nowrap class="fontCol2"><%=rsOra("qty")%></td>
								<td nowrap class="fontCol2"><%=rsOra("startdp")%></td>
								<td nowrap class="fontCol2"><%=formatNumber(rsOra("INITIALVALUE"), 2, true, true, true)%></td>
								<td nowrap class="fontCol2"><%=rsOra("bvalue")%></td>
								<td nowrap class="fontCol2"><%=rsOra("vremark")%>&nbsp;</td>
				
			  </tr>				
				<%
								i = i + 1 
							rsOra.MoveNext
						loop
						rsOra.Close
				%>      																	
		</table>
	  </td>
  </tr>
</table>
<%
	call  showFlowExcel(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),request.querystring("y2"), request.querystring("runNo"), requester, session("uid"),"remark,legend")
%>
</html>
