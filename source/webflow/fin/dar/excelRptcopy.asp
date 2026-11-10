<%@ Language=VBScript %>

<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->

<%
    Response.Buffer = true
	Response.ContentType ="Application/vnd.excel" 
	Response.AddHeader "Content-Disposition", "attachment; filename=Disposal.xls" 
'landscape
%>


<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<title>Entertainment Form Excel</title>
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
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>DISPOSAL ASSETS REQUEST FORM</title>

<style>
	body{
		font-family:sans-serif;
	}
	.table-main {
		margin-left:15px;
	}
	.table-main th{
		border: 1px solid write;
		background-color:#FA9BC9;
	}
	.table-main td{
		padding-left:5px;
		padding-right:5px;
		border: 1px solid #FA9BC9;
		background-color:#EED2DB;
	}
</style>
</head>
<body>
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/flowUtil2.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/date2.inc"--> 
<!--#include file="../../inc/date.inc"-->
<!--#include file="../../inc/misc.inc"-->
<!--#include file="inc/showAttFile.inc"-->
<!--#include file="../inc/flowExcel.inc"-->
<script language="JavaScript">
	function CheckData(act)
	{
		return true;
	}
	function createRow(id, value){
		document.getElementById(id).innerHTML = document.getElementById(id).innerHTML + value;
	}
</script>
<%
	Dim rsOra, sqlOra
	Dim mode, requester
	Dim totalAdv, totalExp, cntLine, i
	Dim ptype, gAMEC, gTotal, sqlSub, totalAmount
	Dim totalQty , totalIniValue
	

	set rsOra = createObject("adodb.recordset")
	
	totalAdv = 0
	totalExp = 0
	cntLine = 1
	i = 1
	if session("uid") = request.queryString("empNo") then
		apv = request.queryString("empNo")
	else
		apv = session("uid")
	end if		
	
	sqlOra = "select cExtdata from flow where nfrmno = '" & request.querystring("no") & "' and vorgno = '" & request.querystring("orgNo") & "'" _
	& " and cyear = '" & request.querystring("y") & "' and cyear2='" & request.querystring("y2") & "' and nrunno = '" & request.querystring("runNo") & "'" _
	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') and cstepst='" & STEP_READY & "'"
'	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') "
'	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') and cstepst='" & STEP_READY & "'"
'	response.write sqlOra
'	response.end
	rsOra.open sqlOra, connOra

	if 	not IsNull(rsOra("cExtdata")) and (not rsOra.EoF) then
		stepExt = rsOra("cExtdata")
	end if
	rsOra.close

	sqlOra = "select  distinct form.dreqdate, a.sempno inputer, a.sname inpName " _
	& " , b.sempno reqNo, b.sname reqName, sec.ssec, dep.sdept, div.sdiv, pos.sposition, d.writeoffdesc, fxa_darform.recipients , fxa_darform.tsale , fxa_darform.cause , fxa_darform.writeoffno" _
	& " from fxa_darform, form, amec.aemployee a, amec.aemployee b, amec.psection sec, amec.pdepartment dep, amec.pdivision div, amec.pposition pos, fxa_causewriteoff  d  " _
	& " where b.sseccode = sec.sseccode and b.sdepcode= dep.sdepcode and b.sdivcode = div.sdivcode and b.sposcode = pos.sposcode " _
	& " and fxa_darform.nfrmno = form.nfrmno  and fxa_darform.vorgno = form.vorgno " _
	& " and fxa_darform.cyear = form.cyear and fxa_darform.cyear2 = form.cyear2 " _
	& " and fxa_darform.nrunno = form.nrunno and a.sempno = form.vinputer and b.sempno = form.vreqno " _
	& " and fxa_darform.writeoffno = d.writeoffno " _
	& " and form.nfrmno = '" & request.querystring("no") & "'" _
	& " and form.vorgno = '" & request.querystring("orgNo")  & "'" _
	& " and form.cyear = '" & request.querystring("y") & "'" _
	& " and form.cyear2 = '" & request.querystring("y2")  & "'" _
	& " and form.nrunno = '" & request.querystring("runNo")  & "'"
'	response.write sqlOra
'	response.end

	rsOra.open sqlOra, connOra
	mode = getMode(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid"))
'	stepExt = getExtdata(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid")) 	
	requester = rsOra("reqNo")
%>

<table>
  <tr><td rowspan=15>&nbsp;</td></tr>  
  <tr><td colspan=2 align=center class="fontHead"><b>DISPOSAL ASSETS REQUEST FORM</b></td></tr>
  <tr>
  	<td class="fontCol1">Form No : </td><td class="fontCol1"><%=toFormNumber(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),  request.querystring("y2") , request.querystring("runNo"))%></td>
  </tr>
  <tr><td class="fontCol1">Input By : </td><td class="fontCol1"><%="(" & rsOra("inputer") & ") " & rsOra("inpName") %></td></tr>	
  <tr><td class="fontCol1">Request By :</td><td class="fontCol1"><%="(" & rsOra("reqNo") & ") " & rsOra("reqName") %></td></tr>	
  <tr><td class="fontCol1">DIV/DEPT/SEC :</td><td class="fontCol1"><%=rsOra("sdiv")%>/<%=rsOra("sdept")%>/<%=rsOra("ssec")%></td></tr>	
  <tr><td class="fontCol1">Date Request :</td><td class="fontCol1"><%=rsOra("dreqdate")%></td></tr>	
   <tr><td class="fontCol1">CAUSE OF WRITE-OFF ASSETS :</td><td class="fontCol1"><%=left(rsOra("writeoffdesc"),instr(rsOra("writeoffdesc"),"/")-1)%></td></tr>	
 
<%
	rsOra.close	

			sqlOra = "select a.grpcode dgrpcode, a.loccode dloccode, a.assetno dassetno, to_number(a.nqty) nqty, a.vunit, a.vremark, to_number(b.initialvalue) inivalue, b.* , c.locname, d.writeoffdesc " _
				& " a.writeoffno , a.recipients , a.tsale , a.cause" _
				& "  from fxa_darform a, fxa_faMst b, fxa_locMst c, fxa_causewriteoff  d " _
				& "  where a.assetno = b.assetno and a.loccode = c.loccode and a.writeoffno = d.writeoffno " _
				& " and a.nfrmno = '" & request.querystring("no") & "'" _
				& " and a.vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and a.cyear = '" & request.querystring("y") & "'" _
				& " and a.cyear2 = '" & request.querystring("y2") & "'" _
				& " and a.nrunno = '" & request.querystring("runNo") & "'"  _
				& " order by a.assetno"

				
				
				
				
	rsOra.open sqlOra, connOra	
	
%>

		 
		  <tr align= center bgcolor="#ec858f" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
		  	<td class="fontCol2">ITEM</td>
			<td class="fontCol2" >GROUP</td>
			<td class="fontCol2" >STATUS</td>
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

<% do while not rsOra.EoF	%>
								<tr align="center" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
								<td nowrap class="fontCol2"><%=i%></td>
								<td nowrap class="fontCol2"><%=rsOra("dgrpcode")%></td>
								<td nowrap class="fontCol2"><%=rsOra("status")%></td>
								<td nowrap class="fontCol2"><%=rsOra("dassetno")%></td>
								<td nowrap class="fontCol2"><%=rsOra("loccode")%></td>
								<td nowrap class="fontCol2"><%=rsOra("locname")%></td>
								<td nowrap class="fontCol2"><%=rsOra("assdesc")%></td>
								<td nowrap class="fontCol2"><%=rsOra("nqty")%></td>
								<td nowrap class="fontCol2"><%=rsOra("startdp")%></td>
								<td nowrap class="fontCol2"><%=rsOra("initialvalue")%></td>
								<td nowrap class="fontCol2"><%=rsOra("bvalue")%></td>
								<!--<td valign=top><%'=left(rsOra("writeoffdesc"),instr(rsOra("writeoffdesc"),"/")-1)%></td>-->
									
								<td nowrap class="fontCol2"><%=rsOra("vremark")%>&nbsp;</td>
							</tr>
							
				<%
					totalQty = totalQty + rsOra("nqty")
					totalIniValue = totalIniValue + rsOra("inivalue")					
					i = i + 1 
					rsOra.MoveNext
				loop
				rsOra.close
				%>	
				 <tr align="center" style="PADDING-RIGHT: 10px; border-top: #000000 1px solid; border-left: #000000 1px solid;border-right: #000000 1px solid;border-bottom: #000000 1px solid; line-height: 10px;">
				<td nowrap class="fontCol2" colspan=7 align=right ><strong>Grand Total</strong></td><td nowrap class="fontCol2"><%=totalQty%></td><td nowrap class="fontCol2">&nbsp;</td><td nowrap class="fontCol2"><%=totalIniValue%></td><td nowrap class="fontCol2">&nbsp;</td><td nowrap class="fontCol2" colspan=2>&nbsp;</td></tr>
</table>
<%
	call  showFlowExcel(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),request.querystring("y2"), request.querystring("runNo"), requester, session("uid"),"remark,legend")
%>
</body>
</html>