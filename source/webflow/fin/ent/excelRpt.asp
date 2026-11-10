<%@ Language=VBScript %>

<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->

<%
    Response.Buffer = true
	Response.ContentType ="Application/vnd.excel" 
	Response.AddHeader "Content-Disposition", "attachment; filename=Entertain.xls" 
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
<title>Request For Entertainment Form</title>

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
	Dim payto, purpose, paydate, place, description
	Dim ptype, gAMEC, gTotal, sqlSub, totalAmount
	

	set rsOra = createObject("adodb.recordset")
	
	totalAdv = 0
	totalExp = 0
	cntLine = 1
	if session("uid") = request.queryString("empNo") then
		apv = request.queryString("empNo")
	else
		apv = session("uid")
	end if		
	
	sqlOra = "select cExtdata from flow where nfrmno = '" & request.querystring("no") & "' and vorgno = '" & request.querystring("orgNo") & "'" _
	& " and cyear = '" & request.querystring("y") & "' and cyear2='" & request.querystring("y2") & "' and nrunno = '" & request.querystring("runNo") & "'" _
	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') "
'	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') and cstepst='" & STEP_READY & "'"
'	response.write sqlOra
'	response.end
	rsOra.open sqlOra, connOra

	if 	not IsNull(rsOra("cExtdata")) and (not rsOra.EoF) then
		stepExt = rsOra("cExtdata")
	end if
	rsOra.close

	sqlOra = "select  form.dreqdate, a.sempno inputer, a.sname inpName " _
	& " , b.sempno reqNo, b.sname reqName, sec.ssec, dep.sdept, div.sdiv, pos.sposition " _
	& " from rqtEntForm, form, amec.aemployee a, amec.aemployee b, amec.psection sec, amec.pdepartment dep, amec.pdivision div, amec.pposition pos " _
	& " where b.sseccode = sec.sseccode and b.sdepcode= dep.sdepcode and b.sdivcode = div.sdivcode and b.sposcode = pos.sposcode " _
	& " and rqtEntForm.nfrmno = form.nfrmno  and rqtEntForm.vorgno = form.vorgno " _
	& " and rqtEntForm.cyear = form.cyear and rqtEntForm.cyear2 = form.cyear2 " _
	& " and rqtEntForm.nrunno = form.nrunno and a.sempno = form.vinputer and b.sempno = form.vreqno " _
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
  <tr><td colspan=2 align=center class="fontHead"><b>Request For Entertainment Form</b></td></tr>
  <tr>
  	<td class="fontCol1">Form No : </td><td class="fontCol1"><%=toFormNumber(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),  request.querystring("y2") , request.querystring("runNo"))%></td>
  </tr>
  <tr><td class="fontCol1">Input By : </td><td class="fontCol1"><%="(" & rsOra("inputer") & ") " & rsOra("inpName") %></td></tr>	
  <tr><td class="fontCol1">Request By :</td><td class="fontCol1"><%="(" & rsOra("reqNo") & ") " & rsOra("reqName") %></td></tr>	
  <tr><td class="fontCol1">DIV/DEPT/SEC :</td><td class="fontCol1"><%=rsOra("sdiv")%>/<%=rsOra("sdept")%>/<%=rsOra("ssec")%></td></tr>	
  <tr><td class="fontCol1">Date Request :</td><td class="fontCol1"><%=rsOra("dreqdate")%></td></tr>	
<%
	rsOra.close				
	sqlOra = "select a.*, b.sname from rqtEntForm a, amec.aemployee b where " _
	& " a.payto = b.sempno" _
	& " and a.nfrmno = '" & request.querystring("no") & "'" _
	& " and a.vorgno = '" & request.querystring("orgNo")  & "'" _
	& " and a.cyear = '" & request.querystring("y") & "'" _
	& " and a.cyear2 = '" & request.querystring("y2") & "'" _
	& " and a.nrunno = '" & request.querystring("runNo") & "'"
	rsOra.open sqlOra, connOra	
	
	payto = "(" & rsOra("payto") & ") " & rsOra("sname")
	purpose = rsOra("purpose")
	paydate = rsOra("paydate")
	place = rsOra("place")
	description = rsOra("description")
%>
  <tr><td class="fontCol1">PAY TO : </td><td class="fontCol1"><%=payto%></td></tr>	
  <tr>
  	<td colspan=2>
		<table border=1>
			<tr><td colspan=2 class="fontCol">PARTICIPANT(s)</td></tr>		
			<!--<tr><td class="fontCol2">AMEC</td><td class="fontCol2">GUEST</td></tr>-->
<%
				rsOra.close
				
				sqlOra = "select count(*) as A from rqtEntPersons where " _
				& " nfrmno = '" & request.querystring("no") & "'" _
				& " and vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and cyear = '" & request.querystring("y") & "'" _
				& " and cyear2 = '" & request.querystring("y2") & "'" _
				& " and nrunno = '" & request.querystring("runNo") & "' group by ptype "
				'response.write sqlOra
				rsOra.open sqlOra, connOra
				i = 1
				do while not rsOra.EoF
					if i = 1 then 
						gAMEC = rsOra("A") 
					else 
						if gAMEC >= rsOra("A") then
							ptype = "AMEC"
							sqlSub = "asc"
						else 
							ptype = "GUEST"
							sqlSub = "desc"
						end if 
						gTotal = gAMEC + rsOra("A")
					end if 
					i = i + 1
					rsOra.MoveNext
				loop
				rsOra.close
				i = 1
				j = 1
				
				sqlOra = "select * from rqtEntPersons where " _
				& " nfrmno = '" & request.querystring("no") & "'" _
				& " and vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and cyear = '" & request.querystring("y") & "'" _
				& " and cyear2 = '" & request.querystring("y2") & "'" _
				& " and nrunno = '" & request.querystring("runNo") & "' order by ptype " & sqlSub
				
				sqlOra = "select ptype, wm_concat(pname) pname from rqtEntPersons" _
				& " where nfrmno = '" & request.querystring("no") & "'" _
				& " and vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and cyear = '" & request.querystring("y") & "'" _
				& " and cyear2 = '" & request.querystring("y2") & "'" _
				& " and nrunno = '" & request.querystring("runNo") & "'" _
				& " group by ptype order by ptype " & sqlSub
				
				rsOra.open sqlOra, connOra
			'	pname = split(rsOra("pname"),",")
				
				
				
			do while not rsOra.EoF 
					'if rsOra("ptype") = ptype then 
					'	if rsOra("ptype") = "AMEC" then
						if rsOra("ptype") = "AMEC" then
							pname = split(rsOra("pname"),",")
							response.write "<tr><td class='fontCol1' valign=top rowspan=" & ubound(pname)+1 & ">AMEC</td>"
							for i=0 to ubound(pname)
								response.write "<td class='fontCol1'>" & i+1 & "  " & pname(i)  & "</td></tr>"
							next
						else
							pname = split(rsOra("pname"),",")						
							response.write "<tr><td class='fontCol1' valign=top rowspan=" & ubound(pname)+1 & ">GUEST</td>"
							for i=0 to ubound(pname)
								response.write "<td class='fontCol1'>" & i+1 & "  " & pname(i)  & "</td></tr>"
							next
							'response.write "<td class='fontCol1'>" & rsOra("pname") & "</td></tr>"
						end if
							'response.write "<tr><td id='pAmec" & i & "' align='left' width='50%' class='fontCol2'>" & i & ".&nbsp;" & rsOra("pname") & "&nbsp;</td>"
							'response.write "<td id='pGuest" & i & "' align='left' width='50%' class='fontCol2'>" & i & ".&nbsp;</td></tr>"
					'	else 
					'		response.write "<tr><td id='pAmec" & i & "' align='left' width='40%' class='fontCol2'>" & i & ".&nbsp;</td>"
					'		response.write "<td id='pGuest" & i & "' align='left' width='40%' class='fontCol2'>" & i & ".&nbsp;" & rsOra("pname") & "&nbsp;</td></tr>"
					'	end if 
					'else
						'if rsOra("ptype") = "AMEC" then
						'	'response.write "<td>" & j & "</td><td>" &  rsOra("pname") & "<td>"
						'	response.write "<td><script>createRow('pAmec" & j & "', '" & rsOra("pname") & "');</script><t/d>"
						'	j = j + 1
						'else 
						'	'response.write "<td>" & j & "</td><td>" &  rsOra("pname") & "<td>"						
						'	response.write "<script>createRow('pGuest" & j & "', '" & rsOra("pname") &  "');</script>"
						'	j = j + 1
						'end if
					'end if 
					i = i + 1 
					rsOra.MoveNext
				loop
				rsOra.close
			'	response.end
				
%>
				<tr>
					<td colspan="2" align="center" class='fontCol2'><b>TOTAL</b>&nbsp;&nbsp;<u><%=gTotal%></u>&nbsp;&nbsp;PERSONS</td>
				</tr>			
		</table>
  	</td>
  </tr>
	<tr>
	<td class="fontCol1">PURPOSE : </td><td class="fontCol1"><%=purpose%>&nbsp;</td>
	</tr>
	<tr>
		<td class="fontCol1">DATE : </td><td class="fontCol1"><%=paydate%>&nbsp;</td>
	</tr>
	<tr>
		<td class="fontCol1">PLACE : </td><td class="fontCol1"><%=place%>&nbsp;</td>
	</tr>
	<tr>
		<td class="fontCol1">DESCRIPTION : </td><td class="fontCol1"><%=description%>&nbsp;</td>
	</tr>  
		<tr>
			<td style="padding:0px;">
				<table cellspacing="1" cellpadding="1" border=1>
					<tr><td colspan=2 class="fontCol">EXPENSE</td></tr>
					<tr>
						<td align="center" class="fontCol2">PARTICULARS</td><td  align="center" class="fontCol2">PRICE</td>
					</tr>
					<%
					totalAmount = 0
					sqlOra = "select * from rqtEntList where " _
					& " nfrmno = '" & request.querystring("no") & "'" _
					& " and vorgno = '" & request.querystring("orgNo")  & "'" _
					& " and cyear = '" & request.querystring("y") & "'" _
					& " and cyear2 = '" & request.querystring("y2") & "'" _
					& " and nrunno = '" & request.querystring("runNo") & "' order by id"
					rsOra.open sqlOra, connOra
					do while not rsOra.EoF
						response.write "<tr><td align=center class='fontCol2'>" & rsOra("description") & "</td><td align='center'  class='fontCol2'>" & rsOra("namount") & "</td></th>"
						totalAmount = totalAmount + CDbl(rsOra("namount"))
						rsOra.MoveNext
					loop
					%>
					<tr>
						<td align=center class="fontCol2">TOTAL(BAHT)</td><td align="center" class="fontCol2"><%=totalAmount%></td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
		  <td width=1% nowrap class="fontCol1">Attached File : </td>
		  <td class="fontCol1">
			<%
				 call showAttFile(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), "ent", mode ,server.urlencode(request.querystring("bp")))
			%>
		  </td>
		</tr>
	</table>
<%
		call  showFlowExcel(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),request.querystring("y2"), request.querystring("runNo"), requester, session("uid"),"remark,legend")
%>			
</body>
</html>