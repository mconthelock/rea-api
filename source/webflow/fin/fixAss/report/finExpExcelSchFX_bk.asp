<%@ Language=VBScript %>
<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->
<!--#include virtual="/sys/inc/userUtil.inc"-->
<!--#include virtual="/form/inc/formUtil.inc"-->
<!--#include virtual="/sys/inc/date.inc"-->
<!--#include virtual="/sys/inc/date2.inc"-->
<!--#include virtual="/sys/inc/misc.inc"-->
<!--#include virtual="/sys/inc/report.inc"-->
<!--#include virtual="/sys/inc/reportData.inc"-->
<!--#include virtual="/sys/inc/search.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include virtual="/sys/inc/pageUtil.inc"-->
<%
	Response.Buffer = true
	Response.ContentType ="Application/vnd.excel" 
	Response.AddHeader "Content-Disposition", "attachment; filename=Report-Assest List.xls" 
'landscape
%>

<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40"-->
<head>
<title>FIN Report</title>
<META http-equiv="Content-Type" content="text/html; charset=windows-874">
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

-->	  
</style>
<head>
<title>Asset Report</title>
<META http-equiv="Content-Type" content="text/html; charset=windows-874">

</head>
<body>
<%
	dim rs
	dim sql
	dim chk,count
	dim reqDate
	count=1
	set rs=CreateObject("Adodb.recordset")
	sql = request("sqlOra")
'	response.write request("sqlOra")
'	response.end
	rs.Open sql, connOra

	if not rs.EoF then	
%>
      <table width="100%" border="1" cellspacing="1" cellpadding="2">
        <tr align="center"> 
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">No.</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Group No.</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Description</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">L/C</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Description L/C</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Asset No.</font></b></td>    
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Description</font></b></td>	
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">Doc. Date</font></b></td>	
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">INITIAL VALUE</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">INV.NO</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">MODEL NO</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">S/N</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">PO No.</font></b></td>
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">UNIT</font></b></td>    
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">SUPPLIER NAME</font></b></td>    
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">PR NO</font></b></td>    	
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">REQUEST BY</font></b></td>    		
    <td nowrap bgcolor="#AACCEE"><b><font color="#000000">PICTURE</font></b></td>    			
  </tr>
  <% do while not rs.EOF %>
  <tr  bgcolor="#FFFFFF"> 
    <td nowrap align="center"><font color="#000000"><%=count%></font></td>
    <td nowrap align="center"><font color="#000000"><%=rs("grpCode")%>&nbsp;</font></td>
    <td nowrap align="center"><font color="#000000"><%=rs("grpdesc")%>&nbsp;</font></td>    
    <td nowrap align="center"><font color="#000000"><%=rs("locCode")%>&nbsp;</font></td>
    <td nowrap align="center"><font color="#000000"><%=rs("locName")%>&nbsp;</font></td>	
    <td nowrap align="center"><font color="#000000"><%=rs("assetno")%>&nbsp;</font></td>	
    <td nowrap align="center"><font color="#000000"><%=rs("assDesc")%>&nbsp;</font></td>		
    <td nowrap align="center"><font color="#000000"><%=rs("docdate")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("initialvalue")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("invno")%>&nbsp;</font></td>			
   <td nowrap align="center"><font color="#000000"><%=rs("modelno")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("snno")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("pono")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("unit")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("supname")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("prno")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000"><%=rs("reqby")%>&nbsp;</font></td>			
    <td nowrap align="center"><font color="#000000">&nbsp;</font></td>					
  </tr>
  <%
				rs.MoveNext
				count=count+1
  		loop
%>
      </table>
<%	
	else
		response.write "data not found"
	end if	
%>
</body>
</html>
