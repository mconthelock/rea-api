<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../oraUtil.inc"  -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
	'/*****************************************************************/
	'/*COPYRIGHT (C) 2015 Mitsubishi Elevator Asia	*/
	'/* PROGRAM NO  */
	'/* Revision No  */
	'/* DEV by  (Kanittha - WSD) 	*/
	'/* Release on  	*/ 
	'/* Reason = “”*/  
	'/*****************************************************************/
%>
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
'response.write Request("stprodname")
'response.end
	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sedesc,nprice,simage,dadddate,sempno
	Dim sdrawnum, sspec, scountry, sdimen, sbrand, smarkern,seprodnamen
	Dim img_sprodcode,img_sexptype,img_svendcode,img_nprice,sbread,sCatID,Img_image
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year ,sErrImage
	Dim sEName,Mode,sCurcode,sUnitcode,sfilename,sAcccode,sCattype,chkProduct,hazardno,hazardstatus
	Dim l(7)
	Dim s(7)



	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sProdcode=Request.QueryString("sProdcode")
	sExptype=Request.QueryString("sExptype")
	sCattype=Request.QueryString("sCattype")
	sCatID=Request.QueryString("sCatID")
	sVendcode=Trim(Request.QueryString("sVendcode"))
	sCurcode=Request.QueryString("sCurcode")
	hazardno=Request.QueryString("hazardno")
	hazardstatus=Request.QueryString("hazardstatus")
	
	if sVendcode<> "" then
		Session("sVendcode")=sVendcode
	elseif Request("sExptype_hid") <> "" then
		Session("sVendcode")=sVendcode
	else
		sVendcode=Session("sVendcode")
	End if

	Mode = Request.QueryString("Mode")

	If IsEmpty(Request.Form("EVENT")) Then
		If  sProdcode <> "" and sExptype <> "" Then
			sMode = "EDIT"
			Get_Product(sprodcode)
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Product(Request.Form("EVENT"))
	End If		

' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Product(sprodcode)
	Dim sSQL, Rs , x
	Dim  templv , temp
	Dim  i
	
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM PProduct WHERE sprodcode='" & sprodcode & "'" &_
"and sexptype='" & sexptype & "'"
   Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
 ' response.write "test = "&conn
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sprodcode = Rs("sprodcode")
		sexptype = Rs("sexptype")
		sCattype = Rs("sCattype")
		if Request("mode")<>"Search" then
			sVendcode = Rs("svendcode")
		end if
		sprodid = Rs("sprodid")
		stprodname = Rs("stprodname")
		seprodname = Rs("seprodname")
	'	smodel = Rs("smodel")
		sAcccode=RS("sAcccode")
		stdesc = Rs("stdesc")
		sedesc = Rs("sedesc")
		sCatID = Rs("sCatID")
		nprice = Rs("nprice")
		simage = Rs("simage")
		dadddate = Rs("dAdddate")
		hazardstatus = Rs("hazardstatus")
		if hazardstatus = "1" then
			hazardno = Rs("hazardno")
		end if
	'hazardno = Rs("hazardno")
	'	sdrawnum = Rs("sdrawnum")
	'	sspec = Rs("sspec")
	'	scountry = Rs("scountry")
	'	sdimen = Rs("sdimen")
	'	sbrand = Rs("sbrand")
	'	smarkern = Rs("smarkern")
	
		if Rs("seprodnamen") <> "" and IsNull(Rs("seprodnamen")) = False then
			seprodnamen = Rs("seprodnamen")
		else 
			seprodnamen = seprodname
		end if
		
		temp = Rs("slevel")
	
		
		if temp = "" or IsNull(temp) then
			l(0) = "seprodname"
			s(0)= seprodnamen
		else
			templv = Split(temp,",")
			i = 0
			for each x in templv
				if  x = "seprodname" then
					l(i) = x
					s(i) = oraFormatStr(seprodnamen)
				else
					l(i) = x
					s(i) = oraFormatStr(Rs(x))
				end if 
				
				'"l"&i = 1
				's&i = Rs(x)
				i = i + 1
			next
		end if
		'response.write "ssssssss = "&s(0)
		'response.end
		d_day=day(dadddate)
		d_month=month(dadddate)
		d_year=year(dadddate)
		sempno = RS("sempno")
		sCurcode=Rs("sCurcode")
		sUnitcode=Rs("sUnitcode")
		hazardno = Rs("hazardno")
		hazardstatus = Rs("hazardstatus")
		sbread = " readonly"
	Else
		sMode = "ADD"
	End If
End Sub	

%>
<%
' This routines presents the form to the user. Depending on the mode, the
' form fields are filled with data or blank.
Sub Show_Form()
	' Set the title of the page depending on if an error occurred during
	' form submission or not
%> <html> <head> 
<title>Product</title>

<meta http-equiv="Content-Type" content="text/html; charset=windows-874">

<link rel="stylesheet" href="../styles.css">
<SCRIPT LANGUAGE="JavaScript">
<!--

function OpenNewWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

-->
</SCRIPT>
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
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  Product </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2">


<script type="text/javascript">
	
	function checkSave()
	{	
		var prdcode = document.getElementById("SPRODCODE").value;
		if(gtrim(prdcode) == "" || prdcode == null)
		{
			alert("Please input Product Code !!");
			return false;
		}
		var chk_productname = "No";
		var selt = "";
		var seltinput = ""
		for(i =0;i<8;i++)
		{
			i = i+1;
			selt = document.getElementById('l'+i).value;
			if(selt == "seprodname")
			{
				chk_productname = "Yes";
				seltinput = document.getElementById('s'+i).value;
				if(seltinput == null || gtrim(seltinput) == "")
				{
					chk_productname = "No";
				}
			}
			i = i-1;
		}
		
		if(chk_productname == "No"){
			alert("Please Choose & input Product Name !!");
			return false;
		}


		var vendor = document.getElementById("sEname").value;
		var uprice = document.getElementById("NPRICE").value;
		var Currency = document.getElementById("sCurrency").value;
		var Unit = document.getElementById("sUnit").value;
		var hazardstatus = document.getElementsByName('hazardstatus');
		var hazardnoz = document.getElementById('hazardno').value
		var chkHazard = "";
		
		for (var i = 0, length = hazardstatus.length; i < length; i++) {
			if (hazardstatus[i].checked == true){
				chkHazard = hazardstatus[i].value;
			}
		}
		
		if(chkHazard != "Yes" && chkHazard != "No"){
			alert("Please Choose Hazardous Substance !!");
			return false;
		}else if(chkHazard == "Yes"){
			if(hazardnoz == "null"){
				alert("Please Choose Hazardous SubStance!!!");
				return false;
			}
		}
		
		//alert(Currency);

		if(gtrim(vendor) == "" || vendor == null) 
		{
			alert("Please fill verdor !!");
			return false;
		}else if(uprice == "" || uprice == null){
			alert("Please fill UPrice !!");
			return false;
		}else if(Currency == "no"){
			alert("Please Choose Currency !!");
			return false;
		}else if(Unit == "no"){
			alert("Please Choose Unit !!");
			return false;
		}
		
	}
	
	function checkText()
	{
		var elem = document.getElementById('SPRODCODE').value;
		
		if(!elem.match(/^([a-z0-9\_.-])+$/i)) 
		{
			if(elem == ""){}else
			{
				alert("Input a-Z, A-Z, 0-9 , - and _(underscore) Only !!!");
				document.getElementById('SPRODCODE').value = elem.substring(0,elem.length -1);
				document.getElementById('SPRODCODE').focus();
			}
		} 
	}
	function checkText2()
	{
		var elem = document.getElementById('STPRODNAME').value;
		var str = '|';
		var chk;
		chk = (elem.indexOf(str));
		if(chk == -1){

		}else{
			alert("Can't Input | (pipe character) !!!");
			document.getElementById('STPRODNAME').value = document.getElementById('STPRODNAME').value.replace("|", "");
			document.getElementById('STPRODNAME').focus();
		}	
	}

	
	function chkText(name)
	{	
		var chk;
		var numb = name.substring(1);
		var selt = document.getElementById('l'+numb).value;
		if(selt == "seprodname"){
			var elem = document.getElementById(name).value;
			chk = (elem.indexOf("|"));
			if(chk == -1){

			}else{
				alert("Can't Input | (pipe character) !!!");
				document.getElementById(name).value = document.getElementById(name).value.replace("|", "");
				document.getElementById(name).focus();
			}
			//if(!elem.match(/^([a-z0-9\_.])+$/i)) 
			//{
			//	if(elem == ""){}else
			//	{
			//		alert("Input a-z, A-Z, 0-9 and _(underscore) Only !!!");
			//		document.getElementById(name).value = elem.substring(0,elem.length -1);
			//		document.getElementById(name).focus();
			//	}
			//}
		}
	}
	
	
	 
	 function OnChangeRadio (radio) {
			if(radio.value == "Yes")
			{
				document.getElementById('Hazardno').disabled = false;
			}else{
				document.getElementById('Hazardno').disabled = true;
				document.getElementById('Hazardno').value = "null";
			}
        }
		function gtrim(str)
		{
			whitespace = new String(" \t\n\r");
			var s = new String(str);
			if (whitespace.indexOf(s.charAt(0)) != -1) {
			var j=0, i = s.length;
			while (j < i && whitespace.indexOf(s.charAt(j)) != -1)
				j++;
				s = s.substring(j, i);
			}
			return s;
		}
		
		
		
</script>
				
                  <form method="post" action="../datamanager/Imagesave.asp" enctype="multipart/form-data" name="frmName" 
				  id="frmName" onsubmit="return checkSave()">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">*Vendor 
                          : </font></b></td>
                        <td width="71%"> 
                          <input type="text" name="sEname" id="sEname"size="45" value="<%Show_Vender(sVendcode)%>" maxlength="60">
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%">| <a href="JavaScript:OpenNewWindow('../Datamanager/SearchVender.asp?sProdcode=<%=sProdcode%>&sExptype=<%=sExptype%>&sCurcode=<%=sCurcode%>&sUnitcode=<%=sUnitcode%>','SearchProduct','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0 ,width=600,height=500')">Search 
                          Vendor</a> | </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">*Expense 
                          Type : </font></b></td>
                        <td width="71%"> 
<%								Show_Pexptype(sexptype)  %>


						 </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Expense Group
                          : </font></b></td>
                        <td width="71%"> 
						<%	Show_PCattype(sCattype) %>
						 </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Expense Sub-Group : </font></b></td>
                        <td width="71%"> 
                            <% Show_PCatalog(sCatID) %> 
                        </td>
                      </tr>
<%if sMode="ADD" then%>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%">|

<a href="JavaScript:OpenNewWindow('../Datamanager/SearchExpense.asp?sProdcode=<%=sProdcode%>&sExptype=<%=sExptype%>&svendcode=','SearchExp','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">Search 
                          Expense Type</a> | </td>
                      </tr>
<%End if%>
                      <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">*Product 
                          Code : </font></b></td>
                        <td width="71%"> <%If sMode="EDIT" then 
Response.write SPRODCODE
Else
%> 

                          <input type="text" name="SPRODCODE" id="SPRODCODE" size="20" maxlength="12" onblur="checkText()" value="<%=SPRODCODE%>" <%=sbread%>>
                          <%
Response.write img_sprodcode
End if
%> 

</td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Product Code Data 
                          (Character,Don't use &quot; ' , ; &lt; &gt; / \ | &amp; 
                          % # @ + ) :Not more than 12 Characters .</font></td>
                      </tr>
					  <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">*Choose Product 
                          Name : </font></b></td>
                        <td width="79%"></td>
                      </tr>
					  	<tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><select name="l1" id="l1"size=1>
                            <%List_Detail(l(0))%> 
                          </select> : </font></b></td>
                        <td width="71%"> 
								
									<input type="text" name="s1" id="s1"size="45" onblur="chkText(this.name)" value="<%=s(0)%>">
                        </td>
              
                     	<tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><select name="l2"id="l2"size=1>
                            <%List_Detail(l(1))%> 
                          </select>  : </font></b></td>
                        <td width="71%"> 
									<input type="text" name="s2" id="s2" size="45" onblur="chkText(this.name)" value="<%=s(1)%>">
                        </td>
					  
					      </tr>
					     <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><select name="l3" id="l3"size=1>
                            <%List_Detail(l(2))%> 
                          </select>  : </font></b></td>
                        <td width="71%"> 
									<input type="text" name="s3" id="s3"size="45" onblur="chkText(this.name)" value="<%=s(2)%>" >
                        </td>
                      </tr>
					  	 <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><select name="l4" id="l4"size=1>
                            <%List_Detail(l(3))%> 
                          </select> : </font></b></td>
                        <td width="71%"> 
									<input type="text" name="s4" id="s4" size="45" onblur="chkText(this.name)" value="<%=s(3)%>">
                        </td>
                      </tr>
					    	<tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><select name="l5" id="l5" size=1>
                            <%List_Detail(l(4))%> 
                          </select>  : </font></b></td>
                        <td width="71%"> 
									<input type="text" name="s5" id="s5" size="45" onblur="chkText(this.name)" value="<%=s(4)%>">
                        </td>
                      </tr>
					   <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><select name="l6" id="l6" size=1>
                            <%List_Detail(l(5))%> 
                          </select>  : </font></b></td>
                        <td width="71%"> 
									<input type="text" name="s6" id="s6"size="45" onblur="chkText(this.name)" value="<%=s(5)%>"  >
                        </td>
                      </tr>
					    <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><select name="l7" id="l7" size=1>
                            <%List_Detail(l(6))%> 
                          </select> 
                          : </font></b></td>
                        <td width="71%"> 
									<input type="text" name="S7" id="s7"size="45" onblur="chkText(this.name)" value="<%=s(6)%>">
                        </td>
                      </tr>
					    <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><select name="l8" id="l8" size=1>
                            <%List_Detail(l(7))%> 
                          </select>  : </font></b></td>
                        <td width="71%"> 
									<input type="text" name="s8" id="s8"size="45" onblur="chkText(this.name)" value="<%=s(7)%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Name (Thai) : </font></b></td>
                        <td width="71%"> 
                          <input type="text" name="STPRODNAME" id="STPRODNAME"size="45" onblur="checkText2()" value="<%=stprodname%>"  maxlength="90">
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right" valign="top">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">Product Name 
                          	(Not more than 90 Characters).
							:Don't use |
							</font></td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor's ProductID 
                          : </font></b></td>
                        <td width="71%"> 
                          <input type="text" name="SPRODID" size="30" value="<%=sprodid%>" maxlength="30">
                        </td>
                      </tr>
                    
                 
                      <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Account Code. 
                          : </font></b></td>
                        <td width="71%"> 
                          <input type="text" name="SACCCODE" size="20" maxlength="20" value="<%=sAcccode%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right" valign="top">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">Account Code (Character) 
                          :Not more than 20 Characters .</font></td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"  color="#CC3333"> 
                          Description (Thai): </font></b></td>
                        <td width="71%"> 
                          <textarea name="STDESC" cols="45" rows="3"><%=stdesc%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">Product Description 
                          (Character,Thai Language) :Not more than 90 Characters 
                          .</font></td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          Description (English): </font></b></td>
                        <td width="71%"> 
                          <textarea name="SEDESC" cols="45" rows="3"><%=sedesc%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">Product Description 
                          (Character,English Language) :Not more than 90 Characters 
                          .</font></td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"  color="#CC3333">Product Image
                          : </font></b></td>
                        <td width="71%">
						<a href="../datamanager/productimages/<%=year(date)%>/<%=simage%>" TARGET="_Blank"><%=simage%> </a>&nbsp;&nbsp;
							<%if simage<>"" then%><input type="checkbox" name="del" value="del">Del<%end if%>
                          <input type="file" name="sImage" size="18" maxlength="18" value="<%=sImage%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">Product Image  
                          </font></td>
                      </tr>

                      <tr> 
                        <td width="29%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"  color="#CC3333">*UPrice 
                          : </font></b></td>
                        <td width="71%"> 
                          <input type="text" name="NPRICE" id="NPRICE"size="18" maxlength="18" value="<%=nprice%>" 
						  onKeyUp="if(this.value*1!=this.value){alert('Number Only!!');this.value='';}" >
						  <%=img_nprice%>
						</td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">Product Price (Numeric) 
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">*Currency 
                          : </font></b></td>
                        <td width="71%"> 
                          <select name="sCurrency" id="sCurrency" size=1>
							<option value="no">Please Choose Currency!!</option>
                            <%List_Currency(sCurcode)%> 
                          </select>
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">Select From List </font></td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">*Unit 
                          : </font></b></td>
                        <td width="71%"> 
                          <select name="sUnit" id="sUnit" size=1>
							<option value="no">Please Choose Unit!!</option>
                            <%List_Unit(sUnitcode)%> 
                          </select>
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">Measure Unit (Select 
                          From List) </font></td>
                      </tr>
					  

					    <tr> 
                        <td width="29%" valign="top" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">*Hazardous Substance 
                          : </font></b></td>
                        <td width="71%"> 
						<input type="radio" name="hazardstatus" id="hazardstatus" onclick="OnChangeRadio(this)" value="Yes" <%if IsNull(hazardno) = false then %>checked="checked"<% end if%>>Yes
							<select name="hazardno" size=1 id="hazardno">
								<option value="null">Please Choose !!!</option>
								<%List_Hazardno(hazardno)%>
							</select>
						
						  <br>
						 <input type="radio" name="hazardstatus" id="hazardstatus" onclick="OnChangeRadio(this)" value="No" <%if IsNull(hazardno) = true then %>checked="checked"<% end if%>> No
                        </td>
                      </tr>
					  
					  
                      <!--                   <tr> 
                        <td width="19%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">AddDate</font></td>
                        <td width="81%">วันที่ 
                          <select name="dadddate_day">
						  
                            <%   ' if d_day="" then d_day=day(date)
												'		Create_daylist(d_day) %> 
                          </select>
                          เดือน 
                          <select name="dadddate_month">
                            <% 'if d_month="" then d_month=month(date)
														'		Create_MonthList(d_month)%> 
                          </select>
                          พศ. <%' if d_year="" then d_year=year(date) %> 
                          <input type="text" name="dadddate_year" size="8" maxlength="4" value="<%=d_year%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Empno</font></td>
                        <td width="81%"> 
                          <input type="text" name="SEMPNO" size="30" maxlength="30" value="<%=sempno%>">
                        </td>
                      </tr> --> 
                    </table>
                    <br>

					<input type=HIDDEN name='sVendcode' value="<%=sVendcode%>">
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
					<input type=HIDDEN name='SEXPTYPE' value="<%=SEXPTYPE%>">
					<input type=HIDDEN name='SCATTYPE' value="<%=SCATTYPE%>">
					<input type=HIDDEN name='SCATID' value="<%=SCATID%>">
<%if sMode="EDIT" and bError=false then%>
					<input type=HIDDEN id='SPRODCODE' name='SPRODCODE' value="<%=SPRODCODE%>">
<%End if%>

                    <input type="submit" name="Submit" value="Save">
                    <input type="reset" name="Reset" value="Reset">
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp; </td>
              </tr>
            </table>
<font color="red"><blink><b>*Required Field.</b></blink></font>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="../datamanager/datamanager.asp">Data Manager</a></font></p>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../security/images/y04.jpg" width="12" height="12"></td>
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
<% End Sub %>

<%
Sub Update_Product(sUpdateMode)
	Dim  Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sprodcode = replaceChars(Request.Form("sprodcode"))
		

		sexptype = Request.Form("sexptype")
		svendcode = Request.Form("svendcode")
		sprodid = Request.Form("sprodid")
		stprodname = Request.Form("stprodname")
		seprodname = Request.Form("seprodname")
		smodel = Request.Form("smodel")
		sAcccode = Request.Form("sAcccode")
		stdesc = Request.Form("stdesc")
		sedesc = Request.Form("sedesc")
		nprice = Request("nprice")
		sCatID = Request.Form("sCatID")
		simage=Request.Form("simage")
		sempno = Session("sempno") 'Request.Form("sempno")
		sCurcode=Request("sCurrency")
		sUnitcode=Request("sUnit")
		hazardno=Request("hazardno")
		hazardstatus=Request("hazardstatus")
		
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty
chk_Product()
	If (Trim(sprodcode) = "") Or chkProduct=1 Then
		bError = True
		img_sprodcode = sErrImage
	End If

'	If (Trim(sexptype) = "") Then
'		bError = True
'		img_sexptype = sErrImage
'	End If

'	If (Trim(svendcode) = "") Then
'		bError = True
'		img_svendcode = sErrImage
'	End If
 if  Not IsNumeric(NPRICE) then  
 		nprice=0
		img_nprice = sErrImage
	End If

	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  PProduct WHERE sprodcode='" & sprodcode &"'and sexptype='" & sexptype &"'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "pproduct", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("sprodcode") = Ucase(sprodcode)
		Rs("sexptype") = sexptype
		Rs("scatID") = scatID
		Rs("svendcode") = svendcode
		Rs("sprodid") = sprodid
		Rs("stprodname") = stprodname
		Rs("seprodname") = seprodname
		Rs("smodel") = smodel
		Rs("sAcccode") = sAcccode
		Rs("stdesc") = stdesc
		Rs("sedesc") = sedesc
		Rs("nprice") = FormatNumber(nprice,4)
		Rs("simage") = sfilename
	   	Rs("dadddate") = date
		Rs("sempno") = sempno
		RS("sCurcode")=sCurcode
		RS("sUnitcode")=sUnitcode
		RS("hazardno")=hazardno
		RS("hazardstatus")=hazardstatus
	  Rs.Update	

	  Rs.Close
	  Set Rs = Nothing
'if sMode = "ADD" then
'     Response.Redirect "pproductdesc.asp?sprodcode="&sprodcode&"&sexptype="&sexptype&"&sCatid="&sCatID
'else

     Response.Redirect "pproductedt.asp?sprodcode="&sprodcode&"&sexptype="&sexptype&"sCatid="&sCatID&"sStatus=del"
'End if
     'Response.Redirect "javascript:history.back()"
	else

		Show_Form()
	End If	' if Error 
End Sub	

Sub List_Detail(detail)
	 Dim pn,dn,sp,sc,sd,sb,sm,al,md
	 pn = ""
	 dn = ""
	 sp = ""
	 sc = ""
	 sd = ""
	 sb = ""
	 sm = ""
	 al = ""
	 md = ""
	 if detail = "seprodname" then
		pn = "selected"
	 elseif detail = "sdrawnum" then
		dn = "selected"
	 elseif detail = "sspec" then
		sp = "selected"
	 elseif detail = "scountry" then
		sc = "selected"
	 elseif detail = "sdimen" then
		sd = "selected"
	 elseif detail = "sbrand" then
		sb = "selected"
	 elseif detail = "smarkern" then
		sm = "selected"
	 elseif detail = "smodel" then
		md = "selected"
	 else
		al = "selected"
	 end if
	 Response.write "<option value='' "&al&">Choose Detail</option>"
	 Response.write "<option value='seprodname' "&pn&">Product Name</option>"
	 Response.write "<option value='sdrawnum' "&dn&">Drawing Number</option>"
	 Response.write "<option value='sspec' "&sp&">Specification</option>"
	 Response.write "<option value='scountry' "&sc&">Country of origin</option>"
     Response.write "<option value='sdimen' "&sd&">Dimension</option>"
	 Response.write "<option value='sbrand' "&sb&">Brand Name</option>"
	 Response.write "<option value='smarkern' "&sm&">Maker Name</option>"
	 Response.write "<option value='smodel' "&md&">Model Name</option>"
End Sub


Sub List_PCatalog(sCatID)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pCatalog"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sCatID") = sCatID Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("sCatID")& "'" & cSelect &">"& Rs2("sCatalog") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	set Rs2=Nothing
End Sub

Sub List_PCattype(sCattype)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pCattype"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sCattype") = sCattype Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("sCattype")& "'" & cSelect &">"& Rs2("sCatdesc") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	set Rs2=Nothing
End Sub

Sub Show_pCatalog(sCatID)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pCatalog Where sCatID='" & sCatID & "'" &_
		" Order By sCatalog"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sCatID") = sCatID Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write(Rs2("sCatalog"))
			Rs2.MoveNext
		Loop
	End IF
End Sub
Sub Show_pCattype(sCattype)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pCattype Where sCattype='" & sCattype & "'" &_
		" Order By sCatDesc"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sCattype") = sCattype Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write(Rs2("sCatdesc"))
			Rs2.MoveNext
		Loop
	End IF
End Sub
%>
<%

Sub replaceChars(sstr)
	Dim elem 
	Dim str
	str = sstr
	for Each elem In Array("/", "\", ":", "*", "?", "<", ">", "|", "#", Chr(34),"'",",",";","&","&","%","@","+")
		str =  Replace(str, elem, "")
	Next
	
	replaceChars = str
	
End Sub


Sub chk_product()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pproduct Where sprodcode='" & sprodcode & "'" &_
		" and sExptype='" & Request("sExptype") & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		chkproduct=1
	else
		chkproduct=0
	End IF
End Sub
%>
