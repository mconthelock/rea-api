<!-- #include file="adovbs.inc" -->
<!-- #include file="conn.asp" -->

<html>
<body>

<font size=+2>ÕË“π‰ø≈Ï text</font><hr>

<%
dim blank
blank="00"
  Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
  MsgFile = Server.MapPath ("vendor1.csv")
  On Error Resume Next 
  Set InStream= FileObject.OpenTextFile (MSGFile, 1, False, False)

 Set Rs = Server.CreateObject("ADODB.RecordSet")
 Rs.Open "pVender", conn, adOpenKeySet, adLockPessimistic, adCmdTable

  Do While InStream.AtEndOfStream <> True
	  str = Instream.ReadLine
	  'Create a recordset

	  Rs.AddNew

	  npos = InStr(Str,",")
      if npos > 0 then
         str1 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if  
	  npos = InStr(Str,",")
      if npos > 0 then
         str2 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if  
	  npos = InStr(Str,",")
      if npos > 0 then
         str3 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if  
	  npos = InStr(Str,",")
      if npos > 0 then
         str4 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if
	  npos = InStr(Str,",")
      if npos > 0 then
         str5 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if
	  npos = InStr(Str,",")
      if npos > 0 then
         str6 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if   
	  npos = InStr(Str,",")
      if npos > 0 then
         str7 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if   
	  npos = InStr(Str,",")
      if npos > 0 then
         str8 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if   
	  npos = InStr(Str,",")
      if npos > 0 then
         str9 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if   
	  npos = InStr(Str,",")
      if npos > 0 then
         str10 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if
	  npos = InStr(Str,",")
      if npos > 0 then
         str11 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if   
   	  npos = InStr(Str,",")
      if npos > 0 then
         str12 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if
	  npos = InStr(Str,",")
      if npos > 0 then
         str13 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if
	  npos = InStr(Str,",")
      if npos > 0 then
         str14 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if   
	  npos = InStr(Str,",")
      if npos > 0 then
         str15 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if
	  npos = InStr(Str,",")
      if npos > 0 then
         str16 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if   
   	  npos = InStr(Str,",")
      if npos > 0 then
         str17 = mid(Str,1,npos-1)
		 str = mid(Str,npos+1,(len(str)-npos))
      end if   

	  'Response.write Str &"<br>"
  	   'Response.write Str2 &"<br>"
	IF STR1 <> "" THEN
								response.write "<font color='green'>" & str1 & "</font><br>"
							     RS("sVendcode") = str1
								  RS("sEName") = str2
								  RS("sTName") = str2
								RS("sTaddress")= str3 & str4
									RS("sEaddress")= str3 & str4
									RS("scity")= str5
								  RS("sState") = str6
								  RS("sPost") = str7
								  RS("sTelno") = str8
							     RS("sSearchKey") = str9
								  RS("sContact") = str10
								  RS("sPurAddr1") = str11
								  RS("sPurAddr2") = str12
									RS("sPurCity")= str13
								  RS("sPurState") = str14
								  RS("sPurPost") = str15
								  RS("sPurContact") = str16
								  RS("sPurPhone") = str17
								  RS("dAddDate") = date()
								  Rs.Update	
	END IF      
  Loop
  Rs.Close
  Set Rs = Nothing

  Set Instream=Nothing

%> 


</body>
</html>
