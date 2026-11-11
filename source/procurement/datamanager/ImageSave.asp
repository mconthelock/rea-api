<% Option Explicit %>
<!-- #include file ="../adovbs.inc"  -->
<!-- #include file ="../Conn.asp" -->
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
	'########## DERMPAN KHOOMRAKSA ##########
	'I changed formatNumber(xxx, y) to formatNumber(xxx, 3) on 11-NOV-2002
	'########## END DERMPAN KHOOMRAKSA ##########
	
	' Author: Ekkapan Kongchoo

     dim objsmartupload,rsimagelib,rsOra	                      'for obj
     dim spath,fd,file,sql,smode,check1,check2,chkEdit          'other variable
     dim scorpcode,simgid,sfilename,stimgdesc,seimgdesc   'for variable in table     
	 dim sImageType,sCatID,sCurcode,sUnitcode,sAcccode,sCattype,chkProduct,hazardno,hazardstatus
	 Dim chkfilename,sEname, hazardsql
     Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sedesc,nprice,simage,dadddate,sempno
	 Dim sdrawnum, sspec, scountry, sdimen, sbrand, smarkern, seprodnamen
'     set objsmartupload = server.createobject("aspSmartUpload.SmartUpload")'create Obj     
'     objsmartupload.MaxFileSize = 1000000                                   'Max filesize 
'     objsmartupload.Upload                                     'upload to objsmartupload

'     ObjSmartUpload.DeniedFilesList = "bat,exe,com,asp,dll"    

	set objsmartupload  = server.createObject("SoftArtisans.FileUp")	 	

	Response.CharSet = "windows-874" 	
	'objsmartupload.CodePage = "windows-874"
     smode = objsmartupload.form("EVENT")
'	 	response.end
     'data from pre page  
	 dim i
	 dim chksename , cntproname , cntmodel, cntdwn , cntspec , cntcountry , cntdimen, cntbrand, cntmarker
	 dim slevel
	 
	 chksename = 0
	 cntproname = 0
	 cntmodel = 0
	 cntdwn = 0
	 cntspec = 0
	 cntcountry = 0
	 cntdimen = 0
	 cntbrand = 0
	 cntmarker = 0
	 seprodname = ""
	 For i = 1 To 8
		
			if objsmartupload.Form("l"&i) <> "" then
				if seprodname = "" then
					seprodname = Ucase(objsmartupload.Form("s"&i))
					slevel = objsmartupload.Form("l"&i)
				else
					seprodname = seprodname&" "&Ucase(objsmartupload.Form("s"&i))	
					slevel = slevel&","&objsmartupload.Form("l"&i)
				end if
			end if
		  	if objsmartupload.Form("l"&i) = "seprodname" then
				chksename = 1
				seprodnamen = Left(Ucase(objsmartupload.Form("s"&i)),255)
				cntproname = cntproname + 1
			end if
			if objsmartupload.Form("l"&i) = "smodel" then
				cntmodel= cntmodel + 1
				smodel = Left(Ucase(objsmartupload.Form("s"&i)),30)
			end if
			if objsmartupload.Form("l"&i) = "sdrawnum" then
				cntdwn = cntdwn + 1
				sdrawnum = Left(Ucase(objsmartupload.Form("s"&i)),50)
			end if
			if objsmartupload.Form("l"&i) = "sspec" then
				cntspec = cntspec + 1
				sspec = Left(Ucase(objsmartupload.Form("s"&i)),50)
			end if
			if objsmartupload.Form("l"&i) = "scountry" then
				cntcountry = cntcountry + 1
				scountry = Left(Ucase(objsmartupload.Form("s"&i)),30)
			end if
			if objsmartupload.Form("l"&i) = "sdimen" then
				cntdimen = cntdimen + 1
				sdimen = Left(Ucase(objsmartupload.Form("s"&i)),30)
			end if
			if objsmartupload.Form("l"&i) = "sbrand" then
				cntbrand = cntbrand + 1
				sbrand = Left(Ucase(objsmartupload.Form("s"&i)),50)
			end if
			if objsmartupload.Form("l"&i) = "smarkern" then
				cntmarker = cntmarker + 1
				smarkern = Left(Ucase(objsmartupload.Form("s"&i)),50)
			end if
			'response.write "l"&i&" = "&objsmartupload.Form("l"&i)&"<br/>"
			'response.write "s"&i&" = "&objsmartupload.Form("s"&i)&"<br/>"
	 Next
	' response.write "seprodname = "&seprodname&"<br/>"
	'response.write slevel
	 'response.end
	 
     if smode = "ADD" then
'  	   scorpcode  = session("gbcorpcode")
'       simgid     = mid(year(Date),3,2)&month(Date)&Day(date)&Hour(Now)&Minute(Now)&Second(now)
		sprodcode = replaceChars(Ucase(rtrim(objsmartupload.Form("sprodcode"))))
		sexptype = objsmartupload.Form("sexptype")
		sCattype = objsmartupload.Form("sCattype")
		svendcode = objsmartupload.Form("svendcode")
		sprodid = objsmartupload.Form("sprodid")
		stprodname = objsmartupload.Form("stprodname")
		sAcccode = objsmartupload.Form("sAcccode")
		stdesc = objsmartupload.Form("stdesc")
		sedesc = objsmartupload.Form("sedesc")
		nprice = objsmartupload.Form("nprice")
		sCatID = objsmartupload.Form("sCatID")
		simage=objsmartupload.Form("simage")
		sempno = Session("sempno") 'Request.Form("sempno")
		sCurcode=objsmartupload.Form("sCurrency")
		sUnitcode=objsmartupload.Form("sUnit")
		hazardstatus=objsmartupload.Form("hazardstatus")
		if hazardstatus = "Yes" then
			hazardstatus = "1"
			hazardno=objsmartupload.Form("hazardno")
		else
			hazardstatus = "0"
		end if
     else
'       scorpcode  = objsmartupload.form("scorpcode")
'       simgid     = objsmartupload.form("simgid")
		sprodcode = rtrim(objsmartupload.Form("sprodcode"))
		sexptype = objsmartupload.Form("sexptype")
		sCattype = objsmartupload.Form("sCattype")
		svendcode = objsmartupload.Form("svendcode")
		sEname = objsmartupload.Form("sEname")
		sprodid = objsmartupload.Form("sprodid")
		stprodname = objsmartupload.Form("stprodname")
		sAcccode = objsmartupload.Form("sAcccode")
		stdesc = objsmartupload.Form("stdesc")
		sedesc = objsmartupload.Form("sedesc")
		nprice = objsmartupload.Form("nprice")
		sCatID = objsmartupload.Form("sCatID")
		simage=objsmartupload.Form("simage")
		sempno = Session("sempno") 'Request.Form("sempno")
		sCurcode=objsmartupload.Form("sCurrency")
		sUnitcode=objsmartupload.Form("sUnit")
		'hazardno=objsmartupload.Form("hazardno")
		hazardstatus=objsmartupload.Form("hazardstatus")
		if hazardstatus = "Yes" then
			hazardstatus = "1"
			hazardno=objsmartupload.Form("hazardno")
		else
			hazardstatus = "0"
		end if
     end if
'     stimgdesc  = objsmartupload.form("stimgdesc")
'     seimgdesc  = objsmartupload.form("seimgdesc")

	if 	objsmartupload.Form("del")="del" then
		del()
	end if

	
	 spath = request.serverVariables("APPL_PHYSICAL_PATH")      'ª×èÍ ·Ò§physic ¢Í§ server
'	 response.write spath
''    spath = spath & "procurement\datamanager\ProductImages"          'check path corp
  '  response.write spath
    spath = spath & "datamanager\ProductImages\" & year(date) 			'check path corp
'	response.write spath

	objsmartupload.path = spath
'     fd = objsmartupload.save(spath)           'save file to server at spath

'	 for each file in objsmartupload.files
'	 	 sfilename = file.filename             'ª×èÍ¢Í§ file ·Õèload ÁÒ
' 	 next	
'######################################
	If Not objsmartupload.Form("simage").IsEmpty Then
			dim multiPart, element, fileName, name
			For each multiPart in objsmartupload.Form
				For each element in objsmartupload.FormEx(multiPart)
						If IsObject( element ) then
							If not ( element.IsEmpty ) then
								sfilename = mid(objsmartupload.Form("simage").UserFilename,instrrev(objsmartupload.Form("simage").UserFilename,"\")+1)
								sfilename = name &"" & sfilename
								sfilename = replace(sfilename," ","")
								response.write spath & "\" &  sfilename
								objsmartupload.Form("simage").SaveAs spath & "\" &  sfilename
'response.write fileName
'response.end
							End if 
						End if	
					Next
			Next					
	End if
'######################################
	 check_filename() 

if chkfilename=0 then
     set rsimagelib = server.createobject("ADODB.Recordset")
	 set rsOra = server.createobject("ADODB.Recordset")
     if smode = "ADD" then 
		If (Trim(sprodcode) <> "" ) Then
			sql = "SELECT * FROM  pproduct WHERE sprodcode ='" & Ucase(sprodcode) &"'"
			rsOra.Open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
			if Not rsOra.EOF then
					Response.write "Duplicate Product Code "
					Response.end
			end if
			rsOra.close
		End if
'comment :because change Compoent
'         fd = objsmartupload.save(spath)           'save file to server at spath

    	 'create recordset & save to DB
         sql = "select * from pproduct"
         rsimagelib.open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
	     rsimagelib.addnew
			rsimagelib("sprodcode") = Ucase(sprodcode)
			if sExptype="" then
				response.redirect "productError.asp"
			else
				rsimagelib("sexptype") = sexptype
			end if
				if sCattype="" then
					sCattype="00"
				End if
			rsimagelib("sCattype") = sCattype
				if sCatID="" then
					sCatID="00"
				End if
			rsimagelib("scatID") = scatID
			rsimagelib("svendcode") = svendcode
			rsimagelib("sprodid") = sprodid
			rsimagelib("stprodname") = stprodname
			
			
			
			
			if hazardstatus = 1 then
				if IsNull(hazardno) = true then
				hazardsql = " hazardno = '" & hazardno & "', hazardstatus='1' "
					response.redirect "productError.asp"
				end if
				hazardsql = " hazardno = '" & hazardno & "', hazardstatus='1' "
			else 
				hazardno = null
				hazardsql = " hazardno = null ,hazardstatus='0' "
			end if
			
			if chksename = 0 then 
				response.redirect "productError.asp"
			End if
			if cntproname > 1 or cntmodel > 1 or cntdwn > 1 or cntspec > 1 or cntcountry > 1 or cntdimen > 1 or cntbrand > 1 or cntmarker > 1 then
				response.redirect "dupfield.asp"
			End if
			if Len(seprodname) > 255 then
				response.redirect "chklength.asp"
			end if
			rsimagelib("seprodnamen") = seprodnamen
			rsimagelib("slevel") = slevel
			rsimagelib("seprodname") = seprodname
			rsimagelib("sAcccode") = sAcccode
			rsimagelib("stdesc") = stdesc
			rsimagelib("sedesc") = sedesc
			if nprice="" then 
				nprice=0
			End if
			rsimagelib("nprice") = FormatNumber(nprice,4)
'comment :because change Compoent			
			rsimagelib("simage") = sfilename

		   	rsimagelib("dadddate") = date
			rsimagelib("sempno") = sempno
			rsimagelib("sCurcode")=sCurcode
			rsimagelib("hazardno")=hazardno
			rsimagelib("hazardstatus")=hazardstatus
			rsimagelib("sUnitcode")=sUnitcode
			rsimagelib("smodel") = smodel
			rsimagelib("sdrawnum") = sdrawnum
			rsimagelib("sspec") = sspec
			rsimagelib("scountry") = scountry
			rsimagelib("sdimen") = sdimen
			rsimagelib("sbrand") = sbrand
			rsimagelib("smarkern") = smarkern
			'For i = 1 To 8
			'	if objsmartupload.Form("l"&i) <> "" and objsmartupload.Form("l"&i) <> "seprodname" then
			'		rsimagelib(objsmartupload.Form("l"&i)) = Ucase(objsmartupload.Form("s"&i))
			'	end if
			'Next
	     rsimagelib.update    
		'response.write "seprodname = "&seprodname&"<br/>"	
		'response.end
     Response.Redirect "../datamanager/pproductdesc.asp?sprodcode="&sprodcode&"&sexptype="&sexptype&"&sCatid="&sCatID&"&sCattype="&sCattype
	 response.end
     else 'smode = edit
		 'update Table
		 
		  	sql = "update pproduct set " _
			& " sCattype='" & sCattype &"',scatID='" & scatID & "', svendcode='" & svendcode & "'" _
			& " ,	sprodid='" & sprodid & "',stprodname='" & stprodname & "', seprodname='" & Ucase(seprodname) & "'" _
			& " , smodel='" & smodel & "', sAccode='" & sAcccode & "',stdesc='" & stdesc & "',sedesc='" & sedesc & "'" _
			& " , nprice='" & FormatNumber(nprice,4) & "',simage='" & Lcase(sfilename) & "', dadddate= '" & date & "'" _
			& " , sCurcode = '" & sCurcode & "', sUnitcode='" & sUnitcode & "'," & hazardsql _
			& "  where sprodcode='" & Ucase(sprodcode) & "' and sexptype='" & sexptype & "'"
'			 response.write stprodname
'			 response.end
'			response.write sql
'			response.end	
		 chkEdit = 0
		 sql = "select count(logid) as id from logproduct where sprodcode='" & sprodcode & "'"
		 rsOra.open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
		 if rsOra("id") = "0" then
			chkEdit = 1
		 end if
		 'rsOra.close
		' response.write sql&rsOra("id")
		' response.end
         sql = "SELECT * from pproduct WHERE sprodcode='" & sprodcode & "'" &_
         " and sexptype='" & sexptype & "'"
         rsimagelib.open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
          If Not rsimagelib.EOF Then
'			rsimagelib("stimgdesc") = stimgdesc
'			rsimagelib("seimgdesc") = seimgdesc
			
			rsimagelib("sprodcode") = Ucase(sprodcode)
			rsimagelib("sexptype") = sexptype
			rsimagelib("sCattype") = sCattype
			rsimagelib("scatID") = scatID
				if sEname="" then
					sVendcode=""
				End if
			rsimagelib("svendcode") = svendcode
			rsimagelib("sprodid") = sprodid
			rsimagelib("stprodname") = stprodname
			if chksename=0 then 
				response.redirect "productError.asp"
			End if	 
			if cntproname > 1 or cntmodel > 1 or cntdwn > 1 or cntspec > 1 or cntcountry > 1 or cntdimen > 1 or cntbrand > 1 or cntmarker > 1 then
				response.redirect "dupfield.asp"
			End if
			if Len(seprodname) > 255 then
				response.redirect "chklength.asp"
			end if
			rsimagelib("seprodnamen") = seprodnamen
			rsimagelib("slevel") = slevel
			rsimagelib("seprodname") = seprodname
			rsimagelib("sAcccode") = sAcccode
			rsimagelib("stdesc") = stdesc
			rsimagelib("sedesc") = sedesc
			if nprice="" then 
				nprice=0
			End if
		
			if cdbl(rsimagelib("nprice"))  <> cdbl(nprice) then 
				if chkEdit = 1 and rsimagelib("nprice") <> "" then
				    if rsimagelib("sempno") <> "" then
						KeepLogAddP rsimagelib("sempno"),rsimagelib("dadddate"),Ucase(sprodcode),FormatNumber(rsimagelib("nprice"),4)
					else
					    KeepLogAddP Session("sempno"),rsimagelib("dadddate"),Ucase(sprodcode),FormatNumber(rsimagelib("nprice"),4)
					end if
				end if
				
				KeepLogProduct Ucase(sprodcode),rsimagelib("sexptype"),FormatNumber(nprice,4)
				
			end if
			rsimagelib("nprice") = FormatNumber(nprice,4)
			
		if sfilename<>"" then
			delfile()
			rsimagelib("simage") = Lcase(sfilename)
			rsimagelib("dadddate") = date
		end if
			rsimagelib("sempno") = sempno
			rsimagelib("sCurcode")=sCurcode	
			rsimagelib("sUnitcode")=sUnitcode
			rsimagelib("smodel") = smodel
			rsimagelib("sdrawnum") = sdrawnum
			rsimagelib("sspec") = sspec
			rsimagelib("scountry") = scountry
			rsimagelib("sdimen") = sdimen
			rsimagelib("sbrand") = sbrand
			rsimagelib("smarkern") = smarkern
			rsimagelib("hazardstatus")=hazardstatus
			rsimagelib("hazardno")=hazardno

			'rsimagelib("hazardno")=hazardno
			'For i = 1 To 8
			'	if objsmartupload.Form("l"&i) <> "" and objsmartupload.Form("l"&i) <> "seprodname" then
			'		rsimagelib(objsmartupload.Form("l"&i)) = Ucase(objsmartupload.Form("s"&i))
			'	end if
			'Next
	     rsimagelib.update 
		'response.write "seprodname = "&seprodname&"<br/>"	
		'response.end		 
	     Response.Redirect "../datamanager/pproductdesc.asp?sprodcode="&sprodcode&"&sexptype="&sexptype&"&sCatid="&sCatID&"&sCattype="&sCattype

	     End IF

     end if

	 rsimagelib.close
	 set rsimagelib = nothing
     'response.redirect("../datamanager/pproductdesc.asp")


else
	response.redirect "../datamanager/printError.asp"
end if
%>

<% 'check blank 
   'set objSmartUpload & Upload & set sMode

   Sub check_blank()
     dim str

     check1 = false
     check2 = false 
            
     if trim(sfilename)  = "" then 
        check1 = true
     end if
     if trim(stimgdesc) = "" then 
        check2 = true
     end if 

        
    if (sMode = "add") and (check1 or check2) then 
        str = "bbimagehome.asp?"         
        if not check2 then
           str = str & "stimgdesc=" & stimgdesc & "&"    
        end if
        if (trim(seimgdesc) <> "") then
           str = str & "seimgdesc=" & "seimgdesc" & "&"    
        end if
        str = str & "event=" & sMode
        response.redirect(str)        
    end if

    if (sMode = "edit") and (check2) then 
        str = "sbimagehome.asp?"         
        if not check2 then
           str = str & "stimgdesc=" & stimgdesc & "&"    
        end if
        if (trim(seimgdesc) <> "") then
           str = str & "seimgdesc=" & seimgdesc & "&"    
        end if
        str = str & "simgid=" & simgid & "&"
        str = str & "event=" & sMode
        response.redirect(str)        
    end if


   End Sub
%>

<%
   Sub check_filename()
    Dim Rs,sSql,str

    sSql = "SELECT * FROM pProduct where simage = '" & lcase(sfilename) & "'"
    set Rs = Server.CreateObject("ADODB.Recordset")
    Rs.Open sSql,Conn,adOpenKeySet,adLockPessimistic,adCmdtext
   
	if not Rs.eof then
 		chkfilename=1
	else
 		chkfilename=0           	
   	end if

   End Sub

%>
<%
Sub Del()
	Dim sSQL,Rs
    Dim objFile,spath,filepath,sprno,nitem,sitemno
'		sprno = Request("sprno")

      '''''''''''''''''''''''''''''''''Dalete File
      set objfile = CreateObject("Scripting.FileSystemObject")
      sPath = request.ServerVariables("APPL_PHYSICAL_PATH")
''      sPath = sPath  & "procurement/datamanager/productimages/"
       sPath = sPath  & "datamanager\productimages\" & year(date) & "\"
      sSQL = "SELECT simage FROM pproduct WHERE sprodcode='" & sprodcode &"'" &_
		" and sExptype='" & sExptype & "'"
      Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
      
      sPath = sPath & rs("simage")
	  Set Rs = Nothing

      'filepath = Server.MapPath(sPath)
	'	  Set objfile = CreateObject("Scripting.FileSystemObject")
      If (objfile.FileExists(sPath)) Then
		objfile.DeleteFile(sPath)
      End If

      'objfile.DeleteFile(sPath) 


      '''''''''''''''''''''''''''''''''Delete DataBase
	  ' Create a recordset and open it for adding records
	  sSQL = "Select simage FROM pproduct WHERE  sprodcode='" & sProdcode &"'" &_
		" and sExptype='"& sExptype &"'"   
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS.EOF then
			RS("sImage")=""
			RS.Update
		End if

	  Set Rs = Nothing
End Sub
%>
<%
Sub Delfile()
	Dim sSQL,Rs
    Dim objFile,spath,filepath,sprno,nitem,sitemno
'		sprno = Request("sprno")

      '''''''''''''''''''''''''''''''''Dalete File
      set objfile = CreateObject("Scripting.FileSystemObject")
      sPath = request.ServerVariables("APPL_PHYSICAL_PATH")
''      sPath = sPath  & "procurement/datamanager/productimages/"
		sPath = sPath  & "datamanager\productimages\" & year(date) & "\"
      sSQL = "SELECT simage FROM pproduct WHERE sprodcode='" & sprodcode &"'" &_
		" and sExptype='" & sExptype & "'"
      Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
      
      sPath = sPath & rs("simage")
	  Set Rs = Nothing

      'filepath = Server.MapPath(sPath)
	'	  Set objfile = CreateObject("Scripting.FileSystemObject")
      If (objfile.FileExists(sPath)) Then
		objfile.DeleteFile(sPath)
      End If

      'objfile.DeleteFile(sPath) 

End Sub


Sub KeepLogAddP(sempno,adddate,procode,price)
	Dim Rs2 , sSQL , id , section , scode , dcode , divcode
	Dim Rs3 , level , pur , sname
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Set Rs3 = Server.CreateObject("ADODB.RecordSet")
	
	sSQL = "select sname, sdivcode , sdepcode , sseccode  from aemployee where sempno = '"&sempno&"'"
	Rs3.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	
	scode = Rs3("sseccode")
	dcode = Rs3("sdepcode")
	divcode = Rs3("sdivcode")
	sname = Rs3("sname")
	Rs3.Close
	if scode <> "00" then
		sSQL = "select ssec from psection where sseccode = '"&scode&"'"
		Rs3.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
		section = Rs3("ssec")
		Rs3.Close
	elseif dcode <> "00" then
		sSQL = "select sdept from pdepartment where sdepcode = '"&dcode&"'"
		Rs3.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
		section = Rs3("sdept")
		Rs3.Close
	else
		sSQL = "select sdiv from pdivision where sdivcode = '"&divcode&"'"
		Rs3.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
		section = Rs3("sdiv")
		Rs3.Close
	end if
	level = divcode
	if dcode <> "00" then
		level = level&","&dcode
	end if
	if scode <> "00" then
		level = level&","&scode
	end if
	sSQL = "select sname from aemployee where sempno = (select sempno FROM vpAppstep where nstep = '30' and sflow = (select sflow from pAppflow where slevel ='"&level&"') )"
    'response.write sSQL
	'response.end
	Rs3.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	pur = Rs3("sname")
	Rs3.Close
	sSQL = "select (count(logid)) as id from logproduct"
	Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs2.EOF Then
		If Rs2("id") = "0" Then
			id =  1
		ELSE
			Rs2.close
			sSQL = "select max(logid) as id from logproduct"
			Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			id = cint(Rs2("id")) + 1
		END IF
		'response.write id
		Rs2.close
	'	sSQL = "insert into logvendor values('"&id&"',to_date(to_char(sysdate,'dd/mm/yyyy'),'dd/mm/yyyy'),'"&Session("sEmpno")&"','"&Session("sName")&"','"&status&"','"&comment&"','"&vendcode&"','"&vendname&"')"
	'	response.write sSQL
	'	Rs2.Open sSQL, conn
		Rs2.Open "logproduct", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		Rs2.AddNew
		Rs2("logid") = id
		Rs2("logdate") = adddate
		Rs2("sempno") = sempno
		Rs2("sname") = sname
		Rs2("ssec") = section
		Rs2("sprodcode") = procode
		Rs2("price") = price
		Rs2("pur") = pur
		Rs2.Update	
		'response.end
	END IF
	Rs2.Close
	'Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	'response.write sSQL
	'if Not Rs.EOF Then
	'	response.write "IFFFFFFFFFFFFF"
	'END IF
	'if Rs2.EOF then
	'	response.write "IFffffffffffffff"
	'else
	'	response.writ "ELSEeeeeeeeeeee"
	'end if
End Sub



Sub KeepLogProduct(procode,sexptype,price)
	Dim Rs2 , sSQL , id , section
	Dim Rs3 , level , pur
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Set Rs3 = Server.CreateObject("ADODB.RecordSet")
	sSQL = "select ssec from psection where sseccode = '"&Session("sSeccode")&"'"
	Rs3.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	section = Rs3("ssec")
	Rs3.Close
	level = Session("sDivcode")
	if Session("sDepcode") <> "00" then
		level = level&","&Session("sDepcode")
	end if
	if Session("sSeccode") <> "00" then
		level = level&","&Session("sSeccode")
	end if
	sSQL = "select sname from aemployee where sempno = (select sempno FROM vpAppstep where nstep = '30' and sflow = (select sflow from pAppflow where slevel ='"&level&"') )"
	Rs3.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	if Not Rs3.EOF then
		pur = Rs3("sname")
	
	end if
	Rs3.Close
	
	sSQL = "select (count(logid)) as id from logproduct"
	Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs2.EOF Then
		If Rs2("id") = "0" Then
			id =  1
		ELSE
			Rs2.close
			sSQL = "select max(logid) as id from logproduct"
			Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			id = cint(Rs2("id")) + 1
		END IF
		'response.write id
		Rs2.close
	'	sSQL = "insert into logvendor values('"&id&"',to_date(to_char(sysdate,'dd/mm/yyyy'),'dd/mm/yyyy'),'"&Session("sEmpno")&"','"&Session("sName")&"','"&status&"','"&comment&"','"&vendcode&"','"&vendname&"')"
	'	response.write sSQL
	'	Rs2.Open sSQL, conn
		Rs2.Open "logproduct", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		Rs2.AddNew
		Rs2("logid") = id
		Rs2("logdate") = Now
		Rs2("sempno") = Session("sEmpno")
		Rs2("sname") = Session("sName")
		Rs2("ssec") = section
		Rs2("sprodcode") = procode
		Rs2("sexptype") = sexptype
		Rs2("price") = price
		Rs2("pur") = pur
		Rs2.Update	
		'response.end
	END IF
	Rs2.Close
	'Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	'response.write sSQL
	'if Not Rs.EOF Then
	'	response.write "IFFFFFFFFFFFFF"
	'END IF
	'if Rs2.EOF then
	'	response.write "IFffffffffffffff"
	'else
	'	response.writ "ELSEeeeeeeeeeee"
	'end if
End Sub

function replaceChars(sstr)
	Dim elem 
	Dim str
	str = sstr
	for Each elem In Array("'",",",";","<",">","/","\","|","&","%","#","@","+", Chr(34))
		str =  Replace(str, elem, "")
	Next
	
	replaceChars = str
	'return str
End function

%>
