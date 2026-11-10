<%@EnableSessionState=True %>
<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/oraUtil.inc"-->
<%
Dim rsOra, sqlOra, Ename, Hmail, Bodymail
	if session("uid") <> "" and session("InvalidAccess") = 3 then 
		set rsOra = createObject("adodb.recordset")
		if request("btnSend") = "SEND" then
			if request("radReason") = "I didn't login. (ไม่ได้ Login ผิดเอง)" then
				sqlOra = "update amec.usrAccessLog set REASON = '" & oraFormatStr(request("radReason")) & "' where sempno = '" & session("uid") & "' and system_access = 'WEBFLOW' and login_flag = '0' and datetime = (select max(datetime) from amec.usrAccessLog where sempno = '" & session("uid") & "' and system_access = 'WEBFLOW' and login_flag = '0')"
			elseif request("radReason") = "I did login. (Login ผิดเอง)" and Not IsNull(request("radReasonSub")) then
				sqlOra = "update amec.usrAccessLog set REASON = '" & oraFormatStr(request("radReason")) & " : " & oraFormatStr(request("radReasonSub")) & "' where sempno = '" & session("uid") & "' and system_access = 'WEBFLOW' and login_flag = '0' and datetime = (select max(datetime) from amec.usrAccessLog where sempno = '" & session("uid") & "' and system_access = 'WEBFLOW' and login_flag = '0')"
			end if 
			rsOra.open sqlOra, connOra
			
			'''Send mail
			'get mail emp
			sqlOra = "select sname from amec.aemployee where sempno = '" & session("uid") & "'"
			rsOra.open sqlOra, connOra
			Ename = rsOra("sname")
			rsOra.close
			
			'get mail head 
			sqlOra = "select distinct seq.headno, a.srecmail as hmail from sequenceorg seq, amec.aemployee a where seq.headno = a.sempno and seq.empno = '" & session("uid") & "'"
			rsOra.open sqlOra, connOra
			Hmail = rsOra("hmail")
			rsOra.close
			
			Bodymail = "WEBFLOW : Log-in failure checking system (ระบบตรวจสอบการ Login ผิด)" & vbcrlf _
			& "User id : " & session("uid") & " (" & Ename & ")" & vbcrlf _
			& "Last login (Login ครั้งสุดท้าย) : " & request("txtLastLogin") & vbcrlf _
			& "Last password change (เปลี่ยนรหัสผ่านครั้งสุดท้าย) : " & request("txtLastChgPass") & vbcrlf _
			& "The reason of login failure (เหตุผลที่ Login ผิด) : " & request("radReason")
			if request("radReason") = "I did login. (Login ผิดเอง)"  then 
				Bodymail = Bodymail & " Becuase " & request("radReasonSub")
			end if
			Bodymail = Bodymail & vbcrlf & vbcrlf &  "Best regards," & vbcrlf & "IS Department." & vbcrlf & "Auto Send mail System."
			
			call send_mail("WEBFLOW ADMIN <webflow_admin@mitsubishielevatorasia.co.th>", Hmail, "WEBFLOW : Log-in failure checking system (ระบบตรวจสอบการ Login ผิด)", Bodymail)
			
			sqlOra = "select srecmail from amec.aemployee where sempno =  '" & session("uid") & "'"
			rsOra.open sqlOra, connOra
			if not rsOra.EoF then
				if not IsNull(rsOra("srecmail")) then
					Hmail = rsOra("srecmail")
					'response.write "xxxxx"
					'response.end
					call send_mail("WEBFLOW ADMIN <webflow_admin@mitsubishielevatorasia.co.th>", Hmail, "WEBFLOW : Log-in failure checking system (ระบบตรวจสอบการ Login ผิด)", Bodymail)
				end if
			end if 
			rsOra.close
			
			session.abandon
			response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/default.asp"
		else 
			sqlOra = "select max(a.datetime) as Ldate, b.dupdate as Cdate from amec.usrAccessLog a, amec.aemployee b where a.sempno = b.sempno and a.sempno = '" & session("uid") & "' and a.system_access = 'WEBFLOW' and a.login_flag = '0' group by b.dupdate "
			rsOra.open sqlOra, connOra
		end if 
	else
		response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/default.asp"
	end if
	
	Function send_mail(mfrom, mto, msubject, mbody) 
		Dim iMsg
		Dim iConf
		Dim Flds
		Dim strHTML

		Const cdoSendUsingPickup = 1

		set iMsg = CreateObject("CDO.Message")
		set iConf = CreateObject("CDO.Configuration")

		Set Flds = iConf.Fields
		With Flds
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp"
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			.Update
		End With

		With iMsg
		   Set .Configuration = iConf
		   .BodyPart.Charset = "utf-8"
		   .To = mto
		   .From = mfrom
		   .Subject = msubject 
		   .TextBody = mbody
		   .Send
		End With

		' Clean up variables.
		Set iMsg = Nothing
		Set iConf = Nothing
		Set Flds = Nothing
		Set myMail=CreateObject("CDO.Message")
	End Function 
%>
<html>
<head>
<meta charset="utf-8" />
<style>
	body {
		background-color:#E3F3F9;
	}
	form {
		padding-left:15px;
		font-family: Angsana New;
		font-size:25px;	
		background-color:#C7E2E4;
	}
	div {
		cursor:default;
	}
	.table-main	{
		font-family: Angsana New;
		font-size:25px;
	}
	.table-sub td {
		padding-left:15px;
		font-family: Angsana New;
		font-size:25px;
	}
</style>
<script src="../js/jquery.min.js"></script>
<script>
	function selectRad(id){
		document.getElementById(id).checked = true;
		if(id == "reason2"){
			$(".table-sub tbody").empty();
			var dataRow = "<tr><td style='color:red;'>Please choose your reason amoung the following. (กรุณาเลือกเหตุผลต่อไปนี้)</td></tr>";
			var reason = new Array("I forgot my password. (ลืมรหัสผ่าน)", "I don't know the password. (ไม่รู้รหัสผ่าน)", "I used user ID and password of others. (ใช้ user id, password ของคนอื่น)", "I typed the wrong password. (พิมพ์รหัสผ่านผิด)", "I didn't login from my computer. (I set remember password on my computer) (ไม่ได้ Login เครื่องตัวเองใช้ auto password ที่ computer ตัวเอง)", "I didn't know the login was failed. (ไม่รู้ว่า Login ผิด)");
			
			for (var i = 0; i < reason.length; i++){
				dataRow += "<tr><td><div onclick=\"selectRad(\'reason2" + (i+1) + "\');\"><input type='radio' id='reason2" + (i+1) + "' name='radReasonSub' value=\"" + reason[i].replace("'","\'") + "\">2." + (i+1) + ") " + reason[i] + "</div></td></tr>";
			}
			$(".table-sub tbody").append(dataRow);
		}else if(id == "reason1"){
			$(".table-sub tbody").empty();
		}
	}
	function validateForm(){
		if (document.getElementById("reason1").checked == false && document.getElementById("reason2").checked == false) {
			alert("Please choose the reason of your login failure. (กรุณาเลือกเหตุผลที่ Login ผิด)");
			return false;
		}
		if(document.getElementById("reason2").checked == true){
			var j = 0;
			for (var i = 1;i <= 6;i++){
				if(document.getElementById("reason2" + i).checked == true){
					j = 1;
					break;
				}
			}
			if (j != 1){
				alert("Please choose your reason. (กรุณาเลือกเหตุผลในข้อที่ 2)");
				return false;
			}
		}
	}	
</script>
</head>
<body>
<form name="frmMain" action="loginfail.asp" method="post" onsubmit="return validateForm();">
<b><u><font color="green">WEBFLOW : Log-in failure checking system (ระบบตรวจสอบการ Login ผิด)</font></u></b><br>
<b><font color="red">User id : </font><font color="blue"><u><%=session("uid")%></font></u></b><br>
<b><font color="red">Last login (Login ครั้งสุดท้าย) : </font></b><font color="blue"><u><%=Replace(rsOra("Ldate")," ", "  ")%></u></font><br>
<b><font color="red">Last password change (เปลี่ยนรหัสผ่านครั้งสุดท้าย): </font></b><font color="blue"><u><%=rsOra("Cdate")%></u></font><br>
<input type="hidden" name="txtLastLogin" value="<%=Replace(rsOra("Ldate")," ", "  ")%>" />
<input type="hidden" name="txtLastChgPass" value="<%=rsOra("Cdate")%>" />
	<table class="table-main">
		<tr>
			<td style="color:blue;">Please choose the reason of your login failure. (โปรดระบุเหตุผลที่ Login ผิด)</td>
		</tr>
		<tr>
			<td><div onclick="selectRad('reason1');"><input type="radio" id="reason1" name="radReason" value="I didn't login. (ไม่ได้ Login ผิดเอง)">1) I didn't login. (ไม่ได้ Login ผิดเอง)</div></td>
		</tr>
		<tr>
			<td><div onclick="selectRad('reason2');"><input type="radio" id="reason2" name="radReason" value="I did login. (Login ผิดเอง)">2) I did login. (Login ผิดเอง)</div></td>
		</tr>
	</table>
	<table class="table-sub">
		<tbody></tbody>
	</table><br>
	<u><b><font color="red" size="5">Remark</b></u><br>
	<font color="blue" size="5"><b>If you forgot or don't know your password, please do the following.</b> (หากลืมรหัสผ่านหรือไม่รู้รหัสผ่าน ให้ทำตามขั้นตอนต่อไปนี้)<br><b>1) Please click forgot password at WEBFLOW login page. (ให้กดปุ่ม Forgot password ที่หน้า webflow) <br>2) If you have email, new password will be sent to your email after submitting. Please check your email.</b><br>&nbsp;&nbsp;&nbsp;(หากท่านมีอีเมล์ password ใหม่จะส่งไปยังอีเมล์ของท่าน หลังจากกดปุ่ม Submit กรุณาเข้าไปเช็คที่อีเมล์ของท่าน
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<b>If you don't have email, new password will be sent to email of your manager after submitting. Please contact your manager for new password.</b><br>&nbsp;&nbsp;&nbsp;(หากท่านไม่มีอีเมล์ password ใหม่จะส่งไปยังอีเมล์ของผู้จัดการต้นสังกัดของท่าน หลังจากกดปุ่ม Submit ให้ติดต่อที่ผู้จัดการต้นสังกัดของท่าน เพื่อรับ password ใหม่)</font><br><br>
	<center><input type="submit" name="btnSend" value="SEND" /></center><br>
</form>
</body>
</html>
