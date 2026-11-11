<%
	Dim mail
	Dim receiver
	

	set mail = server.createObject("aspSmartMail.smartMail")

'	Mail.From="adminprocurement@MitsubishiElevatorAsia.co.th"
'	Mail.To= "kallaya@MitsubishiElevatorAsia.co.th"
'	Mail.Subject="test"
'	Mail.TextBody = " Waiting for your approve"
'	Mail.Send
'	Set Mail=nothing
	

	with mail
		.server = "smtp"
		.senderName = "adminprocurement@MitsubishiElevatorAsia.co.th"
		.senderAddress = "kallaya@MitsubishiElevatorAsia.co.th"
		.recipients.add "kallaya@MitsubishiElevatorAsia.co.th"

		.subject = "test"
		.body = "test"
		.sendMail
	end with


	set mail = nothing


%>