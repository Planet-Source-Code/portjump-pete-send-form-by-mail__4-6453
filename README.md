<div align="center">

## Send Form By Mail


</div>

### Description

Sends form results to you via email
 
### More Info
 
Put mailer.asp in your root, edit the email addresses, and SMTP Server name where defined in the mailer.asp and put this HTML tag in your form: <FORM method="post" action="/mailer.asp">

Code Assumes you are using Bamboo Mail on your server. You can create a custom 'Thank You' page if your message was sent successfully.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[PortJump Pete](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/portjump-pete.md)
**Level**          |Intermediate
**User Rating**    |4.1 (29 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Validation/ Processing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/validation-processing__4-16.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/portjump-pete-send-form-by-mail__4-6453/archive/master.zip)





### Source Code

```
If you have problems with the code below, check out version 2 -
I updated this code some time ago but it's been greatly overlooked!
http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=7537&lngWId=4
<!--- mailer.asp --->
<%
For Each x In Request.Form
 message=message & x & ": " & Request.Form(x) & CHR(10)
Next
set smtp = Server.CreateObject("Bamboo.SMTP")
smtp.Server = "mail.YOURSEVER.com"
smtp.Rcpt = "YOU@YOURSERVER.com"
smtp.From = "user@YOURSERVER.com"
smtp.FromName = "user@YOURSERVER.com"
smtp.Subject = "Message Recieved From Web Form"
smtp.Message = message
on error resume next
smtp.Send
if err then
 response.Write err.Description
else
 response.redirect "/thanks.asp"
end if
set smtp = Nothing
%>
```

