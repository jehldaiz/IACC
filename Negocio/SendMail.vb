Imports System.Configuration
Imports System.Text

Public Class EnviarCorreo
    Public Function SendMail(para As String, cc As String, cuerpo As String) As Boolean
        Dim Message = New System.Net.Mail.MailMessage()
        Dim SMTP = New System.Net.Mail.SmtpClient()

        ''CONFIGURACIÓN DEL STMP
        ''----------------------------------------------------'("cuenta de correo", "contraseña")
        SMTP.Credentials = New System.Net.NetworkCredential(ConfigurationManager.AppSettings("Usuario"), ConfigurationManager.AppSettings("Password"))
        SMTP.Host = ConfigurationManager.AppSettings("Smtp")
        SMTP.Port = Convert.ToInt32(ConfigurationManager.AppSettings("Puerto"))
        SMTP.EnableSsl = True

        '' CONFIGURACION DEL MENSAJE
        Message.[To].Add(para)
        ' ' Acá se escribe la cuenta de correo al que se le quiere enviar el e-mail
        'Message.Bcc.Add(cc)

        ''----------------------------------------------"Quien lo envía","Nombre de quien lo envía"  

        Message.From = New System.Net.Mail.MailAddress(ConfigurationManager.AppSettings("MailOrigen"), "Control Centralizado", Encoding.UTF8)
        ''Quien envía el e-mail
        Message.Subject = ConfigurationManager.AppSettings("Asunto")
        ''Motivo o Asunto del e-mail
        Message.SubjectEncoding = Encoding.UTF8
        ''Codificacion
        Message.Body = cuerpo
        ''contenido del mail
        Message.IsBodyHtml = True

        Try
            SMTP.Send(Message)
            Return True
        Catch ex As System.Net.Mail.SmtpException
            Return False
        End Try
    End Function

End Class
