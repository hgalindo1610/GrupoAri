Imports System.Net.Mail

Module Correos
    Public Envia, Nenvia, Nremitente, asunto, cuerpo, clave As String
    Public Sub EnvioMail()

        Dim correo As New MailMessage
        Dim smtp As New SmtpClient()
        Envia = "sistema@grupoarihn.com"
        Nenvia = "Sistema Grupo Ari"
        clave = "S1stemas2016"
        correo.From = New MailAddress(Envia, Nenvia, System.Text.Encoding.UTF8)
        correo.To.Add(Nremitente)
        correo.SubjectEncoding = System.Text.Encoding.UTF8
        correo.Subject = asunto
        correo.Body = cuerpo
        correo.BodyEncoding = System.Text.Encoding.UTF8
        correo.IsBodyHtml = True
        correo.Priority = MailPriority.High

        smtp.Credentials = New System.Net.NetworkCredential(Envia, clave)
        smtp.Port = 587
        smtp.Host = "mail.grupoarihn.com"
        smtp.EnableSsl = False

        smtp.Send(correo)

    End Sub

End Module
