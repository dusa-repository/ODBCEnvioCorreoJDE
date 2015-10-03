Imports System.Net.Mail
Imports System.IO
Imports System.Xml
Module Principal

    Dim lineaLogger As String
    Dim logger As StreamWriter
    Dim cadenaConexion As String
    Dim cadenaConexionSVM812 As String
    Dim directorio As String
    Dim actualizar As String
    Dim prueba As String
    Dim emailprueba As String
    Dim subject As String
    Dim body As String

    Sub Main()

        
        Dim pgm As String
        Dim grupo As String
        Dim version As String
        Dim job As String
        Dim extension As String
        Dim nombreReporteControl As String
        Dim file_log_path As String

        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()



        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.txt") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.txt")
                fs1.Close()
            End If

            logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.txt", True)

            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion.xml", FileMode.Open, FileAccess.Read)
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            cadenaConexion = diccionario.Item("DSN1")
            cadenaConexionSVM812 = diccionario.Item("DSN2")
            directorio = diccionario.Item("DIRECTORIO")
            actualizar = diccionario.Item("ACTUALIZAR")
            prueba = diccionario.Item("ACTIVAR_PRUEBA")
            emailprueba = diccionario.Item("EMAIL_PRUEBAS")

            Dim cnn As New Odbc.OdbcConnection(cadenaConexionSVM812)
            Dim rs As New Odbc.OdbcCommand("SELECT * FROM F986110 WHERE JCUSER='VENTAS' ", cnn)
            Dim reader As Odbc.OdbcDataReader

        

            cnn.Open()
            reader = rs.ExecuteReader
            While reader.Read()
                nombreReporteControl = reader("JCFNDFUF2")
                job = reader("JCJOBNBR")
                pgm = Left(nombreReporteControl, InStr(1, nombreReporteControl, "_") - 1)
                version = Mid(nombreReporteControl, InStr(1, nombreReporteControl, "_") + 1, InStr(InStr(1, nombreReporteControl, "_") + 1, nombreReporteControl, "_") - InStr(1, nombreReporteControl, "_") - 1)
                extension = Right(nombreReporteControl.Trim, 3)
                grupo = obtenerDireccionGrupo(pgm, version)
                If grupo.Trim <> "" Then

                    If prueba.Trim = "S" Then
                        EnvioCorreo("reporteautomatico@dusa.com.ve", "", emailprueba.Trim, directorio.Trim & nombreReporteControl.Trim & "." & extension)

                        If File.Exists(directorio.Trim & nombreReporteControl.Trim & ".CSV") Then
                            EnvioCorreo("reporteautomatico@dusa.com.ve", "", emailprueba.Trim, directorio.Trim & nombreReporteControl.Trim & ".CSV")
                            escribirLog("ENVIO DE:" & nombreReporteControl.Trim & ".CSV" & " a :" & grupo)
                        End If


                    Else

                        EnvioCorreo("reporteautomatico@dusa.com.ve", "", grupo.Trim, directorio.Trim & nombreReporteControl.Trim & "." & extension)

                        If File.Exists(directorio.Trim & nombreReporteControl.Trim & ".CSV") Then
                            EnvioCorreo("reporteautomatico@dusa.com.ve", "", grupo.Trim, directorio.Trim & nombreReporteControl.Trim & ".CSV")
                            escribirLog("ENVIO DE:" & nombreReporteControl.Trim & ".CSV" & " a :" & grupo)
                        End If

                    End If


                    escribirLog("ENVIO DE:" & nombreReporteControl.Trim & "." & extension & " a :" & grupo)

                    If actualizar.Trim = "S" Then
                        Dim comando As New Odbc.OdbcCommand

                        comando.Connection = cnn
                        comando.CommandText = " UPDATE F986110 SET JCUSER='EMAIL' WHERE JCJOBNBR=" & job & " "
                        comando.ExecuteNonQuery()

                    End If

                End If

            End While
            reader.Close()
            cnn.Close()

        Catch ex As Exception
            escribirLog(ex.ToString)
        End Try


    End Sub


    Private Sub EnvioCorreo(ByVal sender As String, ByVal password As String, ByVal recipients As String, ByVal attach As String)

        Dim correo As New MailMessage
        Dim smtp As New SmtpClient()




        Try
            correo.From = New MailAddress(sender, "REPORTE AUTOMATICO", System.Text.Encoding.UTF8)
            correo.To.Add(recipients)
            correo.SubjectEncoding = System.Text.Encoding.UTF8
            correo.Subject = subject
            correo.Body = body
            correo.BodyEncoding = System.Text.Encoding.UTF8
            correo.IsBodyHtml = False
            correo.Priority = MailPriority.High
            Try
                Dim attachment As New Net.Mail.Attachment(attach)
                correo.Attachments.Add(attachment)
            Catch ex As Exception
                escribirLog(ex.ToString)
            End Try
            smtp.Credentials = New System.Net.NetworkCredential(sender, password)
            smtp.Port = 2525
            smtp.Host = "172.23.20.66"
            ' smtp.EnableSsl = True
            smtp.UseDefaultCredentials = False
            smtp.Send(correo)

        Catch ex As Exception
            escribirLog(ex.ToString)
        End Try

    End Sub



    Private Function obtenerDireccionGrupo(ByVal nombreReporte As String, ByVal version As String) As String


        Dim cnnAux As New Odbc.OdbcConnection(cadenaConexion)
        Dim rsAux As New Odbc.OdbcCommand("SELECT * FROM F55MAIL1 WHERE EDPGM='" & nombreReporte.Trim & "' AND EDJDEVERS='" & version & "' ", cnnAux)
        Dim readerAux As Odbc.OdbcDataReader
        Dim grupo As String
        grupo = ""
        body = ""
        subject = ""
        Try

            cnnAux.Open()
            readerAux = rsAux.ExecuteReader
            While readerAux.Read()
                subject = readerAux("EDSUBJECT")
                body = readerAux("EDESUBJECT")
                grupo = readerAux("EDGEMAIL")
            End While
            readerAux.Close()
            cnnAux.Close()

        Catch ex As Exception
            escribirLog(ex.ToString)

        End Try


        Return grupo

    End Function

    Public Function obtenerNodosHijosDePadre(ByVal nombreNodoPadre As String, ByVal xmldoc As XmlDataDocument) As Dictionary(Of String, String)
        Dim diccionario As New Dictionary(Of String, String)
        Dim nodoPadre As XmlNodeList
        Dim i As Integer
        Dim h As Integer
        nodoPadre = xmldoc.GetElementsByTagName(nombreNodoPadre)
        For i = 0 To nodoPadre.Count - 1
            For h = 0 To nodoPadre(i).ChildNodes.Count - 1
                If Not diccionario.ContainsKey(nodoPadre(i).ChildNodes.Item(h).Name.Trim()) Then
                    diccionario.Add(nodoPadre(i).ChildNodes.Item(h).Name.Trim(), nodoPadre(i).ChildNodes.Item(h).InnerText.Trim())
                End If
            Next
        Next
        Return diccionario
    End Function

    Public Sub escribirLog(ByVal mensaje As String)

        Dim time As DateTime = DateTime.Now
        Dim format As String = "dd/MM/yyyy HH:mm "

        lineaLogger = "[" & time.ToString(format) & "] " & mensaje '& vbNewLine
        logger.WriteLine(lineaLogger)
        logger.Flush()

        Console.WriteLine("[" & time.ToString(format) & "] " & mensaje)
    End Sub

End Module
