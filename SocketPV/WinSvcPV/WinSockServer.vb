Imports System
Imports System.Threading
Imports System.Net.Sockets
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Net.Mail
Imports System.Globalization

Public Class WinSockServer
    Dim FILEPATH As String = "C:\tmp\PV\"
    Dim FILENAME As String = "File"
    Dim FileFlag As Boolean = False
    Dim strFileConst As String = ""
    'Dim WithEvents objWinSockServer As New WinSockServer()
    'Dim STR_ACK As String = "<?xml version='1.0' encoding='ISO-8859-1'?><SIHOT-Document><TN>4711</TN><OC>ACK</OC><RC>0</RC><ORG>disp_4711</ORG></SIHOT-Document>"
#Region "ESTRUCTURAS"
    Private Structure InfoDeUnCliente

        'Esta estructura permite guardar la información sobre un cliente
        Public Socket As Socket 'Socket utilizado para mantener la conexion con el cliente
        Public Thread As Thread 'Thread utilizado para escuchar al cliente
        Public UltimosDatosRecibidos As String 'Ultimos datos enviados por el cliente
    End Structure

#End Region

#Region "VARIABLES"
    Private tcpLsn As TcpListener

    Private Clientes As New Hashtable() 'Aqui se guarda la informacion de todos los clientes conectados

    Private tcpThd As Thread

    Private IDClienteActual As Net.IPEndPoint 'Ultimo cliente conectado

    Private m_PuertoDeEscucha As String

    Private Stm As Stream 'Utilizado para enviar datos al Servidor y recibir datos del mismo

#End Region



#Region "EVENTOS"

    Public Event NuevaConexion(ByVal IDTerminal As Net.IPEndPoint)

    Public Event DatosRecibidos(ByVal IDTerminal As Net.IPEndPoint, ByVal TNString As String)

    Public Event ConexionTerminada(ByVal IDTerminal As Net.IPEndPoint)

#End Region



#Region "PROPIEDADES"

    Property PuertoDeEscucha() As String

        Get

            PuertoDeEscucha = m_PuertoDeEscucha

        End Get



        Set(ByVal Value As String)

            m_PuertoDeEscucha = Value

        End Set

    End Property

#End Region



#Region "METODOS"



    Public Sub Escuchar()

        tcpLsn = New TcpListener(PuertoDeEscucha)
        'Inicio la escucha
        tcpLsn.Start()

        'Creo un thread para que se quede escuchando la llegada de un cliente
        tcpThd = New Thread(AddressOf EsperarCliente)
        tcpThd.Start()

    End Sub
    Public Function ObtenerDatos(ByVal IDCliente As Net.IPEndPoint) As String

        Dim InfoClienteSolicitado As InfoDeUnCliente

        'Obtengo la informacion del cliente solicitado
        InfoClienteSolicitado = Clientes(IDCliente)

        ObtenerDatos = InfoClienteSolicitado.UltimosDatosRecibidos

    End Function
    Public Sub Cerrar(ByVal IDCliente As Net.IPEndPoint)

        Dim InfoClienteActual As InfoDeUnCliente



        'Obtengo la informacion del cliente solicitado

        InfoClienteActual = Clientes(IDCliente)



        'Cierro la conexion con el cliente

        InfoClienteActual.Socket.Close()

    End Sub
    Public Sub Cerrar()
        Dim InfoClienteActual As InfoDeUnCliente

        'Recorro todos los clientes y voy cerrando las conexiones

        For Each InfoClienteActual In Clientes.Values

            Call Cerrar(InfoClienteActual.Socket.RemoteEndPoint)

        Next

    End Sub
    Public Sub EnviarDatos(ByVal IDCliente As Net.IPEndPoint, ByVal Datos As String)
        'Envía un mensaje al cliente especificado.
        Dim Cliente As InfoDeUnCliente
        Try
            'Obtengo la informacion del cliente al que se le quiere enviar el mensaje
            Cliente = Clientes(IDCliente)

            'Le envio el mensaje
            Cliente.Socket.Send(Encoding.ASCII.GetBytes(Datos))
        Catch se As SocketException
            SendEmail(se)
            'MsgBox(se)
        End Try

    End Sub
    Public Sub EnviarDatos(ByVal Datos As String)
        'Envía un mensaje a todas los clientes.

        Dim Cliente As InfoDeUnCliente
        'Recorro todos los clientes conectados, y les envio el mensaje recibido
        'en el parametro Datos

        For Each Cliente In Clientes.Values
            EnviarDatos(Cliente.Socket.RemoteEndPoint, Datos)
        Next

    End Sub
#End Region



#Region "FUNCIONES PRIVADAS"

    Private Sub EsperarCliente()

        Dim InfoClienteActual As InfoDeUnCliente
        Try
            With InfoClienteActual
                While True

                    'Cuando se recibe la conexion, guardo la informacion del cliente
                    'Guardo el Socket que utilizo para mantener la conexion con el cliente

                    .Socket = tcpLsn.AcceptSocket() 'Se queda esperando la conexion de un cliente
                    'Guardo el el RemoteEndPoint, que utilizo para identificar al cliente
                    IDClienteActual = .Socket.RemoteEndPoint
                    'Creo un Thread para que se encargue de escuchar los mensaje del cliente
                    .Thread = New Thread(AddressOf LeerSocket)

                    'Agrego la informacion del cliente al HashArray Clientes, donde esta la

                    'informacion de todos estos

                    SyncLock Me
                        Clientes.Add(IDClienteActual, InfoClienteActual)
                    End SyncLock


                    'Genero el evento Nueva conexion
                    RaiseEvent NuevaConexion(IDClienteActual)

                    'Inicio el thread encargado de escuchar los mensajes del cliente
                    .Thread.Start()
                End While
            End With
        Catch se As SocketException
            'SendEmail(se.Message)
            SendEmailGoogle()
        Catch e As Exception
            'SendEmail(e.Message)
            SendEmailGoogle()
        End Try



    End Sub
    Private Sub LeerSocket()
        Dim IDReal As Net.IPEndPoint 'ID del cliente que se va a escuchar
        Dim Recibir() As Byte 'Array utilizado para recibir los datos que llegan
        Dim InfoClienteActual As InfoDeUnCliente 'Informacion del cliente que se va escuchar
        Dim Ret As Integer = 0

        IDReal = IDClienteActual
        InfoClienteActual = Clientes(IDReal)

        'SyncLock Me
        If (InfoClienteActual.Socket Is Nothing) Then
            'Try to restart the socket again Form1_load
            'Send an email that the socket failed or write in the log file
            'SendEmail(InfoClienteActual)
            SendEmailGoogle()

            'RaiseEvent ReiniciaSocket()
            Exit Sub
        End If
        With InfoClienteActual
            While True
                If .Socket.Connected Then
                    Recibir = New Byte(1850000) {}
                    Try
                        'Me quedo esperando a que llegue un mensaje desde el cliente
                        Ret = .Socket.Receive(Recibir, Recibir.Length, SocketFlags.None)
                        If Ret > 0 Then
                            'Guardo el mensaje recibido
                            .UltimosDatosRecibidos = Encoding.ASCII.GetString(Recibir)
                            'Check if the EOF comes in the string.
                            'If .UltimosDatosRecibidos.Contains("") Then
                            'dcLimpiamos el caracter 0X04
                            Dim sb As New StringBuilder(Encoding.ASCII.GetString(Recibir))

                            'dcI clean the end of file from the archive
                            sb = sb.Replace(sb.ToString(), CleanInput(sb.ToString()))

                            'Dim goodText As String = sb.ToString()

                            Clientes(IDReal) = InfoClienteActual
                            Dim TNStart, TNEnd As Integer
                            Dim TNTag As String
                            TNStart = 0
                            TNEnd = 0
                            TNStart = InStr(1, sb.ToString, "<TN>") + 3
                            TNEnd = ((InStr(1, sb.ToString, "</TN>") - 1) - (InStr(1, sb.ToString, "<TN>") + 3))


                            TNTag = .UltimosDatosRecibidos.Substring(TNStart, TNEnd)


                            '////
                            'BufferDeLectura = New Byte(155000) {}
                            'Me quedo esperando a que llegue algun mensaje
                            'Stm.Read(Recibir, 0, Recibir.Length)
                            'Genero el evento DatosRecibidos, ya que se han recibido datos desde el Servidor
                            'EscribeArchivo(Encoding.ASCII.GetString(Recibir))
                            'MsgBox(" Socket:" + InfoClienteActual.Socket.ToString + " Thread:" + InfoClienteActual.Thread.ToString)
                            If EscribeArchivo(sb.ToString, TNTag) Then
                                'Genero el evento de la recepcion del mensaje
                                RaiseEvent DatosRecibidos(IDReal, TNTag)
                            End If

                            'For the search EOT 
                            'Else
                            'Error in the string
                            'EscribeArchivoError(.UltimosDatosRecibidos)
                            'End If
                        Else
                            'Genero el evento de la finalizacion de la conexion
                            RaiseEvent ConexionTerminada(IDReal)
                            Exit While
                        End If

                    Catch se As SocketException
                        'MsgBox(se)
                        'SendEmail(se)
                        SendEmailGoogle()

                    Catch e As Exception
                        If Not .Socket.Connected Then
                            'Genero el evento de la finalizacion de la conexion
                            RaiseEvent ConexionTerminada(IDReal)
                            Exit While
                        End If

                    End Try
                Else
                    'If socket is not connected
                    SendEmailGoogle()
                End If

            End While

            Call CerrarThread(IDReal)

        End With
        'End SyncLock

    End Sub

    Private Function CleanInput(ByVal inputXML As String) As String
        ' Note - This will perform better if you compile the Regex and use a reference to it.
        ' That assumes it will still be memory-resident the next time it is invoked.
        ' Replace invalid characters with empty strings.
        'Return Regex.Replace(inputXML, "[^><\w\.@-]", "")
        Return Regex.Replace(inputXML, "[\4\0]", "")
    End Function
    Private Sub TestFilter()
        ' In this string, I embedded control chars that do not print here.
        Dim badText As String = "<MyText><TB1>" & ChrW(4) & "abcdef</TB1><TB2>abddef" & ChrW(1) & "</TB2><TB3>abcdef" & ChrW(5) & "</TB3></MyText>"
        Dim sb As New StringBuilder(badText)
        sb = sb.Replace(sb.ToString(), CleanInput(sb.ToString()))
        Dim goodText As String = sb.ToString()
    End Sub



    Private Sub CerrarThread(ByVal IDCliente As Net.IPEndPoint)

        Dim InfoClienteActual As InfoDeUnCliente

        'Cierro el thread que se encargaba de escuchar al cliente especificado

        InfoClienteActual = Clientes(IDCliente)

        Try
            InfoClienteActual.Thread.Abort()



        Catch e As Exception

            SyncLock Me

                'Elimino el cliente del HashArray que guarda la informacion de los clientes
                Clientes.Remove(IDCliente)
            End SyncLock
        End Try
    End Sub

    Private Function EscribeArchivo(ByVal strDatos As String, ByVal TNTag As String) As Boolean
        'Dim strFile As String = "f:\TempSockets.txt"
        'Throw New Exception("here")
        Dim retval As Boolean = False
        Dim varTimeDate As DateTime
        varTimeDate = DateTime.Now
        Dim strFile As String = FILEPATH + FILENAME + varTimeDate.ToString("MM_dd_yyyy_hh_mm_ss") + "_Live_PV_" & TNTag & ".xml"
        'Vallarta
        'MPV

        'Dim strFile As String
        Dim swFile As System.IO.StreamWriter

        If (Regex.Matches(strDatos, "SIHOT-Document").Count.ToString() <> "2") Then
            strDatos = strDatos
        End If

        strFileConst = ""
        'SendEmail()

        If (strDatos.Contains("</SIHOT-Document>") = False) Then
            'then a new document is going to start assign the name of the file with a different name
            'if diffetent than EOT/04  keep using the same filename and insert the string in the same file
            'keep using the same file name
            If FileFlag = False Then
                strFileConst = strFile
                FileFlag = False 'Darinel
            End If
            Try
                swFile = System.IO.File.AppendText(strFileConst)
                swFile.WriteLine(strDatos)
                swFile.Close()
                retval = False
                'SendEmail()
            Catch ex As Exception
                'RaiseEvent DatosRecibidos(ex.Message)
            End Try
            'To fix the issue when we have only one thread and finishes the thread with SIHOT-Document
        ElseIf (strDatos.Contains("</SIHOT-Document>") = True) Then
            If FileFlag = False Then
                strFileConst = strFile
                FileFlag = False
            End If
            Try
                swFile = System.IO.File.AppendText(strFileConst)
                swFile.WriteLine(strDatos)
                swFile.Close()
                strFileConst = ""
                retval = True
            Catch ex As Exception
                'RaiseEvent DatosRecibidos(ex.Message)
            End Try
            strFile = ""
            'This is the last thread from the multiple threads messages
        ElseIf (strDatos.Contains("</SIHOT-Document>") = True And FileFlag = True) Then
            Try
                swFile = System.IO.File.AppendText(strFileConst)
                swFile.WriteLine(strDatos)
                swFile.Close()
                strFileConst = ""
                FileFlag = False
                retval = True
            Catch ex As Exception
                'RaiseEvent DatosRecibidos(ex.Message)
            End Try
            strFile = ""
            'this is only one thread to add  to the file
        ElseIf (strDatos.Contains("</SIHOT-Document>") = True And FileFlag = False) Then
            Try
                swFile = System.IO.File.AppendText(strFile)
                swFile.WriteLine(strDatos)
                swFile.Close()
                'send confirmation for he file and stream that has been processed
                'objWinSockServer.EnviarDatos(STR_ACK)
                strFileConst = ""
                FileFlag = False
                retval = True

            Catch ex As Exception
                'RaiseEvent DatosRecibidos(ex.Message)
            End Try
            strFile = ""
            Return retval
        End If
    End Function

    Private Sub EscribeArchivoError(ByVal strDatos As String)
        'Dim strFile As String = "f:\TempSockets.txt"
        'Throw New Exception("here")
        Dim varTimeDate As DateTime
        varTimeDate = DateTime.Now
        Dim strFile As String = FILEPATH + FILENAME + varTimeDate.ToString("MM_dd_yyyy_hh_mm_ss") + "_Error_Live_Vallarta.xml"
        'MCT
        'MPV

        'Dim strFile As String
        Dim swFile As System.IO.StreamWriter

        If (strDatos.Contains("</SIHOT-Document>") = False) Then
            'then a new document is going to start assign the name of the file with a different name
            'if diffetent than EOT/04  keep using the same filename and insert the string in the same file
            'keep using the same file name
            If FileFlag = False Then
                strFileConst = strFile
                FileFlag = True
            End If
            Try
                swFile = System.IO.File.AppendText(strFileConst)
                swFile.WriteLine(strDatos)
                swFile.Close()
            Catch ex As Exception
                'RaiseEvent DatosRecibidos(ex.Message)
            End Try
            'To fix the issue when we have only one thread and finishes the thread with SIHOT-Document
        ElseIf (strDatos.Contains("</SIHOT-Document>") = True) Then
            If FileFlag = False Then
                strFileConst = strFile
                FileFlag = False
            End If
            Try
                swFile = System.IO.File.AppendText(strFileConst)
                swFile.WriteLine(strDatos)
                swFile.Close()
                strFileConst = ""
            Catch ex As Exception
                'RaiseEvent DatosRecibidos(ex.Message)
            End Try
            strFile = ""
            'This is the last thread from the multiple threads messages
        ElseIf (strDatos.Contains("</SIHOT-Document>") = True And FileFlag = True) Then
            Try
                swFile = System.IO.File.AppendText(strFileConst)
                swFile.WriteLine(strDatos)
                swFile.Close()
                strFileConst = ""
                FileFlag = False
            Catch ex As Exception
                'RaiseEvent DatosRecibidos(ex.Message)
            End Try
            strFile = ""
            'this is only one thread to add  to the file
        ElseIf (strDatos.Contains("</SIHOT-Document>") = True And FileFlag = False) Then
            Try
                swFile = System.IO.File.AppendText(strFile)
                swFile.WriteLine(strDatos)
                swFile.Close()
                'send confirmation for he file and stream that has been processed
                'objWinSockServer.EnviarDatos(STR_ACK)
                strFileConst = ""
                FileFlag = False
            Catch ex As Exception
                'RaiseEvent DatosRecibidos(ex.Message)
            End Try
            strFile = ""
        End If
    End Sub

    Public Sub SendEmail(ByVal e As Object)
        'when adding info
        Dim dt As DateTime = DateTime.Now
        Dim dfi As New DateTimeFormatInfo()
        Dim ci As New CultureInfo("en-US")

        Dim sEmailTo As String = "darinelc@solmeliavc.com"
        Dim sEmailFrom As String = "errors@solmeliavc.com"
        'Dim sEmailBCC As String = "darinelcm@yahoo.com"

        Dim mm As New MailMessage()

        Dim sbBody As New StringBuilder(1000)
        Try
            sbBody.Append("<Strong>Vallarta EMAIL ALERT</Strong><br><br>")
            sbBody.Append("<i>NOTIFICATION OF Service failure for Vallarta</i><br><br>")
            sbBody.Append("Email Submitted ON: <B>" & dt.ToString("f", ci) & " (EST)</B><br><br>")
            sbBody.Append("From: <font color='blue'> Solmelia Vacation Club (an automated service)</font><br>")
            sbBody.Append("ATTENTION<br><br>")
            'sbBody.Append("Comments:<br>" + txtComments.Text.Replace("\n", "<br>") + "<br>");
            sbBody.Append("This is an email update to notify you that the Vallarta process failed: <br><br> ")
            sbBody.Append(e.ToString())

            mm.From = New MailAddress(sEmailFrom.Trim())
            mm.[To].Add(sEmailTo)
            ' mm.Bcc.Add(sEmailBCC)
            mm.Subject = "SMVC - Vallarta Socket Error "
            mm.Body = sbBody.ToString()
            mm.IsBodyHtml = True

            'Dim smtp As New SmtpClient(ConfigurationManager.AppSettings("SMTP.Server"))
            Dim smtp As New SmtpClient("smvca1orlfl001.solmeliavc.com")

            smtp.Send(mm)
            'litMessage.Text ="<strong>"+ mycount+ "</strong> Email send succesulfully. Please continue.";
        Catch errores As Exception
            MsgBox(errores.Message)

        End Try

    End Sub


    Public Sub SendEmailGoogle()
        Dim client = New SmtpClient("smtp.gmail.com", 587) With { _
       .Credentials = New NetworkCredential("pmsalertbysol@gmail.com", "smvcpassword"), _
       .EnableSsl = True _
      }
        'client.Send("pmsalertbysol@gmail.com", "4072563998@txt.att.net", "Alert-PR", "PR Socket Problem")
        client.Send("pmsalertbysol@gmail.com", "4072563998@mymetropcs.com", "Alert-PV", "PV Socket Problem")

    End Sub

#End Region
End Class