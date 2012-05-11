Public Class ServicePV

    Dim WithEvents objWinSockServer As New WinSockServer()
    Dim PUERTODEESCUCHA As String = 14782
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        With objWinSockServer


            'Establezco el puerto donde escuchar

            '.PuertoDeEscucha = 8050
            .PuertoDeEscucha = PUERTODEESCUCHA

            'Comienzo la escucha
            .Escuchar()

        End With
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.

    End Sub


    Public Sub WinSockServer_DatosRecibidos(ByVal IDTerminal As System.Net.IPEndPoint, ByVal TNString As String) Handles objWinSockServer.DatosRecibidos

        Dim STR_ACK As String

        STR_ACK = "<?xml version='1.0' encoding='ISO-8859-1'?><SIHOT-Document><TN>" & TNString & "</TN><OC>ACK</OC><RC>0</RC><ORG>disp_4711</ORG></SIHOT-Document>"
        objWinSockServer.EnviarDatos(IDTerminal, STR_ACK)

    End Sub

End Class
