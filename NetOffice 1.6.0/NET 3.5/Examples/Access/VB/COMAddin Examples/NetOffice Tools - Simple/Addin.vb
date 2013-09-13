Imports System
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.AccessApi.Tools
Imports NetOffice.AccessApi

'/*
'  * This project shows you the COMAddin base class from the NetOffice tools.
'  * Its designed to reduce infrastructure code from your own.
'  * You have to set some attributes and thats all. 
'  * You see also the host application is available as class instance property. no need for dispose here because the base class do this for you while shutdown.
'*/

<COMAddin("NetOfficeVB35 Sample Access Addin", "This Addin shows you the COMAddin base class from the NetOffice Tools", 3)> _
<GuidAttribute("BD432A44-89E6-492e-9CB7-35D57CBC2B64"), ProgIdAttribute("SimpleAccessVB35.Addin")> _
Public Class Addin
    Inherits COMAddin

    Private Sub Addin_OnStartupComplete(ByRef custom As System.Array) Handles Me.OnStartupComplete

        ' get the host application version
        Dim hostVersion As String = Me.Application.Version
        MessageBox.Show("Host Application Version: " + hostVersion)

    End Sub

    Private Sub Addin_OnDisconnection(ByVal RemoveMode As NetOffice.Tools.ext_DisconnectMode, ByRef custom As System.Array) Handles Me.OnDisconnection


    End Sub

End Class
