Imports System.ServiceProcess

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class vini_service
    Inherits System.ServiceProcess.ServiceBase

    'UserService remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' Point d'entrée principal du processus
    <MTAThread()> _
    <System.Diagnostics.DebuggerNonUserCode()> _
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        ' Plusieurs services NT s'exécutent dans le même processus. Pour ajouter
        ' un autre service à ce processus, modifiez la ligne suivante
        ' pour créer un second objet service. Par exemple,
        '
        '   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New vini_service}

        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub

    'Requise par le Concepteur de composants
    Private components As System.ComponentModel.IContainer

    ' REMARQUE : la procédure suivante est requise par le Concepteur de composants
    ' Elle peut être modifiée à l'aide du Concepteur de composants.  
    ' Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        'Me.Timer1 = New System.Threading.Timer(AddressOf ImporterCommandes)
        '
        'Timer1
        '
        '
        'vini_service
        '
        Me.ServiceName = "Service1"

    End Sub
    Friend WithEvents Timer1 As System.Threading.Timer

End Class
