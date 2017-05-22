Option Strict Off
Option Explicit On 
Imports vini_DB
Imports System.Resources
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Module Module1

    Public fMainForm As frmMDI
    Public frmActive As System.Windows.Forms.Form
    '    Public locRM As ResourceManager
    Public currentuser As aut_user



    Public Sub Main()
        '        Dim fLogin As New frmLogin
        Dim strGlobalConnection As String
        StartTrace("vini_internet.trace")

        initConstantes()

        'Connection à la Base de Données
        'Si la valeur GlobalConnection = True
        'Connection en début de prog à la base de données
        'Sinon Chaque opération nécéessite une Connection à la base de données
        Try
            strGlobalConnection = System.Configuration.ConfigurationManager.AppSettings.GetValues("dbGlobalConnection")(0)
        Catch ex As Exception
            strGlobalConnection = "False"
        End Try
        If strGlobalConnection = "True" Then
            Persist.shared_connect()
        End If

        '        fLogin.ShowDialog()
        '       If Not fLogin.OK Then
        'Login Failed so exit app
        '     End
        '      End If
        '    currentuser = fLogin.user
        '   fLogin = Nothing

        Param.LoadcolParams()
        Dim colTrp As Collection = Transporteur.colTransporteur
        fMainForm = New frmMDI
        System.Windows.Forms.Application.Run(fMainForm)
        'fMainForm.L()

        Persist.shared_disconnect()

    End Sub



    '    Public Function LoadResString(ByVal strCode) As String

    '        If locRM Is Nothing Then
    '            locRM = New ResourceManager("vin_app.vin_app", GetType(frmMain).Assembly)
    '        End If
    '        Return locRM.GetString(strCode)
    '    End Function

    Public Function isStringNum(ByVal pstr As String) As Boolean
        If pstr = "" Then
            Return True
        Else
            Return IsNumeric(pstr)
        End If
    End Function
    Public Sub initcboRegion(ByVal cbo As ComboBox)
        Dim objParam As Param
        cbo.DisplayMember = "Valeur"
        cbo.ValueMember = "id"
        For Each objParam In Param.colRegion
            cbo.Items.Add(objParam)
        Next
    End Sub
    Public Sub initcboCouleur(ByVal cbo As ComboBox)
        Dim objParam As Param
        cbo.DisplayMember = "Valeur"
        cbo.ValueMember = "id"
        For Each objParam In Param.colCouleur
            cbo.Items.Add(objParam)
        Next
    End Sub
    Public Sub initcboCodeTRP(ByVal cbo As ComboBox)
        Dim objTRP As Transporteur
        cbo.DisplayMember = "Code"
        cbo.ValueMember = "id"
        For Each objTRP In Transporteur.colTransporteur
            cbo.Items.Add(objTRP)
        Next
    End Sub
    Public Sub initcboConditionnent(ByVal cbo As ComboBox)
        Dim objParam As Param
        cbo.DisplayMember = "Code"
        cbo.ValueMember = "id"
        For Each objParam In Param.colConditionnement
            cbo.Items.Add(objParam)
        Next
    End Sub
    Public Sub initcboContenant(ByVal cbo As ComboBox)
        Dim objContenant As contenant
        cbo.DisplayMember = "libelle"
        cbo.ValueMember = "id"
        For Each objContenant In contenant.colContenant
            cbo.Items.Add(objContenant)
        Next
    End Sub
    Public Sub initcboTypeMvt(ByVal cbo As ComboBox)
        cbo.Items.Add("1-Inventaire")
        cbo.Items.Add("2-CommandeClient")
        cbo.Items.Add("3-BonAppro")
        cbo.Items.Add("4-Regul")
    End Sub
    Public Sub InitComboModeReglement(ByVal objCombo As ComboBox)
        Dim objModeReglement As Param
        objCombo.DisplayMember = "Valeur"
        objCombo.ValueMember = "id"
        For Each objModeReglement In Param.colModeReglement
            objCombo.Items.Add(objModeReglement)
        Next
    End Sub


End Module