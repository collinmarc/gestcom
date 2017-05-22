Option Strict Off
Option Explicit On 
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports vini_DB
Imports System.Resources
Module Module1

    Public fMainForm As frmMain
    Public frmActive As System.Windows.Forms.Form
    '    Public locRM As ResourceManager
    Public currentuser As aut_user



    Public Sub Main()
        Dim fLogin As New frmLogin
        Dim colTrp As Collection
        Dim ofrmSettings As frmSettings


        initConstantes()
        If Not Persist.shared_connect() Then
            ofrmSettings = New frmSettings()
            ofrmSettings.ShowDialog()
            Exit Sub
        Else
            Persist.shared_disconnect()
        End If
        Param.LoadcolParams()
        colTrp = Transporteur.colTransporteur


        'Vérification du symbole décimal
        If System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator <> "." Then
            MessageBox.Show("Vous devez mettre le symbole décimal à . (dans Panneau de configuration/options régionnales) pour fonctionner correctement")
        End If



        'Connection à la Base de Données
        'Si la valeur GlobalConnection = True
        'Connection en début de prog à la base de données
        'Sinon Chaque opération nécéessite une Connection à la base de données
        If GLOBALCONNECTION = "True" Then
            Persist.shared_connect()
        End If

        fLogin.ShowDialog()
        If Not fLogin.OK Then
            'Login Failed so exit app
            StopTrace()
            End
        End If
        currentuser = fLogin.user
        StartTrace(currentuser.code & "-Gestcom.trace")
        'Modification de répertoire d'échange FTP avec le nom de l'utilisateur
        IMPORT_DIRECTORY = IMPORT_DIRECTORY + "/" + currentuser.code
        If Not System.IO.Directory.Exists(IMPORT_DIRECTORY) Then
            System.IO.Directory.CreateDirectory(IMPORT_DIRECTORY)
        End If
        'Modification de répertoire Temporaire de FAX avec le nom de l'utilisateur
        FAX_SCMD_PATH = FAX_SCMD_PATH + currentuser.code + "/"
        If Not System.IO.Directory.Exists(FAX_SCMD_PATH) Then
            System.IO.Directory.CreateDirectory(FAX_SCMD_PATH)
        End If


        fLogin = Nothing

        fMainForm = New frmMain
        fMainForm.CurrentUser = currentuser
        fMainForm.ShowDialog()

        Persist.shared_disconnect()

        StopTrace()
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
    Public Sub initcboCodeTRP(ByVal cbo As ComboBox, Optional ByVal pAddAll As Boolean = False)
        Dim objTRP As Transporteur
        cbo.DisplayMember = "Code"
        cbo.ValueMember = "id"
        For Each objTRP In Transporteur.colTransporteur
            cbo.Items.Add(objTRP)
        Next
        If pAddAll Then
            objTRP = New Transporteur()
            objTRP.code = "*/TOUS"
            objTRP.code = "*/TOUS"
            objTRP.nom = "Tous les transporteurs"
            objTRP.AdresseLivraison.fax = ""
            objTRP.AdresseFacturation.fax = ""
            cbo.Items.Add(objTRP)
        End If
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
    Public Sub InitComboModeReglement(ByVal objCombo As ComboBox)
        Dim objModeReglement As Param
        objCombo.DisplayMember = "Valeur"
        objCombo.ValueMember = "id"
        For Each objModeReglement In Param.colModeReglement
            objCombo.Items.Add(objModeReglement)
        Next
    End Sub

    'Public Sub flx_formatcols(ByVal pflx As AxMSFlexGridLib.AxMSFlexGrid, ByVal pcle As String)
    '    '================================================================================
    '    'Function : flx_formatcols
    '    'Description : Format des listes d'éléménts multicollonnes
    '    '================================================================================
    '    Dim i As Integer

    '    pflx.SelectionMode = MSFlexGridLib.SelectionModeSettings.flexSelectionByRow
    '    pflx.AllowUserResizing = MSFlexGridLib.AllowUserResizeSettings.flexResizeColumns
    '    pflx.Cols = System.Configuration.ConfigurationManager.AppSettings.GetValues(pcle & "_COLS")(0)
    '    For i = 0 To pflx.Cols - 1
    '        pflx.set_ColWidth(i, System.Configuration.ConfigurationManager.AppSettings.GetValues(pcle & "_COL" & i)(0))
    '    Next i

    'End Sub 'formaflx
    'Public Sub flx_setRow(ByVal pflx As AxMSFlexGridLib.AxMSFlexGrid, ByVal pitemindex As Integer, ByVal pstrResume As String, Optional ByVal pbExtraColor As Boolean = False)
    '    '================================================================================
    '    'Function : flx_setRow
    '    'Description : Affiche une ligne dans un MSMFEXGRID
    '    '================================================================================
    '    Dim i As Integer
    '    Dim tabShortResume As String()
    '    Dim nNBrecol As Integer

    '    tabShortResume = Split(pstrResume, "|")
    '    nNBrecol = tabShortResume.GetUpperBound(0)
    '    pflx.Row = pitemindex
    '    'Mise à jour des cols
    '    For i = 0 To nNBrecol
    '        pflx.Col = i
    '        pflx.Text = tabShortResume(i)
    '        If pbExtraColor Then
    '            pflx.CellBackColor = GESTSCMD_EXTRACOLOR
    '        Else
    '            pflx.CellBackColor = GESTSCMD_STANDARDCOLOR
    '        End If
    '    Next i
    '    'Dern col = index dans la collection
    '    pflx.Col = pflx.Cols - 1
    '    pflx.Text = pitemindex

    'End Sub 'flx_setRow
    'Public Sub flx_selectRow(ByVal pflx As AxMSFlexGridLib.AxMSFlexGrid, ByVal pindex As Integer)
    '    pflx.Row = pindex
    '    pflx.RowSel = pindex
    '    pflx.Col = 0
    '    pflx.ColSel = pflx.Cols - 1
    '    pflx.Focus()
    '    '        pflx.HighLight = MSFlexGridLib.HighLightSettings.flexHighlightAlways
    'End Sub

    'Public Sub setReportConnection(ByVal objReport As ReportDocument)

    '    Dim myConnectionInfo As ConnectionInfo = New ConnectionInfo()
    '    myConnectionInfo.ServerName = Persist.dbConnection.DataSource
    '    myConnectionInfo.DatabaseName = Persist.dbConnection.Database
    '    myConnectionInfo.IntegratedSecurity = True

    '    Dim mySections As Sections = objReport.ReportDefinition.Sections
    '    For Each mySection As Section In mySections
    '        Dim myReportObjects As ReportObjects = mySection.ReportObjects
    '        For Each myReportObject As ReportObject In myReportObjects
    '            If myReportObject.Kind = ReportObjectKind.SubreportObject Then
    '                Dim mySubreportObject As SubreportObject = CType(myReportObject, SubreportObject)
    '                Dim subReportDocument As ReportDocument = mySubreportObject.OpenSubreport(mySubreportObject.SubreportName)
    '                setReportConnection(subReportDocument)
    '            End If
    '        Next
    '    Next
    '    Dim myTables As Tables = objReport.Database.Tables
    '    For Each myTable As CrystalDecisions.CrystalReports.Engine.Table In myTables
    '        Dim myTableLogonInfo As TableLogOnInfo = myTable.LogOnInfo
    '        myTableLogonInfo.ConnectionInfo = myConnectionInfo
    '        myTable.ApplyLogOnInfo(myTableLogonInfo)
    '    Next
    'End Sub

End Module