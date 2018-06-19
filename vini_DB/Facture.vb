Imports System.IO
Imports System.Data.OleDb

Public MustInherit Class Facture
    Inherits Commande

    Private m_colReglements As Collection
    Private m_bcolRegelementLoaded As Boolean
    Private m_bExportInternet As Boolean
    Protected m_dateEcheance As Date
    Protected m_IDModeReglement As Integer



    Public Sub New(ByVal poTiers As Tiers, ByVal poEtat As EtatCommande)
        MyBase.New(poTiers, poEtat)
        m_colReglements = New Collection()
        m_bcolRegelementLoaded = False
        m_bExportInternet = False
        m_dateEcheance = DATE_DEFAUT
        dateFacture = Now()
    End Sub


    ''' <summary>
    ''' date facture = dateCommande
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property dateFacture() As Date
        Get
            Return dateCommande.ToString("d")
        End Get
        Set(ByVal Value As Date)
            If Not dateFacture.Equals(Value) Then
                dateCommande = Value
                CalcDateEcheance()
                RaiseUpdated()
            End If
        End Set
    End Property
    Public ReadOnly Property colReglements() As Collection
        Get
            Return m_colReglements
        End Get
    End Property

    Public Property bcolReglementLoaded() As Boolean
        Get
            Return m_bcolRegelementLoaded
        End Get
        Set(ByVal value As Boolean)
            m_bcolRegelementLoaded = value
        End Set
    End Property
    Public Property dEcheance() As Date
        Get
            Return m_dateEcheance
        End Get
        Set(ByVal value As Date)
            If Not value.Equals(m_dateEcheance) Then
                m_dateEcheance = value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Property idModeReglement() As Integer
        Get
            Return m_IDModeReglement
        End Get
        Set(ByVal value As Integer)
            If Not value.Equals(m_IDModeReglement) Then
                m_IDModeReglement = value
                If dateCommande <> DATE_DEFAUT Then
                    CalcDateEcheance()
                End If
                RaiseUpdated()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Rend ou met a jour l'indicateur d'export internet 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property bExportInternet() As Boolean
        Get
            Return m_bExportInternet
        End Get
        Set(ByVal value As Boolean)
            If value <> m_bExportInternet Then
                m_bExportInternet = value
                RaiseUpdated()
            End If
        End Set
    End Property
    Public Function CalcDateEcheance() As Boolean
        Dim bReturn As Boolean
        Dim objParam As ParamModeReglement

        Try
            If (idModeReglement <> 0) Then
                objParam = New ParamModeReglement()
                objParam.load(idModeReglement)
                If Not String.IsNullOrEmpty(objParam.code) Then
                    dEcheance = objParam.calDateEcheance(Me.dateFacture)
                End If
            End If
            bReturn = True
        Catch ex As Exception
            setError("Facture.CalcDateEcheance", ex.ToString())
            bReturn = False
        End Try
        Return bReturn
    End Function
    '=======================================================================
    'Fonction : Equals()
    'Description : Teste l'équivalence entre 2 objets
    'Détails    :  
    'Retour : Vari si l'objet passé en paramétre est equivalent à 
    '=======================================================================
    Public Overrides Function Equals(ByVal obj As Object) As Boolean
        Dim bReturn As Boolean
        Dim objCommande As Facture

        Try

            bReturn = MyBase.Equals(obj)
            objCommande = obj
            bReturn = bReturn And (m_bExportInternet.Equals(objCommande.bExportInternet))

            bReturn = bReturn And (m_IDModeReglement.Equals(objCommande.idModeReglement))
            bReturn = bReturn And (m_dateEcheance.ToShortDateString().Equals(objCommande.dEcheance.ToShortDateString()))

            Return bReturn

        Catch ex As Exception
            bReturn = False
        End Try

        Return bReturn

    End Function 'Equals

    Public Overrides Function checkForDelete() As Boolean
        Return True
    End Function

    Public Overrides Function ControleMvtStock() As Microsoft.VisualBasic.Collection
        Return Nothing
    End Function

    Public Overrides Function genereMvtStock() As Boolean
        Return False
    End Function


    Public Overrides Function supprimeMvtStock() As Boolean
        Return False
    End Function
    ''' <summary>
    ''' Export Quadra
    ''' </summary>
    ''' <param name="pstFileName"></param>
    ''' <param name="pExportType"></param>
    ''' <remarks></remarks>
    Protected Sub ExporterFacture(ByVal pstFileName As String, ByVal pExportType As String)
        Dim strLine As String
        Dim oTA As vini_DB.dsVinicomTableAdapters.EXPORTPARAMTableAdapter
        Dim oTAConstantes As vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter
        Dim nbLignes As Integer
        Dim nLigne As Integer
        Dim nbChamps As Integer
        Dim nChamps As Integer
        Dim oDT As vini_DB.dsVinicom.EXPORTPARAMDataTable
        Dim oRow As vini_DB.dsVinicom.EXPORTPARAMRow
        Dim oDTConstantes As vini_DB.dsVinicom.CONSTANTESDataTable
        Dim strValeur As String
        Try
            'Chargement des informatios du tiers
            oTiers.load()
            strLine = String.Empty
            oTA = New vini_DB.dsVinicomTableAdapters.EXPORTPARAMTableAdapter()
            oTA.Connection = Persist.oleDBConnection
            oTAConstantes = New vini_DB.dsVinicomTableAdapters.CONSTANTESTableAdapter()
            oTAConstantes.Connection = Persist.oleDBConnection

            oDTConstantes = oTAConstantes.GetData()
            'Recupération du nombre de lignes
            nbLignes = oTA.getNbLignes(pExportType)
            For nLigne = 1 To nbLignes
                ' Pour chaque ligne
                'Lecture des paramètres
                oDT = oTA.GetDataBy_TYPE_NLIGNE(pExportType, nLigne)
                'Recupéation du nombre de champs
                nbChamps = oTA.getNbChamps(pExportType, nLigne)
                For nChamps = 1 To nbChamps
                    'Pour chaque Champs
                    oRow = oDT.FindByEXP_TYPEEXP_NLIGNEEXP_NCHAMPS(pExportType, nLigne, nChamps)
                    If oRow IsNot Nothing Then
                        If Trim(oRow.EXP_TYPECHAMPS).Equals("C") Then
                            'Exportation d'une Constante
                            'Si la longueur est égale à 0 => Trim
                            If oRow.EXP_LONGUEUR = 0 Then
                                strLine = strLine + Trim(oRow.EXP_VALEUR)
                            Else
                                strLine = strLine + Left(oRow.EXP_VALEUR + Space(oRow.EXP_LONGUEUR), oRow.EXP_LONGUEUR)
                            End If
                        End If
                        If Trim(oRow.EXP_TYPECHAMPS).Equals("V") Then
                            'Exportation d'une Valeur
                            strValeur = getAttributeValue(oRow.EXP_VALEUR, oDTConstantes)
                            'Si la longueur est égale à 0 => Trim
                            If oRow.EXP_LONGUEUR = 0 Then
                                strLine = strLine + Trim(strValeur)
                            Else
                                strLine = strLine + Left(strValeur + Space(oRow.EXP_LONGUEUR), oRow.EXP_LONGUEUR)
                            End If
                        End If
                    End If
                Next 'champs
                strLine = strLine + vbCrLf

            Next 'Lignes

            File.AppendAllText(pstFileName, strLine)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function getAttributeValue(ByVal pstrAttributeName As String, ByVal pDTConstantes As dsVinicom.CONSTANTESDataTable) As String
        Dim strReturn As String
        strReturn = String.Empty

        Try

            Select Case pstrAttributeName
                Case "CODECOMPTA"
                    strReturn = oTiers.CodeCompta
                Case "DATEFACTURE"
                    strReturn = Format(dateFacture, "ddMMyy")
                Case "DATEECHEANCE"
                    strReturn = Format(dEcheance, "ddMMyy")
                Case "LIBELLE"
                    If totalHT > 0 Then
                        strReturn = "F:" + Trim(code) + " " + getTXTString(oTiers.rs)
                    Else
                        strReturn = "A:" + Trim(code) + " " + getTXTString(oTiers.rs)
                    End If
                Case "TTC"
                    strReturn = Math.Abs(totalTTC).ToString("0000000000.00").Replace(".", "")
                Case "HT"
                    strReturn = Math.Abs(totalHT).ToString("0000000000.00").Replace(".", "")
                Case "TVA"
                    strReturn = Math.Abs((totalTTC - totalHT)).ToString("0000000000.00").Replace(".", "")
                Case "NUMFACT"
                    strReturn = code
                Case "COMPTETVA"
                    If pDTConstantes IsNot Nothing Then
                        If (Not pDTConstantes.Rows(0).IsNull("CST_SOC_COMPTETVA")) Then
                            strReturn = pDTConstantes.Rows(0)("CST_SOC_COMPTETVA")
                        End If
                    End If
                Case "COMPTEPRODUIT"
                    If pDTConstantes IsNot Nothing Then
                        If (Not pDTConstantes.Rows(0).IsNull("CST_SOC_COMPTEPRODUIT")) Then
                            strReturn = pDTConstantes.Rows(0)("CST_SOC_COMPTEPRODUIT")
                        End If
                    End If
                Case "COMPTETVA2"
                    If pDTConstantes IsNot Nothing Then
                        If (Not pDTConstantes.Rows(0).IsNull("CST_SOC2_COMPTETVA")) Then
                            strReturn = pDTConstantes.Rows(0)("CST_SOC2_COMPTETVA")
                        End If
                    End If
                Case "COMPTEPRODUIT2"
                    If pDTConstantes IsNot Nothing Then
                        If (Not pDTConstantes.Rows(0).IsNull("CST_SOC2_COMPTEPRODUIT")) Then
                            strReturn = pDTConstantes.Rows(0)("CST_SOC2_COMPTEPRODUIT")
                        End If
                    End If
                Case "COMPTEPRODUIT2COL"
                    If pDTConstantes IsNot Nothing Then
                        If (Not pDTConstantes.Rows(0).IsNull("CST_SOC2_COMPTEPRODUIT_COL")) Then
                            strReturn = pDTConstantes.Rows(0)("CST_SOC2_COMPTEPRODUIT_COL")
                        End If
                    End If

                Case "MODEREGLEMENT"
                    Dim oParam As New ParamModeReglement
                    oParam.load(Me.idModeReglement)
                    strReturn = oParam.code
                Case "LIBMODEREGLEMENT"
                    Dim oParam As New ParamModeReglement
                    oParam.load(Me.idModeReglement)
                    strReturn = oParam.valeur

                Case "BANQUE"
                    strReturn = Me.oTiers.banque
                Case "RIB"
                    strReturn = Me.oTiers.rib1 & Me.oTiers.rib2 & Me.oTiers.rib3 & Me.oTiers.rib4

                Case "MODEREGLEMENT2"
                    Dim oParam As New ParamModeReglement
                    oParam.load(Me.idModeReglement)
                    Dim oTabCorrespondance As String()
                    oTabCorrespondance = File.ReadAllLines("./ModeReglmtQuadra.csv")
                    For Each line As String In oTabCorrespondance
                        Dim Ligne As String()
                        Ligne = line.Split(";")
                        If Trim(Ligne(0)) = oParam.code Then
                            strReturn = Ligne(1)
                        End If
                    Next

 
                Case "MODEREGLEMENT4"
                    Dim oParam As New ParamModeReglement
                    oParam.load(Me.idModeReglement)
                    Dim oTabCorrespondance As String()
                    oTabCorrespondance = File.ReadAllLines("./ModeReglmtQuadra.csv")
                    For Each line As String In oTabCorrespondance
                        Dim Ligne As String()
                        Ligne = line.Split(";")
                        If Trim(Ligne(0)) = oParam.code Then
                            strReturn = Ligne(2)
                        End If
                    Next

                Case "SENSD"
                    If totalHT > 0 Then
                        strReturn = "D"
                    Else
                        strReturn = "C"
                    End If


                Case "SENSC"
                    If totalHT > 0 Then
                        strReturn = "C"
                    Else
                        strReturn = "D"
                    End If
            End Select
        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
            strReturn = ""
        End Try

        Return strReturn
    End Function
    Private Function getTXTString(ByVal pInputString As String) As String
        Dim strReturn As String = String.Empty
        Dim c As Char
        Dim tab As Char() = pInputString.ToCharArray()

        For Each c In tab
            If Char.IsLetterOrDigit(c) Then
                strReturn = strReturn + c
            Else
                strReturn = strReturn + " "
            End If
        Next
        Return strReturn
    End Function

    Public Sub loadReglements()
        m_colReglements = getListeReglement(DATE_DEFAUT, DATE_DEFAUT, id, typeDonnee)
        m_bcolRegelementLoaded = True
    End Sub
    ''' <summary>
    ''' Rend le solde non réglé de la facture
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getSolde() As Decimal
        Dim montantReglement As Decimal = 0D
        Dim solde As Decimal = 0D
        Dim oReglement As Reglement

        Try
            loadReglements()

            For Each oReglement In m_colReglements
                montantReglement = montantReglement + oReglement.Montant
            Next

            solde = totalTTC - montantReglement
        Catch ex As Exception
            setError(System.Environment.StackTrace, ex.Message)
        End Try
        Return solde
    End Function

    Public Function AddReglement(ByVal pDateReglement As Date, ByVal pMontant As Decimal, ByVal pRef As String) As Reglement
        Dim objReglement As Reglement
        objReglement = New Reglement(Me)

        objReglement.Montant = pMontant
        objReglement.Reference = pRef
        objReglement.DateReglement = pDateReglement.ToShortDateString()
        m_colReglements.Add(objReglement)

        Return objReglement
    End Function




End Class
