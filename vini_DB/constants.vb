
Option Explicit On 
Module constants
    'Code Adresses
    Public Const ADRLIV As String = "Liv"
    Public Const ADRFACT As String = "Fact"
    'Code Commentaire
    Public Const COMLIV As String = "Liv"
    Public Const COMFACT As String = "Fact"
    Public Const COMCMD As String = "cmd"
    Public Const COMLIBRE As String = "libre"
    'Code Paramètres
    Public Const PAR_MODE_RGLMT As String = "R"
    Public Const PAR_TYPE_CLIENT As String = "C"
    Public Const PAR_COULEUR As String = "O"
    Public Const PAR_REGION As String = "G"
    Public Const PAR_CONDITIONNEMENT As String = "D"
    Public Const PAR_TVA As String = "V"
    Public Const PAR_TYPECOMMANDE As String = "T"
    Public Const PAR_TYPETRANSPORT As String = "P"
    Public Const PAR_ETATCOMMANDE As String = "E"
    'Libellé des Etats
    Public Const ETAT_ENCOURS As String = "En Cours"
    Public Const ETAT_VALIDEE As String = "Validée"
    Public Const ETAT_TRANSMISE As String = "Transmise"
    Public Const ETAT_ECLATEE As String = "Eclatée"
    Public Const ETAT_LIVREE As String = "Livrée"
    Public Const ETAT_SS_GENEREE As String = "Générée"
    Public Const ETAT_SS_VALIDEE As String = "Validée"
    Public Const ETAT_SS_TRANSMISEFAX As String = "Transmise Fax"
    Public Const ETAT_SS_EXPORTEEINT As String = "Exportee Int"
    Public Const ETAT_SS_RAPPROCHEE As String = "Rapprochee"
    Public Const ETAT_SS_RAPPROCHEEINT As String = "RapprocheeInt"
    Public Const ETAT_SS_PROVIONNEE As String = "Provisionnee"
    Public Const ETAT_SS_FACTUREE As String = "Facturee"
    Public Const ETAT_EXPORTEE As String = "Exportée"
    Public Const ETAT_GENEREE As String = "Générée"

    'Libellé des Etats MvtSotk
    Public Const ETAT_MVTSTK_NON_FACTURE As String = "non facturé"
    Public Const ETAT_MVTSTK_FACTURE As String = "facturé"

    'Libellé des Etats Reglement
    Public Const ETAT_RGLMT_SAISI As String = "saisi"
    Public Const ETAT_RGLMT_EXPORTE As String = "exporté"

    Public Const DATE_DEFAUT As Date = #1/1/2000#

    Public Const CST_LGFACTTRP_NUM_GO As String = "999999" ' Numéro de ligne pour la ligne GO
    Public Const CST_LGFACTTRP_LIB_GO As String = "MAJORATION GAZOLE "

    Public Sub WaitnSeconds(ByVal n As Integer)
        System.Threading.Thread.Sleep(n * 1000)
    End Sub

    Public Function ConvertJJMMAAAToDate(ByVal s As String) As Date
        Dim dReturn As Date
        dReturn = Date.MinValue
        Try

            If s.Length = 6 Then
                'Date au format jjmmaa
                dReturn = New Date(CInt(Mid(s, 5, 2)), CInt(Mid(s, 3, 2)), CInt(Mid(s, 1, 2)))
            End If
            If s.Length = 8 Then
                'Date au format jjmmaaaa
                dReturn = New Date(CInt(Mid(s, 5, 4)), CInt(Mid(s, 3, 2)), CInt(Mid(s, 1, 2)))
            End If
        Catch ex As Exception
            dReturn = Date.MinValue
        End Try
        Return dReturn


    End Function
End Module
