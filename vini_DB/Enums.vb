Public Module vncEnums
    Public Enum CalculCommQte
        CALCUL_COMMISSION_QTE_LIVREE = 0
        CALCUL_COMMISSION_QTE_CMDE = 1
    End Enum
    Public Enum CalculCommScmd
        CALCUL_COMMISSION_HT_CALCULE = 0
        CALCUL_COMMISSION_HT_FACTURE = 1
        CALCUL_COMMISSION_BASE_COMM = 2
    End Enum
    Public Enum userRole
        ADMIN = 0
        COMMERCIAL = 1
        COMPTABILITE = 2
        INVITE = 3
    End Enum
    Public Enum vncfrmAction
        FRMNOTHING = 0
        FRMNEW = 1
        FRMLOAD = 2
        FRMSAVE = 3
        FRMDEL = 4
        FRMREFRESH = 5
    End Enum
    Public Enum vncTypeDonnee
        CLIENT = 1
        FOURNISSEUR = 2
        PRODUIT = 3
        COMMANDECLIENT = 4
        LGCOMMANDE = 5
        BA = 6
        LGPRECOM = 7
        MVTSTK = 8
        SSCOMMANDE = 9
        FACTCOMM = 10
        LGFACTCOMM = 11
        FACTTRP = 12
        LGFACTTRP = 13
        FACTCOL = 14
        LGFACTCOL = 15
        REGLEMENT = 16
        PRECOMMANDE = 17
        FACTCOMM_NONREGLEE = 100 'utilisée dans RechercheDB
        FACTTRP_NONREGLEE = 101
        FACTCOL_NONREGLEE = 102
    End Enum
    Public Enum vncEtatCommande
        vncRien = 0
        vncEnCoursSaisie = 1
        vncValidee = 2
        vncEclatee = 3
        vncTransmise = 4
        vncRapprochee = 5
        vncLivree = 6

        vncSCMDGeneree = 10
        vncSCMDtransmiseFax = 11
        vncSCMDExporteeInt = 15
        vncSCMDRapprochee = 12
        vncSCMDRapprocheeInt = 16
        '        vncSCMDProvisionnee = 13
        vncSCMDFacturee = 14

        vncBAEnCours = 100
        vncBALivre = 101

        vncFactComGeneree = 200
        '       vncFactComTransmise = 201
        vncFactComExportee = 202 ' Exportée vers Quadra


        vncFactTRPGeneree = 300
        '        vncFactTRPTransmise = 301
        vncFactTRPExportee = 302

        vncFactCOLGeneree = 400
        '        vncFactCOLTransmise = 401
        vncFactCOLExportee = 402
    End Enum

    Public Enum vncTypeCommande
        vncCmdClientPlateforme = 1
        vncCmdClientDirecte = 2
        vncCmdFournisseur = 3
    End Enum


    Public Enum vncTypeTransport
        vncTrpFranco = 1
        vncTrpDepart = 2
        vncTrpAvance = 3
    End Enum
    Public Enum vncActionEtatCommande
        vncActionValider = 1
        vncActionAnnValider = 2
        vncActionLivrer = 4
        vncActionAnnLivrer = 5
        vncActionEclater = 6
        vncActionAnnEclater = 7
        vncActionTransmettre = 8
        vncActionAnnTransmettre = 9
        vncActionMinCommande = 0
        vncActionMaxCommande = 9

        vncActionSCMDFaxer = 21
        vncActionSCMDAnnTransmettre = 22
        vncActionSCMDRapprocher = 23
        vncActionSCMDAnnRapprocher = 24
        vncActionSCMDFacturer = 27
        vncActionSCMDAnnFacturer = 28
        vncActionSCMDExportInternet = 29
        vncActionSCMDImportInternet = 30
        vncActionSCMDAnnImportInternet = 31

        vncActionBALivrer = 100
        vncActionBAAnnLivrer = 101
        vncActionMinBA = 100
        vncActionMaxBA = 101

        vncActionFactComGenerer = 200
        vncActionFactComAnnGenerer = 201
        vncActionFactComExporter = 220
        vncActionFactComAnnExporter = 221
        vncActionMinFactCom = 200
        vncActionMaxFactCom = 221

        vncActionFactTRPGenerer = 300
        vncActionFactTRPAnnGenerer = 301
        vncActionFactTRPExporter = 320
        vncActionFactTRPAnnExporter = 321
        vncActionMinFactTRP = 300
        vncActionMaxFactTRP = 321

        vncActionFactCOLGenerer = 400
        vncActionFactCOLAnnGenerer = 401
        vncActionFactCOLExporter = 420
        vncActionFactCOLAnnExporter = 421
        vncActionMinFactCOL = 400
        vncActionMaxFactCOL = 421
    End Enum
    Public Enum vncGenererSupprimer
        vncRien = 0
        vncGenerer = 1
        vncSupprimer = 2
    End Enum
    Public Enum vncTypeMvt
        vncMvtInventaire = 1
        vncMvtCommandeClient = 2
        vncmvtBonAppro = 3
        vncmvtRegul = 4
    End Enum
    Public Enum vncTypeProduit
        vncPlateforme = 1
        vncFournisseur = 2
        vncTous = 3
    End Enum

    Public Enum vncEtatMVTSTK
        vncMVTSTK_Tous = 99
        vncMVTSTK_nFact = 0
        vncMVTSTK_Fact = 1
    End Enum

    Public Enum vncActionFactColisage
        vncActionFacturer = 1
        vncActionAnnFacturer = 2
    End Enum

    Public Enum vncEtatReglement
        vncRGLMT_Tous = 99
        vncRGLMT_Saisi = 0
        vncRGLMT_Export = 1
    End Enum
    Public Enum vncActionReglement
        vncActionExporter = 1
        vncActionAnnExporter = 2
    End Enum
    Public Class vncObjectProperties
        Public Shared CMD_ID As String = "CMD_ID"
        Public Shared CMD_CODE As String = "CMD_CODE"
        Public Shared CMD_dateCommande As String = "CMD_DATECOMMANDE"
        Public Shared CMD_etat As String = "CMD_ETAT"
        '        Public Shared CMD_colLignes As String = "CMD_CODE"
        Public Shared CMD_TotalHT As String = "CMD_TOTALHT"
        Public Shared CMD_TotalTTC As String = "CMD_TOTALTTC"
        Public Shared CMD_TransporteurCODE As String = "CMD_TRANSPORTEURCODE"
        Public Shared CMD_dateValidation As String = "CMD_DATEVALIDATION"
        Public Shared CMD_dateEnlevement As String = "CMD_DATEENLEV"
        Public Shared CMD_dateLivraison As String = "CMD_DATELIVRAISON"
        Public Shared CMD_RefLivraison As String = "CMD_REFLIVRAISON"
        Public Shared CMD_TiersID As String = "CMD_TIERS_ID"
        Public Shared CMD_TiersCODE As String = "CMD_TIERS_CODE"
        Public Shared CMD_TIERS_NOM As String = "CMD_TIERS_NOM"
        Public Shared CMD_TIERS_RS As String = "CMD_TIERS_RS"
        Public Shared CMD_TIERS_ADLIV1 As String = "CMD_TIERS_ADLIV1"
        Public Shared CMD_TIERS_ADLIV2 As String = "CMD_TIERS_ADLIV2"
        Public Shared CMD_TIERS_ADLIVCP As String = "CMD_TIERS_ADLIVCP"
        Public Shared CMD_TIERS_ADLIVVILLE As String = "CMD_TIERS_ADLIVVILLE"
        Public Shared CMD_TIERS_ADLIVTEL As String = "CMD_TIERS_ADLIVTEL"
        Public Shared CMD_TIERS_ADLIVFAX As String = "CMD_TIERS_ADLIVFAX"
        Public Shared CMD_TIERS_ADLIVPORT As String = "CMD_TIERS_ADLIVPORT"
        Public Shared CMD_TIERS_ADLIVEMAIL As String = "CMD_TIERS_ADLIVEMAIL"
        Public Shared CMD_TIERS_ADFACT1 As String = "CMD_TIERS_ADFACT1"
        Public Shared CMD_TIERS_ADFACT2 As String = "CMD_TIERS_ADFACT2"
        Public Shared CMD_TIERS_ADFACTCP As String = "CMD_TIERS_ADFACTCP"
        Public Shared CMD_TIERS_ADFACTVILLE As String = "CMD_TIERS_ADFACTVILLE"
        Public Shared CMD_TIERS_ADFACTTEL As String = "CMD_TIERS_ADFACTTEL"
        Public Shared CMD_TIERS_ADFACTFAX As String = "CMD_TIERS_ADFACTFAX"
        Public Shared CMD_TIERS_ADFACTPORT As String = "CMD_TIERS_ADFACTPORT"
        Public Shared CMD_TIERS_ADFACTEMAIL As String = "CMD_TIERS_ADFACTEMAIL"
        Public Shared CMD_qteColis As String = "CMD_QTECOLIS"
        Public Shared CMD_qtePalettesNonPreparees As String = "CMD_QTEPALNONPREP"
        Public Shared CMD_qtePalettesPreparees As String = "CMD_QTEPALPREP"
        Public Shared CMD_poids As String = "CMD_POIDS"
        Public Shared CMD_puPalettesNonPreparees As String = "CMD_PUPALNONPREP"
        Public Shared CMD_puPalettesPreparees As String = "CMD_PUPALPREP"
        Public Shared CMD_MontantTransport As String = "CMD_MTTRANSPORT"
        Public Shared CMD_LettreVoiture As String = "CMD_LETTREVOITURE"
        Public Shared CMD_CoutTransport As String = "CMD_COUTRANSPORT"
        Public Shared CMD_RefFactTrp As String = "CMD_REFFACTTRP"
        Public Shared CMD_ComentCom As String = "CMD_COM_COM"
        Public Shared CMD_ComentLiv As String = "CMD_COM_LIV"
        Public Shared CMD_ComentFact As String = "CMD_COM_FACT"
        Public Shared CMD_ComentLibre As String = "CMD_COM_LIBRE"
        Public Shared LGCMD_num As String = "LGCMD_NUM"
        Public Shared LGCMD_PRD_CODE As String = "LGCMD_PRD_CODE"
        Public Shared LGCMD_qteCom As String = "LGCMD_QTECOM"
        Public Shared LGCMD_qteLiv As String = "LGCMD_QTELIV"
        Public Shared LGCMD_qteFact As String = "LGCMD_QTEFACT"
        Public Shared LGCMD_prixU As String = "LGCMD_PU"
        Public Shared LGCMD_prixHT As String = "LGCMD_PRIXHT"
        Public Shared LGCMD_prixTTC As String = "LGCMD_PRICTTC"
        Public Shared LGCMD_bGratuit As String = "LGCMD_BGRATUIT"
        Public Shared LGCMD_poids As String = "LGCMD_POIDS"
        Public Shared LGCMD_qteColis As String = "LGCMD_QTECOLIS"
        Public Shared LGCMD_TxComm As String = "LGCMD_TXCOMM"
        Public Shared LGCMD_MtComm As String = "LGCMD_MTCOMM"

    End Class
End Module
