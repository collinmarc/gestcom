Imports System.Xml.Serialization
Imports System.Collections
Imports System.IO
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions
<Serializable()>
Public Class cmdprestashop
    Public id As String
    Public name As String
    Public origine As String
    Public customer_id As String
    Public company As String
    Public livraison_company As String
    Public livraison_name As String
    Public livraison_firstname As String
    Public livraison_adress1 As String
    Public livraison_adress2 As String
    Public livraison_postalcode As String
    Public livraison_city As String
    Public lignes As List(Of ligneprestashop)
    <XmlIgnore()>
    Public motif As String

    Public Sub New()
        id = ""
        name = ""
        origine = Dossier.VINICOM
        customer_id = ""
        company = ""
        livraison_company = ""
        livraison_name = ""
        livraison_firstname = ""
        livraison_adress1 = ""
        livraison_adress2 = ""
        livraison_postalcode = ""
        livraison_city = ""
        lignes = New List(Of ligneprestashop)
        motif = ""

    End Sub

    ''' <summary>
    ''' Lecture d'une chaine XML et retour d'une Commande
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function readXML(pXmlString As String) As cmdprestashop
        Dim oReturn As New cmdprestashop
        Dim strXMl As String = pXmlString
        Try
            Dim oCmd As New cmdprestashop
            strXMl = strXMl.Replace("=3D", "=")
            strXMl = strXMl.Replace("[", "<")
            strXMl = strXMl.Replace("]", ">")
            strXMl = strXMl.Replace("#", "=")

            Dim oImap As New ImapVB.Imap
            Dim strOut As String = ""
            strOut = oImap.DecodequotedprintableString(strXMl, "utf-8")
            Dim oSr As New StringReader(strOut)
            '            Dim objStreamReader As New StreamReader(oSr)
            Dim x As New XmlSerializer(GetType(cmdprestashop))
            oReturn = x.Deserialize(oSr)
            'Suppression des Space dans les codes Produits
            For Each oLg As ligneprestashop In oReturn.lignes
                oLg.reference = Trim(oLg.reference)
                oLg.reference.Replace(" ", "")
            Next
            '"           objStreamReader.Close()
        Catch ex As Exception
            Debug.Print("cmdPrestashop.readXML: " & ex.Message)
            If ex.InnerException IsNot Nothing Then
                Debug.Print("cmdPrestashop.readXML: " & ex.InnerException.Message)
            End If
            Debug.Print("cmdPrestashop.readXML: " & pXmlString)
            Debug.Print("cmdPrestashop.readXML: " & strXMl)

        End Try
        Return oReturn
    End Function

    Public Shared Function FTO_writeXml(pCmd As cmdprestashop, pFilename As String) As Boolean
        Dim bReturn As Boolean
        Dim objStreamWriter As StreamWriter = Nothing

        Try
            Dim ns As New XmlSerializerNamespaces()
            ns.Add("", "")

            objStreamWriter = New StreamWriter(pFilename)
            Dim x As New XmlSerializer(pCmd.GetType)
            x.Serialize(objStreamWriter, pCmd, ns)
            objStreamWriter.Close()
            bReturn = True
        Catch ex As Exception
            Debug.Print("ParamDiag.WriteXML: " & ex.Message)
            If ex.InnerException IsNot Nothing Then
                Debug.Print("ParamDiag.WriteXML: " & ex.InnerException.Message)
            End If
            bReturn = False
            If objStreamWriter IsNot Nothing Then
                objStreamWriter.Close()
            End If


        End Try
        Return bReturn
    End Function
    ''' <summary>
    ''' Vérification de la Commande importée
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function check() As Boolean
        Dim bReturn As Boolean
        Try
            bReturn = True
            If String.IsNullOrEmpty(id) Then
                motif = "idCommande non renseigné"
                bReturn = False
            End If
            If bReturn Then
                If String.IsNullOrEmpty(name) Then
                    motif = "nameCommande non renseigné"
                    bReturn = False
                End If
            End If
            If bReturn Then
                If Not String.IsNullOrEmpty(customer_id) Then
                    Dim oClient As Client
                    oClient = Client.createandloadPrestashop(customer_id)
                    If oClient Is Nothing Then
                        motif = "Client Inconnu (" + customer_id + ") "
                        bReturn = False
                    End If
                End If
            End If
            If bReturn Then
                Dim oProduit As Produit
                For Each oLg As ligneprestashop In lignes
                    oProduit = Produit.createandloadbyKey(oLg.reference)
                    If oProduit Is Nothing Then
                        motif = "Produit Inconnu (" + oLg.reference + ")"
                        bReturn = False
                    End If
                Next
            End If
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function

    Public Sub AddLigne(pRef As String, pQte As Integer, pPrix As Decimal)
        lignes.Add(New ligneprestashop(pRef, pQte, pPrix))
    End Sub
    ''' <summary>
    ''' Création d'une commande Client à partir d'une commande prestaShop
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function createCommandeClient() As CommandeClient
        Dim oReturn As CommandeClient
        oReturn = Nothing
        Try
            Dim oClient As Client
            Dim oProduit As Produit
            Dim nLigne As Integer
            oClient = Client.createandloadPrestashop(customer_id)
            If oClient IsNot Nothing Then
                oReturn = New CommandeClient(oClient)
                oReturn.IDPrestashop = id
                oReturn.NamePrestashop = name
                oReturn.Origine = origine
                oReturn.caracteristiqueTiers.nom = company
                oReturn.caracteristiqueTiers.rs = livraison_company
                oReturn.RaisonSocialeLivraison = livraison_company
                oReturn.NomLivraison = livraison_name & " " & livraison_firstname
                oReturn.caracteristiqueTiers.AdresseLivraison.nom = livraison_name & " " & livraison_firstname
                oReturn.caracteristiqueTiers.AdresseLivraison.rue1 = livraison_adress1
                oReturn.caracteristiqueTiers.AdresseLivraison.rue2 = livraison_adress2
                oReturn.caracteristiqueTiers.AdresseLivraison.cp = livraison_postalcode
                oReturn.caracteristiqueTiers.AdresseLivraison.ville = livraison_city

                oReturn.caracteristiqueTiers.AdresseFacturation.nom = oReturn.caracteristiqueTiers.AdresseLivraison.nom
                oReturn.caracteristiqueTiers.AdresseFacturation.rue1 = oReturn.caracteristiqueTiers.AdresseLivraison.rue1
                oReturn.caracteristiqueTiers.AdresseFacturation.rue2 = oReturn.caracteristiqueTiers.AdresseLivraison.rue2
                oReturn.caracteristiqueTiers.AdresseFacturation.cp = oReturn.caracteristiqueTiers.AdresseLivraison.cp
                oReturn.caracteristiqueTiers.AdresseFacturation.ville = oReturn.caracteristiqueTiers.AdresseLivraison.ville

                oReturn.caracteristiqueTiers.idModeReglement = oClient.idModeReglement
                oReturn.caracteristiqueTiers.idModeReglement1 = oClient.idModeReglement1
                oReturn.caracteristiqueTiers.idModeReglement2 = oClient.idModeReglement2
                oReturn.caracteristiqueTiers.idModeReglement3 = oClient.idModeReglement3
                oReturn.caracteristiqueTiers.rib1 = oClient.rib1
                oReturn.caracteristiqueTiers.rib2 = oClient.rib2
                oReturn.caracteristiqueTiers.rib3 = oClient.rib3
                oReturn.caracteristiqueTiers.rib4 = oClient.rib4
                nLigne = 10
                For Each oLg As ligneprestashop In lignes
                    oProduit = Produit.createandloadbyKey(oLg.reference)
                    If oProduit IsNot Nothing Then
                        oReturn.AjouteLigne(nLigne, oProduit, oLg.quantite, oLg.prixunitaire)
                        nLigne = nLigne + 10
                    End If
                Next
            End If
        Catch ex As Exception
            oReturn = Nothing
        End Try
        Return oReturn
    End Function
End Class
Public Class ligneprestashop
    Public reference As String
    Public quantite As Integer
    Public prixunitaire As Decimal
    Public Sub New()

    End Sub
    Public Sub New(preference As String, pQte As Integer, pPrix As Decimal)
        reference = preference
        quantite = pQte
        prixunitaire = pPrix
    End Sub
End Class


