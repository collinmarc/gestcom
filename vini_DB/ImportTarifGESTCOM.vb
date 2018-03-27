Imports Microsoft.Office.Interop
Imports vini_DB

Public Class ImportTarifGESTCOM
    Inherits Observable

    Private m_FileName As String
    Private m_SheetName As String
    Private m_numColTarif As Integer
    Private m_numColCode As Integer
    Public Sub New(pFileName As String, pSheetName As String, pNumColCode As Integer, pNumColTarif As Integer)
        m_FileName = pFileName
        m_SheetName = pSheetName
        m_numColCode = pNumColCode
        m_numColTarif = pNumColTarif

    End Sub
    Public Overrides ReadOnly Property shortResume As String
        Get
            Return "Import Tarif Gestcom"
        End Get
    End Property

    Private _Message As String
    Public Property message() As String
        Get
            Return _Message
        End Get
        Set(ByVal value As String)
            _Message = value
        End Set
    End Property
    Public Function getNbreLignes() As Integer
        Dim nReturn As Integer = 0
        Try
            If System.IO.File.Exists(m_FileName) Then

                Dim objApp As New Excel.Application
                objApp.Visible = False
                objApp.Workbooks.Open(m_FileName)
                Dim oSheet As Excel.Worksheet
                oSheet = objApp.Worksheets.Item(m_SheetName)
                nReturn = oSheet.UsedRange.Rows.Count
                objApp.Workbooks(1).Close()
                objApp.Quit()
            End If

        Catch ex As Exception
            nReturn = 0
        End Try
        Return nReturn
    End Function
    Public Function ImportTarif() As Boolean
        Dim bReturn As Boolean = False
        Try
            If System.IO.File.Exists(m_FileName) Then

                Dim objApp As New Excel.Application
                objApp.Visible = False
                objApp.Workbooks.Open(m_FileName)
                Dim oSheet As Excel.Worksheet
                oSheet = objApp.Worksheets(m_SheetName)

                Dim nRow As Integer = 0
                For Each oRow As Excel.Range In oSheet.UsedRange.Rows
                    nRow = nRow + 1
                    If Not String.IsNullOrEmpty(oSheet.Cells(nRow, m_numColCode).Value) Then
                        Dim oProduit As Produit
                        oProduit = Produit.createandloadbyKey(oSheet.Cells(nRow, m_numColCode).Value)
                        If oProduit IsNot Nothing Then
                            Try

                                Dim tarif As Decimal = Convert.ToDecimal(oSheet.Cells(nRow, m_numColTarif).Value)
                                If tarif > 0 Then
                                    oProduit.TarifA = tarif
                                    oProduit.TarifB = tarif
                                    oProduit.TarifC = tarif
                                    oProduit.save()
                                    Me.message = "Ligne " & nRow & ":" & oProduit.code
                                End If
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                    Notifier()
                Next
                objApp.Workbooks(1).Close()
                objApp.Quit()
                bReturn = True
            End If
        Catch ex As Exception
            bReturn = False
        End Try
        Return bReturn
    End Function



End Class
