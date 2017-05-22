'Option Strict Off
'Option Explicit On
'Public Class ftpcolItem

'    Private mcolItem As Collection

'    'UPGRADE_NOTE: Class_Initializea été mis à niveau vers Class_Initialize_Renamed. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
'    Private Sub Class_Initialize_Renamed()

'        'instantiate the collection
'        mcolItem = New Collection

'    End Sub
'    Public Sub New()
'        MyBase.New()
'        Class_Initialize_Renamed()
'    End Sub

'    'UPGRADE_NOTE: Class_Terminatea été mis à niveau vers Class_Terminate_Renamed. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
'    Private Sub Class_Terminate_Renamed()

'        'destroy the collection
'        'UPGRADE_NOTE: L'objet mcolItem ne peut pas être détruit tant qu'il n'est pas récupéré par le garbage collector (ramasse-miettes). Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
'        mcolItem = Nothing

'    End Sub
'    Protected Overrides Sub Finalize()
'        Class_Terminate_Renamed()
'        MyBase.Finalize()
'    End Sub

'    Friend Function Add(ByVal strName As String, ByVal lngAttributes As Integer) As ftpItem

'        Dim objNewItem As ftpItem

'        'instantiate a new item
'        objNewItem = New ftpItem

'        objNewItem.Name = strName
'        objNewItem.Attributes = lngAttributes

'        mcolItem.Add(objNewItem, strName)

'        'return newly created item
'        Add = objNewItem

'        'clean up after ourselves
'        'UPGRADE_NOTE: L'objet objNewItem ne peut pas être détruit tant qu'il n'est pas récupéré par le garbage collector (ramasse-miettes). Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
'        objNewItem = Nothing

'    End Function

'    Friend Sub Clear()

'        Dim lngCount As Integer

'        If Not IsNothing(mcolItem) Then
'            If mcolItem.Count > 0 Then
'                For lngCount = mcolItem.Count() To 1 Step -1
'                    mcolItem.Remove(lngCount)
'                Next lngCount
'            End If
'        End If

'    End Sub

'    Friend ReadOnly Property Count() As Integer
'        Get
'            Return mcolItem.Count()
'        End Get
'    End Property

'    Friend ReadOnly Property Item(ByVal vntItem As Object) As ftpItem
'        Get
'            Item = mcolItem.Item(vntItem)
'        End Get
'    End Property

'    Friend Function nRemove(ByVal lngIndex As Integer) As Object

'        If (lngIndex >= 1) And (lngIndex <= mcolItem.Count()) Then

'            mcolItem.Remove(lngIndex)
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet nRemove. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            nRemove = 0

'        Else

'            'bad return code
'            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet nRemove. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
'            nRemove = -1

'        End If

'    End Function
'End Class