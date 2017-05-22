Option Strict Off
Option Explicit On
'UPGRADE_WARNING: L'instanciation de classe a été modifiée en public. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1043"'
<System.Runtime.InteropServices.ProgId("clsItem_NET.clsItem")> Public Class clsItem
	
	Private mstrItemName As String
	Private mlngAttributes As Integer
	
	Private Const FILE_ATTRIBUTE_READONLY As Short = &H1s
	Private Const FILE_ATTRIBUTE_HIDDEN As Short = &H2s
	Private Const FILE_ATTRIBUTE_SYSTEM As Short = &H4s
	Private Const FILE_ATTRIBUTE_DIRECTORY As Short = &H10s
	Private Const FILE_ATTRIBUTE_ARCHIVE As Short = &H20s
	Private Const FILE_ATTRIBUTE_NORMAL As Short = &H80s
	Private Const FILE_ATTRIBUTE_TEMPORARY As Short = &H100s
	Private Const FILE_ATTRIBUTE_COMPRESSED As Short = &H800s
	Private Const FILE_ATTRIBUTE_OFFLINE As Short = &H1000s
	
	
	Friend Property Name() As String
		Get
			Return mstrItemName
		End Get
		Set(ByVal Value As String)
			mstrItemName = Value
		End Set
	End Property
	
	
	Friend Property Attributes() As Integer
		Get
			Return mlngAttributes
		End Get
		Set(ByVal Value As Integer)
			mlngAttributes = Value
		End Set
	End Property
	
	'UPGRADE_NOTE: ReadOnlya été mis à niveau vers ReadOnly_Renamed. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Friend ReadOnly Property ReadOnly_Renamed() As Boolean
		Get
			Return (mlngAttributes And FILE_ATTRIBUTE_READONLY) = FILE_ATTRIBUTE_READONLY
		End Get
	End Property
	
	Friend ReadOnly Property Hidden() As Boolean
		Get
			Return (mlngAttributes And FILE_ATTRIBUTE_HIDDEN) = FILE_ATTRIBUTE_HIDDEN
		End Get
	End Property
	
	'UPGRADE_NOTE: Systema été mis à niveau vers System_Renamed. Cliquez ici : 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'
	Friend ReadOnly Property System_Renamed() As Boolean
		Get
			Return (mlngAttributes And FILE_ATTRIBUTE_SYSTEM) = FILE_ATTRIBUTE_SYSTEM
		End Get
	End Property
	
	Friend ReadOnly Property Directory() As Boolean
		Get
			Return (mlngAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY
		End Get
	End Property
	
	Friend ReadOnly Property Archive() As Boolean
		Get
			Return (mlngAttributes And FILE_ATTRIBUTE_ARCHIVE) = FILE_ATTRIBUTE_ARCHIVE
		End Get
	End Property
	
	Friend ReadOnly Property Normal() As Boolean
		Get
			Return (mlngAttributes And FILE_ATTRIBUTE_NORMAL) = FILE_ATTRIBUTE_NORMAL
		End Get
	End Property
	
	Friend ReadOnly Property Temporary() As Boolean
		Get
			Return (mlngAttributes And FILE_ATTRIBUTE_TEMPORARY) = FILE_ATTRIBUTE_TEMPORARY
		End Get
	End Property
	
	Friend ReadOnly Property Compressed() As Boolean
		Get
			Return (mlngAttributes And FILE_ATTRIBUTE_COMPRESSED) = FILE_ATTRIBUTE_COMPRESSED
		End Get
	End Property
	
	Friend ReadOnly Property Offline() As Boolean
		Get
			Return (mlngAttributes And FILE_ATTRIBUTE_OFFLINE) = FILE_ATTRIBUTE_OFFLINE
		End Get
	End Property
End Class