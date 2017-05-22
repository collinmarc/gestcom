Imports System.Configuration
Imports System.Diagnostics
Imports vini_DB

Public Module VNC_Constants

    Public CONSTANTES_LOADED As Boolean = False
    Public Const SCROLLBARWIDTH As Short = 90
    Public Const ERR_APP As String = "ViniGestCom"
    Public Const ERR_DEBUG As Integer = 1
    Public Const DATE_DEFAUT As Date = #1/1/2000#
    Public TYPE_APP As Short = 0 ' 0 = Application Courtier (Vinicom) , 1 = Application Grossiste (Hobivin)
    Public GLOBALCONNECTION As String = "False"


    Public FAX_STK_BSENDCOVERPAGE As Boolean
    Public FAX_STK_SUBJECT As String
    Public FAX_STK_PATH As String
    Public FAX_STK_NOTES As String

    Public FAX_PRECOMMANDE_BSENDCOVERPAGE As Boolean
    Public FAX_PRECOMMANDE_SUBJECT As String
    Public FAX_PRECOMMANDE_PATH As String
    Public FAX_PRECOMMANDE_NOTES As String

    Public FAX_DETAILCOMMANDE_PATH As String
    Public FAX_DETAILCOMMANDE_SUBJECT As String
    Public FAX_DETAILCOMMANDE_BSENDCOVERPAGE As Boolean
    Public FAX_DETAILCOMMANDE_NOTES As String

    Public FAX_BL_PATH As String
    Public FAX_BL_SUBJECT As String
    Public FAX_BL_BSENDCOVERPAGE As Boolean
    Public FAX_BL_NOTES As String

    Public FAX_SCMD_PATH As String
    Public FAX_SCMD_SUBJECT As String
    Public FAX_SCMD_BSENDCOVERPAGE As Boolean
    Public FAX_SCMD_NOTES As String

    Public FAX_BA_PATH As String
    Public FAX_BA_SUBJECT As String
    Public FAX_BA_BSENDCOVERPAGE As Boolean
    Public FAX_BA_NOTES As String

    Public FAX_BLBA_PATH As String
    Public FAX_BLBA_SUBJECT As String
    Public FAX_BLBA_BSENDCOVERPAGE As Boolean
    Public FAX_BLBA_NOTES As String

    Public FAX_JAL_PATH As String
    Public FAX_JAL_SUBJECT As String
    Public FAX_JAL_BSENDCOVERPAGE As Boolean
    Public FAX_JAL_NOTES As String

    Public FAX_JAA_PATH As String
    Public FAX_JAA_SUBJECT As String
    Public FAX_JAA_BSENDCOVERPAGE As Boolean
    Public FAX_JAA_NOTES As String

    Public PATHTOREPORTS As String

    Public FAX_NOM_INTERLOCUTEUR As String
    Public FAX_TEL_INTERLOCUTEUR As String

    Public Const COMMANDECLIENT_LISTEPRODUIT_PRECOMMANDE As Boolean = True ' la liste des produits lors de la saisie de commande    True => Precommande False=> Tous les produits
    Public GESTSCMD_STANDARDCOLOR As System.Drawing.Color = System.Drawing.Color.Empty
    Public GESTSCMD_EXTRACOLOR As System.Drawing.Color = System.Drawing.Color.Gray

    Public IMPORT_IDSCMD As Integer
    Public IMPORT_REFFACTFOURN As Integer
    Public IMPORT_DATEFACTFOURN As Integer
    Public IMPORT_TOTALHTFACTURE As Integer
    Public IMPORT_TOTALTTCFACTURE As Integer
    '    Public IMPORT_TAUXCOMMISSION As Integer
    Public IMPORT_DIRECTORY As String
    Public EXPORTFTP_FILENAME As String
    Public IMPORTFTP_FILENAME As String

    Public Sub initConstantes()
        Try
            Try
                TYPE_APP = ConfigurationManager.AppSettings.GetValues("TYPE_APP")(0)
            Catch ex As Exception
                TYPE_APP = 0
            End Try

            PATHTOREPORTS = My.MySettings.Default.PathToReport
            Persist.ConnectionString = My.Settings.ConnectionString
            Persist.setReportCnx(My.Settings.ReportCnxUser, My.Settings.ReportCnxPassword)
            Try
                GLOBALCONNECTION = My.MySettings.Default.dbGlobalconnection
            Catch
                GLOBALCONNECTION = "False"
            End Try


            FAX_NOM_INTERLOCUTEUR = ConfigurationManager.AppSettings.GetValues("FAX_NOM_INTERLOCUTEUR")(0)
            FAX_TEL_INTERLOCUTEUR = ConfigurationManager.AppSettings.GetValues("FAX_TEL_INTERLOCUTEUR")(0)
            FAX_PRECOMMANDE_BSENDCOVERPAGE = ConfigurationManager.AppSettings.GetValues("FAX_PRECOMMANDE_BSENDCOVERPAGE")(0)
            FAX_PRECOMMANDE_SUBJECT = ConfigurationManager.AppSettings.GetValues("FAX_PRECOMMANDE_SUBJECT")(0)
            FAX_PRECOMMANDE_PATH = ConfigurationManager.AppSettings.GetValues("FAX_PRECOMMANDE_PATH")(0)
            FAX_PRECOMMANDE_NOTES = ConfigurationManager.AppSettings.GetValues("FAX_PRECOMMANDE_NOTES")(0)

            FAX_STK_BSENDCOVERPAGE = ConfigurationManager.AppSettings.GetValues("FAX_STK_BSENDCOVERPAGE")(0)
            FAX_STK_SUBJECT = ConfigurationManager.AppSettings.GetValues("FAX_STK_SUBJECT")(0)
            FAX_STK_PATH = ConfigurationManager.AppSettings.GetValues("FAX_STK_PATH")(0)
            FAX_STK_NOTES = ConfigurationManager.AppSettings.GetValues("FAX_STK_NOTES")(0)

            FAX_DETAILCOMMANDE_PATH = ConfigurationManager.AppSettings.GetValues("FAX_DETAILCOMMANDE_PATH")(0)
            FAX_DETAILCOMMANDE_SUBJECT = ConfigurationManager.AppSettings.GetValues("FAX_DETAILCOMMANDE_SUBJECT")(0)
            FAX_DETAILCOMMANDE_BSENDCOVERPAGE = ConfigurationManager.AppSettings.GetValues("FAX_DETAILCOMMANDE_BSENDCOVERPAGE")(0)
            FAX_DETAILCOMMANDE_NOTES = ConfigurationManager.AppSettings.GetValues("FAX_DETAILCOMMANDE_NOTES")(0)

            FAX_BL_PATH = ConfigurationManager.AppSettings.GetValues("FAX_BL_PATH")(0)
            FAX_BL_SUBJECT = ConfigurationManager.AppSettings.GetValues("FAX_BL_SUBJECT")(0)
            FAX_BL_BSENDCOVERPAGE = ConfigurationManager.AppSettings.GetValues("FAX_BL_BSENDCOVERPAGE")(0)
            FAX_BL_NOTES = ConfigurationManager.AppSettings.GetValues("FAX_BL_NOTES")(0)

            FAX_SCMD_PATH = ConfigurationManager.AppSettings.GetValues("FAX_SCMD_PATH")(0)
            FAX_SCMD_SUBJECT = ConfigurationManager.AppSettings.GetValues("FAX_SCMD_SUBJECT")(0)
            FAX_SCMD_BSENDCOVERPAGE = ConfigurationManager.AppSettings.GetValues("FAX_SCMD_BSENDCOVERPAGE")(0)

            FAX_BA_PATH = ConfigurationManager.AppSettings.GetValues("FAX_BA_PATH")(0)
            FAX_BA_SUBJECT = ConfigurationManager.AppSettings.GetValues("FAX_BA_SUBJECT")(0)
            FAX_BA_BSENDCOVERPAGE = ConfigurationManager.AppSettings.GetValues("FAX_BA_BSENDCOVERPAGE")(0)
            FAX_BA_NOTES = ConfigurationManager.AppSettings.GetValues("FAX_BA_NOTES")(0)

            FAX_BLBA_PATH = ConfigurationManager.AppSettings.GetValues("FAX_BLBA_PATH")(0)
            FAX_BLBA_SUBJECT = ConfigurationManager.AppSettings.GetValues("FAX_BLBA_SUBJECT")(0)
            FAX_BLBA_BSENDCOVERPAGE = ConfigurationManager.AppSettings.GetValues("FAX_BLBA_BSENDCOVERPAGE")(0)
            FAX_BLBA_NOTES = ConfigurationManager.AppSettings.GetValues("FAX_BLBA_NOTES")(0)

            FAX_JAL_PATH = ConfigurationManager.AppSettings.GetValues("FAX_JAL_PATH")(0)
            FAX_JAL_SUBJECT = ConfigurationManager.AppSettings.GetValues("FAX_JAL_SUBJECT")(0)
            FAX_JAL_BSENDCOVERPAGE = ConfigurationManager.AppSettings.GetValues("FAX_JAL_BSENDCOVERPAGE")(0)
            FAX_JAL_NOTES = ConfigurationManager.AppSettings.GetValues("FAX_JAL_NOTES")(0)

            FAX_JAA_PATH = ConfigurationManager.AppSettings.GetValues("FAX_JAA_PATH")(0)
            FAX_JAA_SUBJECT = ConfigurationManager.AppSettings.GetValues("FAX_JAA_SUBJECT")(0)
            FAX_JAA_BSENDCOVERPAGE = ConfigurationManager.AppSettings.GetValues("FAX_JAA_BSENDCOVERPAGE")(0)
            FAX_JAA_NOTES = ConfigurationManager.AppSettings.GetValues("FAX_JAA_NOTES")(0)

            IMPORT_IDSCMD = ConfigurationManager.AppSettings.GetValues("IMPORT_IDSCMD")(0)
            IMPORT_REFFACTFOURN = ConfigurationManager.AppSettings.GetValues("IMPORT_REFFACTFOURN")(0)
            IMPORT_DATEFACTFOURN = ConfigurationManager.AppSettings.GetValues("IMPORT_DATEFACTFOURN")(0)
            IMPORT_TOTALHTFACTURE = ConfigurationManager.AppSettings.GetValues("IMPORT_TOTALHTFACTURE")(0)
            IMPORT_TOTALTTCFACTURE = ConfigurationManager.AppSettings.GetValues("IMPORT_TOTALTTCFACTURE")(0)
            'IMPORT_TAUXCOMMISSION = ConfigurationManager.AppSettings.GetValues("IMPORT_TAUXCOMMISSION")(0)
            IMPORT_DIRECTORY = ConfigurationManager.AppSettings.GetValues("IMPORT_DIRECTORY")(0)
            EXPORTFTP_FILENAME = ConfigurationManager.AppSettings.GetValues("EXPORTFTP_FILENAME")(0)
            IMPORTFTP_FILENAME = ConfigurationManager.AppSettings.GetValues("IMPORTFTP_FILENAME")(0)
        Catch ex As Exception

        End Try
        CONSTANTES_LOADED = True
    End Sub

    Public Sub WaitnSeconds(ByVal n As Integer)
        If TimeOfDay >= #11:59:55 PM# Then
        End If
        Dim Start, Finish, TotalTime As Double
        Start = Microsoft.VisualBasic.DateAndTime.Timer
        Finish = Start + n   ' Set end time for 5-second duration.
        Do While Microsoft.VisualBasic.DateAndTime.Timer < Finish
            ' Do other processing while waiting for 5 seconds to elapse.
        Loop
        TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - Start
    End Sub


    Public Sub StartTrace(strFileNAme As String)
        Dim oTraceListener As TextWriterTraceListener

        If Not My.Computer.FileSystem.DirectoryExists("./LOGS") Then
            My.Computer.FileSystem.CreateDirectory("./LOGS")
        End If
        oTraceListener = New TextWriterTraceListener("./LOGS/" & strFileNAme)
        Trace.Listeners.Add(oTraceListener)
        Trace.AutoFlush = True
        Trace.WriteLine("Start at " + Now())
    End Sub




End Module
