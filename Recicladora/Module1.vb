Imports System.Data
Imports System.Data.OleDb

Module Module1

    Public licenciaIngresada As String

    Public Const WM_CAP As Short = &H400S
    Public Const WM_CAP_DRIVER_CONNECT As Integer = WM_CAP + 10
    Public Const WM_CAP_DRIVER_DISCONNECT As Integer = WM_CAP + 11
    Public Const WM_CAP_EDIT_COPY As Integer = WM_CAP + 30
    Public Const WM_CAP_GET_STATUS As Integer = WM_CAP + 54
    Public Const WM_CAP_DLG_VIDEOFORMAT As Integer = WM_CAP + 41
    Public Const WM_CAP_SET_PREVIEW As Integer = WM_CAP + 50
    Public Const WM_CAP_SET_PREVIEWRATE As Integer = WM_CAP + 52
    Public Const WM_CAP_SET_SCALE As Integer = WM_CAP + 53
    Public Const WS_CHILD As Integer = &H40000000
    Public Const WS_VISIBLE As Integer = &H10000000
    Public Const SWP_NOMOVE As Short = &H2S
    Public Const SWP_NOSIZE As Short = 1
    Public Const SWP_NOZORDER As Short = &H4S
    Public Const HWND_BOTTOM As Short = 1
    Public DeviceID As Integer = 0 ' Current device ID
    Public hHwnd As Integer ' Handle to preview window
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer,
        ByRef lParam As CAPSTATUS) As Boolean
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Boolean,
       ByRef lParam As Integer) As Boolean
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
         (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer,
         ByRef lParam As Integer) As Boolean
    Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Integer,
        ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer,
        ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

    Declare Function DestroyWindow Lib "user32" (ByVal hndw As Integer) As Boolean
    Structure POINTAPI
        Dim x As Integer
        Dim y As Integer
    End Structure
    Public Structure CAPSTATUS
        Dim uiImageWidth As Integer                    '// Width of the image
        Dim uiImageHeight As Integer                   '// Height of the image
        Dim fLiveWindow As Integer                     '// Now Previewing video?
        Dim fOverlayWindow As Integer                  '// Now Overlaying video?
        Dim fScale As Integer                          '// Scale image to client?
        Dim ptScroll As POINTAPI                    '// Scroll position
        Dim fUsingDefaultPalette As Integer            '// Using default driver palette?
        Dim fAudioHardware As Integer                  '// Audio hardware present?
        Dim fCapFileExists As Integer                  '// Does capture file exist?
        Dim dwCurrentVideoFrame As Integer             '// # of video frames cap'td
        Dim dwCurrentVideoFramesDropped As Integer     '// # of video frames dropped
        Dim dwCurrentWaveSamples As Integer            '// # of wave samples cap'td
        Dim dwCurrentTimeElapsedMS As Integer          '// Elapsed capture duration
        Dim hPalCurrent As Integer                     '// Current palette in use
        Dim fCapturingNow As Integer                   '// Capture in progress?
        Dim dwReturn As Integer                        '// Error value after any operation
        Dim wNumVideoAllocated As Integer              '// Actual number of video buffers
        Dim wNumAudioAllocated As Integer              '// Actual number of audio buffers
    End Structure
    Declare Function capCreateCaptureWindowA Lib "avicap32.dll" _
         (ByVal lpszWindowName As String, ByVal dwStyle As Integer,
         ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer,
         ByVal nHeight As Short, ByVal hWndParent As Integer,
         ByVal nID As Integer) As Integer
    Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Short,
        ByVal lpszName As String, ByVal cbName As Integer, ByVal lpszVer As String,
        ByVal cbVer As Integer) As Boolean

    Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long

    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long,
                                        ByVal nIDEvent As Long) As Long

    Private Declare Function SendMessageLongRef Lib "user32" _
        Alias "SendMessageA" (
        ByVal hwnd As Long,
        ByVal wMsg As Long,
        ByVal wParam As Long,
        ByRef lParam As Long) As Long

    Private Declare Function FindWindow Lib "user32" _
        Alias "FindWindowA" (
        ByVal lpClassName As String,
        ByVal lpWindowName As String) As Long

    Private Declare Function FindWindowEx Lib "user32" _
        Alias "FindWindowExA" (
        ByVal hWnd1 As Long,
        ByVal hWnd2 As Long,
        ByVal lpsz1 As String,
        ByVal lpsz2 As String) As Long

    Private Declare Function SetTimer Lib "user32" (
        ByVal hwnd As Long,
        ByVal nIDEvent As Long,
        ByVal uElapse As Long,
        ByVal lpTimerFunc As Long) As Long


    Private m_ASC As Long

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (
                    ByVal hwnd As Long,
                    ByVal wMsg As Long,
                    ByVal wParam As Long,
                    lParam As VariantType) As Long
    Private Const CB_ERR = -1, CB_SELECTSTRING = &H14D, CB_SHOWDROPDOWN = &H14F, CBN_SELENDOK = 9
    ' ************************************************END AUTOFILL



    Public textoTeclado As String 'Captura la informacion del teclado en la forma que invoca
    Public timerKeyboard As Boolean 'Indica al Timer cuando apagarse

    Public Declare Function GetWindowLong Lib "user32" Alias _
              "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

    Public Declare Function SetWindowLong Lib "user32" Alias _
              "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long,
              ByVal dwNewLong As Long) As Long

    Public Const GWL_STYLE = (-16)
    Public Const WS_MAXIMIZEBOX = &H10000
    Public Const WS_MINIMIZEBOX = &H20000

    ' VARIABLES PARA ODBC
    Public rs As New ADODB.Recordset
    Public rs1 As New ADODB.Recordset
    Public DBCon As ADODB.Connection 'variable de conexion
    Public Cmd As ADODB.Command
    Public mystream As New ADODB.Stream
    ' END VARIABLES PARA ODBC

    'Variables Globales
    Public formaPadre As String
    Public botonPadre As String
    Public modulo As Integer
    Public moduloTabla As String
    Public tipoMovimiento As Integer
    Public statusRegistro As Integer
    Public busquedaRegistroNumero As Integer ' Variable para buscar un numero registro segun la forma de ejecucion
    Public nombreUsuario As String
    Public numeroUsuario As Integer ' Guarda nombre y numero de usuario logeado al sistema
    Public fechaActual As String
    Public horaActual As String
    Public valorSeleccionado As Boolean
    Public LookupCaller As Form ' Forma padre
    ' Variables de conexion Base de datos
    Public dataBaseNameMysql As String ' Nombre de la base de datos
    Public dataBaseConector As String ' Nombre del conector
    Public dataBaseIp As String
    Public dataBaseUser As String
    Public dataBasePassword As String

    Sub connection_Db()

        DBCon = New ADODB.Connection
        DBCon.CursorLocation = 2
        '    DBCon.CursorLocation = adUseServer
        On Error GoTo lineaError
        DBCon.Open("Driver={" & dataBaseConector & "};Server=" & dataBaseIp & "; Database=" & dataBaseNameMysql & "; User=" & dataBaseUser & ";Password=" & dataBasePassword & " ;Option=3;")
        Cmd = New ADODB.Command
        Cmd.ActiveConnection = DBCon
        Cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
        Cmd.CommandTimeout = 120
        Exit Sub
lineaError:
        MsgBox("Hay un error en los parametros de conexion, favor de verificar", vbCritical, "Error")
        End
    End Sub

    Function get_Numbers(ByVal registro As String) As Integer
        Dim i As Integer
        Dim numero As Integer 'NUMERO DE PROVEEDOR A BUSCAR

        For i = 1 To 5
            If IsNumeric(Mid(registro, i, 1)) Then
                numero = numero & Mid(registro, i, 1)
            End If
        Next i
        get_Numbers = numero
    End Function

    Function get_String(ByVal registro As String) As String
        Dim i As Integer
        Dim n As Integer
        Dim result As String = "" 'NUMERO DE PROVEEDOR A BUSCAR
        n = Len(registro)
        For i = 1 To n
            If Not IsNumeric(Mid(registro, i, 1)) Then
                result &= Mid(registro, i, 1)
            End If
        Next i
        get_String = result
    End Function

    Function numeroValido(ByVal numero As String) As Boolean
        If IsNumeric(numero) Then
            numeroValido = True
        Else
            numeroValido = False
        End If
    End Function

    Public Sub LookUp(caller As Form)
        If LookupCaller Is Nothing Then
            caller.Enabled = False
            LookupCaller = caller
        Else
            '        MsgBox Me.Caption & " already in use." & vbCrLf & "Please complete prvious request."
        End If
        '    Me.Show
        '    Me.ZOrder
    End Sub

    Public Sub emptyForm(caller As Form)
        If Not LookupCaller Is Nothing Then
            LookupCaller.Enabled = True
            caller.Hide()
            caller.Close()
            caller.Dispose()
        Else
            caller.Show()
        End If
        LookupCaller = Nothing
    End Sub

    Public Function licencia(texto As String) As String
        Select Case texto
            Case "50l0unm35"
                licencia = "G3I0L"
            Case "50l0unan0"
                licencia = "B3E6R5T"
            Case "ilimi7ad0"
                licencia = "O3J6A5R0A0M"
            Case "r3ac7ivaci0n"
                licencia = "VALIDO"
            Case Else
                licencia = "NULL"
        End Select
    End Function
End Module
