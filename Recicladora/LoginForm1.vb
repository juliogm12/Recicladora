Public Class LoginForm1

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Public usuario As String = ""
    Dim texto, puerto As String
    Dim accesoValido As Boolean

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Cmd.CommandText = "SELECT COUNT(*), usuario, tipo_usuario AS cuenta FROM re_usuarios WHERE usuario = '" & UsernameTextBox.ToString & "' AND BINARY PASSWORD = '" & PasswordTextBox.ToString & "'"
        rs = Cmd.Execute
        If rs("cuenta").Value = 1 Then
            usuario = rs("usuario").Value
            Dim forma As New MDIParent1
            If rs("tipo_usuario").Value = 1 Then

            End If
            forma.Show()
            Me.Close()
        Else
            MsgBox("El usuario y/o contraseña no son correctos. Favor de verificar.", vbCritical, "Error")
        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub LoginForm1_Load(sender As Object, e As EventArgs) Handles Me.Load
        'PARAVALIDACION
        'GUARDAR LA FECHA ACTUAL Y CCOMPARARLA CON EL DIA ACTUAL, SI DIA DE ACTUAL NULO ENTONCES GUARDAR DIAACTUAL
        'SI DIA ACTUAL IGUAL A DIA HOY NO GUARDAR
        'SI LA DIFERENCIA ENTRE DIA ACTUAL Y DIA DE HOY ES MAYOR A 4 DIAS PEDIR Licencia DE REACTIVACION
        'finDeLicenciamiento: Abgabeende
        'FechaInicioLicenciamiento DatumHeimLicenciaierung
        'fNoModificable Datumnichtänderbar
        'Licencia Licencia

        dataBaseNameMysql = GetSetting("Recicladora", "Properties", "Base de datos", "")
        dataBaseConector = GetSetting("Recicladora", "Properties", "Conector", "")
        dataBaseIp = GetSetting("Recicladora", "Properties", "IP", "")
        dataBaseUser = GetSetting("Recicladora", "Properties", "Usuario", "")
        dataBasePassword = GetSetting("Recicladora", "Properties", "Password", "")
        connection_Db()
        On Error Resume Next
        puerto = GetSetting("BasculasReg", "Properties", "Puerto", "")
        Dim restante, diferencia As Long
        Dim resultado As String
        Dim licenciaFin As String
        On Error Resume Next
        '*****LLENADO DE COMBO 1 CON NOMBRE DE PROVEEDOR


        On Error Resume Next
        '    cerrarVentana = False
        Dim hoy As Date = Format(Now, "yyyy-mm-dd") 'DateValue(
        Dim fechaActualNoModificable As Date = DateValue((GetSetting("BasculasReg", "Properties", "Datumnichtanderbar", "")))
        Dim LicenciaRegEdit = (GetSetting("BasculasReg", "Properties", "Lizenz", ""))
        Dim finLicencia As Date = (GetSetting("BasculasReg", "Properties", "Abgabeende", finLicencia))

        If LicenciaRegEdit = "" Then
            Dim ingreseLicenicia As New Form1
            ingreseLicenicia.ShowDialog()

            texto = licenciaIngresada
            resultado = licencia(texto)
            If resultado = "NULL" Then
                MsgBox("LICENCIA INCORRECTA", 16, "ERROR")
                End
            Else
                Dim Final As Double
                For Index = 1 To Len(resultado)
                    If Mid(resultado, Index, 1) Like "[0-9]" Then
                        Final = Final & Mid(resultado, Index, 1)
                    End If
                Next



                Call SaveSetting("BasculasReg", "Properties", "Lizenz", resultado)
                Call SaveSetting("BasculasReg", "Properties", "DatumHeimLicenciaierung", hoy)
                finLicencia = Format(DateAdd("D", Final, hoy), "yyyy-mm-dd")
                Call SaveSetting("BasculasReg", "Properties", "Abgabeende", finLicencia)
                restante = DateDiff("d", hoy, finLicencia)
                accesoValido = True
            End If
        Else
            If fechaActualNoModificable = "" Then
                Call SaveSetting("BasculasReg", "Properties", "Abgabeende", finLicencia)
                restante = DateDiff("d", hoy, finLicencia)
                accesoValido = True
            Else
                diferencia = DateDiff("d", hoy, fechaActualNoModificable)
                If diferencia < -4 Or diferencia > 4 Then
                    MsgBox("Pida una reactivacion", vbInformation)
                    Dim react As New Form1
                    react.ShowDialog()
                    texto = licenciaIngresada
                    resultado = licencia(texto)
                    If resultado = "VALIDO" Then
                        licenciaFIN = (GetSetting("BasculasReg", "Properties", "Abgabeende", ""))
                        restante = DateDiff("d", hoy, licenciaFIN)
                        accesoValido = True
                    Else
                        MsgBox("REACTIVACION INVALIDA PONGASE EN CONTACTO CON EL ADMINISTRADOR", 16, "ERROR")
                    End If
                Else
                    'DIAS RESTANTES
                    licenciaFIN = (GetSetting("BasculasReg", "Properties", "Abgabeende", ""))
                    restante = DateDiff("d", hoy, licenciaFIN)
                    If restante > 15 Then
                        accesoValido = True
                    ElseIf restante <= 5 And restante >= 1 Then
                        MsgBox("La licencia se vence en " & restante & " días", vbInformation, "Licencia")
                        accesoValido = True
                    ElseIf restante <= 0 Then
                        MsgBox("La licencia ha finalizado", vbCritical, "REACTIVACION")
                        Call SaveSetting("BasculasReg", "Properties", "Lizenz", "")
                        Me.Close()
                        Dim mes As New LoginForm1
                        mes.Show()
                    End If
                End If
            End If
        End If
        If accesoValido = True Then
            Call SaveSetting("BasculasReg", "Properties", "Datumnichtanderbar", hoy)
            Dim logo As String
            ' MsgBox "El tiempo restante de Licencia es: " & restante & " días", vbInformation, "Licencia Restante"
            logo = GetSetting("BasculasReg", "Properties", "Logo", "")
            If logo = "" Then Exit Sub
            LogoPictureBox.Image = Image.FromFile(logo)

        Else
            End
        End If
    End Sub
End Class
