Public Class Form3


    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OpenPreviewWindow()
        Dim iHeight As Integer = picCapture.Height
        Dim iWidth As Integer = picCapture.Width
        '
        ' Open Preview window in picturebox
        '
        hHwnd = capCreateCaptureWindowA(DeviceID.ToString, WS_VISIBLE Or WS_CHILD, 0, 0, 1280,
            1024, picCapture.Handle.ToInt32, 0)
        '
        ' Connect to device
        '
        If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, DeviceID, 0) Then
            '
            'Set the preview scale
            '
            SendMessage(hHwnd, WM_CAP_SET_SCALE, True, 0)

            '
            'Set the preview rate in milliseconds
            '
            SendMessage(hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0)

            '
            'Start previewing the image from the camera
            '
            SendMessage(hHwnd, WM_CAP_SET_PREVIEW, True, 0)

            '
            ' Resize window to fit in picturebox
            '
            SetWindowPos(hHwnd, HWND_BOTTOM, 0, 0, picCapture.Width, picCapture.Height,
                    SWP_NOMOVE Or SWP_NOZORDER)

            Button4.Enabled = True
            Button3.Enabled = False
        Else
            '
            ' Error connecting to device close window
            ' 
            DestroyWindow(hHwnd)
            Button4.Enabled = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        OpenPreviewWindow()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim data As IDataObject
        Dim bmap As Bitmap

        '
        ' Copy image to clipboard
        '
        SendMessage(hHwnd, WM_CAP_EDIT_COPY, 0, 0)
        '
        ' Get image from clipboard and convert it to a bitmap
        '
        data = Clipboard.GetDataObject()
        If data.GetDataPresent(GetType(System.Drawing.Bitmap)) Then
            bmap = CType(data.GetData(GetType(System.Drawing.Bitmap)), Image)
            picCapture.Image = bmap
            Button3.Enabled = False
            Button2.Enabled = True
            bmap.Save(Application.StartupPath & "\Imagennnn.bmp")
            ClosePreviewWindow()
        End If
    End Sub

    Private Sub ClosePreviewWindow()
        '
        ' Disconnect from device
        '
        SendMessage(hHwnd, WM_CAP_DRIVER_DISCONNECT, DeviceID, 0)
        '
        ' close window
        '
        DestroyWindow(hHwnd)
    End Sub

    Private Sub Form3_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Button4.Enabled Then
            ClosePreviewWindow()
        End If
    End Sub
End Class