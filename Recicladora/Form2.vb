Public Class Form2
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            Label1.Text = "Número Proveedor"
        ElseIf ComboBox1.SelectedIndex = 1 Then
            Label1.Text = "Nombre Proveedor"
        Else
            Label1.Text = "Número Boleta"
        End If
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
End Class