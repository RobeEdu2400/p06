Public Class Form1
    Dim PI As Double = Math.PI


    Private Sub desactivarPaneles()
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Base As Double
        Dim Altura As Double
        Dim Resultado As Double

        Base = TextBox1.Text
        Altura = TextBox2.Text

        Resultado = Base * Altura / 2

        MsgBox("El area del triangulo es de : " + CStr(Resultado))
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Panel1.Visible = False
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim Radio As Double
        Dim Resultado As Double

        PI = Math.PI
        Radio = TextBox4.Text
        Resultado = PI * (Radio * Radio)
        MsgBox("El area del rectangulo es de : " + CStr(Resultado))

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox4.Text = ""
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Panel2.Visible = False
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TreeView1.ExpandAll()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim Radio As Double
        Dim Resultado As Double

        Radio = TextBox5.Text

        Resultado = 2 * PI * Radio

        MsgBox("La longitud de la circunferencia es de : " + CStr(Resultado))
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TextBox5.Text = ""
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Panel3.Visible = False
    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        Select Case e.Node.Tag
            Case 1
                desactivarPaneles()
                Panel1.Visible = True
                Panel1.Dock = DockStyle.Fill
            Case 2
                desactivarPaneles()
                Panel2.Visible = True
                Panel2.Dock = DockStyle.Fill
            Case 3
                desactivarPaneles()
                Panel3.Visible = True
                Panel3.Dock = DockStyle.Fill
            Case 4
                desactivarPaneles()
                Panel4.Visible = True
                Panel4.Dock = DockStyle.Fill
            Case 5
                desactivarPaneles()
                Panel5.Visible = True
                Panel5.Dock = DockStyle.Fill
            Case 6
                desactivarPaneles()
                Panel6.Visible = True
                Panel6.Dock = DockStyle.Fill
            Case 7
                desactivarPaneles()
                Panel7.Visible = True
                Panel7.Dock = DockStyle.Fill
        End Select
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim Base As Double
        Dim Altura As Double
        Dim Resultado As Double

        Base = TextBox6.Text
        Altura = TextBox3.Text

        Resultado = Base * Altura

        MsgBox("El area del rectangle es de : " + CStr(Resultado))
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox6.Text = ""
        TextBox3.Text = ""
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel4.Visible = False
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim valor1 As Double
        Dim valor2 As Double
        Dim resultado As Double

        valor1 = TextBox8.Text
        valor2 = TextBox7.Text
        resultado = valor1 + valor2

        MsgBox("El resultado es de : " + CStr(resultado))
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TextBox8.Text = ""
        TextBox7.Text = ""
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Panel5.Visible = False
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim valor1 As Double
        Dim valor2 As Double
        Dim resultado As Double

        valor1 = TextBox10.Text
        valor2 = TextBox9.Text
        resultado = valor1 - valor2

        MsgBox("El resultado es de : " + CStr(resultado))
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TextBox10.Text = ""
        TextBox9.Text = ""
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Panel6.Visible = False
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim valor1 As Double
        Dim valor2 As Double
        Dim resultado As Double

        valor1 = TextBox12.Text
        valor2 = TextBox11.Text
        If (valor2 <> 0) Then
            resultado = valor1 / valor2
            MsgBox("El resultado de la división es de: " + CStr(resultado))
        Else
            MsgBox("Introduce un número diferente a 0 en el Nº2")
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        TextBox12.Text = ""
        TextBox11.Text = ""
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Panel7.Visible = False
    End Sub


End Class
