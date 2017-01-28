Public Class Form1
    Dim PI As Double = Math.PI
    'Metodo para desactivar los paneles, cuando otro se active
    Private Sub desactivarPaneles()
        Panel1.Visible = False
        Panel2.Visible = False
        Panel3.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        Panel6.Visible = False
        Panel7.Visible = False
        Panel8.Visible = False
        Panel9.Visible = False
        Panel10.Visible = False
        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Variables
        Dim Base As Double
        Dim Altura As Double
        Dim Resultado As Double
        Try
            'Guardando la información del textbox en las variables, calculando y enseñando por pantalla
            Base = CDbl(TextBox1.Text)
            Altura = CDbl(TextBox2.Text)
            Resultado = Base * Altura / 2
            MsgBox("El area del triangulo es de : " + CStr(Resultado))
        Catch ex As Exception
            'Si no introduce un número le enseñamos un mensaje por pantalla
            MsgBox("Introduce números en los textboxs")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Para resetear los valores
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Cerrar el panel actual
        Panel1.Visible = False
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'Variables
        Dim Radio As Double
        Dim Resultado As Double
        Try
            'Calculando el area y enseñandolo por pantalla
            Radio = CDbl(TextBox4.Text)
            Resultado = PI * (Radio * Radio)
            MsgBox("El area del rectangulo es de : " + CStr(Resultado))
        Catch ex As Exception
            'Si no introduce un valor numerico, error
            MsgBox("Introduce un número en el textbox")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Resetear el valor del textbox
        TextBox4.Text = ""
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Desactivando panel actual
        Panel2.Visible = False
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TreeView1.ExpandAll()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        'Variables
        Dim Radio As Double
        Dim Resultado As Double
        Try
            'Calculando el area
            Radio = CDbl(TextBox5.Text)
            Resultado = 2 * PI * Radio
            MsgBox("La longitud de la circunferencia es de : " + CStr(Resultado))
        Catch ex As Exception
            'Si no introduce un valor númerico, error
            MsgBox("Introduce un número en el textbox")
        End Try

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'Reseteando el valor del textbox
        TextBox5.Text = ""
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        'Desactivando el panel actual
        Panel3.Visible = False
    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        'Dependiendo a que le di click en el treeview enseña un panel o otro
        'DockStyle.fill es para que ocupe el espacio restante
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
            Case 8
                desactivarPaneles()
                Panel8.Visible = True
                Panel8.Dock = DockStyle.Fill
            Case 9
                desactivarPaneles()
                Panel9.Visible = True
                Panel9.Dock = DockStyle.Fill
            Case 10
                desactivarPaneles()
                Panel10.Visible = True
                Panel10.Dock = DockStyle.Fill
            Case 11
                desactivarPaneles()
                Panel11.Visible = True
                Panel11.Dock = DockStyle.Fill
            Case 12
                desactivarPaneles()
                Panel12.Visible = True
                Panel12.Dock = DockStyle.Fill
            Case 13
                desactivarPaneles()
                Panel13.Visible = True
                Panel13.Dock = DockStyle.Fill
            Case 14
                desactivarPaneles()
                Panel14.Visible = True
                Panel14.Dock = DockStyle.Fill
        End Select
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'Variables
        Dim Base As Double
        Dim Altura As Double
        Dim Resultado As Double

        Try
            'Calculando el area y enseñandolo por pantalla
            Base = CDbl(TextBox6.Text)
            Altura = CDbl(TextBox3.Text)
            Resultado = Base * Altura
            MsgBox("El area del rectangle es de : " + CStr(Resultado))
        Catch ex As Exception
            'Si no introduce un número, error
            MsgBox("Introduce un número en los textboxs")
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'Reseteando valores
        TextBox6.Text = ""
        TextBox3.Text = ""
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'Desactivando panel actual
        Panel4.Visible = False
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'Variables
        Dim valor1 As Double
        Dim valor2 As Double
        Dim resultado As Double
        Try
            'Calculando y enseñando el resultado por pantalla
            valor1 = CDbl(TextBox8.Text)
            valor2 = CDbl(TextBox7.Text)
            resultado = valor1 + valor2
            MsgBox("El resultado es de : " + CStr(resultado))
        Catch ex As Exception
            'Si no introduce un número, error
            MsgBox("Introduce números en el textbox")
        End Try
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
        Try
            valor1 = CDbl(TextBox10.Text)
            valor2 = CDbl(TextBox9.Text)
            resultado = valor1 - valor2
            MsgBox("El resultado es de : " + CStr(resultado))
        Catch ex As Exception
            MsgBox("Introduce números en el textbox")
        End Try

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
        Try
            valor1 = CDbl(TextBox12.Text)
            valor2 = CDbl(TextBox11.Text)
            If (valor2 <> 0) Then
                resultado = valor1 / valor2
                MsgBox("El resultado de la división es de: " + CStr(resultado))
            Else
                MsgBox("Introduce un número diferente a 0 en el Nº2")
            End If
        Catch ex As Exception
            MsgBox("Introduce un número en el textbox")
        End Try

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        TextBox12.Text = ""
        TextBox11.Text = ""
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Panel7.Visible = False
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Dim valor1 As Double
        Dim valor2 As Double
        Dim resultado As Double
        Try
            valor1 = CDbl(TextBox14.Text)
            valor2 = CDbl(TextBox13.Text)
            resultado = valor1 * valor2
            MsgBox("El resultado de la multiplicación es de: " + CStr(resultado))
        Catch ex As Exception
            MsgBox("Introduce un número en el textbox")
        End Try
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        TextBox14.Text = ""
        TextBox13.Text = ""
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Panel8.Visible = False
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        TextBox15.Text = StrReverse(TextBox16.Text)
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        TextBox15.Text = ""
        TextBox16.Text = ""
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Panel9.Visible = False
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        Dim frase As String
        Dim totalLetras As Integer = 0
        Dim totalVocales As Integer = 0
        Dim totalConsonantes As Integer = 0
        Dim totalEspacios As Integer = 0

        frase = TextBox18.Text
        totalLetras = frase.Length

        For c = 0 To totalLetras - 1
            Select Case frase.ElementAt(c)
                Case "a"
                    totalVocales = totalVocales + 1
                Case "e"
                    totalVocales = totalVocales + 1
                Case "i"
                    totalVocales = totalVocales + 1
                Case "o"
                    totalVocales = totalVocales + 1
                Case "u"
                    totalVocales = totalVocales + 1
                Case " "
                    totalEspacios = totalEspacios + 1
            End Select
        Next

        totalConsonantes = totalLetras - totalVocales - totalEspacios
        TextBox17.Text = totalVocales
        TextBox19.Text = totalConsonantes

    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        TextBox18.Text = ""
        TextBox17.Text = ""
        TextBox19.Text = ""
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        Panel10.Visible = False
    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        Dim caracter1 As Char
        Dim caracter2 As Char
        Dim frase As String
        Dim contador As Integer = 0
        Dim totalRepeticiones As Integer = 0

        caracter1 = CChar(TextBox21.Text)
        caracter2 = CChar(TextBox20.Text)
        frase = TextBox22.Text

        While contador < frase.Length - 1
            If caracter1 = frase.ElementAt(contador) Then
                contador = contador + 1
                If caracter2 = frase.ElementAt(contador) Then
                    totalRepeticiones = totalRepeticiones + 1
                End If
            End If
            contador = contador + 1
        End While
        TextBox23.Text = totalRepeticiones
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        TextBox21.Text = ""
        TextBox20.Text = ""
        TextBox23.Text = ""
        TextBox22.Text = ""
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        Panel11.Visible = False
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        Dim numeroEntero As String

        Try
            numeroEntero = CInt(TextBox25.Text)
            TextBox24.Text = convertirNumeroLetra(numeroEntero)

        Catch
            MsgBox("Introduce un número entero.")
        End Try
    End Sub

    Public Function convertirNumeroLetra(numero As String) As String
        Dim b, paso As Integer
        Dim expresion, entero, deci, flag As String

        flag = "N"
        For paso = 1 To Len(numero)
            If Mid(numero, paso, 1) = "." Then
                flag = "S"
            Else
                If flag = "N" Then
                    entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero
                Else
                    deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero
                End If
            End If
        Next paso

        If Len(deci) = 1 Then
            deci = deci & "0"
        End If

        flag = "N"
        If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
            For paso = Len(entero) To 1 Step -1
                b = Len(entero) - (paso - 1)
                Select Case paso
                    Case 3, 6, 9
                        Select Case Mid(entero, b, 1)
                            Case "1"
                                If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then
                                    expresion = expresion & "cien "
                                Else
                                    expresion = expresion & "ciento "
                                End If
                            Case "2"
                                expresion = expresion & "doscientos "
                            Case "3"
                                expresion = expresion & "trescientos "
                            Case "4"
                                expresion = expresion & "cuatrocientos "
                            Case "5"
                                expresion = expresion & "quinientos "
                            Case "6"
                                expresion = expresion & "seiscientos "
                            Case "7"
                                expresion = expresion & "setecientos "
                            Case "8"
                                expresion = expresion & "ochocientos "
                            Case "9"
                                expresion = expresion & "novecientos "
                        End Select

                    Case 2, 5, 8
                        Select Case Mid(entero, b, 1)
                            Case "1"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    flag = "S"
                                    expresion = expresion & "diez "
                                End If
                                If Mid(entero, b + 1, 1) = "1" Then
                                    flag = "S"
                                    expresion = expresion & "once "
                                End If
                                If Mid(entero, b + 1, 1) = "2" Then
                                    flag = "S"
                                    expresion = expresion & "doce "
                                End If
                                If Mid(entero, b + 1, 1) = "3" Then
                                    flag = "S"
                                    expresion = expresion & "trece "
                                End If
                                If Mid(entero, b + 1, 1) = "4" Then
                                    flag = "S"
                                    expresion = expresion & "catorce "
                                End If
                                If Mid(entero, b + 1, 1) = "5" Then
                                    flag = "S"
                                    expresion = expresion & "quince "
                                End If
                                If Mid(entero, b + 1, 1) > "5" Then
                                    flag = "N"
                                    expresion = expresion & "dieci"
                                End If

                            Case "2"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "veinte "
                                    flag = "S"
                                Else
                                    expresion = expresion & "veinti"
                                    flag = "N"
                                End If

                            Case "3"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "treinta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "treinta y "
                                    flag = "N"
                                End If

                            Case "4"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "cuarenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "cuarenta y "
                                    flag = "N"
                                End If

                            Case "5"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "cincuenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "cincuenta y "
                                    flag = "N"
                                End If

                            Case "6"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "sesenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "sesenta y "
                                    flag = "N"
                                End If

                            Case "7"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "setenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "setenta y "
                                    flag = "N"
                                End If

                            Case "8"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "ochenta "
                                    flag = "S"
                                Else
                                    expresion = expresion & "ochenta y "
                                    flag = "N"
                                End If

                            Case "9"
                                If Mid(entero, b + 1, 1) = "0" Then
                                    expresion = expresion & "noventa "
                                    flag = "S"
                                Else
                                    expresion = expresion & "noventa y "
                                    flag = "N"
                                End If
                        End Select

                    Case 1, 4, 7
                        Select Case Mid(entero, b, 1)
                            Case "1"
                                If flag = "N" Then
                                    If paso = 1 Then
                                        expresion = expresion & "uno "
                                    Else
                                        expresion = expresion & "un "
                                    End If
                                End If
                            Case "2"
                                If flag = "N" Then
                                    expresion = expresion & "dos "
                                End If
                            Case "3"
                                If flag = "N" Then
                                    expresion = expresion & "tres "
                                End If
                            Case "4"
                                If flag = "N" Then
                                    expresion = expresion & "cuatro "
                                End If
                            Case "5"
                                If flag = "N" Then
                                    expresion = expresion & "cinco "
                                End If
                            Case "6"
                                If flag = "N" Then
                                    expresion = expresion & "seis "
                                End If
                            Case "7"
                                If flag = "N" Then
                                    expresion = expresion & "siete "
                                End If
                            Case "8"
                                If flag = "N" Then
                                    expresion = expresion & "ocho "
                                End If
                            Case "9"
                                If flag = "N" Then
                                    expresion = expresion & "nueve "
                                End If
                        End Select
                End Select
                If paso = 4 Then
                    If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And
                   Len(entero) <= 6) Then
                        If expresion = "un " Then
                            expresion = "mil "
                        Else
                            expresion = expresion & "mil "
                        End If
                    End If
                End If
                If paso = 7 Then
                    If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                        expresion = expresion & "millón "
                    Else
                        expresion = expresion & "millones "
                    End If
                End If
            Next paso

            If deci <> "" Then
                If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                    convertirNumeroLetra = "menos " & expresion & "con " & deci ' & "/100"
                Else
                    convertirNumeroLetra = expresion & "con " & deci ' & "/100"
                End If
            Else
                If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                    convertirNumeroLetra = "menos " & expresion
                Else
                    convertirNumeroLetra = expresion
                End If
            End If
        Else 'si el numero a convertir esta fuera del rango superior e inferior
            convertirNumeroLetra = ""
        End If
    End Function

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        TextBox25.Text = ""
        TextBox24.Text = ""
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        Panel12.Visible = False
    End Sub

    Private Sub Button39_Click(sender As Object, e As EventArgs) Handles Button39.Click
        ListBox1.Items.Clear()
        Dim numero As Integer
        Dim actual As Integer = 1
        Dim anterior As Integer = 1
        Dim suma As Integer
        Try
            numero = CUInt(TextBox27.Text)
            ListBox1.Items.Add(1)
            ListBox1.Items.Add(1)
            For c = 2 To numero - 1
                suma = anterior + actual
                anterior = actual
                actual = suma
                ListBox1.Items.Add(suma)
            Next
        Catch
            MsgBox("Introduce un número entero")
        End Try



    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        TextBox27.Text = ""
        ListBox1.Items.Clear()
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        Panel13.Visible = False
    End Sub

    Private Sub Button42_Click(sender As Object, e As EventArgs) Handles Button42.Click
        Dim numero As Integer
        Dim numeroMultiplicar As Integer
        Dim multiplicar As Integer
        Try
            numero = TextBox26.Text
            numeroMultiplicar = TextBox28.Text
            ListBox2.Items.Clear()

            For c = 1 To numeroMultiplicar
                multiplicar = numero * c
                ListBox2.Items.Add(multiplicar)
            Next
        Catch ex As Exception
            MsgBox("Introduce un número entero")
        End Try

    End Sub

    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        TextBox26.Text = ""
        TextBox28.Text = ""
        ListBox2.Items.Clear()
    End Sub

    Private Sub Button40_Click(sender As Object, e As EventArgs) Handles Button40.Click
        Panel14.Visible = False
    End Sub

    Private Sub InicializarA0ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InicializarA0ToolStripMenuItem.Click
        CType(ContextMenuStrip1.SourceControl, TextBox).Text = 0

    End Sub

    Private Sub IncrementarEn1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IncrementarEn1ToolStripMenuItem.Click
        Try
            CType(ContextMenuStrip1.SourceControl, TextBox).Text = CInt(CType(ContextMenuStrip1.SourceControl, TextBox).Text) + 1
        Catch ex As Exception
            MsgBox("Introduce un valor antes de sumarle 1")
        End Try

    End Sub

    Private Sub DisminuirEn1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisminuirEn1ToolStripMenuItem.Click
        Try
            CType(ContextMenuStrip1.SourceControl, TextBox).Text = CInt(CType(ContextMenuStrip1.SourceControl, TextBox).Text) - 1
        Catch ex As Exception
            MsgBox("Introduce un valor antes de restarle 1")
        End Try
    End Sub

    Private Sub CopiarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopiarToolStripMenuItem.Click
        Dim info As String
        Try
            info = CType(ContextMenuStrip1.SourceControl, TextBox).SelectedText
            My.Computer.Clipboard.SetText(info)
        Catch ex As Exception
            MsgBox("No has copiado nada")
        End Try

    End Sub

    Private Sub CortarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CortarToolStripMenuItem.Click
        Dim info As String
        Try
            info = CType(ContextMenuStrip1.SourceControl, TextBox).SelectedText
            My.Computer.Clipboard.SetText(info)
            CType(ContextMenuStrip1.SourceControl, TextBox).SelectedText = ""
        Catch ex As Exception
            MsgBox("No has cortado nada")
        End Try

    End Sub

    Private Sub PegarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PegarToolStripMenuItem.Click
        Dim info As String
        Try
            info = My.Computer.Clipboard.GetText
            CType(ContextMenuStrip1.SourceControl, TextBox).Text = info
        Catch ex As Exception
            MsgBox("No has pegado nada")
        End Try

    End Sub
End Class
