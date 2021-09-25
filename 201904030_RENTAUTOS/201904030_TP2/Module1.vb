Module Module1

    Public nit(5) As Integer
    Public placa(5) As String
    Public marca(5) As String
    Public dias(5) As Integer
    Public efectivo(5) As Double
    Public tarjeta(5) As Double
    Public pagoparcial(5) As Double
    Public pagototal(5) As Double
    Public vdescuento(5)
    Public vrecargo(5)

    Public fila As Byte = 0

    'variables de calculo

    Public calculopagoparcial As Double = 0
    Public calculodescuento As Double = 0
    Public calculorecargo As Double = 0
    Public calculopagototal As Double = 0

    'funcion para pago parcial'
    Function funcionpagoparcial(dias As Integer, marca As String) As Double
        Select Case Form1.ComboBox1.SelectedIndex
            Case 0
                calculopagoparcial = Val(Form1.TextBox3.Text) * 250
            Case 1
                calculopagoparcial = Val(Form1.TextBox3.Text) * 450
            Case 2
                calculopagoparcial = Val(Form1.TextBox3.Text) * 325
            Case 3
                calculopagoparcial = Val(Form1.TextBox3.Text) * 300
        End Select
        Return calculopagoparcial
    End Function

    'funcion para descuento 
    Function funciondescuento(efectivo As Double) As Double
        If Val(Form1.TextBox4.Text) = 100 Then
            funciondescuento = 5 / 100
        Else
            funciondescuento = 0
        End If
        Return funciondescuento
    End Function

    'funcion para recargo
    Function funcionrecargo(tarjeta As Double) As Double
        If Val(Form1.TextBox5.Text) = 100 Then
            funcionrecargo = 2.5 / 100
        Else
            funcionrecargo = 0
        End If
        Return funcionrecargo
    End Function




End Module
