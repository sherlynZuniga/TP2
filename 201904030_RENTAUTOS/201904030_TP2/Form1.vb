Public Class Form1
    Private Sub ProcesosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProcesosToolStripMenuItem.Click

        'asignando textbox a variables 
        nit(fila) = Val(TextBox1.Text)
        placa(fila) = TextBox2.Text
        dias(fila) = Val(TextBox3.Text)
        efectivo(fila) = Val(TextBox4.Text)
        tarjeta(fila) = Val(TextBox5.Text)

        'validando que la cantidad sea la correcta'

        If Val(TextBox4.Text) + Val(TextBox5.Text) = 100 Then
            pagoparcial(fila) = efectivo(fila) + tarjeta(fila)
            TextBox6.Text = pagoparcial(fila)
        Else
            MsgBox("NO INGRESÓ UNA CANTIDAD CORRECTA")
        End If

        'envio de valores
        If (fila <= 5) Then

            'valores en vectores
            marca(fila) = ComboBox1.SelectedItem
            pagoparcial(fila) = funcionpagoparcial(dias(fila), marca(fila))
            vdescuento(fila) = pagoparcial(fila) * funciondescuento(efectivo(fila))
            vrecargo(fila) = pagoparcial(fila) * funcionrecargo(efectivo(fila))
            pagototal(fila) = pagoparcial(fila) - vdescuento(fila) + vrecargo(fila)

            'mostrar valores en listboxs
            ListBox1.Items.Add(placa(fila))
            ListBox2.Items.Add(marca(fila))
            ListBox3.Items.Add(dias(fila))
            ListBox4.Items.Add(pagoparcial(fila))
            ListBox6.Items.Add(pagototal(fila))
        End If




    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

    End Sub

    Private Sub LimpiarEntradasToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles LimpiarEntradasToolStripMenuItem1.Click

        'limpiar'
        ComboBox1.SelectedIndex = -1
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox3.Items.Clear()
        ListBox4.Items.Clear()
        ListBox6.Items.Clear()


    End Sub

    Private Sub LimpiarVectoresToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles LimpiarVectoresToolStripMenuItem1.Click

        'limpiar vectores


    End Sub

    Private Sub SalirToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem1.Click

        'mostrar formulario2 de salida
        Form2.Show()

    End Sub
End Class
