Imports System.Data.OleDb

Public Class BibliotecaForm
    ' Conexión a la base de datos
    Private connection As OleDbConnection

    ' Evento que se ejecuta cuando el formulario carga
    Private Sub BibliotecaForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Configura la conexión a la base de datos (Biblioteca.accdb)
        Dim dbPath As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Biblioteca.accdb")
        Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}"
        connection = New OleDbConnection(connectionString)

        ' Llenar ComboBox con opciones
        cmbOperaciones.Items.AddRange(New String() {"Agregar Libro", "Modificar Libro", "Eliminar Libro", "Prestar Libro", "Devolver Libro"})
        cmbOperaciones.SelectedIndex = 0 ' Selecciona la primera opción por defecto

        ' Cargar libros al inicio
        ActualizarListaLibros()
    End Sub

    ' Evento del botón "Ejecutar"
    Private Sub btnEjecutar_Click(sender As Object, e As EventArgs) Handles btnEjecutar.Click
        Dim operacionSeleccionada As String = cmbOperaciones.SelectedItem.ToString()

        Select Case operacionSeleccionada
            Case "Agregar Libro" : AgregarLibro()
            Case "Modificar Libro" : ModificarLibro()
            Case "Eliminar Libro" : EliminarLibro()
            Case "Prestar Libro" : PrestarLibro()
            Case "Devolver Libro" : DevolverLibro()
            Case Else : MessageBox.Show("Seleccione una operación válida")
        End Select
    End Sub

    ' Métodos CRUD
    Private Sub AgregarLibro()
        Dim query As String = "INSERT INTO Libros (Titulo, Autor, Editorial, Disponible) VALUES (?, ?, ?, ?)"
        Dim command As New OleDbCommand(query, connection)

        command.Parameters.AddWithValue("?", TXTTitulo.Text)
        command.Parameters.AddWithValue("?", TXTAUTOR.Text)
        command.Parameters.AddWithValue("?", TXTEditorial.Text)
        command.Parameters.AddWithValue("?", True) ' Libro nuevo disponible por defecto

        EjecutarComando(command, "Libro agregado exitosamente.")
        ActualizarListaLibros()
    End Sub

    Private Sub ModificarLibro()
        If lstLibros.SelectedIndex >= 0 Then
            Dim libroID As Integer = ObtenerIDLibroSeleccionado()
            Dim query As String = "UPDATE Libros SET Titulo = ?, Autor = ?, Editorial = ? WHERE ID = ?"
            Dim command As New OleDbCommand(query, connection)

            command.Parameters.AddWithValue("?", TXTTitulo.Text)
            command.Parameters.AddWithValue("?", TXTAUTOR.Text)
            command.Parameters.AddWithValue("?", TXTEditorial.Text)
            command.Parameters.AddWithValue("?", libroID)

            EjecutarComando(command, "Libro modificado exitosamente.")
            ActualizarListaLibros()
        Else
            MessageBox.Show("Seleccione un libro primero.")
        End If
    End Sub

    Private Sub EliminarLibro()
        If lstLibros.SelectedIndex >= 0 Then
            Dim libroID As Integer = ObtenerIDLibroSeleccionado()
            Dim query As String = "DELETE FROM Libros WHERE ID = ?"
            Dim command As New OleDbCommand(query, connection)

            command.Parameters.AddWithValue("?", libroID)

            EjecutarComando(command, "Libro eliminado exitosamente.")
            ActualizarListaLibros()
        Else
            MessageBox.Show("Seleccione un libro primero.")
        End If
    End Sub

    Private Sub PrestarLibro()
        If lstLibros.SelectedIndex >= 0 Then
            Dim libroID As Integer = ObtenerIDLibroSeleccionado()
            Dim query As String = "UPDATE Libros SET Disponible = ? WHERE ID = ?"
            Dim command As New OleDbCommand(query, connection)

            command.Parameters.AddWithValue("?", False) ' Cambiar a "prestado"
            command.Parameters.AddWithValue("?", libroID)

            EjecutarComando(command, "Libro prestado exitosamente.")
            ActualizarListaLibros()
        Else
            MessageBox.Show("Seleccione un libro primero.")
        End If
    End Sub

    Private Sub DevolverLibro()
        If lstLibros.SelectedIndex >= 0 Then
            Dim libroID As Integer = ObtenerIDLibroSeleccionado()
            Dim query As String = "UPDATE Libros SET Disponible = ? WHERE ID = ?"
            Dim command As New OleDbCommand(query, connection)

            command.Parameters.AddWithValue("?", True) ' Cambiar a "disponible"
            command.Parameters.AddWithValue("?", libroID)

            EjecutarComando(command, "Libro devuelto exitosamente.")
            ActualizarListaLibros()
        Else
            MessageBox.Show("Seleccione un libro primero.")
        End If
    End Sub

    ' Actualizar lista de libros en el ListBox
    Private Sub ActualizarListaLibros()
        Dim query As String = "SELECT * FROM Libros"
        Dim adapter As New OleDbDataAdapter(query, connection)
        Dim table As New DataTable()

        Try
            connection.Open()
            adapter.Fill(table)
            lstLibros.Items.Clear()

            For Each row As DataRow In table.Rows
                Dim disponible As String = If(CBool(row("Disponible")), "Disponible", "Prestado")
                lstLibros.Items.Add($"{row("ID")}: {row("Titulo")} - {disponible}")
            Next
        Catch ex As Exception
            MessageBox.Show($"Error al cargar libros: {ex.Message}")
        Finally
            connection.Close()
        End Try
    End Sub

    ' Método para ejecutar comandos SQL
    Private Sub EjecutarComando(command As OleDbCommand, mensajeExito As String)
        Try
            connection.Open()
            command.ExecuteNonQuery()
            MessageBox.Show(mensajeExito)
        Catch ex As Exception
            MessageBox.Show($"Error: {ex.Message}")
        Finally
            connection.Close()
        End Try
    End Sub

    ' Obtener el ID del libro seleccionado
    Private Function ObtenerIDLibroSeleccionado() As Integer
        Dim libroSeleccionado As String = lstLibros.SelectedItem.ToString()
        Dim id As String = libroSeleccionado.Split(":")(0)
        Return Integer.Parse(id)
    End Function

    ' Botón para actualizar la lista de libros manualmente
    Private Sub btnActualizar_Click(sender As Object, e As EventArgs) Handles btnActualizar.Click
        ActualizarListaLibros()
    End Sub
End Class
