Public Class Libro
    ' Atributos (con Property)
    Public Property ID As Integer ' ID del libro (clave primaria en la base de datos)
    Public Property Titulo As String
    Public Property Autor As String
    Public Property Editorial As String
    Public Property Disponible As Boolean = True ' Disponible por defecto

    ' Constructor para nuevos libros
    Public Sub New(titulo As String, autor As String, editorial As String)
        Me.Titulo = titulo
        Me.Autor = autor
        Me.Editorial = editorial
    End Sub

    ' Constructor para cargar desde la base de datos
    Public Sub New(id As Integer, titulo As String, autor As String, editorial As String, disponible As Boolean)
        Me.ID = id
        Me.Titulo = titulo
        Me.Autor = autor
        Me.Editorial = editorial
        Me.Disponible = disponible
    End Sub
End Class

