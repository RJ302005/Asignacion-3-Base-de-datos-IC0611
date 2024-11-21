<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BibliotecaForm
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BibliotecaForm))
        btnEjecutar = New Button()
        lstLibros = New ListBox()
        TXTTitulo = New TextBox()
        TXTAUTOR = New TextBox()
        TXTEditorial = New TextBox()
        cmbOperaciones = New ComboBox()
        Titulo = New Label()
        Autor = New Label()
        Editorial = New Label()
        LbrsDisp = New Label()
        Operaciones = New Label()
        btnActualizar = New Button()
        SuspendLayout()
        ' 
        ' btnEjecutar
        ' 
        btnEjecutar.Location = New Point(62, 181)
        btnEjecutar.Name = "btnEjecutar"
        btnEjecutar.Size = New Size(75, 23)
        btnEjecutar.TabIndex = 0
        btnEjecutar.Text = "Ejecutar"
        btnEjecutar.UseVisualStyleBackColor = True
        ' 
        ' lstLibros
        ' 
        lstLibros.FormattingEnabled = True
        lstLibros.ItemHeight = 15
        lstLibros.Location = New Point(291, 35)
        lstLibros.Name = "lstLibros"
        lstLibros.Size = New Size(308, 199)
        lstLibros.TabIndex = 1
        ' 
        ' TXTTitulo
        ' 
        TXTTitulo.Location = New Point(55, 40)
        TXTTitulo.Name = "TXTTitulo"
        TXTTitulo.Size = New Size(230, 23)
        TXTTitulo.TabIndex = 2
        ' 
        ' TXTAUTOR
        ' 
        TXTAUTOR.Location = New Point(55, 71)
        TXTAUTOR.Name = "TXTAUTOR"
        TXTAUTOR.Size = New Size(230, 23)
        TXTAUTOR.TabIndex = 3
        ' 
        ' TXTEditorial
        ' 
        TXTEditorial.Location = New Point(68, 100)
        TXTEditorial.Name = "TXTEditorial"
        TXTEditorial.Size = New Size(217, 23)
        TXTEditorial.TabIndex = 4
        ' 
        ' cmbOperaciones
        ' 
        cmbOperaciones.FormattingEnabled = True
        cmbOperaciones.Location = New Point(143, 140)
        cmbOperaciones.Name = "cmbOperaciones"
        cmbOperaciones.Size = New Size(112, 23)
        cmbOperaciones.TabIndex = 5
        ' 
        ' Titulo
        ' 
        Titulo.AutoSize = True
        Titulo.Location = New Point(12, 43)
        Titulo.Name = "Titulo"
        Titulo.Size = New Size(37, 15)
        Titulo.TabIndex = 6
        Titulo.Text = "Titulo"
        ' 
        ' Autor
        ' 
        Autor.AutoSize = True
        Autor.Location = New Point(12, 71)
        Autor.Name = "Autor"
        Autor.Size = New Size(37, 15)
        Autor.TabIndex = 7
        Autor.Text = "Autor"
        ' 
        ' Editorial
        ' 
        Editorial.AutoSize = True
        Editorial.Location = New Point(12, 103)
        Editorial.Name = "Editorial"
        Editorial.Size = New Size(50, 15)
        Editorial.TabIndex = 8
        Editorial.Text = "Editorial"
        ' 
        ' LbrsDisp
        ' 
        LbrsDisp.AutoSize = True
        LbrsDisp.Location = New Point(337, 17)
        LbrsDisp.Name = "LbrsDisp"
        LbrsDisp.Size = New Size(105, 15)
        LbrsDisp.TabIndex = 9
        LbrsDisp.Text = "Libros disponibles:"
        ' 
        ' Operaciones
        ' 
        Operaciones.AutoSize = True
        Operaciones.Location = New Point(61, 143)
        Operaciones.Name = "Operaciones"
        Operaciones.Size = New Size(76, 15)
        Operaciones.TabIndex = 10
        Operaciones.Text = "Operaciones:"
        Operaciones.TextAlign = ContentAlignment.TopCenter
        ' 
        ' btnActualizar
        ' 
        btnActualizar.Location = New Point(143, 181)
        btnActualizar.Name = "btnActualizar"
        btnActualizar.Size = New Size(75, 23)
        btnActualizar.TabIndex = 11
        btnActualizar.Text = "Actualizar"
        btnActualizar.UseVisualStyleBackColor = True
        ' 
        ' BibliotecaForm
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = SystemColors.Control
        BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), Image)
        ClientSize = New Size(607, 271)
        Controls.Add(btnActualizar)
        Controls.Add(Operaciones)
        Controls.Add(LbrsDisp)
        Controls.Add(Editorial)
        Controls.Add(Autor)
        Controls.Add(Titulo)
        Controls.Add(cmbOperaciones)
        Controls.Add(TXTEditorial)
        Controls.Add(TXTAUTOR)
        Controls.Add(TXTTitulo)
        Controls.Add(lstLibros)
        Controls.Add(btnEjecutar)
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        MaximizeBox = False
        Name = "BibliotecaForm"
        Text = "Formato de Biblioteca"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnEjecutar As Button
    Friend WithEvents lstLibros As ListBox
    Friend WithEvents TXTTitulo As TextBox
    Friend WithEvents TXTAUTOR As TextBox
    Friend WithEvents TXTEditorial As TextBox
    Friend WithEvents cmbOperaciones As ComboBox
    Friend WithEvents Titulo As Label
    Friend WithEvents Autor As Label
    Friend WithEvents Editorial As Label
    Friend WithEvents LbrsDisp As Label
    Friend WithEvents Operaciones As Label
    Friend WithEvents btnActualizar As Button
End Class
