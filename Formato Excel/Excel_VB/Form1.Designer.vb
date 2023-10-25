<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container()
        Me.PBExportar = New System.Windows.Forms.ProgressBar()
        Me.btnExportar = New System.Windows.Forms.Button()
        Me.btnBuscar = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DGV = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CategoryIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CategoryNameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DescriptionDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PictureDataGridViewImageColumn = New System.Windows.Forms.DataGridViewImageColumn()
        Me.CategoriesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.NorthwindDataSet = New Excel_VB.NorthwindDataSet()
        Me.CategoriesTableAdapter = New Excel_VB.NorthwindDataSetTableAdapters.CategoriesTableAdapter()
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CategoriesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NorthwindDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PBExportar
        '
        Me.PBExportar.Location = New System.Drawing.Point(61, 223)
        Me.PBExportar.Name = "PBExportar"
        Me.PBExportar.Size = New System.Drawing.Size(450, 23)
        Me.PBExportar.TabIndex = 23
        Me.PBExportar.Visible = False
        '
        'btnExportar
        '
        Me.btnExportar.Font = New System.Drawing.Font("Corbel", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExportar.Location = New System.Drawing.Point(411, 78)
        Me.btnExportar.Name = "btnExportar"
        Me.btnExportar.Size = New System.Drawing.Size(100, 35)
        Me.btnExportar.TabIndex = 22
        Me.btnExportar.Text = "Exportar"
        Me.btnExportar.UseVisualStyleBackColor = True
        '
        'btnBuscar
        '
        Me.btnBuscar.Font = New System.Drawing.Font("Corbel", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBuscar.Location = New System.Drawing.Point(411, 40)
        Me.btnBuscar.Name = "btnBuscar"
        Me.btnBuscar.Size = New System.Drawing.Size(100, 32)
        Me.btnBuscar.TabIndex = 21
        Me.btnBuscar.Text = "Buscar"
        Me.btnBuscar.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Lucida Bright", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.Label2.Location = New System.Drawing.Point(72, 86)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(333, 17)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Consulte y exporte las ventas por categoria"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Lucida Bright", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label1.Location = New System.Drawing.Point(103, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(240, 24)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Ventas por Categoria"
        '
        'DGV
        '
        Me.DGV.AllowUserToAddRows = False
        Me.DGV.AllowUserToDeleteRows = False
        Me.DGV.AutoGenerateColumns = False
        Me.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column4, Me.Column3, Me.Column2, Me.CategoryIDDataGridViewTextBoxColumn, Me.CategoryNameDataGridViewTextBoxColumn, Me.DescriptionDataGridViewTextBoxColumn, Me.PictureDataGridViewImageColumn})
        Me.DGV.DataSource = Me.CategoriesBindingSource
        Me.DGV.Location = New System.Drawing.Point(61, 124)
        Me.DGV.Name = "DGV"
        Me.DGV.ReadOnly = True
        Me.DGV.Size = New System.Drawing.Size(450, 339)
        Me.DGV.TabIndex = 18
        '
        'Column1
        '
        Me.Column1.DataPropertyName = "CateogryName"
        Me.Column1.HeaderText = "Categoria"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        '
        'Column4
        '
        Me.Column4.DataPropertyName = "Exp3"
        Me.Column4.HeaderText = "Mes"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        '
        'Column3
        '
        Me.Column3.DataPropertyName = "Exp2"
        Me.Column3.HeaderText = "Año"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        '
        'Column2
        '
        Me.Column2.DataPropertyName = "Exp1"
        Me.Column2.HeaderText = "Monto"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        '
        'CategoryIDDataGridViewTextBoxColumn
        '
        Me.CategoryIDDataGridViewTextBoxColumn.DataPropertyName = "CategoryID"
        Me.CategoryIDDataGridViewTextBoxColumn.HeaderText = "CategoryID"
        Me.CategoryIDDataGridViewTextBoxColumn.Name = "CategoryIDDataGridViewTextBoxColumn"
        Me.CategoryIDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'CategoryNameDataGridViewTextBoxColumn
        '
        Me.CategoryNameDataGridViewTextBoxColumn.DataPropertyName = "CategoryName"
        Me.CategoryNameDataGridViewTextBoxColumn.HeaderText = "CategoryName"
        Me.CategoryNameDataGridViewTextBoxColumn.Name = "CategoryNameDataGridViewTextBoxColumn"
        Me.CategoryNameDataGridViewTextBoxColumn.ReadOnly = True
        '
        'DescriptionDataGridViewTextBoxColumn
        '
        Me.DescriptionDataGridViewTextBoxColumn.DataPropertyName = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.HeaderText = "Description"
        Me.DescriptionDataGridViewTextBoxColumn.Name = "DescriptionDataGridViewTextBoxColumn"
        Me.DescriptionDataGridViewTextBoxColumn.ReadOnly = True
        '
        'PictureDataGridViewImageColumn
        '
        Me.PictureDataGridViewImageColumn.DataPropertyName = "Picture"
        Me.PictureDataGridViewImageColumn.HeaderText = "Picture"
        Me.PictureDataGridViewImageColumn.Name = "PictureDataGridViewImageColumn"
        Me.PictureDataGridViewImageColumn.ReadOnly = True
        '
        'CategoriesBindingSource
        '
        Me.CategoriesBindingSource.DataMember = "Categories"
        Me.CategoriesBindingSource.DataSource = Me.NorthwindDataSet
        '
        'NorthwindDataSet
        '
        Me.NorthwindDataSet.DataSetName = "NorthwindDataSet"
        Me.NorthwindDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CategoriesTableAdapter
        '
        Me.CategoriesTableAdapter.ClearBeforeFill = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(599, 494)
        Me.Controls.Add(Me.PBExportar)
        Me.Controls.Add(Me.btnExportar)
        Me.Controls.Add(Me.btnBuscar)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DGV)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CategoriesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NorthwindDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PBExportar As ProgressBar
    Friend WithEvents btnExportar As Button
    Friend WithEvents btnBuscar As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents DGV As DataGridView
    Friend WithEvents NorthwindDataSet As NorthwindDataSet
    Private WithEvents Column1 As DataGridViewTextBoxColumn
    Private WithEvents Column4 As DataGridViewTextBoxColumn
    Private WithEvents Column3 As DataGridViewTextBoxColumn
    Private WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents CategoriesBindingSource As BindingSource
    Friend WithEvents CategoriesTableAdapter As NorthwindDataSetTableAdapters.CategoriesTableAdapter
    Friend WithEvents CategoryIDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents CategoryNameDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents DescriptionDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents PictureDataGridViewImageColumn As DataGridViewImageColumn
End Class
