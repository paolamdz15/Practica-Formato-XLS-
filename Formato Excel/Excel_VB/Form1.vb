Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel


Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'NorthwindDataSet.Categories' Puede moverla o quitarla según sea necesario.
        Me.CategoriesTableAdapter.Fill(Me.NorthwindDataSet.Categories)
        'TODO: esta línea de código carga datos en la tabla 'NorthwindDataSet.Categories' Puede moverla o quitarla según sea necesario.
        Me.CategoriesTableAdapter.Fill(Me.NorthwindDataSet.Categories)
        'TODO: esta línea de código carga datos en la tabla 'NorthwindDataSet2.Sales_by_Category' Puede moverla o quitarla según sea necesario.


        'TODO: esta línea de código carga datos en la tabla 'NorthwindDataSet.Sales_by_Category' Puede moverla o quitarla según sea necesario.
        btnExportar.Enabled = False

    End Sub

    Private Sub btnBuscar_Click(sender As Object, e As EventArgs) Handles btnBuscar.Click
        Me.CategoriesTableAdapter.Fill(Me.NorthwindDataSet.Categories)
        'TODO: esta línea de código carga datos en la tabla 'NorthwindDataSet1.Sales_by_Category' Puede moverla o quitarla según sea necesario.
        Me.CategoriesTableAdapter.Fill(Me.NorthwindDataSet.Categories)
        'TODO: esta línea de código carga datos en la tabla 'NorthwindDataSet1.Sales_by_Category' Puede moverla o quitarla según sea necesario.
        Me.CategoriesTableAdapter.Fill(Me.NorthwindDataSet.Categories)
        'TODO: esta línea de código carga datos en la tabla 'NorthwindDataSet1.Sales_by_Category' Puede moverla o quitarla según sea necesario.
        Me.CategoriesTableAdapter.Fill(Me.NorthwindDataSet.Categories)
        Me.CategoriesTableAdapter.Fill(Me.NorthwindDataSet.Categories)
        btnExportar.Enabled = True
        PBExportar.Maximum = DGV.Rows.Count
    End Sub
    '  Select Case
    '    [CategoryName]
    '    ,[ProductName]
    '    ,sum([ProductSales])
    '    ,year([orderDate]), MONTH(OrderDate)
    'FROM [Northwind].[dbo].[Sales by Category]
    'group by year(OrderDate), MONTH(OrderDate),[CategoryName]
    '    ,[ProductName]
    Private Sub btnExportar_Click(sender As Object, e As EventArgs)

    End Sub
    Private Sub wait(ByVal seconds As Integer) 'suspende la ejecucion un tiempo especifico de segundos
        For i As Integer = 0 To seconds * 100 'realiza el conteo de segundos por 100 para hacerlos equivalentes a la cantidad de milisegundos a esperar
            System.Threading.Thread.Sleep(10) 'suspende la ejecucion 10 milisegundos
            Application.DoEvents() 'procesa lo que esta en espera
        Next
    End Sub
    'Crear formato en Excel

    'Crear aplicacion para llenar el formato en base a los datos de una base de datos en SQL Server

    'Aplicacion en VB y en c#

    '-Conservar formato, calcular sumatorias, nostrar numero de linea 
    'y abrir en solo lectura
End Class