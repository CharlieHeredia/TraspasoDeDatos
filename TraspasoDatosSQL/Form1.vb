Imports System.IO
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data.Sql
Public Class Form1
    Dim rutaGlobal As String = "" 'SIN USO <--Pruebas 15/12/2017
    Dim tablaGlobal As String = "" 'SIN USO <--Pruebas 15/12/2017
    Dim ds As DataSet = New DataSet()
    Dim dsprueba As DataSet = New DataSet() 'SIN USO <--Pruebas 15/12/2017'
    Dim datatab As DataTable = New DataTable() 'Se utiliza para cargar el esquema de la tabla dfb'
    Public Conexiones As New OleDb.OleDbConnection()
    Public ConexionesSQL As New SqlConnection()
    Dim adaptador As OleDb.OleDbDataAdapter
    Dim adaptadorSQL As SqlDataAdapter
    Dim cmd As SqlCommand
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Limpieza de DataGridView1 y DataSet'
        DataGridView1.DataSource = Nothing
        ds.Reset()
        datatab.Reset()
        '<-- Termina limpieza -->'
        Dim OpenFile As New OpenFileDialog
        OpenFile.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFile.Filter = "DBF files| *.dbf" 'Filtro de documentos .dfb'
        If (OpenFile.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFile.FileName 'Ruta del archivo'
            'Capturar nombre de archivo'
            Dim arreglo() As String = Split(FileName, "\") 'Separar la ruta'
            Dim tabla As String = arreglo(arreglo.Length - 1) 'Obtención de nombre de la tabla con extension'
            Dim ruta As String = ""
            'MsgBox("Ruta del arhivo = " & FileName)
            'MsgBox("Nombre de tabla = " & tabla)
            For i = 0 To arreglo.Length - 2   'Ciclo for para generar unicamente la ruta'
                ruta = ruta + arreglo(i) + "\"
            Next
            'MsgBox("Ruta = " & ruta)

            'Reutilización de arreglo'
            arreglo = Split(tabla, ".")
            tabla = arreglo(arreglo.Length - 2) 'Quitar la extersión de la tabla'
            'MsgBox("Tabla = " & tabla)
            Try
                Dim cadena As String = "Provider = VFPOLEDB.1; Data Source = " & ruta & "; Extended Properties = dBase IV;User Id=;Password="
                Dim con As OleDbConnection = New OleDbConnection()
                con.ConnectionString = cadena
                con.Open()

                Dim consulta As String = "Select * from " & tabla
                Dim adaptador As OleDbDataAdapter = New OleDbDataAdapter(consulta, con)

                adaptador.Fill(ds) 'Cargar la consulta en el DataSet'
                adaptador.FillSchema(datatab, SchemaType.Source) 'Carga del esquema de la tabla dbf'
                con.Close()

                DataGridView1.DataSource = ds.Tables(0) 'Copia de datos al DataGridView'
                'rutaGlobal = ruta 'Copia de la ruta, se utiliza SIN USAR <--Pruebas
                'tablaGlobal = tabla 'Copia del nombre de la tabla dfb  SIN USAR --Pruebas
            Catch ex As Exception
                MsgBox("Problema encontrado: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '<---------------- Datos Para Conexión a SQL Servier --------------------------->'
        Dim ventana As LoginSQL = New LoginSQL
        ventana.ShowDialog()
        '<----------------- Termina recolección de datos ------------------------------->'
        '<----------------- Buscar archivo Log ----------------------------------------->'
        '<----------------- Termina busqueda de archivo Log ---------------------------->'
        If DatosCompletos = True Then 'Validación de datos ingresados para realizar la conexción.'
            Try
                Dim temporal As String = ""
                For Each table In ds.Tables
                    For Each column In table.Columns
                        temporal = temporal + column.ColumnName + "\"
                    Next
                Next
                Dim NombreCampos() = Split(temporal, "\")  '<---------------- OBTENCIÓN DE NOMBRES DE CAMPOS  ------------>
                temporal = ""
                '---------------------------'
                Dim TipoCadena(ds.Tables(0).Columns.Count) As String
                Dim TamCadena(ds.Tables(0).Columns.Count) As String
                'Tipo de dato de columna'
                For i = 0 To ds.Tables(0).Columns.Count - 1 'Ciclo For para recorrer todas las columnas, solo existe una tabla'
                    Dim col As DataColumn = ds.Tables(0).Columns(i) 'Selección de columna'
                    Dim tipo As Type = col.DataType 'Obtiene el tipo de dato'
                    Dim tam As String = datatab.Columns(i).MaxLength 'Obtiene el tamaño del datos tipo String'
                    'MsgBox("Tipo de dato columna: " & tipo.ToString() & " I: " & i + 1 & " Tamaño maximo: " & tam) <--Mensaje de prueba
                    TipoCadena(i) = tipo.ToString()
                    TamCadena(i) = tam
                    'MsgBox("tipoCadena: " & TipoCadena(i) & "TamCadena: " & TamCadena(i) & " Nombre Campo" & NombreCampos(i))
                Next
                ' -------------------------'
                'Lectura de filas'
                'MsgBox("Paso por aqui " & ds.Tables(0).Rows.Count)
                Dim valor As String = ""
                Dim Prime As Boolean = True 'Se utiliza para saber si es la primera vez que entra en el ciclo For each'
                'La primera vez crea la base de datos y la tabla'
                For Each row As DataRow In ds.Tables(0).Rows  'Ciclo For each para recorrer todas las filas de la tabla'
                    GC.Collect()
                    For i = 0 To ds.Tables(0).Columns.Count - 1 'Ciclo For para recorrer todas las columnas'
                        If IsDBNull(row(i)) Then 'Validación de Dato Null'
                            If i <> ds.Tables(0).Columns.Count - 1 Then 'Validar que no sea la ultima columna'
                                valor = valor + "¬"
                            End If
                        Else
                            If i <> ds.Tables(0).Columns.Count - 1 Then 'Validar que no sea la ultima columna'
                                valor = valor + row(i).ToString().Trim() + "¬"
                            Else
                                valor = valor + row(i).ToString().Trim()
                            End If

                        End If
                    Next
                    'MsgBox("Valor: " & valor)
                    Dim arregloCadena() As String = Split(valor, "¬") 'Copia de datos en un arreglo (separados)'
                    '<----------------- Traspaso de datos a SQL SERVER ---------------------------------->'
                    ConexionesSQL = New SqlConnection()
                    'Ciclo For para Crear la Base datos y tabla, despues inserta los datos'
                    Dim camposTabla As String = ""
                    If Prime = True Then
                        'Primera vex que entra en el ciclo'
                        'Se crea la base de datos (en caso de no existir), la tabla'
                        '<--------------------------- CREAR BASE DE DATOS ----------------------------------------->'
                        ConexionesSQL.ConnectionString = "Data Source=" & Instancia & ";User Id=" & UsuarioBD & ";Password=" & Contra
                        '<--- Ciclo For para crear la cadena de campos y tamaño -------->'

                        For i = 0 To ds.Tables(0).Columns.Count - 1
                            If TipoCadena(i) = "System.Int32" Then 'Campo de tipo INT'
                                If i <> ds.Tables(0).Columns.Count - 1 Then 'Validación que no sea el ultimo campo'
                                    camposTabla = camposTabla + NombreCampos(i) + " INT,"
                                Else
                                    camposTabla = camposTabla + NombreCampos(i) + " INT"
                                End If
                            ElseIf TipoCadena(i) = "System.Boolean" Then '<- Campo tipo Boolean'
                                If i <> ds.Tables(0).Columns.Count - 1 Then 'Validación que no sea el ultimo campo'
                                    camposTabla = camposTabla + NombreCampos(i) + " BIT,"
                                Else
                                    camposTabla = camposTabla + NombreCampos(i) + " BIT"
                                End If
                            Else 'Campo de tipo String'
                                If i <> ds.Tables(0).Columns.Count - 1 Then 'Validación que no sea el ultimo campo
                                    camposTabla = camposTabla + NombreCampos(i) + " VARCHAR(" & TamCadena(i) & "),"
                                Else
                                    camposTabla = camposTabla + NombreCampos(i) + " VARCHAR(" & TamCadena(i) & ")"
                                End If
                            End If
                        Next
                        'MsgBox("camposTabla: " & camposTabla)
                        ' <----- Fin de ciclo For ------>
                        ' <------------- VERIFICACIÓN DE EXISTENCIA DE LA BASE DE DATOS ----------------------->'
                        'Dim BaseExiste As Integer = 1 ' 1 indica que la base de datos existe '
                        Try
                            'Se crea la base de datos'
                            cmd = New SqlCommand("CREATE DATABASE " & BaseDatos, ConexionesSQL)
                            ConexionesSQL.Open()
                            cmd.ExecuteNonQuery()
                        Catch ex As Exception
                            'En caso de existir la base de datos entra en esta sección'
                        Finally
                            Prime = False
                            ConexionesSQL.Close()
                        End Try
                        '<------------------------ CREAR TABLA  ------------------------------------------------------->'
                        'ConexionesSQL = New SqlConnection()
                        Try
                            ConexionesSQL.ConnectionString = "Data Source=" & Instancia & ";initial Catalog=" & BaseDatos & ";User Id=" & UsuarioBD & ";Password=" & Contra
                            cmd = New SqlCommand("CREATE TABLE " & Tabla & " (" & camposTabla & ")", ConexionesSQL)
                            ConexionesSQL.Open()
                            cmd.ExecuteNonQuery()
                        Catch ex As Exception
                            'En caso de ocurrir un error entra en esta sección '
                            ' Cuando entra aqui, la tabla ya existe.'
                        Finally
                            ConexionesSQL.Close()
                        End Try
                        '<-------------------------------- INSERTAR PRIMERA FILA ------------------------------------------>'
                        ConexionesSQL.ConnectionString = "Data Source=" & Instancia & ";initial Catalog=" & BaseDatos & ";User Id=" & UsuarioBD & ";Password=" & Contra
                        '<----- Ciclo for para hacer la cadena de datos a insertar -------------------------->'
                        camposTabla = "" '<---- REUTILIZACIÓN DE LA VARIEABLE PARA CREAR LA CADENA ------>'
                        For i = 0 To ds.Tables(0).Columns.Count - 1
                            arregloCadena(i) = arregloCadena(i).Replace("&", "&amp;") 'para reemplazar caracteres
                            arregloCadena(i) = arregloCadena(i).Replace("""", "&quot;")
                            arregloCadena(i) = arregloCadena(i).Replace("<", "&lt;")
                            arregloCadena(i) = arregloCadena(i).Replace(">", "&gt;")

                            arregloCadena(i) = arregloCadena(i).Replace("'", "&#39;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(145), "&#39;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(146), "&#39;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(147), "&quot;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(148), "&quot;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(188), "1/4")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(189), "1/2")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(190), "3/4")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(132), "&quot;")
                            If TipoCadena(i) = "System.Boolean" Then
                                If arregloCadena(i) = "True" Then
                                    arregloCadena(i) = "True"
                                Else
                                    arregloCadena(i) = "False"
                                End If
                            End If
                            If i <> ds.Tables(0).Columns.Count - 1 Then
                                camposTabla = camposTabla + "'" & arregloCadena(i) & "',"
                            Else
                                camposTabla = camposTabla + "'" & arregloCadena(i) & "'"
                            End If
                        Next
                        'MsgBox("camposTabla: " & camposTabla)
                        cmd = New SqlCommand("INSERT INTO " & Tabla & " VALUES (" & camposTabla & ")", ConexionesSQL)
                        ConexionesSQL.Open()
                        cmd.ExecuteNonQuery()
                        ConexionesSQL.Close()
                    Else
                        '<------------- INSERTA DE LA SEGUNDA FILA EN ADELANTE ---------------------------->'
                        ConexionesSQL.ConnectionString = "Data Source=" & Instancia & ";initial Catalog=" & BaseDatos & ";User Id=" & UsuarioBD & ";Password=" & Contra
                        '<----- Ciclo for para hacer la cadena de datos a insertar -------------------------->'
                        camposTabla = "" '<---- REUTILIZACIÓN DE LA VARIEABLE PARA CREAR LA CADENA ------>'
                        For i = 0 To ds.Tables(0).Columns.Count - 1
                            arregloCadena(i) = arregloCadena(i).Replace("&", "&amp;") 'para reemplazar caracteres
                            arregloCadena(i) = arregloCadena(i).Replace("""", "&quot;")
                            arregloCadena(i) = arregloCadena(i).Replace("<", "&lt;")
                            arregloCadena(i) = arregloCadena(i).Replace(">", "&gt;")

                            arregloCadena(i) = arregloCadena(i).Replace("'", "&#39;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(145), "&#39;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(146), "&#39;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(147), "&quot;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(148), "&quot;")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(188), "1/4")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(189), "1/2")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(190), "3/4")
                            arregloCadena(i) = arregloCadena(i).Replace(Chr(132), "&quot;")
                            If i <> ds.Tables(0).Columns.Count - 1 Then
                                camposTabla = camposTabla + "'" & arregloCadena(i) & "',"
                            Else
                                camposTabla = camposTabla + "'" & arregloCadena(i) & "'"
                            End If
                        Next
                        'MsgBox("camposTabla: " & camposTabla)
                        cmd = New SqlCommand("INSERT INTO " & Tabla & " VALUES (" & camposTabla & ")", ConexionesSQL)
                        ConexionesSQL.Open()
                        cmd.ExecuteNonQuery()
                        ConexionesSQL.Close()
                    End If

                    '<---------------- Termina traspaso de datos a SQL SERVER --------------------------->'
                    '<------------------ Limpieza de variables y memoria -------------------------------->'
                    GC.GetTotalMemory(True) '<--- Limpieza de basura ------>'
                    valor = "" 'Limpieza de variable'
                Next
            Catch ex As Exception
                MsgBox("Ocurrio un problema: " & ex.ToString())
            End Try
            DatosCompletos = False 'Limpieza cuando quiera volver a cargar datos
        Else
            MsgBox("No se han cargado los datos a SQL Server.")
        End If


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        'SE UTILIZO PARA PRUEBAS <-- NO EXISTE EL BOTON 15/12/2017'
        
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub
End Class
