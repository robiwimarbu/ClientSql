Imports System.Data
Imports System.Data.SqlClient
'Imports System.Windows.Forms
Public Class ClientSql
    Private myConn As SqlConnection
    Private myAdapter As SqlDataAdapter
    Private myCommand As SqlCommand
    Public catalog As String
    Public host As String
    Public segurity As String
	'funcion para abrir la conexion.
    Public Sub conect()
        Try
            catalog = "myBd"
            host = "servidor"
            segurity = "Integrated Security=SSPI;"
            myConn = New SqlConnection("Initial Catalog=" & catalog & ";" & _
                "Data Source=" & host & ";" & segurity)
            myConn.Open()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub
	'funcion que cierra la conexion.
    Public Sub disconect()
        Try
            myConn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub
	'Executa una consulta de seleccion retornando los datos en una datatable.
    Public Function runSelect(ByVal table As String, ByVal columns As String, ByVal clause As String) As DataTable
        Dim result As New DataTable
        Try
            Dim stringQuery As String
            stringQuery = "SELECT " & columns & " FROM " & table & " WHERE " & clause
            'MessageBox.Show(stringQuery)
            myAdapter = New SqlDataAdapter(stringQuery, myConn)
            myAdapter.Fill(result)
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
        Return result
    End Function
	'Executa una consulta de actualizacion regresando un entero con la cantidad de filas afectadas.
    Public Function runUpdate(ByVal table As String, ByVal stringSet As String, ByVal clause As String) As Integer
        Dim result As Integer
        Dim stringQuery As String
        stringQuery = "UPDATE " & table & " SET " & stringSet & " WHERE " & clause
        'MessageBox.Show(stringQuery)
        Try
            myCommand = New SqlCommand(stringQuery, myConn)
            result = myCommand.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
        Return result
    End Function
	'Executa una consulta de insercion regresando un entero con la cantidad de filas insertadas.
    Public Function runInsert(ByVal table As String, ByVal columns As String, ByVal values As String) As Integer
        Dim result As Integer
        Dim stringQuery As String
        stringQuery = "INSERT INTO " & table & " (" & columns & ")  VALUES (" & values & ")"
        Try
            myCommand = New SqlCommand(stringQuery, myConn)
            result = myCommand.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
        Return result
    End Function
	'Executa una consulta de insercion regresando un entero con el numero del id del nuevo elemento.
    Public Function runInsertValId(ByVal table As String, ByVal columns As String, ByVal values As String) As Integer
        Dim result As Integer = 0
        Dim stringQuery As String
        stringQuery = "INSERT INTO " & table & " (" & columns & ")  VALUES (" & values & ")  SET @id = SCOPE_IDENTITY()"
        'MessageBox.Show(stringQuery)
        Try
            myCommand = New SqlCommand(stringQuery, myConn)
            myCommand.Parameters.Add("@id", SqlDbType.Int).Direction = ParameterDirection.Output
            myCommand.ExecuteNonQuery()
            result = myCommand.Parameters("@id").Value.ToString()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
        Return result
    End Function
	'Executa una consulta de eliminacion regresando un entero con la cantidad de filas eliminadas.
    Public Function runDelete(ByVal table As String, ByVal clause As String) As Integer
        Dim result As Integer
        Dim stringQuery As String
        stringQuery = "DELETE FROM " & table & " WHERE " & clause
        Try
            myCommand = New SqlCommand(stringQuery, myConn)
            result = myCommand.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
        Return result
    End Function
	'Executa la consulta contenida en la cadena  retorna un datatable con los resultados.
    Public Function runFreeQuery(ByVal stringQuery As String) As DataTable
        'MessageBox.Show(stringQuery)
        Dim result As New DataTable
        Try
            myAdapter = New SqlDataAdapter(stringQuery, myConn)
            myAdapter.Fill(result)
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
        Return result
    End Function
End Class
