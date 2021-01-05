Imports System
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration

Public Class AccessDatos
    Private _wvarStoreProcedure As String
    Private _wvarParameterStoreProcedure() As String
    Private _wvarValueStoreProcedure() As String
    Private _wvarDataTypeParameterStoreProcuder() As Object
    Private _wvarSizeParameterStoreProcedure() As Integer
    Private _wvarAffectedRow As Integer
    Private _wvarCount As Integer
    Private _wobjTransaction As SqlTransaction
    Private ReadOnly _wobjConnectionSql As SqlConnection
    Private ReadOnly _wobjCommandSql As SqlCommand
    Private _wobjAdapterSql As SqlDataAdapter
    Private _wvarDataSet As DataSet
    Private _wvarDimensionParameter As Integer

    Public Sub New()
        Dim sCnn As String
        sCnn = ConfigurationManager.ConnectionStrings("conexionOrden").ConnectionString
        _wobjConnectionSql = New SqlConnection(sCnn)
        _wobjCommandSql = New SqlCommand()
        _wvarDataSet = New DataSet()
        _wvarCount = 0
    End Sub

    Public Property StoreProcedure As String
        Get
            Return _wvarStoreProcedure
        End Get
        Set(ByVal value As String)
            _wvarStoreProcedure = value
        End Set
    End Property

    Public Property ParameterStoreProcedure As String()
        Get
            Return _wvarParameterStoreProcedure
        End Get
        Set(ByVal value() As String)
            _wvarParameterStoreProcedure = value
        End Set
    End Property

    Public Property ValueStoreProcedure As String()
        Get
            Return _wvarValueStoreProcedure
        End Get
        Set(ByVal value() As String)
            _wvarValueStoreProcedure = value
        End Set
    End Property

    Public Property DataTypeParameterStoreProcedure As Object()
        Get
            Return _wvarDataTypeParameterStoreProcuder
        End Get
        Set(ByVal value() As Object)
            _wvarDataTypeParameterStoreProcuder = value
        End Set
    End Property

    Public Property SizeParameterStoreProcedure As Integer()
        Get
            Return _wvarSizeParameterStoreProcedure
        End Get
        Set(ByVal value() As Integer)
            _wvarSizeParameterStoreProcedure = value
        End Set
    End Property

    Public Property DimensionParameter As Integer
        Get
            Return _wvarDimensionParameter
        End Get
        Set(ByVal value As Integer)
            _wvarDimensionParameter = value
        End Set
    End Property

    Private Function GetConexion() As SqlConnection
        Return _wobjConnectionSql
    End Function

    Public Function ExecuteStoreProcedureGeneral() As DataSet
        Try
            GetConexion().Open()
            _wobjTransaction = GetConexion().BeginTransaction()
            _wobjCommandSql.Transaction = _wobjTransaction
            _wobjCommandSql.CommandTimeout = 200
            _wobjCommandSql.CommandType = CommandType.StoredProcedure
            _wobjCommandSql.CommandText = StoreProcedure
            _wobjCommandSql.Connection = GetConexion()
            _wobjAdapterSql = New SqlDataAdapter(_wobjCommandSql)
            _wobjAdapterSql.Fill(_wvarDataSet)
            If _wvarDataSet.Tables(0).Rows.Count > 0 Then
                _wobjTransaction.Commit()
                GetConexion().Close()
            End If
        Catch ex As Exception
            _wobjTransaction.Rollback()
            GetConexion().Close()
        Finally
            GetConexion().Close()
        End Try
        Return _wvarDataSet
    End Function

    Public Function ExecuteStoreProcedureSpecific() As DataSet
        Try
            _wvarDataSet = New DataSet
            GetConexion().Open()
            _wobjTransaction = GetConexion().BeginTransaction()
            _wobjCommandSql.Transaction = _wobjTransaction
            _wobjCommandSql.CommandTimeout = 200
            _wobjCommandSql.CommandType = CommandType.StoredProcedure
            _wobjCommandSql.CommandText = StoreProcedure
            _wobjCommandSql.Connection = GetConexion()
            _wobjCommandSql.Parameters.Clear()
            For _wvarCount = 0 To _wvarParameterStoreProcedure.Length - 1
                _wobjCommandSql.Parameters.Add("@" & _wvarParameterStoreProcedure(_wvarCount).ToString(), _wvarDataTypeParameterStoreProcuder(_wvarCount), _wvarSizeParameterStoreProcedure(_wvarCount)).Value = _wvarValueStoreProcedure(_wvarCount)
            Next

            _wobjAdapterSql = New SqlDataAdapter(_wobjCommandSql)
            _wobjAdapterSql.Fill(_wvarDataSet)
            If _wvarDataSet.Tables(0).Rows.Count > 0 Then
                _wobjTransaction.Commit()
                GetConexion().Close()
            End If
        Catch ex As Exception
            _wobjTransaction.Rollback()
            GetConexion().Close()
        Finally
            GetConexion().Close()
        End Try
        Return _wvarDataSet
    End Function

    Public Function ExecuteStoreProcedureAction() As Integer
        Try
            GetConexion().Open()
            _wobjTransaction = GetConexion().BeginTransaction()
            _wobjCommandSql.Transaction = _wobjTransaction
            _wobjCommandSql.CommandTimeout = 200
            _wobjCommandSql.CommandType = CommandType.StoredProcedure
            _wobjCommandSql.CommandText = StoreProcedure
            _wobjCommandSql.Connection = GetConexion()
            For _wvarCount = 0 To _wvarParameterStoreProcedure.Length - 1
                _wobjCommandSql.Parameters.Add("@" & _wvarParameterStoreProcedure(_wvarCount).ToString(), _wvarDataTypeParameterStoreProcuder(_wvarCount), _wvarSizeParameterStoreProcedure(_wvarCount)).Value = _wvarValueStoreProcedure(_wvarCount)
            Next


            _wvarAffectedRow = _wobjCommandSql.ExecuteNonQuery()
            If _wvarAffectedRow > 0 Then
                _wobjTransaction.Commit()
                GetConexion().Close()
            End If
        Catch ex As Exception
            _wobjTransaction.Rollback()
            GetConexion().Close()
        Finally
            GetConexion().Close()
        End Try
        Return _wvarAffectedRow
    End Function


    Public Function ExecuteStoreProcedureActionByte(ByVal data As Byte()) As Integer
        Try

            GetConexion().Open()
            _wobjTransaction = GetConexion().BeginTransaction()
            _wobjCommandSql.Transaction = _wobjTransaction
            _wobjCommandSql.CommandTimeout = 200
            _wobjCommandSql.CommandType = CommandType.StoredProcedure
            _wobjCommandSql.CommandText = StoreProcedure
            _wobjCommandSql.Connection = GetConexion()
            _wobjCommandSql.Parameters.Clear()
            For _wvarCount = 0 To _wvarParameterStoreProcedure.Length - 1
                _wobjCommandSql.Parameters.Add("@" & _wvarParameterStoreProcedure(_wvarCount).ToString(), _wvarDataTypeParameterStoreProcuder(_wvarCount), _wvarSizeParameterStoreProcedure(_wvarCount)).Value = _wvarValueStoreProcedure(_wvarCount)
            Next

            'If (data.Length > 0) Then
            _wobjCommandSql.Parameters.Add("@pData", SqlDbType.VarBinary).Value = data
            'End If

            _wvarAffectedRow = _wobjCommandSql.ExecuteNonQuery()
            If _wvarAffectedRow > 0 Then
                _wobjTransaction.Commit()
                GetConexion().Close()
            End If
        Catch ex As Exception
            _wobjTransaction.Rollback()
            GetConexion().Close()
        Finally
            GetConexion().Close()
        End Try
        Return _wvarAffectedRow
    End Function

    Public Function ValidateDataBaseConnection() As Integer
        Try
            GetConexion().Open()
            Return 0
        Catch ex As Exception
            Return 10
        End Try
    End Function

    Public Sub ReDimParameterStoreProcedure(ByVal wvarDimensionParameter)
        ReDim ParameterStoreProcedure(wvarDimensionParameter)
        ReDim ValueStoreProcedure(wvarDimensionParameter)
        ReDim DataTypeParameterStoreProcedure(wvarDimensionParameter)
        ReDim SizeParameterStoreProcedure(wvarDimensionParameter)
    End Sub

    Public Shared Function GetAccesoDatos(ByVal storedProcedure As String, ByVal parameterIndex As Integer) As AccessDatos

        Dim accesoDatos = New AccessDatos()

        ReDim accesoDatos.ParameterStoreProcedure(parameterIndex)
        ReDim accesoDatos.ValueStoreProcedure(parameterIndex)
        ReDim accesoDatos.DataTypeParameterStoreProcedure(parameterIndex)
        ReDim accesoDatos.SizeParameterStoreProcedure(parameterIndex)
        accesoDatos.StoreProcedure = storedProcedure

        Return accesoDatos
    End Function

    Public Shared Sub SetParameter(ByVal accesoDatos As AccessDatos, ByVal parameterIndex As Integer, ByVal name As String, ByVal value As String, ByVal type As SqlDbType, ByVal size As Integer)
        accesoDatos.ParameterStoreProcedure(parameterIndex) = name
        accesoDatos.ValueStoreProcedure(parameterIndex) = value
        accesoDatos.DataTypeParameterStoreProcedure(parameterIndex) = type
        accesoDatos.SizeParameterStoreProcedure(parameterIndex) = size
    End Sub
End Class
