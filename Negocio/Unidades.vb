Imports System.Data
Imports Datos
Public Class Unidades
    Private _wobjDatos As AccessDatos

    Private _idUnidad As Integer
    Private _nomUnidad As String
    Private _activo As Integer
    Private _correoJefe As String
    Private _correoTaller As String

    Public Property IdUnidad As Integer
        Get
            Return _idUnidad
        End Get
        Set(ByVal value As Integer)
            _idUnidad = value
        End Set
    End Property
    Public Property NomUnidad As String
        Get
            Return _nomUnidad
        End Get
        Set(ByVal value As String)
            _nomUnidad = value
        End Set
    End Property

    Public Property Activo As Integer
        Get
            Return _activo
        End Get
        Set(ByVal value As Integer)
            _activo = value
        End Set
    End Property
    Public Property CorreoJefe As String
        Get
            Return _correoJefe
        End Get
        Set(ByVal value As String)
            _correoJefe = value
        End Set
    End Property
    Public Property CorreoTaller As String
        Get
            Return _correoTaller
        End Get
        Set(ByVal value As String)
            _correoTaller = value
        End Set
    End Property

    Public Function QryGetAllUnidadDes() As DataSet
        _wobjDatos = New AccessDatos()
        Dim wvarDataSetEstado As DataSet
        _wobjDatos.StoreProcedure = "QryGetAllUnidadDes"
        wvarDataSetEstado = _wobjDatos.ExecuteStoreProcedureGeneral()
        Return wvarDataSetEstado
    End Function

    Public Function QryGetAllUnidadDesById() As DataSet
        _wobjDatos = AccessDatos.GetAccesoDatos("QryGetAllUnidadDesById", 0)
        AccessDatos.SetParameter(_wobjDatos, 0, "ID", IdUnidad, SqlDbType.Int, 4)
        Return _wobjDatos.ExecuteStoreProcedureSpecific()
    End Function

    Public Function GetJefeUnidad(ByVal idUnidad As Integer) As DataSet
        _wobjDatos = AccessDatos.GetAccesoDatos("GetJefeUnidad", 0)
        AccessDatos.SetParameter(_wobjDatos, 0, "IdUnidad", idUnidad, SqlDbType.Int, 4)
        Return _wobjDatos.ExecuteStoreProcedureSpecific()
    End Function
End Class
