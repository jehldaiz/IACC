Imports System.Data
Imports Datos
Public Class Servicios
    Private _wobjDatos As AccessDatos

    Private _idServicio As Integer
    Private _nomServicio As String
    Private _activo As Integer

    Public Property IdServicio As Integer
        Get
            Return _idServicio
        End Get
        Set(ByVal value As Integer)
            _idServicio = value
        End Set
    End Property
    Public Property NomServicio As String
        Get
            Return _nomServicio
        End Get
        Set(ByVal value As String)
            _nomServicio = value
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

    Public Function QryGetAllServicios() As DataSet
        _wobjDatos = New AccessDatos()
        Dim wvarDataSetEstado As DataSet
        _wobjDatos.StoreProcedure = "QryGetAllServicios"
        wvarDataSetEstado = _wobjDatos.ExecuteStoreProcedureGeneral()
        Return wvarDataSetEstado
    End Function
End Class
