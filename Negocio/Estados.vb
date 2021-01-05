Imports System.Data
Imports Datos
Public Class Estados
    Private _wobjDatos As AccessDatos

    Private _idestado As Integer
    Private _nomEstado As String
    Private _activo As Integer

    Public Property Idestado As Integer
        Get
            Return _idestado
        End Get
        Set(ByVal value As Integer)
            _idestado = value
        End Set
    End Property
    Public Property NomEstado As String
        Get
            Return _nomEstado
        End Get
        Set(ByVal value As String)
            _nomEstado = value
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

    Public Function QryGetAllEstados() As DataSet
        _wobjDatos = New AccessDatos()
        Dim wvarDataSetEstado As DataSet
        _wobjDatos.StoreProcedure = "QryGetAllEstados"
        wvarDataSetEstado = _wobjDatos.ExecuteStoreProcedureGeneral()
        Return wvarDataSetEstado
    End Function
End Class
