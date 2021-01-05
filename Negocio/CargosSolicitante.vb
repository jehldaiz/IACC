Imports System.Data
Imports Datos
Public Class CargosSolicitante
    Private _wobjDatos As AccessDatos

    Private _idCargoSol As Integer
    Private _nomCargoSol As String
    Private _activo As Integer

    Public Property IdCargoSol As Integer
        Get
            Return _idCargoSol
        End Get
        Set(ByVal value As Integer)
            _idCargoSol = value
        End Set
    End Property
    Public Property NomCargoSol As String
        Get
            Return _nomCargoSol
        End Get
        Set(ByVal value As String)
            _nomCargoSol = value
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

    Public Function QryGetAllCargoSolicitante() As DataSet
        _wobjDatos = New AccessDatos()
        Dim wvarDataSetEstado As DataSet
        _wobjDatos.StoreProcedure = "QryGetAllCargoSolicitante"
        wvarDataSetEstado = _wobjDatos.ExecuteStoreProcedureGeneral()
        Return wvarDataSetEstado
    End Function
End Class
