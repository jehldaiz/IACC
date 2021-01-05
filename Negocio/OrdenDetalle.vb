Imports System.Data
Imports Datos
Public Class OrdenDetalle
    Private _wobjDatos As AccessDatos

    Private _id As Integer
    Private _nOrden As Integer
    Private _fecha As DateTime
    Private _tecnicoUnidad As String
    Private _jefeUnidad As String
    Private _solucionProblema As String
    Private _costoTotal As Integer
    Private _jefeServicio As String
    Private _estadoNom As String
    Private _estadoId As Integer
    Private _usuario As String
    Private _comentario As String

    Public Property Id As Integer
        Get
            Return _id
        End Get
        Set(value As Integer)
            _id = value
        End Set
    End Property

    Public Property NOrden As Integer
        Get
            Return _nOrden
        End Get
        Set(value As Integer)
            _nOrden = value
        End Set
    End Property
    Public Property Fecha As DateTime
        Get
            Return _fecha
        End Get
        Set(value As DateTime)
            _fecha = value
        End Set
    End Property
    Public Property EstadoNom As String
        Get
            Return _estadoNom
        End Get
        Set(value As String)
            _estadoNom = value
        End Set
    End Property
    Public Property EstadoId As Integer
        Get
            Return _estadoId
        End Get
        Set(value As Integer)
            _estadoId = value
        End Set
    End Property
    Public Property TecnicoUnidad As String
        Get
            Return _tecnicoUnidad
        End Get
        Set(value As String)
            _tecnicoUnidad = value
        End Set
    End Property
    Public Property JefeUnidad As String
        Get
            Return _jefeUnidad
        End Get
        Set(value As String)
            _jefeUnidad = value
        End Set
    End Property
    Public Property SolucionProblema As String
        Get
            Return _solucionProblema
        End Get
        Set(value As String)
            _solucionProblema = value
        End Set
    End Property
    Public Property CostoTotal As Integer
        Get
            Return _costoTotal
        End Get
        Set(value As Integer)
            _costoTotal = value
        End Set
    End Property
    Public Property JefeServicio As String
        Get
            Return _jefeServicio
        End Get
        Set(value As String)
            _jefeServicio = value
        End Set
    End Property
    Public Property Usuario As String
        Get
            Return _usuario
        End Get
        Set(value As String)
            _usuario = value
        End Set
    End Property
    Public Property Comentario As String
        Get
            Return _comentario
        End Get
        Set(value As String)
            _comentario = value
        End Set
    End Property

    Public Function QryGetOrdenesDetalleById() As DataSet
        _wobjDatos = AccessDatos.GetAccesoDatos("QryGetOrdenesDetalleById", 0)
        AccessDatos.SetParameter(_wobjDatos, 0, "NOrden", NOrden, SqlDbType.Int, 4)
        Return _wobjDatos.ExecuteStoreProcedureSpecific()
    End Function
    Public Function GetTecnicoMaxRegistro() As DataSet
        _wobjDatos = AccessDatos.GetAccesoDatos("GetTecnicoMaxRegistro", 0)
        AccessDatos.SetParameter(_wobjDatos, 0, "nOrden", NOrden, SqlDbType.Int, 4)
        Return _wobjDatos.ExecuteStoreProcedureSpecific()
    End Function
    Public Function INS_DetalleOrden() As DataSet

        Dim acceso As AccessDatos = AccessDatos.GetAccesoDatos("INS_DetalleOrden", 8)

        AccessDatos.SetParameter(acceso, 0, "OT", NOrden, SqlDbType.Int, 4)
        AccessDatos.SetParameter(acceso, 1, "FECHA", Fecha, SqlDbType.DateTime, 30)
        AccessDatos.SetParameter(acceso, 2, "TECUNIDAD", TecnicoUnidad, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 3, "JEFEUNIDAD", JefeUnidad, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 4, "SOLUCION", SolucionProblema, SqlDbType.NVarChar, 1200)
        AccessDatos.SetParameter(acceso, 5, "ESTADO", EstadoId, SqlDbType.Int, 4)
        AccessDatos.SetParameter(acceso, 6, "USUARIO", Usuario, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 7, "COMENTARIO", Comentario, SqlDbType.NVarChar, 1200)
        AccessDatos.SetParameter(acceso, 8, "SUPERVISOR", JefeServicio, SqlDbType.NVarChar, 255)

        Return acceso.ExecuteStoreProcedureSpecific()
    End Function
End Class

