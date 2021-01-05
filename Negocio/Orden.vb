Imports System.Data
Imports Datos
Public Class Orden
    Private _wobjDatos As AccessDatos

    Private _nOrden As Integer
    Private _fecha As DateTime
    Private _estado As Integer
    Private _estadoNom As String
    Private _servicio As Integer
    Private _servicioNom As String
    Private _modulo As String
    Private _piso As String
    Private _unidadDestino As Integer
    Private _unidadDestinoNom As String
    Private _recinto As String
    Private _descripcionProblema As String
    Private _operador As String
    Private _solicitante As String
    Private _cargoSolicitante As String
    Private _cargoSolicitanteId As Integer
    ' Private _listaDetalleOrden As List(Of OrdenDetalle)


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
    Public Property Estado As Integer
        Get
            Return _estado
        End Get
        Set(value As Integer)
            _estado = value
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
    Public Property Servicio As Integer
        Get
            Return _servicio
        End Get
        Set(value As Integer)
            _servicio = value
        End Set
    End Property
    Public Property ServicioNom As String
        Get
            Return _servicioNom
        End Get
        Set(value As String)
            _servicioNom = value
        End Set
    End Property
    Public Property Modulo As String
        Get
            Return _modulo
        End Get
        Set(value As String)
            _modulo = value
        End Set
    End Property

    Public Property Piso As String
        Get
            Return _piso
        End Get
        Set(value As String)
            _piso = value
        End Set
    End Property
    Public Property UnidadDestino As Integer
        Get
            Return _unidadDestino
        End Get
        Set(value As Integer)
            _unidadDestino = value
        End Set
    End Property
    Public Property UnidadDestinoNom As String
        Get
            Return _unidadDestinoNom
        End Get
        Set(value As String)
            _unidadDestinoNom = value
        End Set
    End Property

    Public Property Recinto As String
        Get
            Return _recinto
        End Get
        Set(value As String)
            _recinto = value
        End Set
    End Property

    Public Property DescripcionProblema As String
        Get
            Return _descripcionProblema
        End Get
        Set(value As String)
            _descripcionProblema = value
        End Set
    End Property
    Public Property Operador As String
        Get
            Return _operador
        End Get
        Set(value As String)
            _operador = value
        End Set
    End Property
    Public Property Solicitante As String
        Get
            Return _solicitante
        End Get
        Set(value As String)
            _solicitante = value
        End Set
    End Property
    Public Property CargoSolicitante As String
        Get
            Return _cargoSolicitante
        End Get
        Set(value As String)
            _cargoSolicitante = value
        End Set
    End Property
    Public Property CargoSolicitanteId As Integer
        Get
            Return _cargoSolicitanteId
        End Get
        Set(value As Integer)
            _cargoSolicitanteId = value
        End Set
    End Property
    Public Function QryGetAllOrdenesByUnidad() As DataSet
        _wobjDatos = AccessDatos.GetAccesoDatos("QryGetAllOrdenesByUnidad", 0)
        AccessDatos.SetParameter(_wobjDatos, 0, "Uunidad", UnidadDestino, SqlDbType.Int, 4)
        Return _wobjDatos.ExecuteStoreProcedureSpecific()
    End Function
    Public Function QryGetOrdenesById() As DataSet
        _wobjDatos = AccessDatos.GetAccesoDatos("QryGetOrdenesById", 0)
        AccessDatos.SetParameter(_wobjDatos, 0, "NOrden", NOrden, SqlDbType.Int, 4)
        Return _wobjDatos.ExecuteStoreProcedureSpecific()
    End Function

    Public Function QryGetAllOrdenes() As DataSet
        _wobjDatos = New AccessDatos()
        Dim wvarDataSetEstado As DataSet
        _wobjDatos.StoreProcedure = "QryGetAllOrdenes"
        wvarDataSetEstado = _wobjDatos.ExecuteStoreProcedureGeneral()
        Return wvarDataSetEstado
    End Function
    Public Function DelUSuario() As DataSet

        Dim acceso As AccessDatos = AccessDatos.GetAccesoDatos("DelUSuario", 0)
        AccessDatos.SetParameter(acceso, 0, "Id", NOrden, SqlDbType.Int, 4)

        Return acceso.ExecuteStoreProcedureSpecific()
    End Function
    Public Function GetOrdenesByFiltro() As DataSet

        _wobjDatos = New AccessDatos()
        Dim wvarDataSetUsuario As DataSet
        ReDim _wobjDatos.ParameterStoreProcedure(3)
        ReDim _wobjDatos.ValueStoreProcedure(3)
        ReDim _wobjDatos.DataTypeParameterStoreProcedure(3)
        ReDim _wobjDatos.SizeParameterStoreProcedure(3)
        _wobjDatos.StoreProcedure = "GetOrdenesByFiltro"
        _wobjDatos.ParameterStoreProcedure(0) = "Numero"
        _wobjDatos.ParameterStoreProcedure(1) = "Servicio"
        _wobjDatos.ParameterStoreProcedure(2) = "Destino"
        _wobjDatos.ParameterStoreProcedure(3) = "Estado"


        _wobjDatos.ValueStoreProcedure(0) = NOrden
        _wobjDatos.ValueStoreProcedure(1) = Servicio
        _wobjDatos.ValueStoreProcedure(2) = UnidadDestino
        _wobjDatos.ValueStoreProcedure(3) = Estado

        _wobjDatos.DataTypeParameterStoreProcedure(0) = SqlDbType.Int
        _wobjDatos.DataTypeParameterStoreProcedure(1) = SqlDbType.Int
        _wobjDatos.DataTypeParameterStoreProcedure(2) = SqlDbType.Int
        _wobjDatos.DataTypeParameterStoreProcedure(3) = SqlDbType.Int


        _wobjDatos.SizeParameterStoreProcedure(0) = 4
        _wobjDatos.SizeParameterStoreProcedure(1) = 60
        _wobjDatos.SizeParameterStoreProcedure(2) = 60
        _wobjDatos.SizeParameterStoreProcedure(3) = 60

        wvarDataSetUsuario = _wobjDatos.ExecuteStoreProcedureSpecific()
        Return wvarDataSetUsuario
    End Function
    Public Function INS_Ordenes() As DataSet

        Dim acceso As AccessDatos = AccessDatos.GetAccesoDatos("INS_Ordenes", 9)

        AccessDatos.SetParameter(acceso, 0, "Fecha", Fecha, SqlDbType.DateTime, 30)
        AccessDatos.SetParameter(acceso, 1, "Servicio", Servicio, SqlDbType.Int, 4)
        AccessDatos.SetParameter(acceso, 2, "Recinto", Recinto, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 3, "Modulo", Modulo, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 4, "Piso", Piso, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 5, "Destino", UnidadDestino, SqlDbType.Int, 4)
        AccessDatos.SetParameter(acceso, 6, "Solicitante", Solicitante, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 7, "DescProblema", DescripcionProblema, SqlDbType.NVarChar, 800)
        AccessDatos.SetParameter(acceso, 8, "Operador", Operador, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 9, "Cargo_sol", CargoSolicitante, SqlDbType.Int, 4)


        Return acceso.ExecuteStoreProcedureSpecific()
    End Function

    Shared Function Tables(ByVal p1 As Integer) As Object
        Throw New NotImplementedException
    End Function

End Class
