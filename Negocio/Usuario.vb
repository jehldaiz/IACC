Imports System.Data
Imports Datos
Public Class Usuario
    Private _wobjDatos As AccessDatos

    Private _userName As String
    Private _nomCompleto As String
    Private _rutUsuario As Integer
    Private _rutDvUsuario As String
    Private _mailUsuario As String
    Private _passUsuario As String
    Private _passDefault As Integer
    Private _fechaCreacion As DateTime
    Private _idPerfil As Integer
    Private _nombrePerfil As String
    Private _idUnidada As Integer
    Private _nombreUnidad As String
    Private _activo As Integer

    Public Property NombreUnidad As String
        Get
            Return _nombreUnidad
        End Get
        Set(value As String)
            _nombreUnidad = value
        End Set
    End Property
    Public Property NombrePerfil As String
        Get
            Return _nombrePerfil
        End Get
        Set(value As String)
            _nombrePerfil = value
        End Set
    End Property
    Public Property PassUsuario As String
        Get
            Return _passUsuario
        End Get
        Set(value As String)
            _passUsuario = value
        End Set
    End Property
    Public Property MailUsuario As String
        Get
            Return _mailUsuario
        End Get
        Set(value As String)
            _mailUsuario = value
        End Set
    End Property
    Public Property RutDvUsuario As String
        Get
            Return _rutDvUsuario
        End Get
        Set(value As String)
            _rutDvUsuario = value
        End Set
    End Property
    Public Property NomCompleto As String
        Get
            Return _nomCompleto
        End Get
        Set(value As String)
            _nomCompleto = value
        End Set
    End Property
    Public Property UserName As String
        Get
            Return _userName
        End Get
        Set(value As String)
            _userName = value
        End Set
    End Property

    Public Property FechaCreacion As DateTime
        Get
            Return _fechaCreacion
        End Get
        Set(value As DateTime)
            _fechaCreacion = value
        End Set
    End Property
    Public Property RutUsuario As Integer
        Get
            Return _rutUsuario
        End Get
        Set(value As Integer)
            _rutUsuario = value
        End Set
    End Property
    Public Property PassDefault As Integer
        Get
            Return _passDefault
        End Get
        Set(value As Integer)
            _passDefault = value
        End Set
    End Property
    Public Property IdPerfil As Integer
        Get
            Return _idPerfil
        End Get
        Set(value As Integer)
            _idPerfil = value
        End Set
    End Property
    Public Property IdUnidada As Integer
        Get
            Return _idUnidada
        End Get
        Set(value As Integer)
            _idUnidada = value
        End Set
    End Property
    Public Property Activo As Integer
        Get
            Return _activo
        End Get
        Set(value As Integer)
            _activo = value
        End Set
    End Property
    Public Function GetUsuario() As DataSet

        _wobjDatos = New AccessDatos()
        Dim wvarDataSetUsuario As DataSet
        ReDim _wobjDatos.ParameterStoreProcedure(0)
        ReDim _wobjDatos.ValueStoreProcedure(0)
        ReDim _wobjDatos.DataTypeParameterStoreProcedure(0)
        ReDim _wobjDatos.SizeParameterStoreProcedure(0)
        _wobjDatos.StoreProcedure = "GetUsuario"
        _wobjDatos.ParameterStoreProcedure(0) = "userName"
        ' _wobjDatos.ParameterStoreProcedure(1) = "contraseña"

        _wobjDatos.ValueStoreProcedure(0) = UserName
        ' _wobjDatos.ValueStoreProcedure(1) = PassUsuario

        _wobjDatos.DataTypeParameterStoreProcedure(0) = SqlDbType.VarChar
        ' _wobjDatos.DataTypeParameterStoreProcedure(1) = SqlDbType.VarChar

        _wobjDatos.SizeParameterStoreProcedure(0) = 255
        ' _wobjDatos.SizeParameterStoreProcedure(1) = 255

        wvarDataSetUsuario = _wobjDatos.ExecuteStoreProcedureSpecific()
        Return wvarDataSetUsuario
    End Function
    Public Function UPD_PassUser() As DataSet

        Dim acceso As AccessDatos = AccessDatos.GetAccesoDatos("UPD_PassUser", 1)

        AccessDatos.SetParameter(acceso, 0, "userName", UserName, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 1, "contraseña", PassUsuario, SqlDbType.NVarChar, 255)

        Return acceso.ExecuteStoreProcedureSpecific()
    End Function
End Class
