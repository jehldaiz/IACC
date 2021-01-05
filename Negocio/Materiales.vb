Imports System.Data
Imports Datos
Public Class Materiales
    Private _wobjDatos As AccessDatos

    Private _id As Integer
    Private _nOrden As Integer
    Private _materialNom As String
    Private _cantidad As Integer
    Private _precio As Integer

    Public Property NOrden As Integer
        Get
            Return _nOrden
        End Get
        Set(value As Integer)
            _nOrden = value
        End Set
    End Property

    Public Property Id As Integer
        Get
            Return _id
        End Get
        Set(value As Integer)
            _id = value
        End Set
    End Property
    Public Property MaterialNom As String
        Get
            Return _materialNom
        End Get
        Set(value As String)
            _materialNom = value
        End Set
    End Property

    Public Property Precio As Integer
        Get
            Return _precio
        End Get
        Set(value As Integer)
            _precio = value
        End Set
    End Property
    Public Property Cantidad As Integer
        Get
            Return _cantidad
        End Get
        Set(value As Integer)
            _cantidad = value
        End Set
    End Property

    Public Function INS_ItemOt() As DataSet

        Dim acceso As AccessDatos = AccessDatos.GetAccesoDatos("INS_ItemOt", 4)

        AccessDatos.SetParameter(acceso, 0, "ID", Id, SqlDbType.Int, 4)
        AccessDatos.SetParameter(acceso, 1, "OT", NOrden, SqlDbType.Int, 4)
        AccessDatos.SetParameter(acceso, 2, "NOMITEM", MaterialNom, SqlDbType.NVarChar, 255)
        AccessDatos.SetParameter(acceso, 3, "CANTIDAD", Cantidad, SqlDbType.Int, 4)
        AccessDatos.SetParameter(acceso, 4, "VALORITEM", Precio, SqlDbType.Int, 4)


        Return acceso.ExecuteStoreProcedureSpecific()
    End Function

    Public Function Delete_ItemOt() As DataSet

        Dim acceso As AccessDatos = AccessDatos.GetAccesoDatos("DeleteMaterialByOT", 1)

        AccessDatos.SetParameter(acceso, 0, "OT", NOrden, SqlDbType.Int, 4)
        AccessDatos.SetParameter(acceso, 1, "ID", Id, SqlDbType.Int, 4)
        Return acceso.ExecuteStoreProcedureSpecific()
    End Function

    Public Function GetMaterialesByOt() As DataSet
        _wobjDatos = AccessDatos.GetAccesoDatos("GetMaterialesByOt", 0)
        AccessDatos.SetParameter(_wobjDatos, 0, "OT", NOrden, SqlDbType.Int, 4)
        Return _wobjDatos.ExecuteStoreProcedureSpecific()
    End Function

    Public Function GetInicioFinOt(ByVal ot As Integer, ByVal est As Integer, ByVal var As String) As DataSet
        _wobjDatos = AccessDatos.GetAccesoDatos("GetInicioFinOt", 2)
        AccessDatos.SetParameter(_wobjDatos, 0, "OT", ot, SqlDbType.Int, 4)
        AccessDatos.SetParameter(_wobjDatos, 1, "ESTADO", est, SqlDbType.Int, 4)
        AccessDatos.SetParameter(_wobjDatos, 2, "VAR", var, SqlDbType.Char, 50)
        Return _wobjDatos.ExecuteStoreProcedureSpecific()
    End Function
End Class
