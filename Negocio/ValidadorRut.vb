Public Class ValidadorRut
    Public wvarRut As Integer
    Public wvarRutDv As String

    Public Property Rut As Integer
        Get
            Return wvarRut
        End Get
        Set(ByVal value As Integer)
            wvarRut = value
        End Set
    End Property
    Public Property RutDv As String
        Get
            Return wvarRutDv
        End Get
        Set(ByVal value As String)
            wvarRutDv = value
        End Set
    End Property
    Public Sub New()
        Rut = 0
        RutDv = ""
    End Sub
    Public Function validarRut() As Boolean

        Dim rutLimpio As String

        Dim digitoVerificador As String

        Dim suma As Integer

        Dim contador As Integer = 2

        Dim valida As Boolean = True

        Dim retorno As Boolean = False


        rutLimpio = Rut.ToString()





        digitoVerificador = RutDv.ToString()


        Dim i As Integer


        For i = rutLimpio.Length - 1 To 0 Step -1


            If contador > 7 Then

                contador = 2

            End If


            Try

                suma += Integer.Parse(rutLimpio(i).ToString()) * contador

                contador += 1

            Catch ex As Exception

                valida = False

            End Try


        Next


        If valida Then

            Dim dv As Integer = 11 - (suma Mod 11)

            Dim DigVer As String = dv.ToString()


            If DigVer = "10" Then

                DigVer = "K"

            End If


            If DigVer = "11" Then

                DigVer = "0"

            End If


            If DigVer = digitoVerificador.ToUpper Then

                retorno = True


            Else

                retorno = False

            End If

        End If

        Return retorno

    End Function
End Class

