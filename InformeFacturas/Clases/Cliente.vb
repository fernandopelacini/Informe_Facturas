Public Class Cliente
#Region "Miembros"
    Public Property idCliente As Integer
    Public Property RazonSocial As String
    Public Property condicionIva As String
    Public Property Cuit As String
    Public Property Direccion As String
    Public Property CodigoPostal As String
    Public Property Localidad As String
    Public Property Provincia As String
    Public Property FormaPago As String
    Public Property FechaAlta As DateTime
    Public Property FechaModificacion As DateTime
    Public Property FechaBaja As DateTime
    Public Property Activo As Boolean
    Public Property Email As String
    Public Property Telefono As String
    Public Property Movil As String

#End Region
#Region "Metodos"

    ''' <summary>
    ''' Destruye el objeto
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose()
        MyBase.Finalize()
    End Sub



#End Region
End Class
