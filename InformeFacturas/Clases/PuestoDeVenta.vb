Public Class PuestoDeVenta
#Region "Propiedades"
    Public Property IDpuestoDeVenta As Integer
    Public Property PuestoVenta As String
    Public Property NumeroSerie As String
    Public Property Modelo As String
    Public Property Marca As String
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
