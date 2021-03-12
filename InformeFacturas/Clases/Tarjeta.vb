Public Class Tarjeta
#Region "Propiedades"
    Public Property IDtarjeta As Integer
    Public Property Descripcion As String
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
