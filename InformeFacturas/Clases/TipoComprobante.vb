Public Class TipoComprobante
#Region "Propiedades"
    Public Property IDtipoComprobante As Integer
    Public Property Descripcion As String
    Public Property Codigo As String
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
