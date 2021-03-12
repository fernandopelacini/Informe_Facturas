Public Class Cuenta
#Region "Propiedades"
    Public Property IDcuenta As Integer
    Public Property Cuenta As String
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

    Public Sub New()
        MyBase.New()
    End Sub

#End Region
End Class
