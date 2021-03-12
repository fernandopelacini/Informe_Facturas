
Public Class Factura
#Region "Propiedades"
    Public Property IDfactura As Integer
    Public Property NumeroFactura As String
    Public Property IDcaea As CAEA
    Public Property IDpuestoVenta As PuestoDeVenta
    Public Property IDcliente As Cliente
    Public Property IDiva As CondicionIVA
    Public Property IDtipoComprobante As TipoComprobante
    Public Property FechaVenta As DateTime
    Public Property CondicionVenta As String
    Public Property Subtotal As Decimal
    Public Property Impuestos As Decimal
    Public Property IvaInscripto As Decimal
    Public Property IvaNoInscripto As Decimal
    Public Property Total As Decimal
    Public Property InformadaEnAFIP As Boolean
    Public Property Activo As Boolean
    Public Property Oblea As String
    Public Property Dominio As String
    Public Property IDusuario As String
    Public Property Categoria As String
    Public Property Titutlar As String 'SE USA EN LUGAR DEL ID DE CLIENTE AHORA. ON HOLD HASTA QUE NO ESTE EL WEBSERVICE
    Public Property AnioVehiculo As String
    Public Property CUIT As String
    Public Property PendienteDeFacturar As Boolean 'Se usa para los remitos cuando aun se no se facturo, el tipo de comprobante va a ser remito y este campo =1
    Public Property PendienteDeCobrar As Boolean ' Se utiliza para cuando se genero la factura pero no se cobro aun, no tiene un recibo por la factura.
    Public Property IDfacturaCancela As Integer 'Se usa para indicar el ID de la factura que cancela uno o varios remitos
    Public Property IDcuenta As Cuenta
    Public Property CantidadCuotas As SByte
    Public Property NumeroTransferencia As String
    Public Property Lote As String
    Public Property IDTarjeta As Tarjeta
    Public Property NumeroTarjeta As String
#End Region

#Region "Metodos"
    ''' <summary>
    ''' Destruye el objeto
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose()
        MyBase.Finalize()
    End Sub

    ''' <summary>
    ''' Inicializa el objeto
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        IDcaea = New CAEA()
        IDpuestoVenta = New PuestoDeVenta()
        IDcliente = New Cliente()
        IDiva = New CondicionIVA()
        IDtipoComprobante = New TipoComprobante()
        IDcuenta = New Cuenta()
        IDTarjeta = New Tarjeta
    End Sub

    ''' <summary>
    ''' Trae las facturas sin informar y que todavia estan dentro del tope permitido por el CAE
    ''' Hay que informar por un lado las FACTURAS A y por otro las B
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ListarFacturasSinInformar(ByRef dsFacturas As dsDatos, ByVal intTipoComprobante As Int16)

        Try
            conexion.Command.Parameters.Clear()
            conexion.ExecuteDataset(TRAER_FACTURAS_PARA_INFORMAR,
                                             TABLA_FACTURAS, _
                                             "@tipocomprobante", _
                                             intTipoComprobante, _
                                             True, _
                                             dsFacturas)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Informa las facturas utilizando el webservice de la afip
    ''' </summary>
    ''' <param name="dsFacturas">Dataset con las facturas a informar</param>
    ''' <param name="oAuthentication">Objeto de authenticacion para el webservice de la AFIP</param>
    ''' <param name="strPuestoVenta">Puesto de venta del que se van a informas las facturas</param>
    ''' <param name="intTipoComprobante">Tipo de comprobante que se va a informar
    ''' 1: Factura A
    ''' 2: Nota de Débito A 
    ''' 3: Nota de Crédito A 
    ''' 6: Factura B 
    ''' 7: Nota de Débito B 
    ''' 8: Nota de Crédito B  
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InformarFacturas(ByVal dsFacturas As dsDatos, _
                                ByVal oAuthentication As AuthenticationRequest, _
                                ByVal strPuestoVenta As String, _
                                ByVal intTipoComprobante As Int16) As Boolean

        Dim objFECAEAResponse As wsfe.FECAEAResponse
        Dim FEAuthRequest As wsfe.FEAuthRequest
        Dim objFECAEACabRequest As wsfe.FECAEACabRequest
        Dim objFECAEARequest As wsfe.FECAEARequest
        Dim wsFactura As wsfe.Service


        wsFactura = New wsfe.Service

        Dim indicemax_arrayFECAEADetRequest As Integer = dsFacturas.Tables(TABLA_FACTURAS).Rows.Count - 1 'porque es en base 0

        objFECAEAResponse = New wsfe.FECAEAResponse
        FEAuthRequest = New wsfe.FEAuthRequest
        objFECAEACabRequest = New wsfe.FECAEACabRequest
        objFECAEARequest = New wsfe.FECAEARequest

        Dim d_arrayFECAEADetRequest As Integer = 0
        Dim arrayFECAEADetRequest(indicemax_arrayFECAEADetRequest) As wsfe.FECAEADetRequest

        FEAuthRequest = Inicializar(oAuthentication, CUIT_VERIFICADORA_DEL_NORTE)

        wsFactura.Url = My.Settings.WSFE_URL

        objFECAEACabRequest.CantReg = dsFacturas.Tables(TABLA_FACTURAS).Rows.Count
        objFECAEACabRequest.PtoVta = strPuestoVenta
        objFECAEACabRequest.CbteTipo = intTipoComprobante

        d_arrayFECAEADetRequest = 0
        For Each oRow As DataRow In dsFacturas.Tables(TABLA_FACTURAS).Rows

            Dim objFECAEADetRequest As wsfe.FECAEADetRequest
            objFECAEADetRequest = New wsfe.FECAEADetRequest

            With objFECAEADetRequest
                .Concepto = 2 'Servicios o 3 Prod y Servicios
                .DocTipo = 80 'CUIT
                .DocNro = oRow(10).ToString() 'Nro Cuit
                .CbteDesde = oRow(1).ToString() 'Comprobante desde
                .CbteHasta = oRow(1).ToString() 'Comprobante hasta
                .CbteFch = Left(Format(CDate(oRow(5).ToString()), formatoEscritura), 8) 'Comprobante fecha
                .ImpTotal = oRow(9).ToString() 'Total
                .ImpTotConc = 0
                .ImpNeto = oRow(7).ToString() 'Subtotal 
                .ImpOpEx = 0 'Exento
                .ImpTrib = 0 'Sumatoria de tributos
                .ImpIVA = oRow(8).ToString 'IVA
                .FchServDesde = Left(Format(CDate(oRow(5).ToString()), formatoEscritura), 8)
                .FchServHasta = Left(Format(CDate(oRow(5).ToString()).AddYears(1), formatoEscritura), 8)
                .FchVtoPago = Left(Format(CDate(oRow(5).ToString()).AddDays(30), formatoEscritura), 8) '15 le sumo a la fecha tope de pago
                .MonId = "PES" ' Pesos
                .MonCotiz = 1.0 'Cotizacion
                .CAEA = oRow(2).ToString()
            End With
            arrayFECAEADetRequest(d_arrayFECAEADetRequest) = objFECAEADetRequest

            d_arrayFECAEADetRequest += 1
        Next


        With objFECAEARequest
            .FeCabReq = objFECAEACabRequest
            .FeDetReq = arrayFECAEADetRequest
        End With

        Try
            objFECAEAResponse = wsFactura.FECAEARegInformativo(FEAuthRequest, objFECAEARequest)
            If objFECAEAResponse.Errors IsNot Nothing Then
                For i = 0 To objFECAEAResponse.Errors.Length - 1
                    ErrorLog.Create("InformeFacturas", objFECAEAResponse.Errors(0).Code.ToString(), objFECAEAResponse.Errors(0).Msg.ToString())
                Next
            End If
            If objFECAEAResponse.Events IsNot Nothing Then
                For i = 0 To objFECAEAResponse.Events.Length - 1
                    ErrorLog.Create("InformeFacturasEvento", objFECAEAResponse.Events(i).Code.ToString(), objFECAEAResponse.Events(i).Msg.ToString())
                Next
            End If

            If objFECAEAResponse.Errors Is Nothing Then Return True

            Return False

        Catch ex As Exception
            ErrorLog.Create("InformeFacturas", ex)
            Return False
        End Try

    End Function

    Private Function Inicializar(ByVal oAuthentication As AuthenticationRequest, ByVal strCuit As String) As wsfe.FEAuthRequest
        Try
            Dim oFactura As wsfe.FEAuthRequest
            oFactura = New wsfe.FEAuthRequest

            oFactura.Token = oAuthentication.Token
            oFactura.Sign = oAuthentication.Sign
            oFactura.Cuit = strCuit
            Return oFactura
        Catch ex As Exception
            Throw ex
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Actualiza el campo InformadaEnAFIP =1 en la base de datos para indicar que la factura ya se informo
    ''' </summary>
    ''' <param name="lngIDfactura">ID de la factura a actualizar</param>
    ''' <remarks></remarks>
    Public Sub ActualizarStatusFacturaInformada(ByVal lngIDfactura As Long)

        Try
            conexion.Command.Parameters.Clear()
            conexion.ExecuteSPNonQuery(ACTUALIZAR_FACTURA_INFORMADA, "@idFactura", lngIDfactura)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region


End Class

Public Class FacturaItems
#Region "Propiedades"
    Public Property IDitemFactura As Integer
    Public Property IDfactura As Factura
    Public Property Descripcion As String
    Public Property Cantidad As Integer
    Public Property PrecioUnitario As Decimal
    Public Property Importe As Decimal
    Public Property IncluyeIVA As Boolean
#End Region

#Region "Metodos"
    ''' <summary>
    ''' Destruye el objeto
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose()
        MyBase.Finalize()
    End Sub

    ''' <summary>
    ''' Inicializa el objeto
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        IDfactura = New Factura()
    End Sub

    ''' <summary>
    ''' Inserta los items de una factura en la base
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NuevoItemFactura() As Boolean
        Try
            Dim oParametros As OleDb.OleDbParameterCollection
            Dim oParameter As OleDb.OleDbParameter

            conexion.Command.Parameters.Clear()
            oParametros = conexion.Command.Parameters

            oParameter = New OleDb.OleDbParameter("@idfactura", IDfactura.IDfactura)
            oParametros.Add(oParameter)
            oParameter = Nothing

            oParameter = New OleDb.OleDbParameter("@descripcion", Descripcion)
            oParametros.Add(oParameter)
            oParameter = Nothing

            oParameter = New OleDb.OleDbParameter("@cantidad", Cantidad)
            oParametros.Add(oParameter)
            oParameter = Nothing

            oParameter = New OleDb.OleDbParameter("@preciounitario", PrecioUnitario)
            oParametros.Add(oParameter)
            oParameter = Nothing

            oParameter = New OleDb.OleDbParameter("@importe", Importe)
            oParametros.Add(oParameter)
            oParameter = Nothing

            oParameter = New OleDb.OleDbParameter("@incluyeIva", IncluyeIVA)
            oParametros.Add(oParameter)
            oParameter = Nothing

            'conexion.ExecuteSPNonQuery(INSERTAR_FACTURA_ITEMS)

            oParametros = Nothing


        Catch ex As Exception
            'No se guarda porque hay rollbck y se pierde el catch superior se encarga
            'ErrorLog.Create(UsuarioLogin, ex)
            'MsgBox(MENSAJE_ERROR, MsgBoxStyle.Critical, APPLICATION_NAME)
            Throw ex
        End Try
        Return True
    End Function

    ''' <summary>
    ''' Trae los items de una factura
    ''' </summary>
    '''<param name="intIDFactura">ID de la factura que se quieren obtener los items</param>
    ''' <param name="dsData">Dataset con los datos cargados</param>
    ''' <remarks></remarks>
    Public Sub TraerDatosFacturaItems(ByVal intIDFactura As Integer, ByRef dsData As dsDatos)
        'Try

        '    If dsData Is Nothing Then dsData = New dsDatos

        '    conexion.Command.Parameters.Clear()

        '    conexion.ExecuteDataset(TRAER_DATOS_FACTURA_ITEMS, TABLA_FACTURA_ITEMS, "@idfactura", intIDFactura, True, dsData)

        'Catch ex As Exception
        '    ErrorLog.Create(UsuarioLogin, ex)
        '    MsgBox(MENSAJE_ERROR, MsgBoxStyle.Critical, APPLICATION_NAME)

        'End Try

    End Sub
#End Region


End Class