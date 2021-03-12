Module Module1

    Sub Main()
        Try
            GenerarConfiguracionInicial()
            Dim oAuthentication As AuthenticationRequest

            oAuthentication = New AuthenticationRequest

            Console.WriteLine("Obteniendo authentication...")
            'paso1:  Obtener token y sign de authentication y fecha vencimiento
            oAuthentication.ObtenerAuthenticationRequest(WEB_SERVICE_FACTURA_ELECTRONICA, oAuthentication)

            If oAuthentication.Token.Length < 2 Then
                Console.WriteLine("Authentication inexistente.")
                Console.WriteLine("Generando nueva Authentication...")
                'Paso1.1: obtener Webserice y certificados de authentication
                Dim oParametros As CertificadoParametro
                oParametros = New CertificadoParametro
                oParametros.ObtenerParametrosCertificado(WEB_SERVICE_FACTURA_ELECTRONICA, oParametros)
                oAuthentication.GenerarTokenYsignature(WEB_SERVICE_FACTURA_ELECTRONICA, oParametros, oAuthentication)
                oParametros.Dispose()
                oAuthentication.InsertarAuthenticationRequest()
                Console.WriteLine("Authentication generado.")
            End If
            'SE INFORMA PRIMERO LAS FACTURAS A 
            InformarFacturas(oAuthentication, TiposDeComprobantes.Factura_A)
            'Despues las B
            InformarFacturas(oAuthentication, TiposDeComprobantes.Factura_B)

            oAuthentication.Dispose()
        Catch ex As Exception
            ErrorLog.Create("InformeFacturas", ex)
            Console.WriteLine(MENSAJE_ERROR & vbCrLf & ex.Message)
        End Try

    End Sub

    Private Sub InformarFacturas(ByVal oAuthentication As AuthenticationRequest, ByVal intComprobanteAInformar As Int16)
        Try
            Dim oFactura As Factura
            Dim dsFacturas As dsDatos

            oFactura = New Factura
            dsFacturas = New dsDatos


            Console.WriteLine("Obteniendo facturas desde la base de datos....")
            oFactura.ListarFacturasSinInformar(dsFacturas, intComprobanteAInformar)

            If dsFacturas.Tables(TABLA_FACTURAS).Rows.Count > 0 Then
                Console.WriteLine("Facturas obtenidas.")

                Dim strPuestoVenta As String = "003"
                strPuestoVenta = dsFacturas.Tables(TABLA_FACTURAS).Rows(0).Item(3).ToString()

                Console.WriteLine("Informando facturas.....")


                If oFactura.InformarFacturas(dsFacturas, oAuthentication, strPuestoVenta, 1) Then

                    For Each oRow As DataRow In dsFacturas.Tables(TABLA_FACTURAS).Rows
                        Console.WriteLine("Actualizando status factura N°: " & oRow(1).ToString())
                        oFactura.ActualizarStatusFacturaInformada(oRow(0)) 'ID de factura
                    Next
                    Console.WriteLine("Proceso finalizado!!!")
                Else

                    Console.WriteLine("El proceso de informacion de facturas ha fallado.")
                End If

            Else
            Console.WriteLine("No se han encontrado facturas para informar")
            End If

            dsFacturas.Dispose()
            oFactura.Dispose()
        Catch ex As Exception
            Throw ex
        End Try


    End Sub

    Private Sub GenerarConfiguracionInicial()
        Try
            Servidor = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Verificadora").GetValue("Servidor")
            BaseDatos = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Verificadora").GetValue("base")
            UserBD = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Verificadora").GetValue("usuario")
            Password = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Verificadora").GetValue("password")
            Password = Desencriptar(Password)
            conexion = New clsSQLDataManagement(clsSQLDataManagement.Providers.SQL, Servidor, BaseDatos, UserBD, Password)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
End Module
