'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fNEntrega_NoFiscal_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fNEntrega_NoFiscal_IKP
    Inherits vis2formularios.frmReporte

#Region "Declaraciones"

	''' <summary>
	''' Representa una factura a ser impresa por impresora fiscal.
	''' </summary>
	''' <remarks></remarks>
	Private Structure strFactura
	
		Dim pcDocumento				As string
		Dim pcCodigoCliente 		As string
		Dim pcNombreCliente 		As string
		Dim pcRifCliente 			As string
		Dim pcDireccionCliente 		As string
		Dim pcTelefonoCliente 		As string
		Dim pcCodigoVendedor		As string
		Dim pcNombreVendedor		As string
		Dim pcCodigoCajero		    As string
		Dim pcNombreCajero		    As string
		Dim pcComentario			As string
		Dim pnPorRecargo			As Decimal
		Dim pnPorDescuento			As Decimal
		Dim pnMonRecargo			As Decimal
		Dim pnMonDescuento			As Decimal
		Dim pnSaldoPendiente		As Decimal
		Dim pnTotalFactura			As Decimal

		Dim pnCobroEfectivo			As Decimal
		Dim pnCobroCheque1			As Decimal
		Dim pnCobroCheque2			As Decimal
		Dim pnCobroTarjeta1			As Decimal
		Dim pnCobroTarjeta2			As Decimal
		Dim pnCobroTransferencia	As Decimal
		Dim pnCobroNotaCredito		As Decimal
		Dim pnCobroTicket			As Decimal
		
		Dim pnTipoTarjeta1			As String
		Dim pnTipoTarjeta2			As String
		
        Dim paDatos                 As Generic.Dictionary(Of String, Object)
		Dim laRenglones As Generic.List(Of strRenglonesFactura)
		
	End Structure

	''' <summary>
	''' Representa un renglón de una variable tipo strFactura.
	''' </summary>
	''' <remarks></remarks>
	PRivate Structure strRenglonesFactura
	
		Dim pcCodigo			As string
		Dim pcNombre			As string
		Dim pcNombreCorto		As string
		Dim pnCantidad			As Decimal
		Dim pnPrecio			As Decimal
		Dim pnPorImpuesto		As Decimal
		Dim pcComentario		As string
		
	End Structure

#End Region

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT      Entregas.Cod_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Entregas.Nom_Cli = '') ")
            loConsulta.AppendLine("                THEN Clientes.Nom_Cli ")
            loConsulta.AppendLine("                ELSE Entregas.Nom_Cli END)      AS  Nom_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Entregas.Rif = '') ")
            loConsulta.AppendLine("                THEN Clientes.Rif ")
            loConsulta.AppendLine("                ELSE Entregas.Rif END)          AS  Rif, ")
            loConsulta.AppendLine("            (CASE WHEN (Entregas.Dir_Fis = '') ")
            loConsulta.AppendLine("                THEN Clientes.Dir_Fis ")
            loConsulta.AppendLine("                ELSE Entregas.Dir_Fis END)      AS  Dir_Fis, ")
            loConsulta.AppendLine("            (CASE WHEN (Entregas.Telefonos = '') ")
            loConsulta.AppendLine("                THEN Clientes.Telefonos ")
            loConsulta.AppendLine("                ELSE Entregas.Telefonos END)    AS Telefonos,")
            loConsulta.AppendLine("            Entregas.Documento, ")
            loConsulta.AppendLine("            Entregas.Fec_Ini, ")
            loConsulta.AppendLine("            Entregas.Cod_For, ")
            loConsulta.AppendLine("            Entregas.Mon_Rec1, ")
            loConsulta.AppendLine("            Entregas.Por_Rec1, ")
            loConsulta.AppendLine("            Entregas.Mon_Des1, ")
            loConsulta.AppendLine("            Entregas.Por_Des1, ")
            loConsulta.AppendLine("            Entregas.Mon_Net, ")
            loConsulta.AppendLine("            Entregas.Cod_Ven, ")
            loConsulta.AppendLine("            Entregas.Comentario, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("            Renglones_Entregas.Cod_Art, ")
            loConsulta.AppendLine("            CASE")
            loConsulta.AppendLine("                WHEN Renglones_Entregas.Notas <> Articulos.Nom_Art THEN Renglones_Entregas.Notas")
            loConsulta.AppendLine("                ELSE (CASE WHEN articulos.Nom_Cor>'' THEN articulos.Nom_Cor ELSE articulos.Nom_Art END)")
            loConsulta.AppendLine("            END		                    												AS Nom_Art,  ")
            loConsulta.AppendLine("            Renglones_Entregas.Renglon, ")
            loConsulta.AppendLine("            (CASE WHEN (Renglones_Entregas.Cod_Uni = Renglones_Entregas.Cod_Uni2) ")
            loConsulta.AppendLine("                THEN Renglones_Entregas.Can_Art1 ")
            loConsulta.AppendLine("                ELSE Renglones_Entregas.Can_Art2 ")
            loConsulta.AppendLine("            END)                                                                        AS Can_Art1,")
            loConsulta.AppendLine("            (CASE WHEN (Renglones_Entregas.Cod_Uni = Renglones_Entregas.Cod_Uni2) ")
            loConsulta.AppendLine("                THEN Renglones_Entregas.Cod_Uni ")
            loConsulta.AppendLine("                ELSE Renglones_Entregas.Cod_Uni2 ")
            loConsulta.AppendLine("            END)                                                                        AS Cod_Uni, ")
            loConsulta.AppendLine("            (CASE WHEN (Renglones_Entregas.Cod_Uni = Renglones_Entregas.Cod_Uni2) ")
            loConsulta.AppendLine("                THEN Renglones_Entregas.Precio1 ")
            loConsulta.AppendLine("                ELSE Renglones_Entregas.Precio1*Renglones_Entregas.Can_Uni2 ")
            loConsulta.AppendLine("            END)*(1-Renglones_Entregas.Por_Des/100)                                     AS Precio1, ")
            loConsulta.AppendLine("            Renglones_Entregas.Mon_Net          AS  Neto, ")
            loConsulta.AppendLine("            Renglones_Entregas.Por_Imp1         AS  Por_Imp, ")
            loConsulta.AppendLine("            Renglones_Entregas.Cod_Imp, ")
            loConsulta.AppendLine("            Renglones_Entregas.Por_Des, ")
            loConsulta.AppendLine("            Renglones_Entregas.Mon_Imp1         AS  Impuesto, ")
            loConsulta.AppendLine("            Renglones_Entregas.Comentario       AS Comentario_Renglon ")
            loConsulta.AppendLine("FROM        Entregas ")
            loConsulta.AppendLine("    JOIN    Renglones_Entregas ON Renglones_Entregas.Documento = Entregas.Documento")
            loConsulta.AppendLine("    JOIN    Clientes ON Clientes.Cod_Cli = Entregas.Cod_Cli")
            loConsulta.AppendLine("    JOIN    Formas_Pagos ON Formas_Pagos.Cod_For = Entregas.Cod_For")
            loConsulta.AppendLine("    JOIN    Vendedores ON Vendedores.Cod_Ven = Entregas.Cod_Ven")
            loConsulta.AppendLine("    JOIN    Articulos ON Articulos.Cod_Art = Renglones_Entregas.Cod_Art")
            loConsulta.AppendLine("WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
			
			'-------------------------------------------------------------------------------------------------------
            ' Verifica si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
						vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
                Return 
            End If
            
            Dim loDatosFactura As New strFactura()
            Dim loFila As DataRow = laDatosReporte.Tables(0).Rows(0)
            
            loDatosFactura.pcDocumento          = CStr(loFila("Documento")).Trim()
            loDatosFactura.pcCodigoCliente 		= CStr(loFila("Cod_Cli")).Trim()
            loDatosFactura.pcNombreCliente 		= CStr(loFila("Nom_Cli")).Trim()
            loDatosFactura.pcRifCliente 		= CStr(loFila("Rif")).Trim()
            loDatosFactura.pcDireccionCliente 	= CStr(loFila("Dir_Fis")).Trim()
            loDatosFactura.pcTelefonoCliente 	= CStr(loFila("Telefonos")).Trim()
            loDatosFactura.pcCodigoVendedor		= CStr(loFila("Cod_Ven")).Trim()
            loDatosFactura.pcNombreVendedor		= CStr(loFila("Nom_Ven")).Trim()
            loDatosFactura.pcCodigoCajero		= goUsuario.pcCodigo.Trim()
            loDatosFactura.pcNombreCajero		= goUsuario.pcNombre.Trim()
            loDatosFactura.pcComentario			= CStr(loFila("Comentario")).Trim()
            loDatosFactura.pnPorRecargo			= CDec(loFila("Por_Rec1"))
            loDatosFactura.pnPorDescuento		= CDec(loFila("Por_Des1"))
            loDatosFactura.pnMonRecargo			= CDec(loFila("Mon_Rec1"))
            loDatosFactura.pnMonDescuento		= CDec(loFila("Mon_Des1"))
            loDatosFactura.pnSaldoPendiente		= 0D
            loDatosFactura.pnTotalFactura		= CDec(loFila("Mon_Net"))
            loDatosFactura.pnCobroEfectivo      = CDec(loFila("Mon_Net"))
            
            loDatosFactura.paDatos = New Generic.Dictionary(Of String, Object)
            loDatosFactura.paDatos.Add("pdFechaDocumento", CDate(loFila("Fec_Ini")))

            loDatosFactura.laRenglones = New Generic.List(Of strRenglonesFactura)
            For Each loFila In laDatosReporte.Tables(0).Rows
                Dim loRenglon As New strRenglonesFactura()

                loRenglon.pcCodigo		= CStr(loFila("Cod_Art")).Trim()
                loRenglon.pcNombre		= CStr(loFila("Nom_Art")).Trim()
                loRenglon.pcNombreCorto	= CStr(loFila("Nom_Art")).Trim()
                loRenglon.pnCantidad	= CDec(loFila("Can_Art1"))
                loRenglon.pnPrecio		= CDec(loFila("Precio1"))
                loRenglon.pnPorImpuesto	= CDec(loFila("Por_Imp"))
                loRenglon.pcComentario	= CStr(loFila("Comentario_Renglon")).Trim()

                loDatosFactura.laRenglones.Add(loRenglon)

            Next loFila

			Dim llFacturaImpresa As Boolean = Me.mImprimirFacturaXml(loDatosFactura)

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

        End Try

    End Sub

''' <summary>
''' Imprime la Nota de Entrega no Fiscal (remota) con los datos indicados. Si la impresión se ejecuta sin errores 
''' devuelve True, en caso contrario devuelve False y un mensaje de error por el parámetro lcMensaje.
''' </summary>
''' <param name="loFactura"></param>
''' <returns></returns>
''' <remarks></remarks>
	Private Function mImprimirFacturaXml(ByVal loFactura As strFactura) As Boolean

        Dim lcMedioPagoEfectivo             As String   = CStr(goUsuario.mObtenerOpcionGlobal("MEDPEFIPOS")).Trim()
        Dim llIncluirVendedorAlImprimir     As Boolean  = CBool(goUsuario.mObtenerOpcionGlobal("INCVENIPOS"))
        Dim llIncluirCajeroAlImprimir       As Boolean  = CBool(goUsuario.mObtenerOpcionGlobal("INCCAJIPOS"))
        Dim lcComentarioInicioFacturaFiscal As String   = CStr(goUsuario.mObtenerOpcionGlobal("COMIFFIPOS")).Trim()
        Dim llIncluirBarrasPieFacturas		As Boolean  = CBool(goUsuario.mObtenerOpcionGlobal("ACTIBFIPOS"))


            'Genera el encabezado del documento
            Dim loSalidaXml As New System.Xml.XmlDocument()

            Dim loRaiz As System.Xml.XmlElement = loSalidaXml.CreateElement("documento_ipos")
            loSalidaXml.AppendChild(loRaiz)

            Dim loEncabezado As System.Xml.XmlElement = loSalidaXml.CreateElement("encabezado")
            loRaiz.AppendChild(loEncabezado)
            
            Dim loNodo As System.Xml.XmlElement
            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("tipo"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection("NOTAENTREGA"))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("documento"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcDocumento))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("tipo_documento"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(""))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("cod_cli"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcCodigoCliente))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("nom_cli"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcNombreCliente))
        
            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("rif"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcRifCliente))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("direccion"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcDireccionCliente))
            
            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("telefono"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcTelefonoCliente))
            
            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("cod_caj"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(goUsuario.pcCodigo))
            
            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("nom_caj"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(goUsuario.pcNombre))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("cod_ven"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcCodigoVendedor))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("nom_ven"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcNombreVendedor))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("comentario"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcComentario))
            
			If (lcComentarioInicioFacturaFiscal.Length > 0) Then
                loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("adicional"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcComentarioInicioFacturaFiscal))
			End If
            
	    '-------------------------------------------------------------------------------------------'
	    ' Renglones de Venta.									                                    '
	    '-------------------------------------------------------------------------------------------'	
            Dim loRenglones As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("renglones"))
 			For lnFila As Integer = 0 To loFactura.laRenglones.Count - 1 
				Dim loRenglon As strRenglonesFactura = loFactura.laRenglones(lnFila)

                Dim loRenglonxml As System.Xml.XmlElement = loRenglones.AppendChild(loSalidaXml.CreateElement("renglon"))

                loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("numero"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(lnFila+1, _
                                                                        goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 0)))

                loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("cod_art"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(loRenglon.pcCodigo))
                
                loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("nom_art"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(loRenglon.pcNombre))

                loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("nom_cor"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(loRenglon.pcNombreCorto))

                loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("can_art"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loRenglon.pnCantidad, _
                                                                        goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

                loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("precio"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loRenglon.pnPrecio, _
                                                                        goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

                loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("cod_imp"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection("x"))

                loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("por_imp"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loRenglon.pnPorImpuesto, _
                                                                        goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

                loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("comentario"))
                loNodo.SetAttribute("antes_del_articulo", "true")
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(loRenglon.pcComentario))

			Next lnFila
		
		
	'-------------------------------------------------------------------------------------------'
	' Si hay descuentos o recargos, entonces aplicarlos ahora.									'
	'-------------------------------------------------------------------------------------------'	
        Dim loDescuento As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("descuento"))
        Dim loRecargo As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("recargo"))


        Dim lcPorcentaje As String = ""
        Dim lcMonto As String = ""
        
        lcPorcentaje = goServicios.mObtenerFormatoCadenaCSV(loFactura.pnPorDescuento, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)
        lcMonto = goServicios.mObtenerFormatoCadenaCSV(loFactura.pnMonDescuento, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)
            
        loDescuento.SetAttribute("porcentaje", lcPorcentaje)
        loDescuento.SetAttribute("monto", lcMonto)
        loDescuento.SetAttribute("global", "true")


        lcPorcentaje = goServicios.mObtenerFormatoCadenaCSV(loFactura.pnPorRecargo, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)
        lcMonto = goServicios.mObtenerFormatoCadenaCSV(loFactura.pnMonRecargo, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)
		
        loRecargo.SetAttribute("porcentaje", lcPorcentaje)
        loRecargo.SetAttribute("monto", lcMonto)
        loRecargo.SetAttribute("global", "true")
        
	'-------------------------------------------------------------------------------------------'
	' Comentario para el subtotal.                                              				'
	'-------------------------------------------------------------------------------------------'	
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("comentario_subtotal"))
        loNodo.SetAttribute("antes_de_subtotal", "false")
        loNodo.AppendChild(loSalidaXml.CreateCDataSection("DESTINO: FACT-" & loFactura.pcDocumento))
            			
		
	'-------------------------------------------------------------------------------------------'
	' Si está activa la impresión del código de barras, entonces lo imprime ahora.				'
	'-------------------------------------------------------------------------------------------'	
		If llIncluirBarrasPieFacturas Then

	        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("barras"))
            loNodo.SetAttribute("imprimir", llIncluirBarrasPieFacturas.ToString().ToLower())
            loNodo.SetAttribute("tipo", "EAN13")
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcDocumento))
			
		End If

 	'-------------------------------------------------------------------------------------------'
	' Si los medios de pago detallados están ACTIVADOS, entonces aplicar las diferentes formas	'
	' de pago según se requiera.																'
	'-------------------------------------------------------------------------------------------'	
        Dim loFormasPago As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("formas_pago"))
        Dim loFormaPago As System.Xml.XmlElement     

 	'-------------------------------------------------------------------------------------------'
	' Se aplica un único pago en efectivo para cerrar el documento.								'
	'-------------------------------------------------------------------------------------------'	
        loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
        loFormaPago.SetAttribute("forma", lcMedioPagoEfectivo)
        loFormaPago.SetAttribute("tipo", "efectivo")
        loFormaPago.SetAttribute("totalizar", "true")

        loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnTotalFactura, _
                                                    goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

		
        'Añade una sección de datos vacia
        Dim loDatos As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("datos"))
        Dim loDato As System.Xml.XmlElement

        loDato = loDatos.AppendChild(loSalidaXml.CreateElement("dato"))
        loDato.SetAttribute("nombre", "pdFechaDocumento")
        loDato.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(CDate(loFactura.paDatos("pdFechaDocumento"))).ToLower()))
            

        'Añade la sección de parámetros 
        Dim loParametros As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("parametros"))
        Dim loParametro As System.Xml.XmlElement

        loParametro = loParametros.AppendChild(loSalidaXml.CreateElement("parametro"))
        loParametro.SetAttribute("nombre", "plIncluirVendedorAlImprimir")
        loParametro.AppendChild(loSalidaXml.CreateCDataSection(CStr(llIncluirVendedorAlImprimir).ToLower()))
            
        loParametro = loParametros.AppendChild(loSalidaXml.CreateElement("parametro"))
        loParametro.SetAttribute("nombre", "plIncluirCajeroAlImprimir")
        loParametro.AppendChild(loSalidaXml.CreateCDataSection(CStr(llIncluirCajeroAlImprimir).ToLower()))

        loParametro = loParametros.AppendChild(loSalidaXml.CreateElement("parametro"))
        loParametro.SetAttribute("nombre", "pcComentarioInicioFacturaFiscal")
        loParametro.AppendChild(loSalidaXml.CreateCDataSection(lcComentarioInicioFacturaFiscal))

        loParametro = loParametros.AppendChild(loSalidaXml.CreateElement("parametro"))
        loParametro.SetAttribute("nombre", "plIncluirBarrasPieFacturas")
        loParametro.AppendChild(loSalidaXml.CreateCDataSection(CStr(llIncluirBarrasPieFacturas).ToLower()))

        Dim laDatosOrigen As New Generic.Dictionary(Of String, Object)
        laDatosOrigen.Add("lcCliente", goCliente.pcCodigo)
        laDatosOrigen.Add("lcEmpresa", goEmpresa.pcCodigo)
        laDatosOrigen.Add("lcSucursal", goSucursal.pcCodigo)
        laDatosOrigen.Add("lcUsuario", goUsuario.pcCodigo)
        Dim loDireccion As New Uri(Me.Request.Url, Me.ResolveClientUrl("../../iPos/Formularios/wbsServicioFiscalIPOS.asmx"))
        laDatosOrigen.Add("lcDireccion", loDireccion.AbsoluteUri)

        loParametro = loParametros.AppendChild(loSalidaXml.CreateElement("parametro"))
        loParametro.SetAttribute("nombre", "pcDatosDeOrigen")
        loParametro.AppendChild(loSalidaXml.CreateCDataSection(cusSeguridad.goSeguridad.mEncriptarDiccionario("eFactoryLPC", laDatosOrigen)))

   	'-------------------------------------------------------------------------------------------'
	' Descarga el archivo.																		'
	'-------------------------------------------------------------------------------------------'	
        Dim lcSalida As String = loSalidaXml.OuterXml.Replace("\","\\").Replace("'","\'")
        lcSalida = Regex.Replace(lcSalida, "[\r\n]+", "\n")
        
        Dim lcNombreArchivo As String = "ipos_" & goEmpresa.pcCodigo & "_NotaEntrega_" & loFactura.pcDocumento & "_" & (Date.now()).ToString("yyyyMMddhhmm") & ".xml"
        lcNombreArchivo = Regex.Replace(lcNombreArchivo, "[\\\/*? ]", "-")
        
        Dim loScript As New StringBuilder()

        loScript.Append("(function(){")
        loScript.Append("var lcContenido = '")
        loScript.Append(lcSalida)
        loScript.Append("'; var loArchivo = new Blob([lcContenido], { type: 'application/xml' }); ")
        'loScript.Append("window.top.mDescargarBlob(loArchivo, '" & lcNombreArchivo & "')")
        loScript.Append("")
        loScript.Append("var loLector = new FileReader();")
        loScript.Append("loLector.readAsDataURL(loArchivo);")
        loScript.Append("loLector.onload = function (event) {")
        loScript.Append("    var loGuardar = document.createElement('a');")
        loScript.Append("    loGuardar.href = event.target.result;")
        loScript.Append("    loGuardar.target = '_blank';")
        loScript.Append("    loGuardar.download = '" & lcNombreArchivo & "';")
        loScript.Append("")
        loScript.Append("    var clicEvent = new MouseEvent('click', {")
        loScript.Append("        'view': window,")
        loScript.Append("        'bubbles': true,")
        loScript.Append("        'cancelable': true")
        loScript.Append("    });")
        loScript.Append("    loGuardar.dispatchEvent(clicEvent);")
        loScript.Append("    (window.URL || window.webkitURL).revokeObjectURL(loGuardar.href);")
        loScript.Append("};")
        loScript.Append("")
        loScript.Append("})();")

        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "descargardocumentoXml", loScript.ToString(), True)

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso Completado", _
                          "Se generó un archivo XML a partir de la Nota de Entrega para ser impreso con eFactory PLC.", _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                           "auto", _
                           "auto")
       
		Return True
		
	End Function

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 23/09/13: Codigo inicial.																'
'-------------------------------------------------------------------------------------------'
' RJG: 24/09/13: Se añadió la fecha al XML generado. 										'
'-------------------------------------------------------------------------------------------'
