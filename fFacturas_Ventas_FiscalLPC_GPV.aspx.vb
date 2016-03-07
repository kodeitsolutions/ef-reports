'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFacturas_Ventas_FiscalLPC_GPV"
'-------------------------------------------------------------------------------------------'
Partial Class fFacturas_Ventas_FiscalLPC_GPV
    Inherits vis2formularios.frmReporte

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

        '****************************************************************************
        ' Valida el tipo de salida
        '****************************************************************************
            Dim lcSalida As String = Strings.Trim(Me.Request.QueryString("salida"))
            
            If String.IsNullOrEmpty(lcSalida) Then lcSalida = "" 

            If (lcSalida = "") OrElse (lcSalida.ToLower() <> "pantalla" ) Then 
                
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _
                                          "Este formato imprime una factura fiscal mediante eFactory LPC: solo puede seleccionar el tipo de salida ""Pantalla"".", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
                                           "500px", "250px")

                Return

            End If

        '****************************************************************************
        ' Busca los datos de la factura
        '****************************************************************************
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Facturas.Cod_Cli                                        AS Cod_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Facturas.Nom_Cli = '') ")
            loConsulta.AppendLine("                THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END)    AS Nom_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Facturas.Rif = '') ")
            loConsulta.AppendLine("                THEN Clientes.Rif ELSE Facturas.Rif END)            AS Rif, ")
            loConsulta.AppendLine("            Clientes.Nit                                            AS Nit, ")
            loConsulta.AppendLine("            REPLACE((CASE WHEN (Facturas.Dir_Fis = '') ")
            loConsulta.AppendLine("                THEN Clientes.Dir_Fis ")
            loConsulta.AppendLine("                ELSE Facturas.Dir_Fis END), CHAR(13), '')           AS Dir_Fis, ")
            loConsulta.AppendLine("            (CASE WHEN (Facturas.Nom_Cli = '') ")
            loConsulta.AppendLine("                THEN Clientes.Telefonos ELSE Facturas.Telefonos END)AS Telefonos, ")
            loConsulta.AppendLine("            Clientes.Fax                                            AS Fax, ")
            loConsulta.AppendLine("            Clientes.Generico                                       AS Generico, ")
            loConsulta.AppendLine("            Facturas.Nom_Cli                                        AS Nom_Gen, ")
            loConsulta.AppendLine("            Facturas.Rif                                            AS Rif_Gen, ")
            loConsulta.AppendLine("            Facturas.Nit                                            AS Nit_Gen, ")
            loConsulta.AppendLine("            Facturas.Dir_Fis                                        AS Dir_Gen, ")
            loConsulta.AppendLine("            Facturas.Telefonos                                      AS Tel_Gen, ")
            loConsulta.AppendLine("            Facturas.Documento                                      AS Documento, ")
            loConsulta.AppendLine("            Facturas.Status                                         AS Status, ")
            loConsulta.AppendLine("            Facturas.Fec_Ini                                        AS Fec_Ini, ")
            loConsulta.AppendLine("            Facturas.Fec_Fin                                        AS Fec_Fin, ")
            loConsulta.AppendLine("            Facturas.Mon_Bru                                        AS Mon_Bru, ")
            loConsulta.AppendLine("            Facturas.Mon_Imp1                                       AS Mon_Imp1, ")
            loConsulta.AppendLine("            Facturas.Por_Imp1                                       AS Por_Imp1, ")
            loConsulta.AppendLine("            Facturas.Mon_Net                                        AS Mon_Net, ")
            loConsulta.AppendLine("            Facturas.Por_Des1                                       AS Por_Des, ")
            loConsulta.AppendLine("            Facturas.Dis_Imp                                        AS Dis_Imp, ")
            loConsulta.AppendLine("            Facturas.Mon_Des1                                       AS Mon_Des, ")
            loConsulta.AppendLine("            Facturas.Por_Rec1                                       AS Por_Rec, ")
            loConsulta.AppendLine("            Facturas.Mon_Rec1                                       AS Mon_Rec, ")
            loConsulta.AppendLine("            Facturas.Cod_For                                        AS Cod_For, ")
            loConsulta.AppendLine("            Formas_Pagos.Nom_For                                    AS Nom_For, ")
            loConsulta.AppendLine("            Facturas.Cod_Ven                                        AS Cod_Ven, ")
            loConsulta.AppendLine("            Facturas.Comentario                                     AS Comentario,")
            loConsulta.AppendLine("            Facturas.Fiscal1                                        AS Fiscal1,")
            loConsulta.AppendLine("            Facturas.Fiscal2                                        AS Fiscal2,")
            loConsulta.AppendLine("            Facturas.Fiscal3                                        AS Fiscal3,")
            loConsulta.AppendLine("            Facturas.Fiscal4                                        AS Fiscal4,")
            loConsulta.AppendLine("            Vendedores.Nom_Ven                                      AS Nom_Ven, ")
            loConsulta.AppendLine("            Renglones_Facturas.Cod_Art                              AS Cod_Art, ")
            loConsulta.AppendLine("		       LEFT(RTRIM(Renglones_Facturas.Cod_Art) +  ' - ' + ")
            loConsulta.AppendLine("		       (CASE WHEN Renglones_Facturas.Notas > '' ")        
            loConsulta.AppendLine("		           THEN Renglones_Facturas.Notas ")               
            loConsulta.AppendLine("		           ELSE Articulos.Nom_Art  END), 38)                   AS Nom_Art,")
            loConsulta.AppendLine("		       LEFT(RTRIM(Renglones_Facturas.Cod_Art) +  ' - ' + ")
            loConsulta.AppendLine("		       (CASE WHEN Renglones_Facturas.Notas > '' ")        
            loConsulta.AppendLine("		           THEN Renglones_Facturas.Notas ")               
            loConsulta.AppendLine("		           ELSE Articulos.Nom_Art  END), 38)                   AS Nom_Cor,")
            loConsulta.AppendLine("            Renglones_Facturas.Renglon                              AS Renglon, ")
            loConsulta.AppendLine("            Renglones_Facturas.Can_Art2                             AS Can_Art1, ")
            loConsulta.AppendLine("            Renglones_Facturas.Cod_Uni2                             AS Cod_Uni, ")
            loConsulta.AppendLine("            Renglones_Facturas.Precio1*Renglones_Facturas.Can_Uni2  AS Precio1,")
            loConsulta.AppendLine("            Renglones_Facturas.Mon_Net                              AS Neto, ")
            loConsulta.AppendLine("            Renglones_Facturas.Por_Imp1                             AS Por_Imp, ")
            loConsulta.AppendLine("            Renglones_Facturas.Cod_Imp                              AS Cod_Imp, ")
            loConsulta.AppendLine("            Renglones_Facturas.Mon_Imp1                             AS Impuesto, ")
            loConsulta.AppendLine("            Renglones_Facturas.Comentario                           AS Comentario_Renglon ")
            loConsulta.AppendLine("FROM        Facturas ")
            loConsulta.AppendLine("    JOIN    Renglones_Facturas")
            loConsulta.AppendLine("        ON  Facturas.Documento  =   Renglones_Facturas.Documento")
            loConsulta.AppendLine("    JOIN    Clientes")
            loConsulta.AppendLine("        ON  Facturas.Cod_Cli    =   Clientes.Cod_Cli")
            loConsulta.AppendLine("    JOIN    Formas_Pagos")
            loConsulta.AppendLine("        ON  Facturas.Cod_For    =   Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN    Vendedores ")
            loConsulta.AppendLine("        ON  Facturas.Cod_Ven    =   Vendedores.Cod_Ven")
            loConsulta.AppendLine("    JOIN    Articulos ")
            loConsulta.AppendLine("        ON  Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art")
            loConsulta.AppendLine("WHERE       " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            
            Dim loServicios As New cusDatos.goDatos()
            
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            Dim laTablaFactura As DataTable = laDatosReporte.Tables(0)
            
            If laTablaFactura.Rows.Count = 0 Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "500px", "250px")
                Return 
            End If 

            Dim loFactura As DataRow = laDatosReporte.Tables(0).Rows(0)

            Dim lnDecimalesParaCantidad As Integer  = goOpciones.pnDecimalesParaCantidad
            Dim lnDecimalesParaMonto As Integer  = goOpciones.pnDecimalesParaMonto
            Dim lnDecimalesParaPorcentaje As Integer  = goOpciones.pnDecimalesParaPorcentaje

        '****************************************************************************
        ' Valida que la factura no haya sido impresa anteriormente, que esté confirmada, 
        ' y que tenga monto mayor a cero.
        '****************************************************************************
            Dim lcDocumento As String = CStr(loFactura("Documento")).Trim()
            Dim lcFiscal1 As String = CStr(loFactura("Fiscal1")).Trim()
            Dim lcFiscal2 As String = CStr(loFactura("Fiscal2")).Trim()
            Dim lcFiscal3 As String = CStr(loFactura("Fiscal3")).Trim()
            Dim lcFiscal4 As String = CStr(loFactura("Fiscal4")).Trim()

            If (lcFiscal1 > "") OrElse (lcFiscal2 > "") OrElse (lcFiscal3 > "") OrElse (lcFiscal4 > "") Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _
                                          "La Factura de Venta " & lcDocumento & " ya fue impresa en una impresora fiscal y no puede imprimirse nuevamente: " & _
                                          "<br/>* Serial Impresora: " & lcFiscal1 & _ 
                                          "<br/>* N° Factura Fiscal: " & lcFiscal2 & _ 
                                          "<br/>* Cierre Z: " & lcFiscal3 & _ 
                                          "<br/>* Fecha y hora: " & lcFiscal4 , _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
                                           "500px", "250px")
                Return 
            End If 

            Dim lcEstatus As String = CStr(loFactura("Status")).Trim()
            If (lcEstatus.ToLower() <> "confirmado") AndAlso _
                (lcEstatus.ToLower() <> "afectado")  AndAlso _
                (lcEstatus.ToLower() <> "procesado") Then

                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _
                                          "Solo se puede imprimir una Factura de Venta fiscal si tiene estatus 'Confirmado', 'Afectado' o 'Procesado'. " & _
                                          "La factura " & lcDocumento & " tiene estatus '" & lcEstatus & "'.", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
                                           "500px", "250px")
                Return 
            End If 

            Dim lcMontoNeto As String = CDec(loFactura("Mon_Net"))
            If (lcMontoNeto <= 0D) Then 

                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _
                                          "Solo se puede imprimir una Factura de Venta fiscal si su monto neto es mayor a cero.", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
                                           "500px", "250px")
                Return 

            End If


        '****************************************************************************
        ' Prepara los datos para imprimir la factura
        '****************************************************************************
		    Dim loDatosFactura As New goPuntoVenta.strFactura()
		    

			loDatosFactura.pcDocumento						= CStr(loFactura("Documento")).Trim()
			loDatosFactura.pcCodigoCliente 					= CStr(loFactura("Cod_Cli")).Trim()
			loDatosFactura.pcNombreCliente 					= CStr(loFactura("Nom_Cli")).Trim()
			loDatosFactura.pcRifCliente 					= CStr(loFactura("Rif")).Trim()
			loDatosFactura.pcDireccionCliente				= CStr(loFactura("Dir_Fis")).Trim()
			loDatosFactura.pcTelefonoCliente                = CStr(loFactura("Telefonos")).Trim()
			loDatosFactura.pcCodigoVendedor					= CStr(loFactura("Cod_Ven")).Trim()
			loDatosFactura.pcNombreVendedor					= CStr(loFactura("Nom_Ven")).Trim()
			loDatosFactura.pcCodigoCajero					= goUsuario.pcCodigo
			loDatosFactura.pcNombreCajero					= goUsuario.pcNombre
			loDatosFactura.pcComentario						= CStr(loFactura("Comentario")).Trim()
			loDatosFactura.pnPorRecargo						= goServicios.mRedondearValor(CDec(loFactura("Por_Rec")), lnDecimalesParaPorcentaje, goServicios.enuTipoRedondeo.KN_PuntoMedio)
			loDatosFactura.pnPorDescuento					= goServicios.mRedondearValor(CDec(loFactura("Por_Des")), lnDecimalesParaPorcentaje, goServicios.enuTipoRedondeo.KN_PuntoMedio)
			loDatosFactura.pnMonRecargo						= goServicios.mRedondearValor(CDec(loFactura("Mon_Rec")), lnDecimalesParaMonto, goServicios.enuTipoRedondeo.KN_PuntoMedio)
			loDatosFactura.pnMonDescuento					= goServicios.mRedondearValor(CDec(loFactura("Mon_Des")), lnDecimalesParaMonto, goServicios.enuTipoRedondeo.KN_PuntoMedio)
			loDatosFactura.pnSaldoPendiente					= 0D
			loDatosFactura.pnTotalFactura					= goServicios.mRedondearValor(CDec(loFactura("Mon_Net")), lnDecimalesParaMonto, goServicios.enuTipoRedondeo.KN_PuntoMedio)

			loDatosFactura.pnCobroEfectivo					= goServicios.mRedondearValor(CDec(loFactura("Mon_Net")), lnDecimalesParaMonto, goServicios.enuTipoRedondeo.KN_PuntoMedio)
			loDatosFactura.pnCobroCheque1					= 0D
			loDatosFactura.pnCobroCheque2					= 0D
			loDatosFactura.pnCobroTarjeta1					= 0D
			loDatosFactura.pnCobroTarjeta2					= 0D
			loDatosFactura.pnCobroTransferencia				= 0D
			loDatosFactura.pnCobroNotaCredito				= 0D
			loDatosFactura.pnCobroTicket					= 0D

			loDatosFactura.pnTipoTarjeta1					= "" 
			loDatosFactura.pnTipoTarjeta2					= ""
			
			loDatosFactura.laRenglones = New Generic.List(Of goPuntoVenta.strRenglonesFactura)
			
			For Each loFila As DataRow In laTablaFactura.Rows
			
				Dim loRenglon As New goPuntoVenta.strRenglonesFactura()
				loRenglon.pcCodigo		= CStr(loFila("cod_art")).Trim()
				loRenglon.pcNombre		= CStr(loFila("nom_art")).Trim()
				loRenglon.pcNombreCorto	= CStr(loFila("nom_cor")).Trim()
				loRenglon.pnCantidad	= goServicios.mRedondearValor(CDec(loFila("can_art1")), lnDecimalesParaCantidad, goServicios.enuTipoRedondeo.KN_PuntoMedio)
				loRenglon.pnPrecio		= goServicios.mRedondearValor(CDec(loFila("precio1")), lnDecimalesParaMonto, goServicios.enuTipoRedondeo.KN_PuntoMedio)
				loRenglon.pnPorImpuesto	= goServicios.mRedondearValor(CDec(loFila("por_imp1")), lnDecimalesParaMonto, goServicios.enuTipoRedondeo.KN_PuntoMedio)
				loRenglon.pcComentario	= CStr(loFila("comentario")).Trim()
				
				loDatosFactura.laRenglones.Add(loRenglon)
				
			Next loFila
            

        '****************************************************************************
        ' Genera el XML para la impresora
        '****************************************************************************
            goPuntoVenta.mObtenerOpcionesIpos()
            If Me.mImprimirFacturaXml(loDatosFactura) Then

                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso Completado", _
                                            "La Factura de Venta " & loDatosFactura.pcDocumento & " fue enviada a eFactory LPC para ser impresa. ", _
                                            vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                            "500px", "250px")

            End If
            
        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
                          "No fue posible imprimir la factura solicitada. Información adicional: <br/>" & loExcepcion.ToString(), _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

''' <summary>
''' Imprime la factura fiscal con los datos indicados. Si la impresión se ejecuta sin errores 
''' devuelve True, en caso contrario devuelve False y un mensaje de error por el parámetro lcMensaje.
''' </summary>
''' <param name="loFactura"></param>
''' <returns></returns>
''' <remarks></remarks>
	Private Function mImprimirFacturaXml(ByVal loFactura As goPuntoVenta.strFactura) As Boolean


            'Genera el encabezado del documento
            Dim loSalidaXml As New System.Xml.XmlDocument()

            Dim loRaiz As System.Xml.XmlElement = loSalidaXml.CreateElement("documento_ipos")
            loSalidaXml.AppendChild(loRaiz)

            Dim loEncabezado As System.Xml.XmlElement = loSalidaXml.CreateElement("encabezado")
            loRaiz.AppendChild(loEncabezado)
            
            Dim loNodo As System.Xml.XmlElement
            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("tipo"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection("FACTURA"))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("documento"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcDocumento))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("tipo_documento"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection("FACT"))

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
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcCodigoCajero))
            
            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("nom_caj"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcNombreCajero))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("cod_ven"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcCodigoVendedor))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("nom_ven"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcNombreVendedor))

            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("comentario"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcComentario))
            
			If ( goPuntoVenta.pcComentarioInicioFacturaFiscal.Length > 0) Then
                loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("adicional"))
                loNodo.AppendChild(loSalidaXml.CreateCDataSection(goPuntoVenta.pcComentarioInicioFacturaFiscal))
			End If
            
	    '-------------------------------------------------------------------------------------------'
	    ' Renglones de Venta.									                                    '
	    '-------------------------------------------------------------------------------------------'	
            Dim loRenglones As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("renglones"))
 			For lnFila As Integer = 0 To loFactura.laRenglones.Count - 1 
				Dim loRenglon As goPuntoVenta.strRenglonesFactura = loFactura.laRenglones(lnFila)

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
		If goPuntoVenta.plIncluirBarrasPieFacturas Then

	        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("barras"))
            loNodo.SetAttribute("imprimir", goPuntoVenta.plIncluirBarrasPieFacturas.ToString().ToLower())
            loNodo.SetAttribute("tipo", "EAN13")
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(loFactura.pcDocumento))
			
		End If

 	'-------------------------------------------------------------------------------------------'
	' Si los medios de pago detallados están ACTIVADOS, entonces aplicar las diferentes formas	'
	' de pago según se requiera.																'
	'-------------------------------------------------------------------------------------------'	
        Dim loFormasPago As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("formas_pago"))
        Dim loFormaPago As System.Xml.XmlElement 
		If	goPuntoVenta.plImprimirMediosDePagoDetallados			AndAlso _
			(loFactura.pnSaldoPendiente < loFactura.pnTotalFactura) Then

 			If (loFactura.pnCobroEfectivo > 0) Then		 											 

                loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
                loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoEfectivo)
                loFormaPago.SetAttribute("tipo", "efectivo")
                loFormaPago.SetAttribute("totalizar", "false")

                loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnCobroEfectivo, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

			End If

			If (loFactura.pnCobroCheque1 > 0D) Then

                loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
                loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoCheque)
                loFormaPago.SetAttribute("tipo", "cheque")
                loFormaPago.SetAttribute("totalizar", "false")

                loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnCobroCheque1, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

			End If

			If (loFactura.pnCobroCheque2 > 0D) Then

                loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
                loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoCheque)
                loFormaPago.SetAttribute("tipo", "cheque")
                loFormaPago.SetAttribute("totalizar", "false")

                loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnCobroCheque2, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

			End If

			If (loFactura.pnCobroTarjeta1 > 0) Then

                loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
				Select Case loFactura.pnTipoTarjeta1 
					Case "DEBITO"
                        loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoTarjetaDebito)
                        loFormaPago.SetAttribute("tipo", "debito")
					Case "CREDITO"
                        loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoTarjetaCredito)
                        loFormaPago.SetAttribute("tipo", "credito")
					Case Else
                        loFormaPago.SetAttribute("forma", "?")
                        loFormaPago.SetAttribute("tipo", "?")
				End Select
                loFormaPago.SetAttribute("totalizar", "false")
                
                loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnCobroTarjeta1, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

		   End If
		   
			
		   If (loFactura.pnCobroTarjeta2 > 0) Then

                loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
				Select Case loFactura.pnTipoTarjeta2
					Case "DEBITO"
                        loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoTarjetaDebito)
                        loFormaPago.SetAttribute("tipo", "debito")
					Case "CREDITO"
                        loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoTarjetaCredito)
                        loFormaPago.SetAttribute("tipo", "credito")
					Case Else
                        loFormaPago.SetAttribute("forma", "?")
                        loFormaPago.SetAttribute("tipo", "?")
				End Select
                loFormaPago.SetAttribute("totalizar", "false")
                
                loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnCobroTarjeta2, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

			End If
		   
			If (loFactura.pnCobroTransferencia > 0) Then

                loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
                loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoDeposito)
                loFormaPago.SetAttribute("tipo", "transferencia")
                loFormaPago.SetAttribute("totalizar", "false")

                loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnCobroTransferencia, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

			End If
		   
			If (loFactura.pnCobroTicket > 0D) Then

                loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
                loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoTickets)
                loFormaPago.SetAttribute("tipo", "ticket")
                loFormaPago.SetAttribute("totalizar", "false")

                loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnCobroTicket, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

			End If
				
			If (loFactura.pnCobroNotaCredito > 0) Then

                loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
                loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoNotaDeCredito)
                loFormaPago.SetAttribute("tipo", "n/cr")
                loFormaPago.SetAttribute("totalizar", "false")

                loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnCobroNotaCredito, _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

			End If

 	'-------------------------------------------------------------------------------------------'
	' Si los medios de pago detallados están DESACTIVADOS, entonces aplicar un único pago en	'
	' efectivo para cerrar la factura.															'
	'-------------------------------------------------------------------------------------------'	
		ElseIf	Not goPuntoVenta.plImprimirMediosDePagoDetallados	AndAlso _
			(loFactura.pnSaldoPendiente < loFactura.pnTotalFactura) Then

	        
            loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
            loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoEfectivo)
            loFormaPago.SetAttribute("tipo", "efectivo")
            loFormaPago.SetAttribute("totalizar", "true")

            loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnTotalFactura, _
                                                        goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

	'-------------------------------------------------------------------------------------------'
	' Si la factura va a quedar sin pagar (Pendiente), entonces cerrar la factura pagando		'
	' completo en Otro Medio de Pago.															'
	'-------------------------------------------------------------------------------------------'	
		Else 'loFactura.pnSaldoPendiente >= loFactura.pnTotalFactura
						
            loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
            loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoOtro)
            loFormaPago.SetAttribute("tipo", "otro")
            loFormaPago.SetAttribute("totalizar", "true")

            loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(loFactura.pnTotalFactura, _
                                                        goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

		End If
		
        'Añade una sección de datos vacia
        Dim loDatos As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("datos"))

        'Añade la sección de parámetros 
        Dim loParametros As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("parametros"))
        Dim loParametro As System.Xml.XmlElement

        loParametro = loParametros.AppendChild(loSalidaXml.CreateElement("parametro"))
        loParametro.SetAttribute("nombre", "plIncluirVendedorAlImprimir")
        loParametro.AppendChild(loSalidaXml.CreateCDataSection(CStr(goPuntoVenta.plIncluirVendedorAlImprimir).ToLower()))
            
        loParametro = loParametros.AppendChild(loSalidaXml.CreateElement("parametro"))
        loParametro.SetAttribute("nombre", "plIncluirCajeroAlImprimir")
        loParametro.AppendChild(loSalidaXml.CreateCDataSection(CStr(goPuntoVenta.plIncluirCajeroAlImprimir).ToLower()))

        loParametro = loParametros.AppendChild(loSalidaXml.CreateElement("parametro"))
        loParametro.SetAttribute("nombre", "pcComentarioInicioFacturaFiscal")
        loParametro.AppendChild(loSalidaXml.CreateCDataSection(CStr(goPuntoVenta.pcComentarioInicioFacturaFiscal.Trim())))

        loParametro = loParametros.AppendChild(loSalidaXml.CreateElement("parametro"))
        loParametro.SetAttribute("nombre", "plIncluirBarrasPieFacturas")
        loParametro.AppendChild(loSalidaXml.CreateCDataSection(CStr(goPuntoVenta.plIncluirBarrasPieFacturas).ToLower()))

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

        Dim lcNombreArchivo As String = "ipos_" & goEmpresa.pcCodigo & "_Factura_" & loFactura.pcDocumento & "_" & (Date.now()).ToString("yyyyMMddhhmm") & ".xml"
        lcNombreArchivo = Regex.Replace(lcNombreArchivo, "[\\\/*? ]", "-")
        
        Dim loScript As New StringBuilder()

        loScript.Append("(function(){")
        loScript.Append("var lcContenido = '")
        loScript.Append(lcSalida)
        loScript.Append("'; var loArchivo = new Blob([lcContenido], { type: 'application/xml' }); ")
        loScript.Append("window.mDescargarBlob(loArchivo, '" & lcNombreArchivo & "')")
        loScript.Append("})();")

        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "descargardocumentoXml", loScript.ToString(), True)
        
		Return True
		
	End Function




    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 15/10/15: Codigo inicial, a partir de fFacturas_Ventas_FiscalLPC.                    '
'-------------------------------------------------------------------------------------------'
