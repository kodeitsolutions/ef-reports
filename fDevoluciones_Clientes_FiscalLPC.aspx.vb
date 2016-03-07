'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDevoluciones_Clientes_FiscalLPC"
'-------------------------------------------------------------------------------------------'
Partial Class fDevoluciones_Clientes_FiscalLPC

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT	Devoluciones_Clientes.Cod_Cli         AS Cod_Cli, ")
            loConsulta.AppendLine("         Clientes.Nom_Cli                      AS Nom_Cli, ")
            loConsulta.AppendLine("         Clientes.Rif                          AS Rif, ")
            loConsulta.AppendLine("         Clientes.Nit                          AS Nit, ")
            loConsulta.AppendLine("         Clientes.Dir_Fis                      AS Dir_Fis, ")
            loConsulta.AppendLine("         Clientes.Telefonos                    AS Telefonos, ")
            loConsulta.AppendLine("         Clientes.Fax                          AS Fax, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Nom_Cli         AS Nom_Gen, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Rif             AS Rif_Gen, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Nit             AS Nit_Gen, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Dir_Fis         AS Dir_Gen, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Telefonos       AS Tel_Gen, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Documento       AS Documento, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Doc_Des1        AS NCR, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Status          AS Status, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Fec_Ini         AS Fec_Ini, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Fec_Fin         AS Fec_Fin, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Mon_Bru         AS Mon_Bru, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Mon_Imp1        AS Mon_Imp1, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Mon_Net         AS Mon_Net, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Cod_For         AS Cod_For, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Por_Des1        AS Por_Des, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Mon_Des1        AS Mon_Des, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Por_Rec1        AS Por_Rec, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Mon_Rec1        AS Mon_Rec, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Dis_Imp         AS Dis_Imp, ")
            loConsulta.AppendLine("         Formas_Pagos.Nom_For                  AS Nom_For,")
            loConsulta.AppendLine("         Devoluciones_Clientes.Cod_Ven         AS Cod_Ven, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Comentario      AS Comentario, ")
            loConsulta.AppendLine("         Devoluciones_Clientes.Fiscal1         AS Fiscal1,")
            loConsulta.AppendLine("         Devoluciones_Clientes.Fiscal2         AS Fiscal2,")
            loConsulta.AppendLine("         Devoluciones_Clientes.Fiscal3         AS Fiscal3,")
            loConsulta.AppendLine("         Devoluciones_Clientes.Fiscal4         AS Fiscal4,")
            loConsulta.AppendLine("         Vendedores.Nom_Ven                    AS Nom_Ven, ")
            loConsulta.AppendLine("         Renglones_DClientes.Cod_Art           AS Cod_Art, ")
            loConsulta.AppendLine("         Articulos.Nom_Art                     AS Nom_Art, ")
            loConsulta.AppendLine("         (CASE WHEN Articulos.Nom_Cor > ''")
            loConsulta.AppendLine("                THEN Articulos.Nom_Cor   ")
            loConsulta.AppendLine("                ELSE Articulos.Nom_Art END)    AS Nom_Cor, ")
            loConsulta.AppendLine("         Renglones_DClientes.Renglon           AS Renglon,  ")
            loConsulta.AppendLine("         Renglones_DClientes.Can_Art1          AS Can_Art1, ")
            loConsulta.AppendLine("         Renglones_DClientes.Cod_Uni           AS Cod_Uni,  ")
            loConsulta.AppendLine("         Renglones_DClientes.Precio1           AS Precio1,  ")
            loConsulta.AppendLine("         Renglones_DClientes.Mon_Net           AS Neto, ")
            loConsulta.AppendLine("         Renglones_DClientes.Por_Des           AS Por_Des_Renglon, ")
            loConsulta.AppendLine("         Renglones_DClientes.Por_Imp1          AS Por_Imp1, ")
            loConsulta.AppendLine("         Renglones_DClientes.Cod_Imp           AS Cod_Imp, ")
            loConsulta.AppendLine("         Renglones_DClientes.Mon_Imp1          AS Impuesto, ")
            loConsulta.AppendLine("         Renglones_DClientes.Comentario        AS Comentario_Renglon, ")
            loConsulta.AppendLine("         COALESCE(Facturas.Documento, '')      AS Factura_Documento, ")
            loConsulta.AppendLine("         COALESCE(Facturas.Fiscal1, '')        AS Factura_Fiscal1, ")
            loConsulta.AppendLine("         COALESCE(Facturas.Fiscal2, '')        AS Factura_Fiscal2, ")
            loConsulta.AppendLine("         COALESCE(Facturas.Fiscal3, '')        AS Factura_Fiscal3, ")
            loConsulta.AppendLine("         COALESCE(Facturas.Fiscal4,  ")
            loConsulta.AppendLine("            '20000101 00:00:00.000')           AS Factura_Fiscal4, ")
            loConsulta.AppendLine("         COALESCE(Facturas.Fiscal5, '')        AS Factura_Fiscal5 ")
            loConsulta.AppendLine("FROM     Devoluciones_Clientes ")
            loConsulta.AppendLine("    JOIN Renglones_DClientes ON Renglones_DClientes.Documento = Devoluciones_Clientes.Documento")
            loConsulta.AppendLine("    JOIN Clientes ON Clientes.Cod_Cli = Devoluciones_Clientes.Cod_Cli")
            loConsulta.AppendLine("    JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Devoluciones_Clientes.Cod_For")
            loConsulta.AppendLine("    JOIN Vendedores ON Vendedores.Cod_Ven = Devoluciones_Clientes.Cod_Ven")
            loConsulta.AppendLine("    JOIN Articulos ON Articulos.Cod_Art = Renglones_DClientes.Cod_Art")
            loConsulta.AppendLine("    LEFT JOIN Facturas")
            loConsulta.AppendLine("        ON  Facturas.Documento = Renglones_dClientes.Doc_Ori")
            loConsulta.AppendLine("        AND Renglones_dClientes.Tip_Ori = 'Facturas'")
            loConsulta.AppendLine("        AND Renglones_dClientes.Renglon = 1")
            loConsulta.AppendLine("WHERE    " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
            Dim laTablaDevolucion As DataTable = laDatosReporte.Tables(0)
            
            If laTablaDevolucion.Rows.Count = 0 Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "500px", "250px")
                Return 
            End If 

            Dim loDevolucion As DataRow = laDatosReporte.Tables(0).Rows(0)

            Dim lnDecimalesParaCantidad As Integer = goOpciones.pnDecimalesParaCantidad
            Dim lnDecimalesParaMonto As Integer = goOpciones.pnDecimalesParaMonto
            Dim lnDecimalesParaPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje

        '****************************************************************************
        ' Valida que la Devolución no haya sido impresa anteriormente, que esté confirmada, 
        ' y que tenga monto mayor a cero.
        '****************************************************************************
            Dim lcDocumento As String = CStr(loDevolucion("Documento")).Trim()
            Dim lcFiscal1 As String = CStr(loDevolucion("Fiscal1")).Trim()
            Dim lcFiscal2 As String = CStr(loDevolucion("Fiscal2")).Trim()
            Dim lcFiscal3 As String = CStr(loDevolucion("Fiscal3")).Trim()
            Dim lcFiscal4 As String = CStr(loDevolucion("Fiscal4")).Trim()

            If (lcFiscal1 > "") OrElse (lcFiscal2 > "") OrElse (lcFiscal3 > "") OrElse (lcFiscal4 > "") Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _
                                          "La Devolución de Venta " & lcDocumento & " ya fue impresa en una impresora fiscal y no puede imprimirse nuevamente: " & _
                                          "<br/>* Serial Impresora: " & lcFiscal1 & _ 
                                          "<br/>* N° N/CR Fiscal: " & lcFiscal2 & _ 
                                          "<br/>* Cierre Z: " & lcFiscal3 & _ 
                                          "<br/>* Fecha y hora: " & lcFiscal4 , _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
                                           "500px", "250px")
                Return 
            End If 

            Dim lcEstatus As String = CStr(loDevolucion("Status")).Trim()
            If (lcEstatus.ToLower() <> "confirmado") AndAlso _
                (lcEstatus.ToLower() <> "afectado")  AndAlso _
                (lcEstatus.ToLower() <> "procesado") Then

                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _
                                          "Solo se puede imprimir una Devolución de Venta fiscal si tiene estatus 'Confirmado', 'Afectado' o 'Procesado'. " & _
                                          "La Devolución de Venta " & lcDocumento & " tiene estatus '" & lcEstatus & "'.", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
                                           "500px", "250px")
                Return 
            End If 

            Dim lcMontoNeto As String = CDec(loDevolucion("Mon_Net"))
            If (lcMontoNeto <= 0D) Then 

                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _
                                          "Solo se puede imprimir una Devolución de Venta fiscal si su monto neto es mayor a cero.", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
                                           "500px", "250px")
                Return 

            End If


        '****************************************************************************
        ' Valida que la Factura de origen sea válida y que tenga información fiscal.
        '****************************************************************************
            Dim lcFacturaOrigen As String = CStr(loDevolucion("Factura_Documento")).Trim()
            Dim lcFacturaFiscal1 As String = CStr(loDevolucion("Factura_Fiscal1")).Trim()
            Dim lcFacturaFiscal2 As String = CStr(loDevolucion("Factura_Fiscal2")).Trim()
            Dim lcFacturaFiscal3 As String = CStr(loDevolucion("Factura_Fiscal3")).Trim()
            Dim lcFacturaFiscal4 As String = CStr(loDevolucion("Factura_Fiscal4")).Trim()

            If (lcFacturaOrigen = "") Then

                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _
                                          "No fue posible encontrar la factura de origen de la " & _
                                          "Devolución de Venta " & lcDocumento & "; esta información es necesaria para imprimir la N/CR fiscal correspondiente.", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
                                           "500px", "250px")
                Return 
            End If 


            If (lcFacturaFiscal1 = "") OrElse (lcFacturaFiscal2 = "") OrElse (lcFacturaFiscal3 = "") OrElse (lcFacturaFiscal4 = "") Then

                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Advertencia", _
                                          "La Factura " & lcFacturaOrigen & ", origen de la " & _
                                          "Devolución de Venta " & lcDocumento & ", no tiene la información fiscal completa;" & _
                                          " esta información es necesaria para imprimir la N/CR fiscal correspondiente." & _
                                          "<br/>* Serial Impresora: " & lcFacturaFiscal1 & _ 
                                          "<br/>* N° Factura Fiscal: " & lcFacturaFiscal2 & _ 
                                          "<br/>* Cierre Z: " & lcFacturaFiscal3 & _ 
                                          "<br/>* Fecha y hora: " & lcFacturaFiscal4, _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
                                           "500px", "250px")
                Return 

            End If 
            
            
        '****************************************************************************
        ' Genera el XML para la impresora
        '****************************************************************************
            goPuntoVenta.mObtenerOpcionesIpos()
            Me.mImprimirDevolucionXML(laTablaDevolucion)
                
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso Completado", _
                                        "La Devolución de Venta " & lcDocumento & " fue enviada a eFactory LPC para ser impresa. ", _
                                        vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                        "500px", "250px")

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

	
''' <summary>
''' Envía la devolución a un archivo XML para descargar.
''' </summary>
''' <remarks></remarks>
	Private Sub mImprimirDevolucionXML(ByRef loDocumento As DataTable)


		
	'-------------------------------------------------------------------------------------------'
	' Busca la información a imprimir.															'
	'-------------------------------------------------------------------------------------------'	
		Dim loEncabezadoDoc As DataRow = loDocumento.Rows(0)
		Dim laRenglones	As DataTable = loDocumento 
		Dim loSeleccion	As New StringBuilder()
		
        Dim lcNumeroDevolucion As String = CStr(loEncabezadoDoc("Documento")).Trim()
		Dim lcNumeroNCR As String = CStr(loEncabezadoDoc("NCR")).Trim()

	'-------------------------------------------------------------------------------------------'
	' Envía el encabezado al XML.   														    '
	'-------------------------------------------------------------------------------------------'	


        Dim loSalidaXml As New System.Xml.XmlDocument()

        Dim loRaiz As System.Xml.XmlElement = loSalidaXml.CreateElement("documento_ipos")
        loSalidaXml.AppendChild(loRaiz)

        Dim loEncabezado As System.Xml.XmlElement = loSalidaXml.CreateElement("encabezado")
        loRaiz.AppendChild(loEncabezado)
                    
        Dim loNodo As System.Xml.XmlElement
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("tipo"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection("DEVOLUCION"))

        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("documento"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcNumeroNCR))

        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("tipo_documento"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection("N/CR"))

        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("cod_cli"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Cod_Cli")).Trim()))

        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("nom_cli"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Nom_Cli")).Trim()))
        
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("rif"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Rif")).Trim()))

        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("direccion"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Dir_Fis")).Trim()))
            
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("telefono"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Telefonos")).Trim()))
            
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("cod_caj"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(goUsuario.pcCodigo))
            
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("nom_caj"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(goUsuario.pcNombre))

        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("cod_ven"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Cod_Ven")).Trim()))

		loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("nom_ven"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Nom_Ven")).Trim()))

        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("comentario"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Comentario")).Trim()))

	    'Encabezado por opciones
		If (goPuntoVenta.pcComentarioInicioNotaDeCredito.Length > 0) Then
            loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("adicional"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(goPuntoVenta.pcComentarioInicioNotaDeCredito))
		End If
		
	'-------------------------------------------------------------------------------------------'
	' Renglones de devolución.									                                '
	'-------------------------------------------------------------------------------------------'	
        Dim loRenglones As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("renglones"))
        Dim lnContador As Integer = 0
 		For lnFila As Integer = 0 To laRenglones.Rows.Count - 1 
			Dim loRenglon As DataRow = laRenglones.Rows(lnFila)
            
            lnContador += 1

            Dim loRenglonxml As System.Xml.XmlElement = loRenglones.AppendChild(loSalidaXml.CreateElement("renglon"))
                
            loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("numero"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(lnFila+1, _
                                                                    goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 0)))

            loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("cod_art"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loRenglon("Cod_Art")).Trim()))
                
            loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("nom_art"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loRenglon("Nom_Art")).Trim()))

            loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("nom_cor"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loRenglon("Nom_Cor")).Trim()))

            loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("can_art"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(CDec(loRenglon("Can_Art1")), _
                                                                    goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

            loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("precio"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(CDec(loRenglon("Precio1")), _
                                                                    goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

            loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("cod_imp"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection("x"))

            loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("por_imp"))
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(CDec(loRenglon("Por_Imp1")), _
                                                                    goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))

            loNodo = loRenglonxml.AppendChild(loSalidaXml.CreateElement("comentario"))
            loNodo.SetAttribute("antes_del_articulo", "false")
            loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loRenglon("Comentario_Renglon")).Trim()))

		Next lnFila

	'-------------------------------------------------------------------------------------------'
	' Si hay descuentos o recargos, entonces aplicarlos ahora.									'
	'-------------------------------------------------------------------------------------------'	
        Dim loDescuento As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("descuento"))
        Dim loRecargo As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("recargo"))

        Dim lcPorcentaje As String = ""
        Dim lcMonto As String = ""

        lcPorcentaje = goServicios.mObtenerFormatoCadenaCSV(CDec(loEncabezadoDoc("Por_Des")), _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)
        lcMonto = goServicios.mObtenerFormatoCadenaCSV(CDec(loEncabezadoDoc("Mon_Des")), _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)
            
        loDescuento.SetAttribute("porcentaje", lcPorcentaje)
        loDescuento.SetAttribute("monto", lcMonto)
        loDescuento.SetAttribute("global", "true")


        lcPorcentaje = goServicios.mObtenerFormatoCadenaCSV(CDec(loEncabezadoDoc("Por_Rec")), _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)
        lcMonto = goServicios.mObtenerFormatoCadenaCSV(CDec(loEncabezadoDoc("Mon_Rec")), _
                                                            goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)
		
        loRecargo.SetAttribute("porcentaje", lcPorcentaje)
        loRecargo.SetAttribute("monto", lcMonto)
        loRecargo.SetAttribute("global", "true")
			


 	'-------------------------------------------------------------------------------------------'
	' Aplica un pago total.																		'
	'-------------------------------------------------------------------------------------------'	
        Dim loFormasPago As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("formas_pago"))
        Dim loFormaPago As System.Xml.XmlElement 
        loFormaPago = loFormasPago.AppendChild(loSalidaXml.CreateElement("forma_pago"))
        loFormaPago.SetAttribute("forma", goPuntoVenta.pcMedioPagoEfectivo)
        loFormaPago.SetAttribute("tipo", "efectivo")
        loFormaPago.SetAttribute("totalizar", "true")

        loFormaPago.AppendChild(loSalidaXml.CreateCDataSection(goServicios.mObtenerFormatoCadenaCSV(CDec(loEncabezadoDoc("Mon_Net")), _
                                                    goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 10)))
		
        
	    'Datos fiscales de origen
        Dim loDatos As System.Xml.XmlElement = loSalidaXml.CreateElement("datos")
        loRaiz.AppendChild(loDatos)

        loNodo = loDatos.AppendChild(loSalidaXml.CreateElement("dato"))
        loNodo.SetAttribute("nombre", "impresora_origen")
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Factura_Fiscal1")).Trim()))

        loNodo = loDatos.AppendChild(loSalidaXml.CreateElement("dato"))
        loNodo.SetAttribute("nombre", "factura_fiscal_origen")
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Factura_Fiscal2")).Trim()))

        loNodo = loDatos.AppendChild(loSalidaXml.CreateElement("dato"))
        loNodo.SetAttribute("nombre", "cierre_z_fiscal_origen")
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Factura_Fiscal3")).Trim()))

        loNodo = loDatos.AppendChild(loSalidaXml.CreateElement("dato"))
        loNodo.SetAttribute("nombre", "fecha_fiscal_origen")
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Factura_Fiscal4")).Trim()))

        loNodo = loDatos.AppendChild(loSalidaXml.CreateElement("dato"))
        loNodo.SetAttribute("nombre", "fiscal5")
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(CStr(loEncabezadoDoc("Factura_Fiscal5")).Trim()))
        
        'Número de Devolución de Ventas, o Ajuste de Inventario
        loNodo = loDatos.AppendChild(loSalidaXml.CreateElement("dato"))
        loNodo.SetAttribute("nombre", "documento_generado")
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcNumeroDevolucion))
        
		Dim lcFechaDetalleSubtotal As String = Strings.Mid(CStr(loEncabezadoDoc("Factura_Fiscal4")).Trim(), 1, 8)
		lcFechaDetalleSubtotal = Strings.Mid(lcFechaDetalleSubtotal, 7, 2) & "-" & Strings.Mid(lcFechaDetalleSubtotal, 5, 2) & "-" & Strings.Mid(lcFechaDetalleSubtotal, 3, 2)

        Dim lcDetalleSubtotal As String = ""
		lcDetalleSubtotal &= "FACTURA: " & Strings.Trim(CStr(loEncabezadoDoc("Factura_Fiscal2")).Trim())
		lcDetalleSubtotal &= "/FEC:" & lcFechaDetalleSubtotal
		lcDetalleSubtotal &= "/IMP:" & Strings.Trim(CStr(loEncabezadoDoc("Factura_Fiscal1")).Trim()) & vbNewLine
		lcDetalleSubtotal &= "ORIGEN: FACT-" & CStr(loEncabezadoDoc("Factura_Documento")).Trim() & "/DESTINO: " 
		lcDetalleSubtotal &= "N/CR-" & lcNumeroNCR

        loNodo = loDatos.AppendChild(loSalidaXml.CreateElement("dato"))
        loNodo.SetAttribute("nombre", "detalle_subtotal")
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcDetalleSubtotal))

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






        Dim lcSalida As String = loSalidaXml.OuterXml.Replace("\","\\").Replace("'","\'")
        lcSalida = Regex.Replace(lcSalida, "[\r\n]+", "\n")

  	'-------------------------------------------------------------------------------------------'
	' Descarga el archivo.																		'
	'-------------------------------------------------------------------------------------------'	
        Dim lcNombreArchivo As String = "ipos_" & goEmpresa.pcCodigo & "_Devolucion_" &  lcNumeroDevolucion & "_" & (Date.now()).ToString("yyyyMMddhhmm") & ".xml"
        lcNombreArchivo = Regex.Replace(lcNombreArchivo, "[\\\/*? ]", "-")
        
        Dim loScript As New StringBuilder()

        loScript.Append("(function(){")
        loScript.Append("var lcContenido = '")
        loScript.Append(lcSalida)
        loScript.Append("'; var loArchivo = new Blob([lcContenido], { type: 'application/xml' }); ")
        loScript.Append("window.mDescargarBlob(loArchivo, '" & lcNombreArchivo & "')")
        loScript.Append("})();")

        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "descargardocumentoXml", loScript.ToString(), True)
        

	End Sub


    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 14/10/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
