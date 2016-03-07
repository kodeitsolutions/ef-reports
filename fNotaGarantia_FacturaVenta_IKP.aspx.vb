'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fNotaGarantia_FacturaVenta_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fNotaGarantia_FacturaVenta_IKP
    Inherits vis2formularios.frmReporte

#Region "Declaraciones"

#End Region

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpFactura(   Documento   CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Fec_Ini     DATETIME, ")
            loConsulta.AppendLine("                            Cod_Cli     CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Nom_Cli     VARCHAR(100) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Rif         VARCHAR(20) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Dir_Fis     VARCHAR(MAX) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Telefonos   VARCHAR(50) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Cod_Ven     CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Nom_Ven     VARCHAR(100) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Comentario  VARCHAR(MAX) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Renglon     INT,")
            loConsulta.AppendLine("                            Cod_Art     CHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nom_Art     VARCHAR(MAX) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Precio1     DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Can_Art1    DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Por_Des     DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Por_Imp1    DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Por_Des_Doc DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Por_Rec_Doc DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Mon_Des_Doc DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Mon_Rec_Doc DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Mon_Net     DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Neto        DECIMAL(28, 10),")
            loConsulta.AppendLine("                            Garantia    VARCHAR(30) COLLATE DATABASE_DEFAULT);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpFactura(Documento, Fec_Ini, Cod_Cli, Nom_Cli, Rif, Dir_Fis, Telefonos,")
            loConsulta.AppendLine("                        Cod_Ven, Nom_Ven, Comentario, ")
            loConsulta.AppendLine("                        Renglon, Cod_Art, Nom_Art, Precio1, Can_Art1, Por_Des, Por_Imp1,")
            loConsulta.AppendLine("                        Por_Des_Doc, Por_Rec_Doc, Mon_Des_Doc, Mon_Rec_Doc, Mon_Net, Neto, Garantia)")
            loConsulta.AppendLine("SELECT  Facturas.Documento, ")
            loConsulta.AppendLine("        Facturas.Fec_Ini, ")
            loConsulta.AppendLine("        Facturas.cod_cli, ")
            loConsulta.AppendLine("        RTRIM(CASE WHEN Facturas.nom_cli<>'' THEN Facturas.nom_cli ELSE clientes.nom_cli END),")
            loConsulta.AppendLine("        RTRIM(CASE WHEN Facturas.rif<>'' THEN Facturas.rif ELSE clientes.rif END),")
            loConsulta.AppendLine("        RTRIM(CASE WHEN Facturas.Dir_Fis<>'' THEN Facturas.Dir_Fis ELSE clientes.Dir_Fis END),")
            loConsulta.AppendLine("        RTRIM(CASE WHEN Facturas.telefonos<>'' THEN Facturas.telefonos ELSE clientes.telefonos END),")
            loConsulta.AppendLine("        Facturas.Cod_Ven,")
            loConsulta.AppendLine("        RTRIM(Vendedores.Nom_Ven),")
            loConsulta.AppendLine("        Facturas.Comentario,")
            loConsulta.AppendLine("        renglones_facturas.renglon,")
            loConsulta.AppendLine("        renglones_facturas.cod_art,")
            loConsulta.AppendLine("        renglones_facturas.notas,")
            loConsulta.AppendLine("        renglones_facturas.precio1,")
            loConsulta.AppendLine("        renglones_facturas.can_art1,")
            loConsulta.AppendLine("        renglones_facturas.Por_Des,")
            loConsulta.AppendLine("        renglones_facturas.por_imp1,")
            loConsulta.AppendLine("        facturas.por_des1,")
            loConsulta.AppendLine("        facturas.por_rec1,")
            loConsulta.AppendLine("        facturas.mon_des1,")
            loConsulta.AppendLine("        facturas.mon_rec1,")
            loConsulta.AppendLine("        facturas.Mon_Net,")
            loConsulta.AppendLine("        0,")
            loConsulta.AppendLine("        Articulos.Garantia")
            loConsulta.AppendLine("FROM Facturas")
            loConsulta.AppendLine("    JOIN clientes ON clientes.cod_cli = facturas.cod_cli")
            loConsulta.AppendLine("    JOIN Vendedores ON Vendedores.Cod_Ven = facturas.Cod_Ven")
            loConsulta.AppendLine("    JOIN renglones_facturas ON renglones_facturas.documento = facturas.documento")
            loConsulta.AppendLine("    JOIN Articulos ON Articulos.Cod_Art = renglones_facturas.Cod_Art")
            loConsulta.AppendLine("WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--Descuento del renglon")
            loConsulta.AppendLine("UPDATE #tmpFactura")
            loConsulta.AppendLine("SET Neto = ROUND(ROUND(Precio1, 2)*(1 - Por_Des/100), 2);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--Descuento y recargo del documento")
            loConsulta.AppendLine("UPDATE #tmpFactura")
            loConsulta.AppendLine("SET Neto = ROUND(Neto*(1 - Por_Des_Doc/100 + Por_Rec_Doc/100), 2);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--Impuesto")
            loConsulta.AppendLine("UPDATE #tmpFactura")
            loConsulta.AppendLine("SET Neto = ROUND(Neto*(1 + Por_Imp1/100), 2);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--Garantia del artículo")
            loConsulta.AppendLine("UPDATE #tmpFactura")
            loConsulta.AppendLine("SET Garantia = '90 días'")
            loConsulta.AppendLine("WHERE Garantia = '';")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      #tmpFactura.*, COALESCE(seriales.serial, '') AS Serial ")
            loConsulta.AppendLine("FROM        #tmpFactura")
            loConsulta.AppendLine("    LEFT JOIN seriales ")
            loConsulta.AppendLine("        ON  seriales.tip_sal = 'facturas'")
            loConsulta.AppendLine("        AND seriales.doc_sal = #tmpFactura.Documento")
            loConsulta.AppendLine("        AND seriales.ren_sal = #tmpFactura.Renglon;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DROP TABLE #tmpFactura;")
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
     
			Dim llFacturaImpresa As Boolean = Me.mGenerarDocumentoIPOSMXL(laDatosReporte.Tables(0))

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
''' Imprime un documento no fiscal (remoto) con los datos indicados. 
''' </summary>
''' <param name="loDocumento"></param>
''' <returns></returns>
''' <remarks></remarks>
	Private Function mGenerarDocumentoIPOSMXL(ByVal loDocumento As DataTable) As Boolean
            
        Dim loFilaInicio        As DataRow  = loDocumento.Rows(0)
        Dim lcDocumento         As String   = CStr(loFilaInicio("Documento")).Trim()
        Dim ldFecha             As Date     = CDate(loFilaInicio("Fec_Ini"))
        Dim lcFecha             As String   = ldFecha.ToString("dd-MM-yyyy")
        Dim lcNombreCliente     As String   = CStr(loFilaInicio("Nom_Cli")).Trim()
        Dim lcRifCliente        As String   = CStr(loFilaInicio("Rif")).Trim()
        Dim lcTelefonoCliente   As String   = CStr(loFilaInicio("Telefonos")).Trim()
            
        'Genera el encabezado del documento
        Dim loSalidaXml As New System.Xml.XmlDocument()

        Dim loRaiz As System.Xml.XmlElement = loSalidaXml.CreateElement("documento_ipos")
        loSalidaXml.AppendChild(loRaiz)

        Dim loEncabezado As System.Xml.XmlElement = loSalidaXml.CreateElement("encabezado")
        loRaiz.AppendChild(loEncabezado)
            
        Dim loNodo As System.Xml.XmlElement
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("tipo"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection("TEXTOPLANO"))
            
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("documento"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcDocumento))
            
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("tipo_documento"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection("Nota de Garantía"))
        
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("fecha"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcFecha))
            
        loNodo = loEncabezado.AppendChild(loSalidaXml.CreateElement("comentario"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection("Formato de Nota de Garantía IKP: Diseñado para impresora fiscal BIXOLON, 56 columnas."))


	'-------------------------------------------------------------------------------------------'
	' Inicio del doccumento de texto: diseñado para una impresora BIXOLON (56 columnas).	    '
	'-------------------------------------------------------------------------------------------'	
        Dim loTextos As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("textos"))
        Const KN_COLUMNAS As Integer = 56
        Dim lcSeparador As String = Strings.StrDup(KN_COLUMNAS, "-")
        
        Dim lcTamaño5 As String = Strings.Space(5)
        Dim lcTamaño16 As String = Strings.Space(16)
        Dim lcTamaño26 As String = Strings.Space(26)
        Dim lcTamaño20 As String = Strings.Space(19)
        Dim lcTamaño40 As String = Strings.Space(40)
        Dim lcTamañoMAX As String = Strings.Space(KN_COLUMNAS)


        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(Strings.Left(lcTamaño20 & "NOTA DE GARANTÍA" & lcTamañoMAX, KN_COLUMNAS)))
        
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcSeparador))
        
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(Strings.Left("Origen: Factura de Venta " & lcDocumento & "/Fecha: " & lcFecha & lcTamañoMAX, KN_COLUMNAS)))
        
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcSeparador))
        
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(Strings.Left("CLIENTE:  " & lcNombreCliente & lcTamañoMAX, KN_COLUMNAS)))
        
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(Strings.Left("C.I./RIF: " & lcRifCliente & lcTamañoMAX, KN_COLUMNAS)))
        
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(Strings.Left("TELÉFONO: " & lcTelefonoCliente & lcTamañoMAX, KN_COLUMNAS)))
        
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcSeparador))
        
        Dim lcEncabezado As String = ""
        lcEncabezado &= "          SERIAL          "
        lcEncabezado &= "    PRECIO     "
        lcEncabezado &= "GARANTÍA HASTA"
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcEncabezado))
        
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcSeparador))
        

        Dim loTexto As System.Xml.XmlElement
        For lnFila As Integer = 0 To loDocumento.Rows.Count - 1
			Dim loRenglon As DataRow = loDocumento.Rows(lnFila)
            
            Dim lcDescripcion As String = Strings.Left(CStr(loRenglon("Nom_art")).Trim() & lcTamaño40, 40) & lcTamaño16

            Dim lcSerial As String = Strings.Left(CStr(loRenglon("Serial")).Trim() & lcTamaño26, 26)

            Dim lcPrecio As String = goServicios.mObtenerFormatoCadena(CDec(loRenglon("Neto")), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 2)
            lcPrecio = Strings.Right(lcTamaño16 & "Bs " & lcPrecio & lcTamaño5, 16)

            Dim lcGarantia As String = Strings.Right(CStr(loRenglon("Garantia")).Trim() & lcTamaño26, 14) 

            loTexto = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
            loTexto.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
            loTexto.AppendChild(loSalidaXml.CreateCDataSection(lcDescripcion))

            loTexto = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
            loTexto.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
            loTexto.AppendChild(loSalidaXml.CreateCDataSection(lcSerial & lcPrecio & lcGarantia))

        Next lnFila 
        
        loNodo = loTextos.AppendChild(loSalidaXml.CreateElement("texto"))
        loNodo.SetAttribute("nombre", "texto" & loTextos.ChildNodes.Count.ToString("000"))
        loNodo.AppendChild(loSalidaXml.CreateCDataSection(lcSeparador))

        'Añade la sección de parámetros 
        Dim loParametros As System.Xml.XmlElement = loRaiz.AppendChild(loSalidaXml.CreateElement("parametros"))
        Dim loParametro As System.Xml.XmlElement

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
        
        Dim lcNombreArchivo As String = "ipos_" & goEmpresa.pcCodigo & "_NotaGarantia_" & lcDocumento & "_" & (Date.now()).ToString("yyyyMMddhhmm") & ".xml"
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
                          "Se generó un archivo XML de Nota de Garantía (a partir de la factura seleccionada) para ser impreso con eFactory PLC.", _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                           "auto", _
                           "auto")
       
		Return True
		
	End Function

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 24/09/13: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
