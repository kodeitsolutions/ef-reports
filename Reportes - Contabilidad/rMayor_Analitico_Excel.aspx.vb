'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMayor_Analitico_Excel"
'-------------------------------------------------------------------------------------------'
Partial Class rMayor_Analitico_Excel

	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try

			Dim lcFechaDesde 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcFechaHasta 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcCuentaContableDesde 	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcCuentaContableHasta 	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			Dim lcCentroCostoDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcCentroCostoHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcCuentaGastoDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcCuentaGastoHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcAuxiliarDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcAuxiliarHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcMonedaDesde 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcMonedaHasta 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcDocumentoDesde		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcDocumentoHasta		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
					
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()
			
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE	@lnCero DECIMAL(28, 10);")
			loComandoSeleccionar.AppendLine("SET		@lnCero = 0;")
			loComandoSeleccionar.AppendLine("DECLARE @llUsarTasa BIT; SET @llUsarTasa = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	ROW_NUMBER() ") 
			loComandoSeleccionar.AppendLine("			OVER (PARTITION BY Renglones_Comprobantes.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("				ORDER BY Renglones_Comprobantes.Documento, Renglones_Comprobantes.Renglon)			AS Posicion, ") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Documento															AS Documento, ")
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Renglon																AS Renglon, ") 
			loComandoSeleccionar.AppendLine("		CAST(Renglones_Comprobantes.Comentario AS VARCHAR(MAX))										AS Comentario, ") 
			loComandoSeleccionar.AppendLine("		CC.Cod_Cue																					AS Cod_Cue, ")
			loComandoSeleccionar.AppendLine("		CC.Nom_Cue																					AS Nom_Cue, ")
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Cen																AS Cod_Cen, ") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Gas																AS Cod_Gas, ") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Aux																AS Cod_Aux, ") 
			loComandoSeleccionar.AppendLine("		ISNULL(Auxiliares.Nom_Aux, '')																AS Nom_Aux, ")
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Fec_Ini																AS Fec_Ini, ") 
			loComandoSeleccionar.AppendLine("		ISNULL(Iniciales.Saldo_Inicial, @lnCero)													AS Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("		(ISNULL(Renglones_Comprobantes.Mon_Deb, @lnCero)") 
			loComandoSeleccionar.AppendLine("				*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)")
			loComandoSeleccionar.AppendLine("			)																						AS Debe,") 
			loComandoSeleccionar.AppendLine("		(ISNULL(Renglones_Comprobantes.Mon_Hab, @lnCero)") 
			loComandoSeleccionar.AppendLine("				*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("			)																						AS Haber")
			loComandoSeleccionar.AppendLine("INTO 	#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables AS CC")
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ") 
			loComandoSeleccionar.AppendLine("		INNER JOIN Comprobantes") 
			loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
			loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
			loComandoSeleccionar.AppendLine("				AND Comprobantes.Tipo = 'Diario' AND Comprobantes.Status <> 'Anulado'") 
			loComandoSeleccionar.AppendLine("				AND Comprobantes.Documento BETWEEN " & lcDocumentoDesde & " AND " & lcDocumentoHasta)
			loComandoSeleccionar.AppendLine("			)") 
			loComandoSeleccionar.AppendLine("		ON	CC.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
			loComandoSeleccionar.AppendLine("		AND (Renglones_Comprobantes.fec_ini BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta & ")")
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN (") 
			loComandoSeleccionar.AppendLine("			SELECT	RC.Cod_Cue						AS Cod_Cue,") 
			loComandoSeleccionar.AppendLine("					SUM(RC.Mon_Deb - RC.Mon_Hab)	AS Saldo_Inicial")
			loComandoSeleccionar.AppendLine("			FROM	Renglones_Comprobantes AS RC ") 
			loComandoSeleccionar.AppendLine("				INNER JOIN Comprobantes AS C") 
			loComandoSeleccionar.AppendLine("					ON (RC.Adicional = C.Adicional ") 
			loComandoSeleccionar.AppendLine("						AND RC.Documento = C.Documento")
			loComandoSeleccionar.AppendLine("						AND C.Tipo = 'Diario' AND C.Status <> 'Anulado'") 
			loComandoSeleccionar.AppendLine("						AND C.Documento BETWEEN " & lcDocumentoDesde & " AND " & lcDocumentoHasta & ")")
			loComandoSeleccionar.AppendLine("			WHERE	(RC.fec_ini < " & lcFechaDesde & " )") 
			loComandoSeleccionar.AppendLine("			AND	RC.Cod_Cue	BETWEEN " & lcCuentaContableDesde	& "	AND	" & lcCuentaContableHasta)
			loComandoSeleccionar.AppendLine("			AND RC.Cod_Cen	BETWEEN " & lcCentroCostoDesde		& "	AND	" & lcCentroCostoHasta)
			loComandoSeleccionar.AppendLine("			AND RC.Cod_Gas	BETWEEN " & lcCuentaGastoDesde		& "	AND	" & lcCuentaGastoHasta)
			loComandoSeleccionar.AppendLine("			AND RC.Cod_Aux	BETWEEN " & lcAuxiliarDesde			& "	AND	" & lcAuxiliarHasta)
			loComandoSeleccionar.AppendLine("			AND RC.Cod_Mon	BETWEEN " & lcMonedaDesde			& "	AND	" & lcMonedaHasta)
			loComandoSeleccionar.AppendLine("			GROUP BY RC.Cod_Cue")
			loComandoSeleccionar.AppendLine("			) AS Iniciales") 
			loComandoSeleccionar.AppendLine("		ON CC.Cod_Cue = Iniciales.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("	LEFT JOIN Auxiliares")
			loComandoSeleccionar.AppendLine("       ON	Renglones_Comprobantes.Cod_Aux = Auxiliares.Cod_Aux") 
			loComandoSeleccionar.AppendLine("WHERE	CC.Movimiento = 1") 
			loComandoSeleccionar.AppendLine("		AND	CC.Cod_Cue						BETWEEN " & lcCuentaContableDesde	& "	AND	" & lcCuentaContableHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcCentroCostoDesde		& "	AND	" & lcCentroCostoHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcCuentaGastoDesde		& "	AND	" & lcCuentaGastoHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcAuxiliarDesde			& "	AND	" & lcAuxiliarHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcMonedaDesde			& "	AND	" & lcMonedaHasta)
			loComandoSeleccionar.AppendLine("ORDER BY CC.Cod_Cue, Posicion")
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("CREATE Clustered INDEX PK_tmpMovimientos_Cuenta_Posicion ON #tmpMovimientos(Cod_Cue, Posicion)")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	#tmpMovimientos.Posicion 								AS Posicion,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Documento 								AS Documento,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Renglon 								AS Renglon,		")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Comentario								AS Comentario,	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Cue 								AS Cod_Cue, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Nom_Cue 								AS Nom_Cue, 	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Cen 								AS Cod_Cen, 	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Gas 								AS Cod_Gas, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Aux 								AS Cod_Aux, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Nom_Aux 								AS Nom_Aux, 	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Fec_Ini 								AS Fec_Ini, 	") 
			'loComandoSeleccionar.AppendLine("		#tmpMovimientos.Saldo_Inicial							AS Original,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Saldo_Inicial							AS Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Debe									AS Debe,		") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Haber									AS Haber,		") 
			loComandoSeleccionar.AppendLine("		(#tmpMovimientos.Debe - #tmpMovimientos.Haber)			AS Saldo_Actual	") 
			loComandoSeleccionar.AppendLine("FROM	#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("GROUP BY	#tmpMovimientos.Posicion,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Documento,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Renglon,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Comentario,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Cue,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Nom_Cue,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Cen,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Gas,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Aux,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Nom_Aux,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Fec_Ini,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Debe,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Haber")
			loComandoSeleccionar.AppendLine("ORDER BY	#tmpMovimientos.Cod_Cue, ")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Posicion") 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes", 1200)
            
            '-------------------------------------------------------------------------------------------'
            ' Selección de opcion por excel (Microsoft Excel - xls)
            '-------------------------------------------------------------------------------------------'
            If (Me.Request.QueryString("salida").ToLower() = "xls") Then
                ' Genera el archivo a partir de la tabla de datos y termina la ejecución
                Me.mGenerarArchivoExcel(laDatosReporte)

            End If

			'-------------------------------------------------------------------------------------------------------
			' Verificando si el select (tabla nº0) trae registros
			'-------------------------------------------------------------------------------------------------------
			If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
				Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
					"No se Encontraron Registros para los Parámetros Especificados. ", _
					vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
					"350px", _
					"200px")
			End If

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMayor_Analitico_Excel", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrMayor_Analitico_Excel.ReportSource = loObjetoReporte

		Catch loExcepcion As Exception

			Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
						  "No se pudo Completar el Proceso: " & loExcepcion.Message, _
						   vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
						   "auto", _
						   "auto")

		End Try

	End Sub
    
    Private Sub mGenerarArchivoExcel(ByVal loDatos As DataSet)

        '-------------------------------------------------------------------------------------------'
        ' Prepara los datos para enviarlos al servicio web de Excel.            '
        '-------------------------------------------------------------------------------------------'
        Dim loSalida As New IO.MemoryStream()
        loDatos.WriteXml(loSalida, XmlWriteMode.WriteSchema)

        '-------------------------------------------------------------------------------------------'
        ' Prepara los parámetros adicionales para enviarlos junto con los datos.'
        '-------------------------------------------------------------------------------------------'
        Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
        Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
        Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje

        Dim loParametros As New NameValueCollection()
        loParametros.Add("lcNombreEmpresa", cusAplicacion.goEmpresa.pcNombre)
        loParametros.Add("lcRifEmpresa", cusAplicacion.goEmpresa.pcRifEmpresa)
        loParametros.Add("lnDecimalesMonto", lnDecimalesMonto.ToString())
        loParametros.Add("lnDecimalesCantidad", lnDecimalesCantidad.ToString())
        loParametros.Add("lnDecimalesPorcentaje", lnDecimalesPorcentaje.ToString())

        Dim loClienteWeb As New System.Net.WebClient()
        loClienteWeb.QueryString = loParametros

        '-------------------------------------------------------------------------------------------'
        ' Envía los datos y parámetros, y espera la respuesta.                  '
        '-------------------------------------------------------------------------------------------'
        Dim loRespuesta As Byte()
        Try
            Dim lcRuta As String = Me.MapPath("~\Framework\Xml\ParametrosGlobales.xml")
            Dim loParam As New System.Xml.XmlDocument()
            loParam.Load(lcRuta)
            Dim lcServicio As String = loParam.DocumentElement.GetAttribute("Servicios")

            loRespuesta = loClienteWeb.UploadData(lcServicio & "/Reportes/rMayor_Analitico_Excel_xlsx.aspx", loSalida.GetBuffer())

        Catch ex As Exception
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 ex.ToString(), vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return
        End Try

        '-------------------------------------------------------------------------------------------'
        ' Vemos si la respuesta es TextoPlano (error) o no (el archivo Excel    '
        ' generado). Si el tipo está vacio : error desconocido.                 '
        '-------------------------------------------------------------------------------------------'
        Dim loTipoRespuesta As String = loClienteWeb.ResponseHeaders("Content-Type")

        If String.IsNullOrEmpty(loTipoRespuesta) Then

            '-------------------------------------------------------------------------------------------'
            'Error no especificado!
            '-------------------------------------------------------------------------------------------'
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: El servicio que genera la salida XSLX no responde.<br/>", _
                                                                 vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        ElseIf loTipoRespuesta.ToLower().StartsWith("text/plain") Then

            Dim lcMensaje As String = UTF32Encoding.UTF8.GetString(loRespuesta)

            lcMensaje = Me.Server.HtmlEncode(lcMensaje)

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 lcMensaje, vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        Else

            '-------------------------------------------------------------------------------------------'
            'Generación exitosa: la respuesta es el archivo en excel para descargar
            '-------------------------------------------------------------------------------------------'
            Me.Response.Clear()
            Me.Response.Buffer = True
            Me.Response.AppendHeader("content-disposition", "attachment; filename=rMayor_Analitico_Excel.xlsx")
            Me.Response.ContentType = "application/excel"
            Me.Response.BinaryWrite(loRespuesta)
            Me.Response.End()

        End If


    End Sub

	Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

		Try

			loObjetoReporte.Close()

		Catch loExcepcion As Exception

		End Try

	End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 29/06/15: Codigo inicial, a partir de rMayor_Analitico.								'
'-------------------------------------------------------------------------------------------'
