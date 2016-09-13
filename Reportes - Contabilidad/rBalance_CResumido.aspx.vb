'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'Imports Microsoft.Office.Interop

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rBalance_CResumido"
'-------------------------------------------------------------------------------------------'
Partial Class rBalance_CResumido

	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Desde DATETIME	")		'Fecha inicio
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Hasta DATETIME	")		'Fecha Fin
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Desde VARCHAR(100)	")	'Cuenta Contable
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Hasta VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Desde VARCHAR(100)	")	'Centro costo
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Hasta VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Desde VARCHAR(100)	")	'Cta. gasto
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Hasta VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Desde VARCHAR(100)	")	'Auxiliares 
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Hasta VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Hasta VARCHAR(100)	")	'moneda
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Desde VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SET		@lcParametro1Desde = " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro1Hasta = " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro2Desde = " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro2Hasta = " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro3Desde = " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro3Hasta = " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro4Desde = " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro4Hasta = " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro5Desde = " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro5Hasta = " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("SET		@lcParametro6Desde = " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro6Hasta = " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("--*********************************************************")
			loComandoSeleccionar.AppendLine("-- Prepara un listado de las cuentas contables a incluir  *")
			loComandoSeleccionar.AppendLine("--*********************************************************")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10);")
			loComandoSeleccionar.AppendLine("SET		@lnCero = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @llUsarTasa BIT; SET @llUsarTasa = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	Cuentas_Contables.Cod_Cue	AS Cod_Cue,")
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < @lcParametro1Desde)")
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.mon_deb - Renglones_Comprobantes.mon_hab, @lnCero)")
			loComandoSeleccionar.AppendLine("					ELSE @lnCero")
			loComandoSeleccionar.AppendLine("			END")
			loComandoSeleccionar.AppendLine("		) AS saldo,")
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=@lcParametro1Desde and Renglones_Comprobantes.Fec_Ini<=@lcParametro1Hasta)")
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.mon_deb, @lnCero)")
			loComandoSeleccionar.AppendLine("					ELSE @lnCero")
			loComandoSeleccionar.AppendLine("			END")
			loComandoSeleccionar.AppendLine("		) AS Debe,")
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=@lcParametro1Desde and Renglones_Comprobantes.Fec_Ini<=@lcParametro1Hasta)")
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Hab, @lnCero)")
			loComandoSeleccionar.AppendLine("					ELSE @lnCero")
			loComandoSeleccionar.AppendLine("			END")
			loComandoSeleccionar.AppendLine("		) AS Haber,")
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=@lcParametro1Desde and Renglones_Comprobantes.Fec_Ini<=@lcParametro1Hasta)")
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.mon_deb - Renglones_Comprobantes.mon_hab, @lnCero)")
			loComandoSeleccionar.AppendLine("					ELSE @lnCero")
			loComandoSeleccionar.AppendLine("			END")
			loComandoSeleccionar.AppendLine("		) AS Monto		")
			loComandoSeleccionar.AppendLine("INTO	#tmpValores")
			loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables ")
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ")
			loComandoSeleccionar.AppendLine("		inner JOIN Comprobantes")
			loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
			loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.documento = Comprobantes.documento")
			loComandoSeleccionar.AppendLine("				AND Comprobantes.Tipo = 'Diario' AND Comprobantes.Status<> 'anulado'")
			loComandoSeleccionar.AppendLine("			)")
			loComandoSeleccionar.AppendLine("		ON Cuentas_Contables.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
			loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.fec_ini<=@lcParametro1Hasta)")
			loComandoSeleccionar.AppendLine("WHERE	Cuentas_Contables.Movimiento=1")
			loComandoSeleccionar.AppendLine("		AND Cuentas_Contables.cod_cue		BETWEEN " & lcParametro1Desde & "	AND	" & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.cod_cen	BETWEEN " & lcParametro2Desde & "	AND	" & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.cod_gas	BETWEEN " & lcParametro3Desde & "	AND	" & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.cod_aux	BETWEEN " & lcParametro4Desde & "	AND	" & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.cod_mon	BETWEEN " & lcParametro5Desde & "	AND	" & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("GROUP BY Cuentas_Contables.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	CC.Cod_Cue																					AS Cod_Cue, ")
			loComandoSeleccionar.AppendLine("		CC.Nom_Cue																					AS Nom_cue,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Saldo, @lnCero)															AS Saldo_Inicial,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Debe, @lnCero)															AS Debe,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Haber, @lnCero)															AS Haber,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Monto, @lnCero)															AS Monto,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Saldo + #tmpValores.Monto, @lnCero)										AS Saldo_Actual")
			loComandoSeleccionar.AppendLine("INTO	#tmpParcial")
			loComandoSeleccionar.AppendLine("FROM	#tmpValores")
			loComandoSeleccionar.AppendLine("	RIGHT JOIN Cuentas_Contables As CC ")
			loComandoSeleccionar.AppendLine("		ON #tmpValores.Cod_Cue = CC.Cod_Cue")
			loComandoSeleccionar.AppendLine("WHERE	CC.Movimiento = 1")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpValores")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		A.Cod_Cue, A.Nom_Cue, ")
			loComandoSeleccionar.AppendLine("			SUM(ISNULL(B.Saldo_Inicial, A.Saldo_Inicial))	AS Saldo_Inicial, ")
			loComandoSeleccionar.AppendLine("			SUM(ISNULL(B.Debe, A.Debe))						AS Mon_Deb, ")
			loComandoSeleccionar.AppendLine("			SUM(ISNULL(B.Haber, A.Haber))					AS Mon_Hab, ")
			loComandoSeleccionar.AppendLine("			SUM(ISNULL(B.Saldo_Actual, A.Saldo_Actual))		AS Monto  ")
			loComandoSeleccionar.AppendLine("FROM		#tmpParcial AS A")
			loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpParcial AS B ON B.Cod_Cue LIKE (RTRIM(A.Cod_Cue)+'%')")
			loComandoSeleccionar.AppendLine("GROUP BY	A.Cod_Cue, A.Nom_Cue  ")
			loComandoSeleccionar.AppendLine("HAVING		ABS(SUM(ISNULL(B.Saldo_Inicial, @lnCero))) ")
			loComandoSeleccionar.AppendLine("		  + ABS(SUM(ISNULL(B.Debe, @lnCero))) ")
			loComandoSeleccionar.AppendLine("		  + ABS(SUM(ISNULL(B.Haber, @lnCero))) > 0")
			loComandoSeleccionar.AppendLine("ORDER BY	A.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpParcial")
			loComandoSeleccionar.AppendLine("")

			Dim loServicios As New cusDatos.goDatos
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rBalance_CResumido", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrBalance_CResumido.ReportSource = loObjetoReporte

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

            loRespuesta = loClienteWeb.UploadData(lcServicio & "/Reportes/rBalance_CResumido_xlsx.aspx", loSalida.GetBuffer())

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
            Me.Response.AppendHeader("content-disposition", "attachment; filename=rBalance_CResumido.xlsx")
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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 11/10/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT:  16/05/11: Mejora de la vista de diseño, Ajuste del Select
'-------------------------------------------------------------------------------------------' 
' RJG:  27/10/11: Corrección en cálculo de saldos inicial y final. Corrección en filtro.	'
'-------------------------------------------------------------------------------------------' 
' RJG: 06/12/11: Se agregó la igualdad de campo Adicional en las uniones entre Comprobantes	'
'				 y sus renglones.															'
'-------------------------------------------------------------------------------------------' 
' RJG: 09/02/12: Cambiado el filtro de comprobantes "='Pendiente'" por "<>'Anulado'".		'
'-------------------------------------------------------------------------------------------' 
' RJG: 27/06/15: Se programó el envío a Excel.		                                        '
'-------------------------------------------------------------------------------------------' 
