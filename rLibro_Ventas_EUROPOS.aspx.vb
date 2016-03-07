'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Ventas_EUROPOS"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Ventas_EUROPOS
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
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosFinales(5)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- FACTURAS DE VENTA")
            loConsulta.AppendLine("SELECT	CAST(Facturas.Fec_Ini AS DATE) 						AS Fecha, ")
            loConsulta.AppendLine("			1													AS Tipo, ")
            loConsulta.AppendLine("			'FACTURA'											AS Tipo_Documento, ")
            loConsulta.AppendLine("			Facturas.Referencia									AS Reporte_Z, ")
            loConsulta.AppendLine("			Facturas.Documento									AS Doc_Ini, ")
            loConsulta.AppendLine("			Facturas.Control									AS Doc_fin,")
            loConsulta.AppendLine("			Renglones_Facturas.Cod_Alm							AS Impresora,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1")
            loConsulta.AppendLine("				THEN 'VARIOS'")
            loConsulta.AppendLine("				ELSE Clientes.Cod_Cli")
            loConsulta.AppendLine("			END)												AS Cod_Cli,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1")
            loConsulta.AppendLine("				THEN 'CLIENTES VARIOS'")
            loConsulta.AppendLine("				ELSE Clientes.Nom_Cli")
            loConsulta.AppendLine("			END)												AS Nom_Cli,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1")
            loConsulta.AppendLine("				THEN 'VARIOS'")
            loConsulta.AppendLine("				ELSE Clientes.Rif")
            loConsulta.AppendLine("			END)												AS Rif,")
            loConsulta.AppendLine("			Facturas.Mon_Net									AS Mon_Net,")
            loConsulta.AppendLine("			Facturas.Mon_Exe									AS Mon_Exe,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 0 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Mon_Net - Facturas.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Base,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 0 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Por_Impuesto,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 0 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Mon_Impuesto,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Mon_Net - Facturas.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Base,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Por_Impuesto,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Mon_Impuesto,")
            loConsulta.AppendLine("			NULL                                                AS Ret_Fecha,")
            loConsulta.AppendLine("			''							                        AS Ret_Comprobante,")
            loConsulta.AppendLine("			0 							                        AS Ret_MontoRetenido")
            loConsulta.AppendLine("FROM		Facturas")
            loConsulta.AppendLine("	JOIN	Renglones_Facturas ")
            loConsulta.AppendLine("		ON	Renglones_Facturas.Documento = Facturas.Documento")
            loConsulta.AppendLine("		AND Renglones_Facturas.Renglon = 1")
            loConsulta.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loConsulta.AppendLine("WHERE    Facturas.Status IN ('Confirmado','Afectado','Procesado')")
            loConsulta.AppendLine("		AND Facturas.Mon_Net > 0")
            loConsulta.AppendLine("     AND Facturas.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Facturas.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("		AND Facturas.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("		AND Facturas.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("		AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Facturas.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Facturas.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("-- NOTAS DE CRÉDITO")
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("SELECT		CAST(Cuentas_Cobrar.Fec_Ini AS DATE) 				AS Fecha, ")
            loConsulta.AppendLine("			2													AS Tipo, ")
            loConsulta.AppendLine("			'NOTA DE CRÉDITO'									AS Tipo_Documento,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Referencia							AS Reporte_Z,  ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Documento							AS Doc_Ini, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Control								AS Doc_fin,")
            loConsulta.AppendLine("			Renglones_Documentos.Cod_Alm						AS Impresora,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1")
            loConsulta.AppendLine("				THEN 'VARIOS'")
            loConsulta.AppendLine("				ELSE Clientes.Cod_Cli")
            loConsulta.AppendLine("			END)												AS Cod_Cli,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1")
            loConsulta.AppendLine("				THEN 'CLIENTES VARIOS'")
            loConsulta.AppendLine("				ELSE Clientes.Nom_Cli")
            loConsulta.AppendLine("			END)												AS Nom_Cli,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1")
            loConsulta.AppendLine("				THEN 'VARIOS'")
            loConsulta.AppendLine("				ELSE Clientes.Rif")
            loConsulta.AppendLine("			END)												AS Rif,")
            loConsulta.AppendLine("			-Cuentas_Cobrar.Mon_Net								AS Mon_Net,")
            loConsulta.AppendLine("			-Cuentas_Cobrar.Mon_Exe								AS Mon_Exe,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 0 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Net - Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Base,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 0 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Por_Impuesto,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 0 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Mon_Impuesto,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 1 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Net - Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Base,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 1 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Por_Impuesto,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 1 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Mon_Impuesto,")
            loConsulta.AppendLine("			NULL                                                AS Ret_Fecha,")
            loConsulta.AppendLine("			''							                        AS Ret_Comprobante,")
            loConsulta.AppendLine("			0 							                        AS Ret_MontoRetenido")
            loConsulta.AppendLine("FROM		Cuentas_Cobrar")
            loConsulta.AppendLine("	JOIN	Renglones_Documentos")
            loConsulta.AppendLine("		ON	Renglones_Documentos.Documento = Cuentas_Cobrar.Documento")
            loConsulta.AppendLine("		AND Renglones_Documentos.Cod_Tip = Cuentas_Cobrar.Cod_Tip")
            loConsulta.AppendLine("		AND Renglones_Documentos.Renglon = 1")
            loConsulta.AppendLine("		AND Renglones_Documentos.Origen = 'Ventas'")
            loConsulta.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loConsulta.AppendLine("WHERE	Cuentas_Cobrar.Status IN ('Pendiente','Afectado','Pagado')")
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Tip = 'N/CR'")
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("		AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- RETENCIONES DE IVA")
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("SELECT		CAST(Cuentas_Cobrar.Fec_Ini AS DATE) 				AS Fecha, ")
            loConsulta.AppendLine("			3													AS Tipo, ")
            loConsulta.AppendLine("			'C/RETENCIÓN'									    AS Tipo_Documento,")
            loConsulta.AppendLine("			''							                        AS Reporte_Z,  ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Num_Com							    AS Doc_Ini, ")
            loConsulta.AppendLine("			''							                        AS Doc_fin,")
            loConsulta.AppendLine("			''							                        AS Impresora,")
            loConsulta.AppendLine("			Clientes.Cod_Cli									AS Cod_Cli,")
            loConsulta.AppendLine("			Clientes.Nom_Cli    								AS Nom_Cli,")
            loConsulta.AppendLine("			Clientes.Rif										AS Rif,")
            loConsulta.AppendLine("			0								                    AS Mon_Net,")
            loConsulta.AppendLine("			0								                    AS Mon_Exe,")
            loConsulta.AppendLine("			0								                    AS NC_Base,")
            loConsulta.AppendLine("			0								                    AS NC_Por_Impuesto,")
            loConsulta.AppendLine("			0								                    AS NC_Mon_Impuesto,")
            loConsulta.AppendLine("			0								                    AS C_Base,")
            loConsulta.AppendLine("			0								                    AS C_Por_Impuesto,")
            loConsulta.AppendLine("			0								                    AS C_Mon_Impuesto,")
            loConsulta.AppendLine("			CAST(Cuentas_Cobrar.Fec_Ini AS DATE)                AS Ret_Fecha,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Num_Com			                    AS Ret_Comprobante,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Net                              AS Ret_MontoRetenido")
            loConsulta.AppendLine("FROM		Cuentas_Cobrar")
            loConsulta.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loConsulta.AppendLine("WHERE	Cuentas_Cobrar.Status IN ('Pendiente','Afectado','Pagado')")
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Tip      =   'RETIVA' ")
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("		AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("ORDER BY	Fecha, Impresora, Tipo, Doc_Ini")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

			'Me.mEscribirConsulta(loConsulta.ToString())
			Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            
          
            
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
            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Ventas_EUROPOS", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibro_Ventas_EUROPOS.ReportSource = loObjetoReporte

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

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 06/05/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
