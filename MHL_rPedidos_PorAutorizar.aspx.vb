'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MHL_rPedidos_PorAutorizar"
'-------------------------------------------------------------------------------------------'
Partial Class MHL_rPedidos_PorAutorizar
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcParametro9Desde As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(9)).Trim().ToUpper()

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()
            

            'Obtiene el URL base para los URLs (relativos) de los documentos/links
            Dim lcUrlReporte AS String = Me.Request.Url.AbsoluteUri 
            Dim lcCarpetaReporte As String = "/Administrativo/Reportes/"
            Dim lnLongitudBase As Integer = lcUrlReporte.IndexOf(lcCarpetaReporte) '+ lcCarpetaReporte.Length

            Dim lcUrlBase As String = lcUrlReporte.Substring(0, lnLongitudBase)
            Dim lcUrlBaseSQL As String = goServicios.mObtenerCampoFormatoSQL(lcUrlBase)
                
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT		Pedidos.Documento, ")
            loConsulta.AppendLine(" 			Pedidos.Status, ")
            loConsulta.AppendLine(" 			Pedidos.Fec_Ini, ")
            loConsulta.AppendLine(" 			Pedidos.Fec_Fin, ")
            loConsulta.AppendLine(" 			Pedidos.Cod_Cli, ")
            loConsulta.AppendLine(" 			Clientes.Nom_Cli,")
            loConsulta.AppendLine(" 			Pedidos.Cod_Ven, ")
            loConsulta.AppendLine(" 			Pedidos.Cod_Tra, ")
            loConsulta.AppendLine(" 			Pedidos.Control, ")
            loConsulta.AppendLine(" 			Pedidos.Mon_Imp1,")
            loConsulta.AppendLine(" 			Pedidos.Mon_Bru,")
            loConsulta.AppendLine(" 			Pedidos.Mon_Net, ")
            loConsulta.AppendLine(" 			CAST((CASE ")
            loConsulta.AppendLine(" 				WHEN CAST(Pedidos.Fec_Ini AS DATE) = CAST(Pedidos.Fec_Fin AS DATE)")
            loConsulta.AppendLine(" 				THEN 1 ELSE 0")
            loConsulta.AppendLine(" 			END)AS BIT)                                                             AS Contado,")
            loConsulta.AppendLine(" 			CAST(COALESCE(Autorizados.Val_Log, 0) AS BIT)                           AS Autorizado,")
            loConsulta.AppendLine(" 			CAST(COALESCE(Pagados.Val_Log, 0) AS BIT)                               AS Pagado,")
            loConsulta.AppendLine(" 			CAST(COALESCE(Aprobados.Val_Log, 0) AS BIT)                             AS Aprobado,")
            loConsulta.AppendLine(" 			CAST(COALESCE(Documentacion.archivo,'[NO DISPONIBLE]') AS VARCHAR(MAX))	AS Nombre_Documento,")
            loConsulta.AppendLine(" 			CAST(CASE WHEN (Documentacion.ruta IS NULL)")
            loConsulta.AppendLine(" 			    THEN ''") 
            loConsulta.AppendLine(" 			    ELSE " & lcUrlBaseSQL & " + REPLACE(RTRIM(Documentacion.ruta), '../..','')")
            loConsulta.AppendLine(" 			END AS VARCHAR(MAX))		                                            AS Ruta_Documento")
            loConsulta.AppendLine("FROM		Pedidos ")
            loConsulta.AppendLine("    JOIN    Clientes ")
            loConsulta.AppendLine("		ON	Clientes.Cod_Cli = Pedidos.Cod_Cli")
            loConsulta.AppendLine("	LEFT JOIN Campos_Propiedades Autorizados")
            loConsulta.AppendLine("		ON	Autorizados.Cod_Reg = Pedidos.Documento")
            loConsulta.AppendLine("			AND Autorizados.Origen = 'Pedidos'")
            loConsulta.AppendLine("			AND Autorizados.Cod_Pro = 'PEDAUTFIN'")
            loConsulta.AppendLine("	LEFT JOIN Campos_Propiedades Pagados")
            loConsulta.AppendLine("		ON	Pagados.Cod_Reg = Pedidos.Documento")
            loConsulta.AppendLine("			AND Pagados.Origen = 'Pedidos'")
            loConsulta.AppendLine("			AND Pagados.Cod_Pro = 'PEDPAGREC'")
            loConsulta.AppendLine("	LEFT JOIN Campos_Propiedades Aprobados")
            loConsulta.AppendLine("		ON	Aprobados.Cod_Reg = Pedidos.Documento")
            loConsulta.AppendLine("			AND Aprobados.Origen = 'Pedidos'")
            loConsulta.AppendLine("			AND Aprobados.Cod_Pro = 'PEDPAGATCB'")
            loConsulta.AppendLine("	LEFT JOIN Documentacion ")
            loConsulta.AppendLine("		ON	Documentacion.Cod_Reg = Pedidos.Documento")
            loConsulta.AppendLine("			AND Documentacion.Origen = 'Pedidos'")
            loConsulta.AppendLine("			AND Documentacion.Cod_Tip = 'INFPEDCLI'")
            loConsulta.AppendLine("WHERE	")
            If (lcParametro9Desde = "SI") Then
                loConsulta.AppendLine(" 	    COALESCE(Autorizados.Val_Log, 0) = 0")
                loConsulta.AppendLine(" 	AND ")
            End If
            loConsulta.AppendLine(" 	    Pedidos.Documento BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("     AND " & lcParametro0Hasta)
            loConsulta.AppendLine("     AND Pedidos.Fec_Ini BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("     AND " & lcParametro1Hasta)
            loConsulta.AppendLine("     AND Pedidos.Cod_Cli BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("     AND " & lcParametro2Hasta)
            loConsulta.AppendLine("     AND Pedidos.Cod_Ven BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("     AND " & lcParametro3Hasta)
            loConsulta.AppendLine("     AND Pedidos.Status IN (" & lcParametro4Desde & ")")
            loConsulta.AppendLine("     AND Pedidos.Cod_Tra BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("     AND " & lcParametro5Hasta)
            loConsulta.AppendLine("     AND Pedidos.Cod_Mon BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("     AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 	AND Pedidos.Cod_mon BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine(" 	AND " & lcParametro6Hasta)
            loConsulta.AppendLine("     AND Pedidos.Cod_Rev between " & lcParametro7Desde)
            loConsulta.AppendLine("    	AND " & lcParametro7Hasta)
            loConsulta.AppendLine("     AND Pedidos.Cod_Suc between " & lcParametro8Desde)
            loConsulta.AppendLine("    	AND " & lcParametro8Hasta)
            loConsulta.AppendLine("ORDER BY        " & lcOrdenamiento)
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos()

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("MHL_rPedidos_PorAutorizar", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMHL_rPedidos_PorAutorizar.ReportSource = loObjetoReporte


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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 23/07/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
