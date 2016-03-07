'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rListado_Alicuotas_CCMV"
'-------------------------------------------------------------------------------------------'
Partial Class rListado_Alicuotas_CCMV
	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT      Articulos.Cod_Art	                                            AS Local, ")
			loComandoSeleccionar.AppendLine("            Articulos.Nom_Art	                                            AS Nom_Art, ")
			loComandoSeleccionar.AppendLine("            CAST(Articulos.Por_Ali AS DECIMAL(38,14))                       AS Por_Ali, ")
			loComandoSeleccionar.AppendLine("            CAST(Articulos.Por_Ali2 AS DECIMAL(38,14))	                    AS Por_Ali2, ")
			loComandoSeleccionar.AppendLine("            CAST(Articulos.Por_Ali3 AS DECIMAL(38,14))                      AS Por_Ali3, ")
			loComandoSeleccionar.AppendLine("            CAST(CAST(Articulos.Por_Ali AS DECIMAL(38,14))  AS VARCHAR(50)) AS Alicuota1, ")
			loComandoSeleccionar.AppendLine("            CAST(CAST(Articulos.Por_Ali2 AS DECIMAL(38,14)) AS VARCHAR(50)) AS Alicuota2, ")
			loComandoSeleccionar.AppendLine("            CAST(CAST(Articulos.Por_Ali3 AS DECIMAL(38,14)) AS VARCHAR(50)) AS Alicuota3, ")
			loComandoSeleccionar.AppendLine("            COALESCE(Datos_Cliente.Nom_Cli, '[NO ASIGNADO]')	AS Nom_Cli,")
			loComandoSeleccionar.AppendLine("            (CASE Articulos.Status ")
			loComandoSeleccionar.AppendLine("            	WHEN 'A' THEN 'Activo'")
			loComandoSeleccionar.AppendLine("            	WHEN 'I' THEN 'Inactivo'")
			loComandoSeleccionar.AppendLine("            	WHEN 'S' THEN 'Suspendido'")
			loComandoSeleccionar.AppendLine("            END) AS Status")
			loComandoSeleccionar.AppendLine("INTO       #tmpArticulos")
			loComandoSeleccionar.AppendLine("FROM		Articulos ")
			loComandoSeleccionar.AppendLine("    LEFT JOIN ( SELECT  precios_clientes.Cod_Art, ")
			loComandoSeleccionar.AppendLine("                        MAX(Clientes_Activos.nom_cli) AS Nom_Cli")
			loComandoSeleccionar.AppendLine("                FROM    precios_clientes ")
			loComandoSeleccionar.AppendLine("                    JOIN(  SELECT  Clientes.Cod_Cli , clientes.Nom_Cli")
			loComandoSeleccionar.AppendLine("                            FROM    Clientes ")
			loComandoSeleccionar.AppendLine("                            JOIN articulos on articulos.cod_art = clientes.cod_cli")
			loComandoSeleccionar.AppendLine("                                AND (articulos.por_ali+articulos.Por_Ali2)>0")
			loComandoSeleccionar.AppendLine("                                AND clientes.status ='A'")
			loComandoSeleccionar.AppendLine("                        ) AS Clientes_Activos")
			loComandoSeleccionar.AppendLine("	                ON  Clientes_Activos.Cod_Cli = precios_clientes.cod_reg")
			loComandoSeleccionar.AppendLine("	                AND precios_clientes.origen = 'clientes' ")
			loComandoSeleccionar.AppendLine("                GROUP BY precios_clientes.Cod_Art")
			loComandoSeleccionar.AppendLine("                ) AS Datos_Cliente ")
			loComandoSeleccionar.AppendLine("	    ON      Datos_Cliente.cod_art = Articulos.cod_art")
			loComandoSeleccionar.AppendLine("WHERE		(Articulos.Por_Ali + Articulos.Por_Ali2)> 0")
			loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep    BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("				AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Sec    BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar    BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine("				AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Cla    BETWEEN " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Tip    BETWEEN " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine("				AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Pro    BETWEEN " & lcParametro6Desde)
			loComandoSeleccionar.AppendLine("				AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT      #tmpArticulos.*,")
			loComandoSeleccionar.AppendLine("            CAST(Totales.Total_1 AS VARCHAR(50)) AS Total_1,")
			loComandoSeleccionar.AppendLine("            CAST(Totales.Total_2 AS VARCHAR(50)) AS Total_2,")
			loComandoSeleccionar.AppendLine("            CAST(Totales.Total_3 AS VARCHAR(50)) AS Total_3")
			loComandoSeleccionar.AppendLine("FROM        #tmpArticulos")
			loComandoSeleccionar.AppendLine("    CROSS JOIN(   SELECT  SUM(por_Ali)    AS Total_1,")
			loComandoSeleccionar.AppendLine("                    SUM(por_Ali2)   AS Total_2,")
			loComandoSeleccionar.AppendLine("                    SUM(por_Ali3)   AS Total_3")
			loComandoSeleccionar.AppendLine("            FROM #tmpArticulos")
			loComandoSeleccionar.AppendLine("            ) As Totales")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("--DROP TABLE #tmpArticulos")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rListado_Alicuotas_CCMV", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrArticulos.ReportSource = loObjetoReporte

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
' RJG: 01/08/08: Código inicial a partir de rArticulos.										'
'-------------------------------------------------------------------------------------------'
' RJG: 29/04/14: Ajustado para mostrar tambien los locale no asignados a clientes.			'
'-------------------------------------------------------------------------------------------'
