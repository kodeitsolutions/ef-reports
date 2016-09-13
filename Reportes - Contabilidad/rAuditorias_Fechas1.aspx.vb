'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAuditorias_Fechas1"
'-------------------------------------------------------------------------------------------'
Partial Class rAuditorias_Fechas1
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
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpAuditorias(Fecha2 NCHAR(10) COLLATE Database_Default, ")
            loConsulta.AppendLine("							Fecha NCHAR(10) COLLATE Database_Default, ")
            loConsulta.AppendLine("							Cod_Usu NCHAR(10) COLLATE Database_Default,")
            loConsulta.AppendLine("							Accion NCHAR(10) COLLATE Database_Default,  ")
            loConsulta.AppendLine("							Cantidad INTEGER,")
            loConsulta.AppendLine("	                        Dias INTEGER) ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT	INTO #tmpAuditorias(Fecha2, Fecha, Cod_Usu, Accion, Cantidad,Dias)")
            loConsulta.AppendLine("SELECT	CONVERT(NCHAR(10), Registro, 112)	AS Fecha2,	")
            loConsulta.AppendLine("			CONVERT(NCHAR(10), Registro, 103)	AS Fecha, ")
            loConsulta.AppendLine("			Cod_Usu								AS Cod_Usu,")
            loConsulta.AppendLine("			Accion								AS Accion,")
            loConsulta.AppendLine("			COUNT(*)							AS Cantidad,")
            loConsulta.AppendLine("			dbo.udf_GetISOWeekDay(Registro)         AS  Dias ")
            loConsulta.AppendLine("FROM 	Auditorias ")
            loConsulta.AppendLine("WHERE	Registro	BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Cod_Usu	BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("			AND " & lcParametro1Hasta)
            loConsulta.AppendLine("		AND Tabla	BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("			AND " & lcParametro2Hasta)
            loConsulta.AppendLine("		AND Opcion	BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("			AND " & lcParametro3Hasta)
            loConsulta.AppendLine("		AND Documento	BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("		AND Codigo	BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("			AND " & lcParametro5Hasta)
            loConsulta.AppendLine("		AND Cod_Emp	BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("			AND " & lcParametro6Hasta)
            loConsulta.AppendLine("GROUP BY	CONVERT(NCHAR(10), Registro, 112),")
            loConsulta.AppendLine("			CONVERT(NCHAR(10), Registro, 103),")
            loConsulta.AppendLine("			Cod_Usu, Accion,dbo.udf_GetISOWeekDay(Registro)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Fecha, ")
            loConsulta.AppendLine("			Fecha2, ")
            loConsulta.AppendLine("         (CASE Dias ")
            loConsulta.AppendLine("                 WHEN 1 THEN 'LUNES'")
            loConsulta.AppendLine("                 WHEN 2 THEN 'MARTES'")
            loConsulta.AppendLine("                 WHEN 3 THEN 'MIERCOLES'")
            loConsulta.AppendLine("                 WHEN 4 THEN 'JUEVES'")
            loConsulta.AppendLine("                 WHEN 5 THEN 'VIERNES'")
            loConsulta.AppendLine("                 WHEN 6 THEN 'SABADO'")
            loConsulta.AppendLine("                 WHEN 7 THEN 'DOMINGO'")
            loConsulta.AppendLine("         END) AS Dia,")
            loConsulta.AppendLine("			COUNT(DISTINCT Cod_Usu) AS CantidadUsuarios,")
            loConsulta.AppendLine("			SUM((CASE WHEN Accion = 'Apertura'	THEN Cantidad ELSE 0 END)) As Apertura, ")
            loConsulta.AppendLine("			SUM((CASE WHEN Accion = 'Agregar'	THEN Cantidad ELSE 0 END)) As Agregar, ")
            loConsulta.AppendLine("			SUM((CASE WHEN Accion = 'Anular'	THEN Cantidad ELSE 0 END)) As Anular, ")
            loConsulta.AppendLine("			SUM((CASE WHEN Accion = 'Eliminar'	THEN Cantidad ELSE 0 END)) As Eliminar, ")
            loConsulta.AppendLine("			SUM((CASE WHEN Accion = 'Modificar'	THEN Cantidad ELSE 0 END)) As Modificar,")
            loConsulta.AppendLine("			SUM((CASE WHEN Accion = 'Reporte'	THEN Cantidad ELSE 0 END)) As Reporte, ")
            loConsulta.AppendLine("			SUM(Cantidad) As TotalDia ")
            loConsulta.AppendLine("FROM		#tmpAuditorias")
            loConsulta.AppendLine("GROUP BY	Fecha2, Fecha,Dias ")
            loConsulta.AppendLine("ORDER BY	" & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())
			
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAuditorias_Fechas1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrAuditorias_Fechas1.ReportSource = loObjetoReporte

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
' JFP: 15/12/08: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
' GCR: 26/03/09: Ajustes al diseño.															'
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros.										'
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/12: Se incluyo el filtro de la empresa.										'
'-------------------------------------------------------------------------------------------'
' RJG: 21/01/13: Se agregó el número de usuarios del día. Se ajustó el SELECT para hacerlo	'
'				 más rápido.																'
'-------------------------------------------------------------------------------------------'
' RJG: 11/03/13: Se corrigió bug en totales de las auditorias.  							'
'-------------------------------------------------------------------------------------------'
' EAG: 28/08/15: Se agrego el campo dia.                          							'
'-------------------------------------------------------------------------------------------'
