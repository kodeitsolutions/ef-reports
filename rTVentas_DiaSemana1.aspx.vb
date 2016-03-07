'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_DiaSemana1"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_DiaSemana1
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		(CASE  DATEPART(WeekDay,Facturas.Fec_Ini) ")
            loComandoSeleccionar.AppendLine("					WHEN 1 Then 'Domingo'")
            loComandoSeleccionar.AppendLine("					WHEN 2 Then 'Lunes'")
            loComandoSeleccionar.AppendLine("					WHEN 3 Then 'Martes'")
            loComandoSeleccionar.AppendLine("					WHEN 4 Then 'Miércoles'")
            loComandoSeleccionar.AppendLine("					WHEN 5 Then 'Jueves'")
            loComandoSeleccionar.AppendLine("					WHEN 6 Then 'Viernes'")
            loComandoSeleccionar.AppendLine("					WHEN 7 Then 'Sábado'")
            loComandoSeleccionar.AppendLine("  			END)								AS Dia, ")
            loComandoSeleccionar.AppendLine("			DATEPART(WEEKDAY,Facturas.Fec_Ini)	AS Dia_Sem, ")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.Can_Art1       	AS Can_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.Mon_Net        	AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.Cos_Ult1       	AS Mon_Cos ")
            loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            loComandoSeleccionar.AppendLine("FROM		Facturas  ")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas	ON Facturas.Documento = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	Articulos			ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE		Facturas.Status					<>  'Anulado' ")
            loComandoSeleccionar.AppendLine("		AND Facturas.Fec_Ini		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Cli      BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Ven      BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Dep     BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND Facturas.Status     IN (" & goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4)) & ")")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		SUM(Mon_Net) As  Tot_Net")
            loComandoSeleccionar.AppendLine("INTO		#curTemporal1 ")
            loComandoSeleccionar.AppendLine("FROM		#curTemporal ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Dia													AS  Dia, ")
            loComandoSeleccionar.AppendLine("			Dia_Sem												AS  Dia_Sem,")
            loComandoSeleccionar.AppendLine(" 			SUM(Can_Art)										AS  Can_Art,")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Net)										AS  Mon_Net,")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Cos)			 							AS  Mon_Cos,")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Net)-SUM(Mon_Cos)							AS  Mon_Gan,")
            loComandoSeleccionar.AppendLine(" 			(((SUM(Mon_Net)-SUM(Mon_Cos))/SUM(Mon_Net))*100)  	AS  Por_Gan,")
            loComandoSeleccionar.AppendLine(" 			CAST(1 AS DECIMAL(28, 10))							AS  Por_Ven ")
            loComandoSeleccionar.AppendLine("INTO		#curTemporal2 ")
            loComandoSeleccionar.AppendLine("FROM		#curTemporal ")
            loComandoSeleccionar.AppendLine("GROUP BY 	Dia, dia_sem ")
            loComandoSeleccionar.AppendLine("ORDER BY 	Dia_Sem")
            loComandoSeleccionar.AppendLine("")															
            loComandoSeleccionar.AppendLine("SELECT		#curTemporal2.Dia           As  Dia, ")
            loComandoSeleccionar.AppendLine("			#curTemporal2.Can_Art       As  Can_Art,")
            loComandoSeleccionar.AppendLine("			#curTemporal2.dia_sem       As  dia_sem,")
            loComandoSeleccionar.AppendLine("			#curTemporal2.Mon_Net       As  Mon_Net,")
            loComandoSeleccionar.AppendLine("			#curTemporal2.Mon_Cos       As  Mon_Cos,")
            loComandoSeleccionar.AppendLine("			#curTemporal2.Mon_Gan       As  Mon_Gan,")
            loComandoSeleccionar.AppendLine("			#curTemporal2.Por_Gan       As  Por_Gan,")
            loComandoSeleccionar.AppendLine("			#curTemporal1.Tot_Net       As  Tot_Net,")															
            loComandoSeleccionar.AppendLine("			(Mon_Net/Tot_Net)*100		As  Por_Ven ")
            loComandoSeleccionar.AppendLine("FROM		#curTemporal2")
            loComandoSeleccionar.AppendLine("	CROSS JOIN #curTemporal1 ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #curTemporal")
            loComandoSeleccionar.AppendLine("DROP TABLE #curTemporal1")
            loComandoSeleccionar.AppendLine("DROP TABLE #curTemporal2")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_DiaSemana1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTVentas_DiaSemana1.ReportSource = loObjetoReporte

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
' JFP: 07/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 14/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' RJG: 19/01/12: Corrección de filtro de fechas. Estandarización y optimización de código.	'
'-------------------------------------------------------------------------------------------'
