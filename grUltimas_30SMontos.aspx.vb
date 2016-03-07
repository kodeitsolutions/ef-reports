'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grUltimas_30SMontos"
'-------------------------------------------------------------------------------------------'
Partial Class grUltimas_30SMontos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT 0 AS Week, 0 AS Mon_Net INTO #temp_semanas")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 1 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 2 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 3 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 4 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 5 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 6 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 7 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 8 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 9 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 10 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 11 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 12 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 13 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 14 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 15 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 16 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 17 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 18 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 19 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 20 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 21 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 22 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 23 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 24 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 25 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 26 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 27 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 28 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 29 AS Week, 0 AS Mon_Net")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		Vendedores.Cod_Ven,")
            loComandoSeleccionar.AppendLine(" 		Vendedores.Nom_Ven,")
            loComandoSeleccionar.AppendLine(" 		#temp_semanas.Week,")
            loComandoSeleccionar.AppendLine(" 		#temp_semanas.Mon_Net")
            loComandoSeleccionar.AppendLine(" INTO #temp_vensem")
            loComandoSeleccionar.AppendLine(" FROM Vendedores,#temp_semanas")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		#temp_vensem.Cod_Ven,")
            loComandoSeleccionar.AppendLine(" 		#temp_vensem.Nom_Ven,")
            loComandoSeleccionar.AppendLine(" 		#temp_vensem.Week,")
            loComandoSeleccionar.AppendLine(" 		CASE ")
            loComandoSeleccionar.AppendLine(" 			WHEN DATEDIFF(WEEK,Facturas.Fec_Ini,GETDATE()) = #temp_vensem.Week")
            loComandoSeleccionar.AppendLine(" 				AND ISNULL(Facturas.Mon_Net,0) > 0 ")
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Status <> 'Anulado' THEN Facturas.Mon_Net ")
            loComandoSeleccionar.AppendLine(" 			ELSE 0 ")
            loComandoSeleccionar.AppendLine(" 		END AS Mon_Net")
            loComandoSeleccionar.AppendLine(" INTO #temp_detven")
            loComandoSeleccionar.AppendLine(" FROM #temp_vensem")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN Facturas ON #temp_vensem.Cod_Ven = Facturas.Cod_Ven")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		#temp_detven.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 		SUM(#temp_detven.Mon_Net) as Mon_Net_Ven ")
            loComandoSeleccionar.AppendLine(" Into #temp_detven2")
            loComandoSeleccionar.AppendLine(" FROM #temp_detven")
            loComandoSeleccionar.AppendLine(" GROUP BY #temp_detven.Cod_Ven")
            loComandoSeleccionar.AppendLine(" ORDER BY #temp_detven.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		#temp_detven.Cod_Ven,")
            loComandoSeleccionar.AppendLine(" 		#temp_detven.Nom_Ven,")
            loComandoSeleccionar.AppendLine(" 		#temp_detven.Week,")
            loComandoSeleccionar.AppendLine(" 		SUM(#temp_detven.Mon_Net) as Mon_Net")
            loComandoSeleccionar.AppendLine(" FROM #temp_detven, #temp_detven2")
            loComandoSeleccionar.AppendLine(" WHERE #temp_detven.cod_ven = #temp_detven2.cod_ven AND #temp_detven2.Mon_Net_Ven > 0")
            loComandoSeleccionar.AppendLine(" GROUP BY #temp_detven.Cod_Ven,#temp_detven.Nom_Ven,#temp_detven.Week")
            loComandoSeleccionar.AppendLine(" ORDER BY #temp_detven.Week,#temp_detven.Cod_Ven")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("grUltimas_30SMontos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrUltimas_30SMontos.ReportSource = loObjetoReporte

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
' Douglas Cortez 27/04/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
