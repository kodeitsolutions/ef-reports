'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grUltimas_30SMontosProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class grUltimas_30SMontosProveedores
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
            loComandoSeleccionar.AppendLine(" 		Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine(" 		Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine(" 		#temp_semanas.Week,")
            loComandoSeleccionar.AppendLine(" 		#temp_semanas.Mon_Net")
            loComandoSeleccionar.AppendLine(" INTO #temp_Prosem")
            loComandoSeleccionar.AppendLine(" FROM Proveedores,#temp_semanas")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		#temp_Prosem.Cod_Pro,")
            loComandoSeleccionar.AppendLine(" 		#temp_Prosem.Nom_Pro,")
            loComandoSeleccionar.AppendLine(" 		#temp_Prosem.Week,")
            loComandoSeleccionar.AppendLine(" 		CASE ")
            loComandoSeleccionar.AppendLine(" 			WHEN DATEDIFF(WEEK,Compras.Fec_Ini,GETDATE()) = #temp_Prosem.Week")
            loComandoSeleccionar.AppendLine(" 				AND ISNULL(Compras.Mon_Net,0) > 0 ")
            loComandoSeleccionar.AppendLine(" 				AND Compras.Status <> 'Anulado' THEN Compras.Mon_Net ")
            loComandoSeleccionar.AppendLine(" 			ELSE 0 ")
            loComandoSeleccionar.AppendLine(" 		END AS Mon_Net")
            loComandoSeleccionar.AppendLine(" INTO #temp_detPro")
            loComandoSeleccionar.AppendLine(" FROM #temp_Prosem")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN Compras ON #temp_Prosem.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT  ")
            loComandoSeleccionar.AppendLine(" 		#temp_detPro.Cod_Pro, ")
            loComandoSeleccionar.AppendLine(" 		SUM(#temp_detPro.Mon_Net) as Mon_Net_Pro ")
            loComandoSeleccionar.AppendLine(" Into #temp_detPro2")
            loComandoSeleccionar.AppendLine(" FROM #temp_detPro")
            loComandoSeleccionar.AppendLine(" GROUP BY #temp_detPro.Cod_Pro")
            loComandoSeleccionar.AppendLine(" ORDER BY #temp_detPro.Cod_Pro ")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 		#temp_detPro.Cod_Pro,")
            loComandoSeleccionar.AppendLine(" 		#temp_detPro.Nom_Pro,")
            loComandoSeleccionar.AppendLine(" 		#temp_detPro.Week,")
            loComandoSeleccionar.AppendLine(" 		SUM(#temp_detPro.Mon_Net) as Mon_Net")
            loComandoSeleccionar.AppendLine(" FROM #temp_detPro, #temp_detPro2")
            loComandoSeleccionar.AppendLine(" WHERE #temp_detPro.cod_Pro = #temp_detPro2.cod_Pro AND #temp_detPro2.Mon_Net_Pro > 0")
            loComandoSeleccionar.AppendLine(" GROUP BY #temp_detPro.Cod_Pro,#temp_detPro.Nom_Pro,#temp_detPro.Week")
            loComandoSeleccionar.AppendLine(" ORDER BY #temp_detPro.Week,#temp_detPro.Cod_Pro")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("grUltimas_30SMontosProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrUltimas_30SMontosProveedores.ReportSource = loObjetoReporte

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
' CMS: 26/06/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
