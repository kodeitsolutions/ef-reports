Imports System.Data
Partial Class MCL_fNotas_Despacho

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT Ajustes.Documento						                            AS Documento, ")
            loComandoSeleccionar.AppendLine("		Ajustes.Status							                            AS Estatus, ")
            loComandoSeleccionar.AppendLine("       Ajustes.Fec_Ini							                            AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("       Renglones_Ajustes.Renglon				                            AS Renglon, ")
            loComandoSeleccionar.AppendLine("       Renglones_Ajustes.Cod_Art				                            AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art						                            AS Nom_Art,")
            loComandoSeleccionar.AppendLine("		Secciones.Nom_Sec					                                AS Tipo,")
            loComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cod_Lot,'')	                            AS Lote,")
            loComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cantidad,Renglones_Ajustes.Can_Art1,0)   AS Cant_Lote,")
            loComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)				                            AS Piezas,")
            loComandoSeleccionar.AppendLine("       COALESCE(R_Produccion_Consumido.Can_Art1, 0)                        AS Cant_Consumido")
            loComandoSeleccionar.AppendLine("INTO #tmpColadas")
            loComandoSeleccionar.AppendLine("FROM Ajustes ")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art")
            loComandoSeleccionar.AppendLine("   JOIN Secciones ON Articulos.Cod_Sec = Secciones.Cod_Sec")
            loComandoSeleccionar.AppendLine("       AND Secciones.Cod_Dep = Articulos.Cod_Dep")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones_Ajustes.Documento")
            loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("       AND Renglones_Ajustes.Cod_Alm = Operaciones_Lotes.Cod_Alm")
            loComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Tip_Doc = 'Ajustes_Inventarios' ")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Ajustes.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Salida'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Ajustes_Inventarios'")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND Renglones_Ajustes.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'AINV-NPIEZ'")
            loComandoSeleccionar.AppendLine("		AND Piezas.Res_Num > 0")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Operaciones_Lotes AS Entrada ON Entrada.Cod_Lot = Operaciones_Lotes.Cod_Lot")
            loComandoSeleccionar.AppendLine("       AND Entrada.Tip_Doc = 'Ajustes_Inventarios' ")
            loComandoSeleccionar.AppendLine("		AND Entrada.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("       AND Entrada.Cod_Art = Operaciones_Lotes.Cod_Art")
            loComandoSeleccionar.AppendLine("       AND Entrada.Cod_Alm = Operaciones_Lotes.Cod_Alm")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Ajustes AS Entrada_Produccion On Entrada.Num_Doc = Entrada_Produccion.Documento")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Ajustes AS Produccion_Consumido ON Produccion_Consumido.Documento = Entrada_Produccion.Caracter1")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Renglones_Ajustes AS R_Produccion_Consumido ON Produccion_Consumido.Documento = R_Produccion_Consumido.Documento")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Tip = 'S10' ")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Art ASC, Lote ASC;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("WITH Totales (Cod_Art, Lote, Cant_Consumido) AS")
            loComandoSeleccionar.AppendLine("(")
            loComandoSeleccionar.AppendLine("   SELECT Cod_Art, Lote, SUM(Cant_Consumido) AS Cant_Consumido")
            loComandoSeleccionar.AppendLine("   FROM #tmpColadas")
            loComandoSeleccionar.AppendLine("   GROUP BY Cod_Art,Lote")
            loComandoSeleccionar.AppendLine(")")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE #tmpColadas ")
            loComandoSeleccionar.AppendLine("SET #tmpColadas.Cant_Consumido = Totales.Cant_Consumido")
            loComandoSeleccionar.AppendLine("FROM #tmpColadas")
            loComandoSeleccionar.AppendLine("	INNER JOIN Totales ON Totales.Cod_Art = #tmpColadas.Cod_Art")
            loComandoSeleccionar.AppendLine("		AND Totales.Lote = #tmpColadas.Lote")
            loComandoSeleccionar.AppendLine("WHERE #tmpColadas.Lote = Totales.Lote")
            loComandoSeleccionar.AppendLine("   AND #tmpColadas.Cod_Art = Totales.Cod_Art")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT DISTINCT * FROM #tmpColadas")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpColadas")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

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


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("MCL_fNotas_Despacho", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMCL_fNotas_Despacho.ReportSource = loObjetoReporte

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