Imports System.Data
Partial Class MCL_rProduccion

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))

        Dim lcEmpresa As String = cusAplicacion.goEmpresa.pcCodigo

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(10) = " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(10) =  " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Renglones_Ajustes.Cod_Art				            AS Cod_Art, ")
            lcComandoSeleccionar.AppendLine("		Articulos.Nom_Art						            AS Nom_Art,")
            lcComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cod_Lot,'')	            AS Lote,")
            lcComandoSeleccionar.AppendLine("       COALESCE(Operaciones_Lotes.Cantidad,0)	            AS Cant_Lote,")
            lcComandoSeleccionar.AppendLine("		COALESCE(Piezas.Res_Num, 0)				            AS Piezas,")
            lcComandoSeleccionar.AppendLine("       COALESCE(R_Produccion_Consumido.Can_Art1, 0)        AS Cant_Consumido,")
            lcComandoSeleccionar.AppendLine("       COALESCE(SUBSTRING(RTRIM(Operaciones_Lotes.Cod_Lot),0,3),'')        AS Horno,")
            lcComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Art_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Art_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Alm_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm  FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Alm_Hasta")
            lcComandoSeleccionar.AppendLine("INTO #tmpColadas")
            lcComandoSeleccionar.AppendLine("FROM Ajustes ")
            lcComandoSeleccionar.AppendLine("   JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art")
            lcComandoSeleccionar.AppendLine("   LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("       AND Articulos.Cod_Art = Operaciones_Lotes.Cod_Art")
            lcComandoSeleccionar.AppendLine("       AND Renglones_Ajustes.Cod_Alm = Operaciones_Lotes.Cod_Alm")
            lcComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Tip_Doc = 'Ajustes_Inventarios' ")
            lcComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Ajustes.Renglon")
            lcComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Ajustes_Inventarios'")
            lcComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            lcComandoSeleccionar.AppendLine("		AND Renglones_Ajustes.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            lcComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'AINV-NPIEZ'")
            lcComandoSeleccionar.AppendLine("		AND Piezas.Res_Num > 0")
            'lcComandoSeleccionar.AppendLine("   LEFT JOIN Ajustes AS Entrada_Produccion On Entrada.Num_Doc = Entrada_Produccion.Documento")
            lcComandoSeleccionar.AppendLine("   LEFT JOIN Ajustes AS Produccion_Consumido ON Produccion_Consumido.Documento = Ajustes.Caracter1")
            lcComandoSeleccionar.AppendLine("   LEFT JOIN Renglones_Ajustes AS R_Produccion_Consumido ON Produccion_Consumido.Documento = R_Produccion_Consumido.Documento")
            lcComandoSeleccionar.AppendLine("       AND R_Produccion_Consumido.Cod_Tip = 'S02'")
            lcComandoSeleccionar.AppendLine("   LEFT JOIN Operaciones_Lotes AS Consumido ON Consumido.Num_Doc = Produccion_Consumido.Documento")
            lcComandoSeleccionar.AppendLine("       AND Consumido.Tip_Doc = 'Ajustes_Inventarios' ")
            lcComandoSeleccionar.AppendLine("		AND Consumido.Tip_Ope = 'Salida'")
            lcComandoSeleccionar.AppendLine("       AND Consumido.Cod_Art = R_Produccion_Consumido.Cod_Art")
            lcComandoSeleccionar.AppendLine("       AND Consumido.Cod_Alm = R_Produccion_Consumido.Cod_Alm")
            lcComandoSeleccionar.AppendLine("		AND Consumido.Ren_Ori = R_Produccion_Consumido.Renglon")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Ajustes.Cod_Tip = 'E02'")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            lcComandoSeleccionar.AppendLine("   AND Ajustes.Status IN ( " & lcParametro3Desde & ")")
            lcComandoSeleccionar.AppendLine("ORDER BY Lote ASC;")

            lcComandoSeleccionar.AppendLine("WITH Totales (Cod_Art, Lote, Cant_Consumido) AS")
            lcComandoSeleccionar.AppendLine("(")
            lcComandoSeleccionar.AppendLine("   SELECT Cod_Art, Lote, SUM(Cant_Consumido) AS Cant_Consumido")
            lcComandoSeleccionar.AppendLine("   FROM #tmpColadas")
            lcComandoSeleccionar.AppendLine("   GROUP BY Cod_Art,Lote")
            lcComandoSeleccionar.AppendLine(")")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("UPDATE #tmpColadas ")
            lcComandoSeleccionar.AppendLine("SET #tmpColadas.Cant_Consumido = Totales.Cant_Consumido")
            lcComandoSeleccionar.AppendLine("FROM #tmpColadas")
            lcComandoSeleccionar.AppendLine("	INNER JOIN Totales ON Totales.Cod_Art = #tmpColadas.Cod_Art")
            lcComandoSeleccionar.AppendLine("		AND Totales.Lote = #tmpColadas.Lote")
            lcComandoSeleccionar.AppendLine("WHERE #tmpColadas.Lote = Totales.Lote")
            lcComandoSeleccionar.AppendLine("   AND #tmpColadas.Cod_Art = Totales.Cod_Art")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT DISTINCT * FROM #tmpColadas ORDER BY Lote ASC")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpColadas")


            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("MCL_rProduccion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMCL_rProduccion.ReportSource = loObjetoReporte

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
' JJD: 06/12/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' GS: 02/02/17: El orden de los grupos en el rpt es por el caso que el mismo lote llegue en recepciones separadas.
'-------------------------------------------------------------------------------------------'

