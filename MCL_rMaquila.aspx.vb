Imports System.Data
Partial Class MCL_rMaquila

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

        Dim lcEmpresa As String = cusAplicacion.goEmpresa.pcCodigo

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Dim lcComandoSeleccionar As New StringBuilder()

        Try

            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcDcto_Desde AS VARCHAR(10) = " & lcParametro1Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcDcto_Hasta AS VARCHAR(10) =  " & lcParametro1Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(10) = " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(10) = " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Desde AS VARCHAR(10) = " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Hasta AS VARCHAR(10) = " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde AS VARCHAR(10) = " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta AS VARCHAR(10) = " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("")

            lcComandoSeleccionar.AppendLine("SELECT Renglones_Ajustes.Cod_Art																		AS	Cod_Art, ")
            lcComandoSeleccionar.AppendLine("		Articulos.Nom_Art																				AS	Nom_Art, ")
            lcComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada' THEN Renglones_Ajustes.Can_Art1 ELSE 0 END)	AS	Can_IniE, ")
            lcComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Renglones_Ajustes.Can_Art1 ELSE 0 END)		AS	Can_IniS ")
            lcComandoSeleccionar.AppendLine("INTO #tmpInicial")
            lcComandoSeleccionar.AppendLine("FROM Ajustes ")
            lcComandoSeleccionar.AppendLine("   JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON  Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
            lcComandoSeleccionar.AppendLine(" WHERE	Ajustes.Status = 'Confirmado' ")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Fec_Ini < @ldFecha_Desde")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Documento BETWEEN @lcDcto_Desde AND @lcDcto_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            lcComandoSeleccionar.AppendLine("GROUP BY Renglones_Ajustes.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("   Articulos.Nom_Art ")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT Ajustes.documento							AS Documento,")
            lcComandoSeleccionar.AppendLine("		Ajustes.Fec_Ini								AS Fec_Ini,")
            lcComandoSeleccionar.AppendLine("		Renglones_Ajustes.Renglon					AS Renglon,")
            lcComandoSeleccionar.AppendLine("		Renglones_Ajustes.Cod_Art					AS Cod_Art, ")
            lcComandoSeleccionar.AppendLine("		Articulos.Nom_Art							AS Nom_Art,")
            lcComandoSeleccionar.AppendLine("		Renglones_Ajustes.Can_Art1					AS Can_Art1,")
            lcComandoSeleccionar.AppendLine("		Renglones_Ajustes.Cod_Alm					AS Cod_Alm,")
            lcComandoSeleccionar.AppendLine("		COALESCE((SELECT (SUM(Can_IniE)- SUM(Can_IniS)) FROM	#tmpInicial ),0) AS Inicial,")
            lcComandoSeleccionar.AppendLine("		MONTH(Ajustes.Fec_Ini)			AS Mes")
            lcComandoSeleccionar.AppendLine("INTO #tmpRecepciones")
            lcComandoSeleccionar.AppendLine("FROM Ajustes ")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Ajustes.Cod_Art = Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Ajustes.Cod_Tip = 'E05'")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Documento BETWEEN @lcDcto_Desde AND @lcDcto_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Status = 'Confirmado'")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	MONTH(Ajustes.Fec_Ini)			AS Mes,")
            lcComandoSeleccionar.AppendLine("		SUM(Renglones_Ajustes.Can_Art1)	AS Cantidad,")
            lcComandoSeleccionar.AppendLine("		CAST(0 AS DECIMAL(28,10))		AS Disponible,")
            lcComandoSeleccionar.AppendLine("       CAST(0 AS DECIMAL(28,10))		AS TotalMes,")
            lcComandoSeleccionar.AppendLine("		(SELECT SUM(can_art1) FROM #tmpRecepciones WHERE Mes = MONTH(Ajustes.Fec_Ini))	AS Recibido,")
            lcComandoSeleccionar.AppendLine("		(SELECT TOP 1 Inicial FROM #tmpRecepciones WHERE Mes = #tmpRecepciones.Mes)		AS Inicial")
            lcComandoSeleccionar.AppendLine("INTO #tmpProcesado")
            lcComandoSeleccionar.AppendLine("FROM Ajustes")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Ajustes.Cod_Art = Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Ajustes.Cod_Tip = 'S02'")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Documento BETWEEN @lcDcto_Desde AND @lcDcto_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Status = 'Confirmado'")
            lcComandoSeleccionar.AppendLine("GROUP BY MONTH(Ajustes.Fec_Ini)")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	MONTH(Ajustes.Fec_Ini)			AS Mes,")
            lcComandoSeleccionar.AppendLine("		SUM(Renglones_Ajustes.Can_Art1)	AS Cantidad,")
            lcComandoSeleccionar.AppendLine("		Articulos.Cod_Uni1				AS Unidad,")
            lcComandoSeleccionar.AppendLine("		CAST(0 AS DECIMAL(28,10))		AS Disponible")
            lcComandoSeleccionar.AppendLine("INTO #tmpDiferencia")
            lcComandoSeleccionar.AppendLine("FROM Ajustes")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Ajustes.Cod_Art = Articulos.Cod_Art	")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN #tmpProcesado ON #tmpProcesado.Mes = MONTH(Ajustes.Fec_Ini)")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Ajustes.Cod_Tip = 'S03'")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Documento BETWEEN @lcDcto_Desde AND @lcDcto_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Status = 'Confirmado'")
            lcComandoSeleccionar.AppendLine("GROUP BY MONTH(Ajustes.Fec_Ini), #tmpProcesado.Disponible, Articulos.Cod_Uni1")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("DECLARE @lnRenglon AS INT = 1")
            lcComandoSeleccionar.AppendLine("DECLARE @lnTotal AS INT = (SELECT MAX(Mes) FROM #tmpProcesado);")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("WHILE (@lnRenglon <= @lnTotal)")
            lcComandoSeleccionar.AppendLine("BEGIN")
            lcComandoSeleccionar.AppendLine("	IF @lnRenglon = 1")
            lcComandoSeleccionar.AppendLine("	BEGIN")
            lcComandoSeleccionar.AppendLine("		UPDATE #tmpProcesado SET Disponible = Inicial + Recibido - Cantidad, TotalMes = Inicial + Recibido WHERE Mes = @lnRenglon")
            lcComandoSeleccionar.AppendLine("	END")
            lcComandoSeleccionar.AppendLine("	ELSE")
            lcComandoSeleccionar.AppendLine("	BEGIN")
            lcComandoSeleccionar.AppendLine("		UPDATE #tmpProcesado ")
            lcComandoSeleccionar.AppendLine("		SET Disponible = (SELECT Disponible FROM #tmpProcesado WHERE Mes = @lnRenglon - 1) + Recibido - Cantidad - COALESCE((SELECT Cantidad FROM #tmpDiferencia WHERE Mes = @lnRenglon),0),")
            lcComandoSeleccionar.AppendLine("			TotalMes = (SELECT Disponible FROM #tmpProcesado WHERE Mes = @lnRenglon - 1) + Recibido WHERE Mes = @lnRenglon")
            lcComandoSeleccionar.AppendLine("	END")
            lcComandoSeleccionar.AppendLine("	")
            lcComandoSeleccionar.AppendLine("	SET @lnRenglon = @lnRenglon + 1 ")
            lcComandoSeleccionar.AppendLine("END")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("UPDATE #tmpDiferencia SET Disponible = COALESCE((SELECT Disponible FROM #tmpProcesado WHERE Mes = #tmpDiferencia.Mes),0)")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT	MONTH(Ajustes.Fec_Ini)			AS Mes,")
            lcComandoSeleccionar.AppendLine("		SUM(Renglones_Ajustes.Can_Art1)	AS Cantidad,")
            lcComandoSeleccionar.AppendLine("		SUM(COALESCE(Piezas.Res_Num,0))	AS Piezas")
            lcComandoSeleccionar.AppendLine("INTO #tmpObtenido")
            lcComandoSeleccionar.AppendLine("FROM Ajustes")
            lcComandoSeleccionar.AppendLine("	JOIN Renglones_Ajustes ON Ajustes.Documento = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Ajustes.Cod_Art = Articulos.Cod_Art")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Renglones_Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("		AND Renglones_Ajustes.Cod_Art = Operaciones_Lotes.Cod_Art")
            lcComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Ajustes.Renglon")
            lcComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Ajustes_Inventarios' AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Ajustes.Documento")
            lcComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Ajustes_Inventarios'")
            lcComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            lcComandoSeleccionar.AppendLine("		AND Renglones_Ajustes.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            lcComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'AINV-NPIEZ'")
            lcComandoSeleccionar.AppendLine("WHERE Renglones_Ajustes.Cod_Tip = 'E02'")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Documento BETWEEN @lcDcto_Desde AND @lcDcto_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            lcComandoSeleccionar.AppendLine("	AND Renglones_Ajustes.Cod_Alm = 'MO.PT'")
            lcComandoSeleccionar.AppendLine("	AND Ajustes.Status = 'Confirmado'")
            lcComandoSeleccionar.AppendLine("GROUP BY MONTH(Ajustes.Fec_Ini)")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("SELECT #tmpRecepciones.Documento,")
            lcComandoSeleccionar.AppendLine("		#tmpRecepciones.Fec_Ini,")
            lcComandoSeleccionar.AppendLine("		#tmpRecepciones.Renglon,")
            lcComandoSeleccionar.AppendLine("		#tmpRecepciones.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("		#tmpRecepciones.Nom_Art,")
            lcComandoSeleccionar.AppendLine("		#tmpRecepciones.Can_Art1,")
            lcComandoSeleccionar.AppendLine("		#tmpRecepciones.Cod_Alm,")
            lcComandoSeleccionar.AppendLine("		#tmpRecepciones.Inicial,")
            lcComandoSeleccionar.AppendLine("		#tmpRecepciones.Mes,")
            lcComandoSeleccionar.AppendLine("		#tmpProcesado.Cantidad                  AS Procesado,")
            lcComandoSeleccionar.AppendLine("		#tmpProcesado.Disponible,")
            lcComandoSeleccionar.AppendLine("       #tmpProcesado.TotalMes,")
            lcComandoSeleccionar.AppendLine("		#tmpObtenido.Piezas,")
            lcComandoSeleccionar.AppendLine("		#tmpObtenido.Cantidad					AS Obtenido,")
            lcComandoSeleccionar.AppendLine("		COALESCE(#tmpDiferencia.Cantidad,0)		AS Diferencia,")
            lcComandoSeleccionar.AppendLine("		COALESCE(#tmpDiferencia.Unidad,'')		AS Und_Diferencia,")
            lcComandoSeleccionar.AppendLine("		COALESCE(#tmpDiferencia.Disponible,0)	AS Disp_Diferencia,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcDcto_Desde <> '' ")
            lcComandoSeleccionar.AppendLine("			 THEN CONCAT(@lcDcto_Desde, ' - ', @lcDcto_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE 'N/E'")
            lcComandoSeleccionar.AppendLine("		END										AS Docs,")
            lcComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Art_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Art_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Dep_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Dep_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Desde AND Cod_Dep = @lcCodDep_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Sec_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Hasta AND Cod_Dep = @lcCodDep_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Sec_Hasta,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Desde <> ''")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Desde)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Alm_Desde,")
            lcComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Hasta <> 'zzzzzzz'")
            lcComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm  FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Hasta)")
            lcComandoSeleccionar.AppendLine("			 ELSE '' END				            AS Alm_Hasta ")
            lcComandoSeleccionar.AppendLine("FROM #tmpRecepciones")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN #tmpProcesado ON #tmpProcesado.Mes = #tmpRecepciones.Mes")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN #tmpObtenido ON #tmpObtenido.Mes = #tmpRecepciones.Mes")
            lcComandoSeleccionar.AppendLine("	LEFT JOIN #tmpDiferencia ON #tmpDiferencia.Mes = #tmpRecepciones.Mes")
            lcComandoSeleccionar.AppendLine("")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpInicial")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpRecepciones")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpProcesado")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpObtenido")
            lcComandoSeleccionar.AppendLine("DROP TABLE #tmpDiferencia")
            lcComandoSeleccionar.AppendLine("")


            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("MCL_rMaquila", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMCL_rMaquila.ReportSource = loObjetoReporte

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

