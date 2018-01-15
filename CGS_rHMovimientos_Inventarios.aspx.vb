'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rHMovimientos_Inventarios "
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rHMovimientos_Inventarios
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
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Desde AS VARCHAR(2) = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Hasta AS VARCHAR(2) = " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodLot_Desde AS VARCHAR(30) = " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodLot_Hasta AS VARCHAR(30) = " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @tmpArticulos AS TABLE(Cod_Art CHAR(30), Saldo DECIMAL(28,10)) ;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO @tmpArticulos ")
            loComandoSeleccionar.AppendLine("SELECT	Cod_Art, 0")
            loComandoSeleccionar.AppendLine("FROM Articulos")
            loComandoSeleccionar.AppendLine("WHERE Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	AND	Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            loComandoSeleccionar.AppendLine("		")
            loComandoSeleccionar.AppendLine("SELECT 'Ajustes'								AS	Operacion, 	")
            loComandoSeleccionar.AppendLine("		CASE WHEN Renglones_Ajustes.Tipo = 'Entrada'")
            loComandoSeleccionar.AppendLine("		     THEN 1 ELSE 2 END					AS	Orden,")
            loComandoSeleccionar.AppendLine("		Ajustes.Documento						AS	Documento,")
            loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Cod_Art				AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote, 	")
            loComandoSeleccionar.AppendLine("       (CASE WHEN Renglones_Ajustes.Tipo = 'Entrada'  THEN COALESCE(Piezas.Res_Num, 0) ELSE 0.0 END)	AS	Piezas_Ent,")
            loComandoSeleccionar.AppendLine("       (CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN COALESCE(Piezas.Res_Num, 0) ELSE 0.0 END)		AS	Piezas_Sal,")
            loComandoSeleccionar.AppendLine("		COALESCE(Longitud.Res_Num, 0)           AS  Longitud,")
            loComandoSeleccionar.AppendLine("		Ajustes.Fec_Ini							AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, 	")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm				        AS  Nom_Alm, 	")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Operaciones_Lotes.Cantidad ELSE 0.0 END)		AS	CanLte_Sal, ")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada'  THEN Operaciones_Lotes.Cantidad ELSE 0.0 END)		AS	CanLte_Ent, ")
            loComandoSeleccionar.AppendLine("       0.0										AS	Saldo_Piezas,")
            loComandoSeleccionar.AppendLine(" 		Articulos.Saldo				            AS	Saldo		")
            loComandoSeleccionar.AppendLine("INTO #curTemporal ")
            loComandoSeleccionar.AppendLine("FROM Ajustes")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
            loComandoSeleccionar.AppendLine(" 	JOIN Almacenes ON Renglones_Ajustes.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Tip_Doc = 'Ajustes_Inventarios'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = Renglones_Ajustes.Tipo")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Ajustes.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Ajustes.Cod_Art")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Ajustes_Inventarios'")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	    AND Mediciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND Renglones_Ajustes.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'AINV-NPIEZ'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Longitud ON Mediciones.Documento = Longitud.Documento")
            loComandoSeleccionar.AppendLine("		AND Longitud.Cod_Var = 'AINV-LARG'")
            loComandoSeleccionar.AppendLine("WHERE Ajustes.Status = 'Confirmado' ")
            loComandoSeleccionar.AppendLine(" 	AND	Renglones_Ajustes.Tipo IN ('Entrada', 'Salida') ")
            loComandoSeleccionar.AppendLine(" 	AND	Ajustes.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine(" 	AND	Renglones_Ajustes.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine(" 	")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("	")
            loComandoSeleccionar.AppendLine("SELECT 'Traslados'								AS	Operacion,	")
            loComandoSeleccionar.AppendLine("		2									    AS	Orden,")
            loComandoSeleccionar.AppendLine(" 		Traslados.Documento						AS	Documento,	")
            loComandoSeleccionar.AppendLine(" 		Renglones_Traslados.Cod_Art				AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote, 	")
            loComandoSeleccionar.AppendLine("       0.0										AS	Piezas_Ent,")
            loComandoSeleccionar.AppendLine("       COALESCE(Piezas.Res_Num, 0)				AS	Piezas_Sal,")
            loComandoSeleccionar.AppendLine("		COALESCE(Longitud.Res_Num, 0)           AS  Longitud,")
            loComandoSeleccionar.AppendLine(" 		Traslados.Fec_Ini						AS	Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 		Traslados.Alm_Ori						AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm				        AS  Nom_Alm, ")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cantidad				AS	CanLte_Sal, ")
            loComandoSeleccionar.AppendLine(" 		0.0										AS	CanLte_Ent, ")
            loComandoSeleccionar.AppendLine("       0.0										AS	Saldo_Piezas,")
            loComandoSeleccionar.AppendLine(" 		Articulos.Saldo				            AS	Saldo		")
            loComandoSeleccionar.AppendLine("FROM Traslados")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
            loComandoSeleccionar.AppendLine(" 	JOIN Almacenes ON Traslados.Alm_Ori = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Traslados.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Traslados'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Salida'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Traslados.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Traslados.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Traslados'")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	    AND Mediciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND Renglones_Traslados.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var IN ('TA-NPIEZ', 'TALM-NPIEZ')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Longitud ON Mediciones.Documento = Longitud.Documento")
            loComandoSeleccionar.AppendLine("		AND Longitud.Cod_Var = 'TA-LARG'")
            loComandoSeleccionar.AppendLine("WHERE Traslados.Status IN ('Confirmado', 'Procesado')	")
            loComandoSeleccionar.AppendLine(" 	AND Traslados.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine(" 	AND Traslados.Alm_Ori BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT 'Traslados'								AS	Operacion,	")
            loComandoSeleccionar.AppendLine("		1									    AS	Orden,")
            loComandoSeleccionar.AppendLine(" 		Traslados.Documento						AS	Documento,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Traslados.Cod_Art				AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote,  	")
            loComandoSeleccionar.AppendLine("       COALESCE(Piezas.Res_Num, 0)				AS	Piezas_Ent,")
            loComandoSeleccionar.AppendLine("       0.0										AS	Piezas_Sal,")
            loComandoSeleccionar.AppendLine("		COALESCE(Longitud.Res_Num, 0)           AS  Longitud,")
            loComandoSeleccionar.AppendLine(" 		Traslados.Fec_Ini						AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine(" 		CASE Traslados.Status					")
            loComandoSeleccionar.AppendLine(" 		    WHEN 'Confirmado'	THEN 'TRANSITO'")
            loComandoSeleccionar.AppendLine(" 		    WHEN 'Procesado'	THEN Traslados.Alm_Des")
            loComandoSeleccionar.AppendLine(" 		END										AS	Cod_Alm, 	")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm				        AS  Nom_Alm, 	")
            loComandoSeleccionar.AppendLine("		0.0										AS	CanLte_Sal, 	")
            loComandoSeleccionar.AppendLine(" 		Operaciones_Lotes.Cantidad				AS	CanLte_Ent, 	")
            loComandoSeleccionar.AppendLine("       0.0										AS	Saldo_Piezas,")
            loComandoSeleccionar.AppendLine(" 		Articulos.Saldo				            AS	Saldo		")
            loComandoSeleccionar.AppendLine("FROM   Traslados")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
            loComandoSeleccionar.AppendLine(" 	JOIN Almacenes ON Traslados.Alm_Des = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Traslados.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Traslados'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Traslados.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Traslados.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Traslados'")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	    AND Mediciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND Renglones_Traslados.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var IN ('TA-NPIEZ', 'TALM-NPIEZ')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Longitud ON Mediciones.Documento = Longitud.Documento")
            loComandoSeleccionar.AppendLine("		AND Longitud.Cod_Var = 'TA-LARG'")
            loComandoSeleccionar.AppendLine("WHERE Traslados.Status IN ('Confirmado', 'Procesado')	")
            loComandoSeleccionar.AppendLine("   AND Traslados.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine(" 	AND (CASE Traslados.Status					")
            loComandoSeleccionar.AppendLine(" 		    WHEN 'Confirmado' THEN 'TRANSITO'")
            loComandoSeleccionar.AppendLine(" 		    WHEN 'Procesado' THEN Traslados.Alm_Des")
            loComandoSeleccionar.AppendLine(" 		END) BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Recepciones'						    AS	Operacion,	")
            loComandoSeleccionar.AppendLine("		1									    AS	Orden,")
            loComandoSeleccionar.AppendLine(" 		Recepciones.Documento				    AS	Documento,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Recepciones.Cod_Art		    AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote, 	")
            loComandoSeleccionar.AppendLine("       COALESCE(Piezas.Res_Num, 0)				AS	Piezas_Ent,")
            loComandoSeleccionar.AppendLine("       0.0										AS	Piezas_Sal,")
            loComandoSeleccionar.AppendLine("		0.0                                     AS  Longitud,")
            loComandoSeleccionar.AppendLine(" 		Recepciones.Fec_Ini					    AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine(" 		Renglones_Recepciones.Cod_Alm		    AS	Cod_Alm, 	")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm				        AS  Nom_Alm, 	")
            loComandoSeleccionar.AppendLine("		0.0									    AS	CanLte_Sal, ")
            loComandoSeleccionar.AppendLine(" 		Operaciones_Lotes.Cantidad			    AS	CanLte_Ent,	")
            loComandoSeleccionar.AppendLine("       0.0										AS	Saldo_Piezas,")
            loComandoSeleccionar.AppendLine(" 		Articulos.Saldo				            AS	Saldo ")
            loComandoSeleccionar.AppendLine("FROM Recepciones")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
            loComandoSeleccionar.AppendLine(" 	JOIN Almacenes ON Renglones_Recepciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Recepciones'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Recepciones.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Recepciones.Cod_Art")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Recepciones'")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	    AND Mediciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND Renglones_Recepciones.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'NREC-NPIEZ'")
            loComandoSeleccionar.AppendLine("WHERE Recepciones.Status IN ('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("   AND Recepciones.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Recepciones.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Ordenes_Trabajo'						AS	Operacion,	")
            loComandoSeleccionar.AppendLine("		1										AS	Orden,")
            loComandoSeleccionar.AppendLine(" 		Encabezados.Documento					AS	Documento,	")
            loComandoSeleccionar.AppendLine(" 		Formulas.Cod_Art						AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote, 	")
            loComandoSeleccionar.AppendLine("       COALESCE(Piezas.Res_Num, 0)				AS	Piezas_Ent,")
            loComandoSeleccionar.AppendLine("       0.0										AS	Piezas_Sal,")
            loComandoSeleccionar.AppendLine("		COALESCE(Longitud.Res_Num, 0)           AS  Longitud,")
            loComandoSeleccionar.AppendLine(" 		Encabezados.Fec_Ini						AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine(" 		Encabezados.Cod_Alm						AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm						AS 	Nom_Alm,	")
            loComandoSeleccionar.AppendLine("		0.0										AS	CanLte_Sal, 	")
            loComandoSeleccionar.AppendLine(" 		Operaciones_Lotes.Cantidad				AS	CanLte_Ent,	")
            loComandoSeleccionar.AppendLine("       0.0										AS	Saldo_Piezas,")
            loComandoSeleccionar.AppendLine("		Articulos.Saldo							AS	Saldo ")
            loComandoSeleccionar.AppendLine("FROM Encabezados")
            loComandoSeleccionar.AppendLine("	JOIN Renglones ON Encabezados.Documento = Renglones.Documento ")
            loComandoSeleccionar.AppendLine("		AND Renglones.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("	JOIN Almacenes ON Encabezados.Cod_Alm = Almacenes.Cod_Alm	")
            loComandoSeleccionar.AppendLine("	JOIN Formulas ON Renglones.Cod_Reg = Formulas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Formulas.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Encabezados.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Encabezados'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Adicional = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Encabezados.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Encabezados'")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	    AND Mediciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND Renglones.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'OTRA-NPIEZ'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Longitud ON Mediciones.Documento = Longitud.Documento")
            loComandoSeleccionar.AppendLine("		AND Longitud.Cod_Var = 'OTRA-LARG'")
            loComandoSeleccionar.AppendLine("WHERE Encabezados.Status IN ('Confirmado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("	AND Encabezados.Origen = 'Ordenes de Trabajo'")
            loComandoSeleccionar.AppendLine("		AND Encabezados.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("		AND Encabezados.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Consumos_Produccion'					AS	Operacion,	")
            loComandoSeleccionar.AppendLine("		2										AS	Orden,")
            loComandoSeleccionar.AppendLine(" 		Encabezados.Documento					AS	Documento,	")
            loComandoSeleccionar.AppendLine(" 		Renglones.Cod_Art						AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote, 	")
            loComandoSeleccionar.AppendLine("       0.0										AS	Piezas_Ent,")
            loComandoSeleccionar.AppendLine("       COALESCE(Piezas.Res_Num, 0)				AS	Piezas_Sal,")
            loComandoSeleccionar.AppendLine("		COALESCE(Longitud.Res_Num, 0)           AS  Longitud,")
            loComandoSeleccionar.AppendLine(" 		Encabezados.Fec_Ini						AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine(" 		Encabezados.Cod_Alm						AS	Cod_Alm, ")
            loComandoSeleccionar.AppendLine("		Almacenes.Nom_Alm						AS 	Nom_Alm,	")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cantidad				AS	CanLte_Sal, 	")
            loComandoSeleccionar.AppendLine(" 		0.0				                        AS	CanLte_Ent,	")
            loComandoSeleccionar.AppendLine("       0.0										AS	Saldo_Piezas,")
            loComandoSeleccionar.AppendLine("		Articulos.Saldo							AS	Saldo ")
            loComandoSeleccionar.AppendLine("FROM Encabezados")
            loComandoSeleccionar.AppendLine("	JOIN Renglones ON Encabezados.Documento = Renglones.Documento ")
            loComandoSeleccionar.AppendLine("		AND Renglones.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("	JOIN Almacenes ON Encabezados.Cod_Alm = Almacenes.Cod_Alm	")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Encabezados.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Encabezados'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Adicional = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Salida'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Mediciones ON Mediciones.Cod_Reg = Encabezados.Documento")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Origen = 'Encabezados'")
            loComandoSeleccionar.AppendLine("       AND Mediciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	    AND Mediciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("		AND Mediciones.Adicional LIKE ('%'+RTRIM(Operaciones_Lotes.Cod_Lot)+'%')")
            loComandoSeleccionar.AppendLine("		AND Renglones.Renglon = SUBSTRING(Mediciones.Adicional, LEN(Mediciones.Adicional), 1)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Piezas ON Mediciones.Documento = Piezas.Documento")
            loComandoSeleccionar.AppendLine("		AND Piezas.Cod_Var = 'CPRO-NPIEZ'")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Mediciones AS Longitud ON Mediciones.Documento = Longitud.Documento")
            loComandoSeleccionar.AppendLine("		AND Longitud.Cod_Var = 'CPRO-LARG'")
            loComandoSeleccionar.AppendLine("WHERE Encabezados.Status IN ('Confirmado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("	AND Encabezados.Origen = 'Consumos Produccion'")
            loComandoSeleccionar.AppendLine("		AND Encabezados.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("		AND Encabezados.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("CREATE CLUSTERED INDEX PK_Fecha ON #curTemporal(Fec_Ini)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#curTemporal.Saldo					AS Saldo,		")
            loComandoSeleccionar.AppendLine("		#curTemporal.Operacion				AS Operacion,	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Orden					AS Orden,")
            loComandoSeleccionar.AppendLine("		#curTemporal.Documento				AS Documento,")
            loComandoSeleccionar.AppendLine("		#curTemporal.Cod_Art				AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		#curTemporal.Lote					AS Lote, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Piezas_Ent				AS Piezas_Ent,")
            loComandoSeleccionar.AppendLine("		#curTemporal.Piezas_Sal				AS Piezas_Sal, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Longitud				AS Longitud, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Saldo_Piezas			AS Saldo_Piezas,")
            loComandoSeleccionar.AppendLine("		#curTemporal.Fec_Ini				AS Fec_Ini, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Cod_Alm				AS Cod_Alm, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Nom_Alm				AS Nom_Alm, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.CanLte_Sal				AS Can_Sal, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.CanLte_Ent				AS Can_Ent, 	")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art					AS Nom_Art,		")
            loComandoSeleccionar.AppendLine("       " & lcParametro6Desde & "           AS Mostrar,")
            loComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Dep_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Dep_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Desde AND Cod_Dep = @lcCodDep_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Sec_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Hasta AND Cod_Dep = @lcCodDep_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Sec_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm  FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Hasta,")
            loComandoSeleccionar.AppendLine("		CONCAT(@lcCodLot_Desde, ' - ', (CASE WHEN @lcCodLot_Hasta = 'zzzzzzz' THEN '' ELSE @lcCodLot_Hasta END))	AS Lotes")
            loComandoSeleccionar.AppendLine("FROM #curTemporal")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ")
            loComandoSeleccionar.AppendLine("       ON Articulos.Cod_Art = #curTemporal.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE #curTemporal.Fec_Ini >= @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Alm ASC, Cod_Art ASC, Fec_Ini ASC, Orden ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #curTemporal")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Calcula el saldo de cada movimiento por artículo
            '-------------------------------------------------------------------------------------------------------
            Dim lcArticulo As String = ""
            Dim lcDocumento As String = ""
            Dim lcOperacion As String = ""
            Dim lcAlmacen As String = ""
            Dim lcLote As String = ""
            Dim lnSaldo As Decimal = 0D
            Dim lnSaldo_Piezas As Decimal = 0D
            For Each loRenglon As DataRow In laDatosReporte.Tables(0).Rows

                If Trim(loRenglon("Cod_Art")).ToLower() <> lcArticulo Or Trim(loRenglon("Cod_Alm")).ToLower() <> lcAlmacen Then
                    If Trim(loRenglon("Cod_Art")).ToLower() <> lcArticulo Then
                        lcArticulo = Trim(loRenglon("Cod_Art")).ToLower()
                    End If
                    If Trim(loRenglon("Cod_Alm")).ToLower() <> lcAlmacen Then
                        lcAlmacen = Trim(loRenglon("Cod_Alm")).ToLower()
                    End If
                    lnSaldo = 0D
                    lnSaldo_Piezas = 0D
                    lcOperacion = ""
                    lcDocumento = ""
                    lcLote = ""
                End If

                If (loRenglon("Documento") <> lcDocumento And loRenglon("Operacion") <> lcOperacion) Or (loRenglon("Documento") <> lcDocumento And loRenglon("Operacion") = lcOperacion) Or (loRenglon("Documento") = lcDocumento And loRenglon("Operacion") = lcOperacion And loRenglon("Lote") <> lcLote) Or (loRenglon("Documento") = lcDocumento And loRenglon("Operacion") = lcOperacion And loRenglon("Lote") = lcLote) Then
                    lcDocumento = loRenglon("Documento")
                    lcOperacion = loRenglon("Operacion")
                    lcLote = loRenglon("Lote")
                    lnSaldo = lnSaldo + CDec(loRenglon("Can_Ent")) - CDec(loRenglon("Can_Sal"))
                Else
                    loRenglon("Can_Ent") = 0D
                    loRenglon("Can_Sal") = 0D
                End If

                lnSaldo_Piezas = lnSaldo_Piezas + CDec(loRenglon("Piezas_Ent")) - CDec(loRenglon("Piezas_Sal"))
                loRenglon("Saldo") = lnSaldo
                loRenglon("Saldo_Piezas") = lnSaldo_Piezas

            Next loRenglon


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rHMovimientos_Inventarios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rHMovimientos_Inventarios.ReportSource = loObjetoReporte

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

