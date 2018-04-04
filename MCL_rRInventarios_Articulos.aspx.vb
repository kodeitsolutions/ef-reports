'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "MCL_rRInventarios_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class MCL_rRInventarios_Articulos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
         Dim lcAlmacenTransito As String
            lcAlmacenTransito = goServicios.mObtenerCampoFormatoSQL(goOpciones.mObtener("CODALMTRA", "C"))



        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(10) =  " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(10) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(10) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Desde AS VARCHAR(10) = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Hasta AS VARCHAR(10) = " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("")
            '-------------------------------------------------------------------------------------------'
            ' Select para ubicar el Stock Inicial a partir de la tabla de Ajustes de Inventarios 
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Ajustes.Cod_Art																		AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art																				AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada' THEN Renglones_Ajustes.Can_Art1 ELSE 0 END)	AS	Can_IniE, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Renglones_Ajustes.Can_Art1 ELSE 0 END)		AS	Can_IniS ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalAjustes ")
            loComandoSeleccionar.AppendLine(" FROM		Ajustes, ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Ajustes.Documento					=	Renglones_Ajustes.Documento ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				=	Renglones_Ajustes.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Ajustes.Status					=	'Confirmado' ")
            loComandoSeleccionar.AppendLine("			AND Renglones_Ajustes.Tipo			IN	('Entrada', 'Salida') ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("			AND Ajustes.Fec_Ini					< @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("			AND Renglones_Ajustes.Cod_Alm		BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            loComandoSeleccionar.AppendLine(" GROUP BY	Renglones_Ajustes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art ")

            '-------------------------------------------------------------------------------------------'
            ' Select para ubicar el Stock Inicial a partir de la tabla de Notas de Recepciones
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Recepciones.Cod_Art		AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art					AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Recepciones.Can_Art1)	AS	Can_IniE, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_IniS ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalRecepciones ")
            loComandoSeleccionar.AppendLine(" FROM		Recepciones, ")
            loComandoSeleccionar.AppendLine("			Renglones_Recepciones, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Recepciones.Documento				=	Renglones_Recepciones.Documento ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				=	Renglones_Recepciones.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Recepciones.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("			AND Recepciones.Fec_Ini				< @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("			AND Renglones_Recepciones.Cod_Alm	BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            loComandoSeleccionar.AppendLine(" GROUP BY	Renglones_Recepciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art ")

            '-------------------------------------------------------------------------------------------'
            ' Select para ubicar el Stock Inicial a partir de la tabla de Traslados en las Entradas
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Traslados.Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art						AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Traslados.Can_Art1)		AS	Can_IniE, ")
            loComandoSeleccionar.AppendLine("			0.00									AS	Can_IniS ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalTrasladosEntradas ")
            loComandoSeleccionar.AppendLine(" FROM		Traslados, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Traslados.Documento				=	Renglones_Traslados.Documento ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art			=	Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Traslados.Status			IN ('Confirmado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("			AND Traslados.Fec_Ini < @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("			AND (CASE Traslados.Status ")
            loComandoSeleccionar.AppendLine("					WHEN 'Confirmado' THEN " & lcAlmacenTransito)
            loComandoSeleccionar.AppendLine("					ELSE Traslados.Alm_Des ")
            loComandoSeleccionar.AppendLine("				END)	BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            loComandoSeleccionar.AppendLine("GROUP BY	Renglones_Traslados.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art")

            '-------------------------------------------------------------------------------------------'
            ' Select para ubicar el Stock Inicial a partir de la tabla de Traslados en las Salidas
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Traslados.Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art						AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			0.00									AS	Can_IniE, ")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Traslados.Can_Art1)		AS	Can_IniS ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalTrasladosSalidas ")
            loComandoSeleccionar.AppendLine(" FROM		Traslados, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Traslados.Documento					=	Renglones_Traslados.Documento ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				=	Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Traslados.Status		IN ('Confirmado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("			AND Traslados.Fec_Ini		< @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("			AND Traslados.Alm_Ori		BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            loComandoSeleccionar.AppendLine(" GROUP BY	Renglones_Traslados.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art")

            '-------------------------------------------------------------------------------------------'
            ' Union de los diferentes select para obtener el stock inicial
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine("SELECT	* ")
            loComandoSeleccionar.AppendLine("INTO		#TablaTemporalCantidades1 ")
            loComandoSeleccionar.AppendLine("FROM		#TablaTemporalAjustes ")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT	* ")
            loComandoSeleccionar.AppendLine("FROM		#TablaTemporalRecepciones ")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT	* ")
            loComandoSeleccionar.AppendLine("FROM		#TablaTemporalTrasladosEntradas ")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT	* ")
            loComandoSeleccionar.AppendLine("FROM		#TablaTemporalTrasladosSalidas ")


            loComandoSeleccionar.AppendLine("SELECT		Cod_Art							AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			SUM(Can_IniE)					AS	Can_IniE, ")
            loComandoSeleccionar.AppendLine("			SUM(Can_IniS)					AS	Can_IniS ")
            loComandoSeleccionar.AppendLine("INTO		#TablaTemporalCantidades2 ")
            loComandoSeleccionar.AppendLine("FROM		#TablaTemporalCantidades1 ")
            loComandoSeleccionar.AppendLine("GROUP BY	Cod_Art, Nom_Art ")

            loComandoSeleccionar.AppendLine("SELECT		Cod_Art							AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			(Can_IniE - Can_IniS)			AS	Can_Ini	")
            loComandoSeleccionar.AppendLine("INTO		#TablaTemporalCantidadesIniciales ")
            loComandoSeleccionar.AppendLine("FROM		#TablaTemporalCantidades2 ")

            '-------------------------------------------------------------------------------------------'
            ' Movimientos de Entradas de Productos 
            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Ajustes (Entradas)
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Ajustes.Cod_Art			AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art					AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Can_Art1			AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalEntradasAjustes1 ")
            loComandoSeleccionar.AppendLine(" FROM		Ajustes, ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Ajustes.Documento					=	Renglones_Ajustes.Documento ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				=	Renglones_Ajustes.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Ajustes.Status					=	'Confirmado' ")
            loComandoSeleccionar.AppendLine("			AND Renglones_Ajustes.Tipo			=	'Entrada' ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("			AND Ajustes.Fec_Ini					BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("			AND Renglones_Ajustes.Cod_Alm		BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")

            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Recepciones (Entradas)
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Recepciones.Cod_Art		AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art					AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Recepciones.Can_Art1		AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalEntradasRecepciones1 ")
            loComandoSeleccionar.AppendLine(" FROM		Recepciones, ")
            loComandoSeleccionar.AppendLine("			Renglones_Recepciones, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Recepciones.Documento				=	Renglones_Recepciones.Documento ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				=	Renglones_Recepciones.Cod_Art ")
            loComandoSeleccionar.AppendLine("			And Recepciones.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("			And Recepciones.Fec_Ini				BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("			And Renglones_Recepciones.Cod_Alm	BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")

            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Traslados (Entradas)
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Traslados.Cod_Art			AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art					AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados.Can_Art1		AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalEntradasTraslados1 ")
            loComandoSeleccionar.AppendLine(" FROM		Traslados, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Traslados.Documento				=	Renglones_Traslados.Documento ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art			=	Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Traslados.Status			IN ('Confirmado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Art				BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("			AND Traslados.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("			AND (CASE Traslados.Status ")
            loComandoSeleccionar.AppendLine("					WHEN 'Confirmado' THEN " & lcAlmacenTransito)
            loComandoSeleccionar.AppendLine("					ELSE Traslados.Alm_Des ")
            loComandoSeleccionar.AppendLine("				END)	BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")

            '-------------------------------------------------------------------------------------------'
            ' Movimientos de Salidas de Productos 
            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Ajustes (Salidas)
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Ajustes.Cod_Art			AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art					AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Can_Art1			AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalSalidasAjustes1 ")
            loComandoSeleccionar.AppendLine(" FROM		Ajustes, ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Ajustes.Documento					=	Renglones_Ajustes.Documento ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				=	Renglones_Ajustes.Cod_Art ")
            loComandoSeleccionar.AppendLine("			And Ajustes.Status					=	'Confirmado' ")
            loComandoSeleccionar.AppendLine("			And Renglones_Ajustes.Tipo			=	'Salida' ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("			And Ajustes.Fec_Ini					BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("			And Renglones_Ajustes.Cod_Alm		BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")

            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Traslados (Salidas)
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Traslados.Cod_Art			AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art					AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			0.00								AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados.Can_Art1		AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalSalidasTraslados1 ")
            loComandoSeleccionar.AppendLine(" FROM		Traslados, ")
            loComandoSeleccionar.AppendLine("			Renglones_Traslados, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Traslados.Documento					=	Renglones_Traslados.Documento ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				=	Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine("			AND Traslados.Status				IN ('Confirmado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("			And Traslados.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("			And Traslados.Alm_Ori BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	        AND Articulos.Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")

            '-------------------------------------------------------------------------------------------'
            ' Union de los selects desde las tablas temporales 
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Nom_Art				AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Can_Com				AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			Can_Ent				AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			Can_Ven				AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			Can_Sal				AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos1 ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalEntradasAjustes1 ")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Nom_Art				AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Can_Com				AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			Can_Ent				AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			Can_Ven				AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			Can_Sal				AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalEntradasRecepciones1 ")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Nom_Art				AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Can_Com				AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			Can_Ent				AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			Can_Ven				AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			Can_Sal				AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalEntradasTraslados1 ")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Nom_Art				AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Can_Com				AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			Can_Ent				AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			Can_Ven				AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			Can_Sal				AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalSalidasAjustes1 ")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Nom_Art				AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Can_Com				AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			Can_Ent				AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			Can_Ven				AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			Can_Sal				AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalSalidasTraslados1 ")

            '-------------------------------------------------------------------------------------------'
            ' Suma de las cantidades luego de la unificacion de los selects
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Nom_Art				AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			SUM(Can_Com)		AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			SUM(Can_Ent)		AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			SUM(Can_Ven)		AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			SUM(Can_Sal)		AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos2 ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalMovimientos1 ")
            loComandoSeleccionar.AppendLine(" GROUP BY	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Nom_Art ")

            '-------------------------------------------------------------------------------------------'
            ' Union del select de movimientos stock inicial y movimients actuales
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	#TablaTemporalMovimientos2.Cod_Art					AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("			#TablaTemporalMovimientos2.Nom_Art					AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("			#TablaTemporalCantidadesIniciales.Can_Ini			AS	Can_Ini,  ")
            loComandoSeleccionar.AppendLine("			#TablaTemporalMovimientos2.Can_Com					AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("			#TablaTemporalMovimientos2.Can_Ent					AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine("			#TablaTemporalMovimientos2.Can_Ven					AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine("			#TablaTemporalMovimientos2.Can_Sal					AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos3 ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalMovimientos2 ")
            loComandoSeleccionar.AppendLine("LEFT JOIN #TablaTemporalCantidadesIniciales ON #TablaTemporalCantidadesIniciales.Cod_Art = #TablaTemporalMovimientos2.Cod_Art")

            '-------------------------------------------------------------------------------------------'
            ' Conversion de los valores NULL en valores numericos
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art												AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			Nom_Art												AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 			(CASE WHEN Can_Ini IS NULL THEN 0 ELSE Can_Ini END)	AS	Can_Ini, ")
            loComandoSeleccionar.AppendLine(" 			(CASE WHEN Can_Com IS NULL THEN 0 ELSE Can_Com END) AS	Can_Com, ")
            loComandoSeleccionar.AppendLine(" 			(CASE WHEN Can_Ent IS NULL THEN 0 ELSE Can_Ent END)	AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine(" 			(CASE WHEN Can_Ven IS NULL THEN 0 ELSE Can_Ven END)	AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine(" 			(CASE WHEN Can_Sal IS NULL THEN 0 ELSE Can_Sal END) AS	Can_Sal ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos4 ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalMovimientos3 ")

            '-------------------------------------------------------------------------------------------'
            ' Calculo del Stock Final
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art												AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			Nom_Art								                AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 			Can_Ini												AS	Can_Ini, ")
            loComandoSeleccionar.AppendLine(" 			Can_Com												AS	Can_Com, ")
            loComandoSeleccionar.AppendLine(" 			Can_Ent												AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine(" 			Can_Ven												AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine(" 			Can_Sal												AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine(" 			(Can_Ini + Can_Com + Can_Ent - Can_Ven - Can_Sal)	AS	Can_Fin ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos5 ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalMovimientos4 ")

            '-------------------------------------------------------------------------------------------'
            ' Calculo del Stock Final
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	#TablaTemporalMovimientos5.Cod_Art					AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			#TablaTemporalMovimientos5.Nom_Art					AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Cod_Uni1									AS	Cod_Uni1, ")
            loComandoSeleccionar.AppendLine(" 			#TablaTemporalMovimientos5.Can_Ini					AS	Can_Ini, ")
            loComandoSeleccionar.AppendLine(" 			#TablaTemporalMovimientos5.Can_Com					AS	Can_Com, ")
            loComandoSeleccionar.AppendLine(" 			#TablaTemporalMovimientos5.Can_Ent					AS	Can_Ent, ")
            loComandoSeleccionar.AppendLine(" 			#TablaTemporalMovimientos5.Can_Ven					AS	Can_Ven, ")
            loComandoSeleccionar.AppendLine(" 			#TablaTemporalMovimientos5.Can_Sal					AS	Can_Sal, ")
            loComandoSeleccionar.AppendLine(" 			#TablaTemporalMovimientos5.Can_Fin					AS	Can_Fin, ")
            loComandoSeleccionar.AppendLine("		    CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            loComandoSeleccionar.AppendLine("		    CASE WHEN @lcCodArt_Desde <> ''")
            loComandoSeleccionar.AppendLine("		    	 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            loComandoSeleccionar.AppendLine("		    	 ELSE '' END				                    AS Art_Desde,")
            loComandoSeleccionar.AppendLine("		    CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("		    	 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            loComandoSeleccionar.AppendLine("		    	 ELSE '' END				                    AS Art_Hasta,")
            loComandoSeleccionar.AppendLine("		    CASE WHEN @lcCodDep_Desde <> ''")
            loComandoSeleccionar.AppendLine("		    	 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Desde)")
            loComandoSeleccionar.AppendLine("		    	 ELSE '' END				                    AS Dep_Desde,")
            loComandoSeleccionar.AppendLine("		    CASE WHEN @lcCodDep_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("		    	 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Hasta)")
            loComandoSeleccionar.AppendLine("		    	 ELSE '' END				                    AS Dep_Hasta,")
            loComandoSeleccionar.AppendLine("		    CASE WHEN @lcCodSec_Desde <> ''")
            loComandoSeleccionar.AppendLine("		    	 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Desde AND Cod_Dep = @lcCodDep_Desde)")
            loComandoSeleccionar.AppendLine("		    	 ELSE '' END				                    AS Sec_Desde,")
            loComandoSeleccionar.AppendLine("		    CASE WHEN @lcCodSec_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("		    	 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Hasta AND Cod_Dep = @lcCodDep_Hasta)")
            loComandoSeleccionar.AppendLine("		    	 ELSE '' END				                    AS Sec_Hasta,")
            loComandoSeleccionar.AppendLine("		    CASE WHEN @lcCodAlm_Desde <> ''")
            loComandoSeleccionar.AppendLine("		    	 THEN (SELECT Nom_Alm FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Desde)")
            loComandoSeleccionar.AppendLine("		    	 ELSE '' END				                    AS Alm_Desde,")
            loComandoSeleccionar.AppendLine("		    CASE WHEN @lcCodAlm_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			     THEN (SELECT Nom_Alm  FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Hasta)")
            loComandoSeleccionar.AppendLine("			     ELSE '' END				                    AS Alm_Hasta")
            loComandoSeleccionar.AppendLine(" FROM		Articulos, #TablaTemporalMovimientos5 ")
            loComandoSeleccionar.AppendLine(" WHERE		Articulos.Cod_Art	=	#TablaTemporalMovimientos5.Cod_Art ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalAjustes  ")
            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalRecepciones")
            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalTrasladosEntradas	")
            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalTrasladosSalidas	")
			loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalCantidades1")
			loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalCantidades2  ")
			loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalCantidadesIniciales  ")
            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalEntradasAjustes1	  ")
            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalEntradasRecepciones1   ")
            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalEntradasTraslados1 ")
            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalSalidasAjustes1  ")
            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalSalidasTraslados1	 ")
            loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalMovimientos1	")
			loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalMovimientos2   ")
			loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalMovimientos3   ")
			loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalMovimientos4	")
			loComandoSeleccionar.AppendLine("DROP TABLE #TablaTemporalMovimientos5  ")   
			

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("MCL_rRInventarios_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvMCL_rRInventarios_Articulos.ReportSource = loObjetoReporte

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
' JJD: 13/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 14/10/08: Continuacion de la programacion
'-------------------------------------------------------------------------------------------'
' CMS:  12/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro "Sucursal"
'-------------------------------------------------------------------------------------------
' CMS:  11/08/09: Verificacion de registros y se agregaro el filtro_ Ubicación
'-------------------------------------------------------------------------------------------'
' RJG: 08/12/10: Ajustado Estatus de Facturas de Venta: Ahora omite las facturas de venta	'
'				 pendientes e incluye las confirmadas.										'
'-------------------------------------------------------------------------------------------'
' MAT: 21/07/11: Adición de nueva columna diferencia. Ajuste de la vista de diseño			'
'-------------------------------------------------------------------------------------------'
' MAT: 26/07/11: Ajuste del Select															'
'-------------------------------------------------------------------------------------------'
' MAT: 15/08/11: Ajuste del Select(Adicion de las tablas Traslados y devoluciones de compra)															'
'-------------------------------------------------------------------------------------------'
' MAT: 29/09/11: Ajuste del Select(Eliminación del filtro Status, no aplica)				'
'-------------------------------------------------------------------------------------------'
' RJG: 15/05/12: Se ajutaron los filtros de sucursal y de almacén.							'
'-------------------------------------------------------------------------------------------'
