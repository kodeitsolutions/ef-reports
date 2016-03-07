'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRInventarios_Marca"
'-------------------------------------------------------------------------------------------'
Partial Class rRInventarios_Marca
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
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)


        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            '-------------------------------------------------------------------------------------------'
            ' Select para ubicar el Stock Inicial a partir de la tabla de Ajustes de Inventarios 
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Ajustes.Cod_Art																			AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("             Articulos.Nom_Art																				AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("             Articulos.Cod_Mar																				As Cod_Mar, ")
            loComandoSeleccionar.AppendLine("             SUM(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada' THEN Renglones_Ajustes.Can_Art1 ELSE 0 END)	AS	Can_IniE,  ")
            loComandoSeleccionar.AppendLine("             SUM(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Renglones_Ajustes.Can_Art1 ELSE 0 END)		AS	Can_IniS  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalAjustes  ")
            loComandoSeleccionar.AppendLine(" FROM		Ajustes,  ")
            loComandoSeleccionar.AppendLine("             Renglones_Ajustes,  ")
            loComandoSeleccionar.AppendLine("             Articulos  ")
            loComandoSeleccionar.AppendLine(" WHERE		Ajustes.Documento					=	Renglones_Ajustes.Documento  ")
            loComandoSeleccionar.AppendLine("             And Articulos.Cod_Art				=	Renglones_Ajustes.Cod_Art  ")
            loComandoSeleccionar.AppendLine("             And Ajustes.Status				=	'Confirmado'  ")
            loComandoSeleccionar.AppendLine("             And Renglones_Ajustes.Tipo		IN	('Entrada', 'Salida')  ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Ajustes.Fec_Ini					< " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And Renglones_Ajustes.Cod_Alm		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY	Renglones_Ajustes.Cod_Art,  ")
            loComandoSeleccionar.AppendLine(" Articulos.Nom_Art, cod_Mar ")

            '-------------------------------------------------------------------------------------------'
            ' Select para ubicar el Stock Inicial a partir de la tabla de Compras 
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Compras.Cod_Art				AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Mar					As Cod_Mar, ")
            loComandoSeleccionar.AppendLine("         SUM(Renglones_Compras.Can_Art1)		AS	Can_IniE,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_IniS  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalCompras  ")
            loComandoSeleccionar.AppendLine(" FROM		Compras,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Compras,  ")
            loComandoSeleccionar.AppendLine(" Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE(Compras.Documento = Renglones_Compras.Documento) ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art					=	Renglones_Compras.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Compras.Status					IN	('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine("         And Renglones_Compras.Tip_Ori			<>	'Recepciones'  ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Compras.Fec_Ini					< " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And Renglones_Compras.Cod_Alm		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY	Renglones_Compras.Cod_Art,  ")
            loComandoSeleccionar.AppendLine(" Articulos.Nom_Art, cod_Mar ")

            '-------------------------------------------------------------------------------------------'
            ' Select para ubicar el Stock Inicial a partir de la tabla de Notas de Recepciones
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Recepciones.Cod_Art		AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Mar					As Cod_Mar, ")
            loComandoSeleccionar.AppendLine("         SUM(Renglones_Recepciones.Can_Art1)	AS	Can_IniE,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_IniS  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalRecepciones  ")
            loComandoSeleccionar.AppendLine(" FROM		Recepciones,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Recepciones,  ")
            loComandoSeleccionar.AppendLine("         Articulos  ")
            loComandoSeleccionar.AppendLine(" WHERE		Recepciones.Documento				=	Renglones_Recepciones.Documento  ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art					=	Renglones_Recepciones.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Recepciones.Status				IN	('Confirmado', 'Afectado', 'Procesado')  ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Recepciones.Fec_Ini				< " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And Renglones_Recepciones.Cod_Alm	BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY	Renglones_Recepciones.Cod_Art,  ")
            loComandoSeleccionar.AppendLine(" Articulos.Nom_Art, Articulos.cod_Mar ")

            '-------------------------------------------------------------------------------------------'
            ' Select para ubicar el Stock Inicial a partir de la tabla de Facturas de Ventas
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Facturas.Cod_Art			AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Mar					As Cod_Mar, ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_IniE,  ")
            loComandoSeleccionar.AppendLine("         SUM(Renglones_Facturas.Can_Art1)	AS	Can_IniS  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalFacturas  ")
            loComandoSeleccionar.AppendLine(" FROM		Facturas,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Facturas,  ")
            loComandoSeleccionar.AppendLine(" Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE Facturas.Documento = Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art					=	Renglones_Facturas.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Facturas.Status					IN	('Confirmado', 'Afectado', 'Procesado')  ")
            loComandoSeleccionar.AppendLine("         And Renglones_Facturas.Tip_Ori		<>	'Entregas'  ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Facturas.Fec_Ini				< " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And Renglones_Facturas.Cod_Alm		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY	Renglones_Facturas.Cod_Art,  ")
            loComandoSeleccionar.AppendLine(" Articulos.Nom_Art, Articulos.cod_Mar ")

            '-------------------------------------------------------------------------------------------'
            ' Select para ubicar el Stock Inicial a partir de la tabla de Notas de Entregas
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Entregas.Cod_Art			AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Mar					As Cod_Mar, ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_IniE,  ")
            loComandoSeleccionar.AppendLine("         SUM(Renglones_Entregas.Can_Art1)	AS	Can_IniS  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalEntregas  ")
            loComandoSeleccionar.AppendLine(" FROM		Entregas,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Entregas,  ")
            loComandoSeleccionar.AppendLine("         Articulos  ")
            loComandoSeleccionar.AppendLine(" WHERE		Entregas.Documento					=	Renglones_Entregas.Documento  ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art					=	Renglones_Entregas.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Entregas.Status					IN	('Confirmado', 'Afectado', 'Procesado')  ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Entregas.Fec_Ini				< " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And Renglones_Entregas.Cod_Alm		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY	Renglones_Entregas.Cod_Art,  ")
            loComandoSeleccionar.AppendLine(" Articulos.Nom_Art, Articulos.cod_Mar ")

            '-------------------------------------------------------------------------------------------'
            ' Union de los diferentes select para obtener el stock inicial
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	*  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalCantidades1  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalAjustes  ")
            loComandoSeleccionar.AppendLine(" UNION ALL  ")
            loComandoSeleccionar.AppendLine(" SELECT	*  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalEntregas  ")
            loComandoSeleccionar.AppendLine(" UNION ALL  ")
            loComandoSeleccionar.AppendLine(" SELECT	*  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalRecepciones  ")
            loComandoSeleccionar.AppendLine(" UNION ALL  ")
            loComandoSeleccionar.AppendLine(" SELECT	*  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalFacturas  ")
            loComandoSeleccionar.AppendLine(" UNION ALL  ")
            loComandoSeleccionar.AppendLine(" SELECT	*  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalEntregas  ")

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art							AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         SUM(Can_IniE)					AS	Can_IniE,  ")
            loComandoSeleccionar.AppendLine("         SUM(Can_IniS)					AS	Can_IniS  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalCantidades2  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalCantidades1  ")
            loComandoSeleccionar.AppendLine(" GROUP BY	Cod_Art, Nom_Art, cod_Mar  ")

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art							AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         (Can_IniE - Can_IniS)			AS	Can_Ini	 ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalCantidadesIniciales  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalCantidades2  ")

            '-------------------------------------------------------------------------------------------'
            ' Movimientos de Entradas de Productos 
            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Ajustes (Entradas)
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Ajustes.Cod_Art			AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Ajustes.Can_Art1			AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalEntradasAjustes1  ")
            loComandoSeleccionar.AppendLine(" FROM		Ajustes,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Ajustes,  ")
            loComandoSeleccionar.AppendLine("         Articulos  ")
            loComandoSeleccionar.AppendLine(" WHERE		Ajustes.Documento					=	Renglones_Ajustes.Documento  ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art					=	Renglones_Ajustes.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Ajustes.Status					=	'Confirmado' ")
            loComandoSeleccionar.AppendLine("         And Renglones_Ajustes.Tipo			=	'Entrada'  ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Ajustes.Fec_Ini					BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Renglones_Ajustes.Cod_Alm		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)

            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Recepciones (Entradas)
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Recepciones.Cod_Art		AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Recepciones.Can_Art1		AS	Can_Com, ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalEntradasRecepciones1  ")
            loComandoSeleccionar.AppendLine(" FROM		Recepciones,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Recepciones,  ")
            loComandoSeleccionar.AppendLine("         Articulos  ")
            loComandoSeleccionar.AppendLine(" WHERE		Recepciones.Documento				=	Renglones_Recepciones.Documento  ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art					=	Renglones_Recepciones.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Recepciones.Status				IN	('Confirmado', 'Afectado', 'Procesado')  ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Recepciones.Fec_Ini				BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Renglones_Recepciones.Cod_Alm	BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)

            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Facturas de Compras (Entradas)
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Compras.Cod_Art			AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Compras.Can_Art1			AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalEntradasCompras1  ")
            loComandoSeleccionar.AppendLine(" FROM		Compras,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Compras,  ")
            loComandoSeleccionar.AppendLine("         Articulos  ")
            loComandoSeleccionar.AppendLine(" WHERE		Compras.Documento				=	Renglones_Compras.Documento  ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art				=	Renglones_Compras.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Compras.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("         And Renglones_Compras.Tip_Ori		<>	'Recepciones'  ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art				BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")
            loComandoSeleccionar.AppendLine("         And Compras.Fec_Ini					BETWEEN '20090101 00:00:00.000' ")
            loComandoSeleccionar.AppendLine("         And '20090711 23:59:59.998' ")
            loComandoSeleccionar.AppendLine("         And Renglones_Compras.Cod_Alm		BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Dep				BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Sec				BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Mar				BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Cla				BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Tip				BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Pro				BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Uni1				BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Suc				BETWEEN '' ")
            loComandoSeleccionar.AppendLine("         And 'zzzzzzz' ")

            '-------------------------------------------------------------------------------------------'
            ' Movimientos de Salidas de Productos 
            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Ajustes (Salidas)
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Ajustes.Cod_Art			AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Ajustes.Can_Art1			AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalSalidasAjustes1  ")
            loComandoSeleccionar.AppendLine(" FROM		Ajustes,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Ajustes,  ")
            loComandoSeleccionar.AppendLine("         Articulos  ")
            loComandoSeleccionar.AppendLine(" WHERE		Ajustes.Documento					=	Renglones_Ajustes.Documento  ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art					=	Renglones_Ajustes.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Ajustes.Status					=	'Confirmado'")
            loComandoSeleccionar.AppendLine("         And Renglones_Ajustes.Tipo			=	'Salida'  ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Ajustes.Fec_Ini					BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Renglones_Ajustes.Cod_Alm		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)

            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Entregas (Salidas)
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Entregas.Cod_Art			AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Entregas.Can_Art1			AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalSalidasEntregas1  ")
            loComandoSeleccionar.AppendLine(" FROM		Entregas,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Entregas,  ")
            loComandoSeleccionar.AppendLine("         Articulos  ")
            loComandoSeleccionar.AppendLine(" WHERE		Entregas.Documento					=	Renglones_Entregas.Documento  ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art				=	Renglones_Entregas.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Entregas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Entregas.Fec_Ini				BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Renglones_Entregas.Cod_Alm		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)

            '-------------------------------------------------------------------------------------------'
            ' Select de la tabla de Facturas de Ventas (Salidas)
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Facturas.Cod_Art			AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Facturas.Can_Art1			AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         0.00								AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalSalidasVentas1  ")
            loComandoSeleccionar.AppendLine(" FROM		Facturas,  ")
            loComandoSeleccionar.AppendLine("         Renglones_Facturas,  ")
            loComandoSeleccionar.AppendLine("         Articulos  ")
            loComandoSeleccionar.AppendLine(" WHERE		Facturas.Documento					=	Renglones_Facturas.Documento  ")
            loComandoSeleccionar.AppendLine("         And Articulos.Cod_Art					=	Renglones_Facturas.Cod_Art  ")
            loComandoSeleccionar.AppendLine("         And Facturas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("         And Renglones_Facturas.Tip_Ori		<>	'Entregas'  ")
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Facturas.Fec_Ini				BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Renglones_Facturas.Cod_Alm		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Dep				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Sec				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Mar				BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Tip				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Pro				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Uni1				BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Suc				BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro10Hasta)

            '-------------------------------------------------------------------------------------------'
            ' Union de los selects desde las tablas temporales 
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art				AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Can_Com				AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         Can_Ent				AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         Can_Ven				AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         Can_Sal				AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos1  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalEntradasAjustes1 ")
            loComandoSeleccionar.AppendLine(" UNION ALL  ")
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art				AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Can_Com				AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         Can_Ent				AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         Can_Ven				AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         Can_Sal				AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalEntradasRecepciones1  ")
            loComandoSeleccionar.AppendLine(" UNION ALL  ")
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art				AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Can_Com				AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         Can_Ent				AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         Can_Ven				AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         Can_Sal				AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalEntradasCompras1  ")
            loComandoSeleccionar.AppendLine(" UNION ALL  ")
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art				AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Can_Com				AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         Can_Ent				AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         Can_Ven				AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         Can_Sal				AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalSalidasAjustes1  ")
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art				AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Can_Com				AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         Can_Ent				AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         Can_Ven				AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         Can_Sal				AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalSalidasEntregas1  ")
            loComandoSeleccionar.AppendLine(" UNION ALL  ")
            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art				AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Can_Com				AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         Can_Ent				AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         Can_Ven				AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         Can_Sal				AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalSalidasVentas1  ")

            '-------------------------------------------------------------------------------------------'
            ' Suma de las cantidades luego de la unificacion de los selects
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art				AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art				AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         SUM(Can_Com)		AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         SUM(Can_Ent)		AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         SUM(Can_Ven)		AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         SUM(Can_Sal)		AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos2  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalMovimientos1  ")
            loComandoSeleccionar.AppendLine(" GROUP BY	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art, #TablaTemporalMovimientos1.Cod_Mar  ")

            '-------------------------------------------------------------------------------------------'
            ' Union del select de movimientos stock inicial y movimients actuales
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	#TablaTemporalMovimientos2.Cod_Art					AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos2.Cod_Mar					As Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos2.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         0.00												AS	Can_Ini,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos2.Can_Com					AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos2.Can_Ent					AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos2.Can_Ven					AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos2.Can_Sal					AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos3  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalMovimientos2  ")

            '-------------------------------------------------------------------------------------------'
            ' Conversion de los valores NULL en valores numericos
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art												AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar												AS	Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art												AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         (CASE WHEN Can_Ini IS NULL THEN 0 ELSE Can_Ini END)	AS	Can_Ini,  ")
            loComandoSeleccionar.AppendLine("         (CASE WHEN Can_Com IS NULL THEN 0 ELSE Can_Com END) AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         (CASE WHEN Can_Ent IS NULL THEN 0 ELSE Can_Ent END)	AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         (CASE WHEN Can_Ven IS NULL THEN 0 ELSE Can_Ven END)	AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         (CASE WHEN Can_Sal IS NULL THEN 0 ELSE Can_Sal END) AS	Can_Sal  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos4  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalMovimientos3  ")

            '-------------------------------------------------------------------------------------------'
            ' Calculo del Stock Final
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Art												AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         Cod_Mar												AS	Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         Nom_Art								                AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Can_Ini												AS	Can_Ini,  ")
            loComandoSeleccionar.AppendLine("         Can_Com												AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         Can_Ent												AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         Can_Ven												AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         Can_Sal												AS	Can_Sal,  ")
            loComandoSeleccionar.AppendLine("         (Can_Ini + Can_Com + Can_Ent - Can_Ven - Can_Sal)	AS	Can_Fin  ")
            loComandoSeleccionar.AppendLine(" INTO		#TablaTemporalMovimientos5  ")
            loComandoSeleccionar.AppendLine(" FROM		#TablaTemporalMovimientos4  ")

            '-------------------------------------------------------------------------------------------'
            ' Calculo del Stock Final
            '-------------------------------------------------------------------------------------------'

            loComandoSeleccionar.AppendLine(" SELECT	#TablaTemporalMovimientos5.Cod_Art					AS	Cod_Art,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos5.Cod_Mar					AS	Cod_Mar,  ")
            loComandoSeleccionar.AppendLine("         (SELECT Nom_Mar FROM Marcas WHERE Cod_Mar = #TablaTemporalMovimientos5.Cod_Mar)	AS	Nom_Mar,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos5.Nom_Art					AS	Nom_Art,  ")
            loComandoSeleccionar.AppendLine("         Articulos.Cod_Uni1									AS	Cod_Uni1,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos5.Can_Ini					AS	Can_Ini,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos5.Can_Com					AS	Can_Com,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos5.Can_Ent					AS	Can_Ent,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos5.Can_Ven					AS	Can_Ven,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos5.Can_Sal					AS	Can_Sal,  ")
            loComandoSeleccionar.AppendLine("         #TablaTemporalMovimientos5.Can_Fin					AS	Can_Fin  ")
            loComandoSeleccionar.AppendLine(" FROM		Articulos, #TablaTemporalMovimientos5")
            loComandoSeleccionar.AppendLine(" WHERE		Articulos.Cod_Art	=	#TablaTemporalMovimientos5.Cod_Art  ")
            loComandoSeleccionar.AppendLine(" ORDER BY    Articulos.Cod_Mar, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRInventarios_Marca", laDatosReporte)

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

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRInventarios_Marca.ReportSource = loObjetoReporte

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
' CMS: 11/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RJG: 08/12/10: Ajustado Estatus de Facturas de Venta: Ahora omite las facturas de venta	'
'				 pendientes e incluye las confirmadas.										'
'-------------------------------------------------------------------------------------------'
