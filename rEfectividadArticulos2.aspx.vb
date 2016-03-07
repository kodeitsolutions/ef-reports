'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEfectividadArticulos2"
'-------------------------------------------------------------------------------------------'
Partial Class rEfectividadArticulos2
    Inherits vis2Formularios.frmReporte

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
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))



            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("SELECT      ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Art,")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Cotizaciones.Can_Art1) AS Can_Art1,")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Cotizaciones.Mon_Net) AS Mon_Net,")
            loComandoSeleccionar.AppendLine("           Articulos.Exi_Act1,")
            loComandoSeleccionar.AppendLine("           Articulos.Exi_Por1,")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Mon")
            loComandoSeleccionar.AppendLine("INTO	#tmpCotizaciones")
            loComandoSeleccionar.AppendLine("FROM	Cotizaciones ")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Cotizaciones ON Renglones_Cotizaciones.Documento = Cotizaciones.Documento")
            loComandoSeleccionar.AppendLine(" 			                    AND Renglones_Cotizaciones.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			                    AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Cotizaciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" 			                    AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			                    AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			                    AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			                    AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			                    AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			                    AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			                    AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			                    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			                    AND Articulos.cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			                    AND " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine("JOIN	Clientes ON Clientes.Cod_Cli = Cotizaciones.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE      Cotizaciones.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.cod_Suc BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Rev BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Renglones_Cotizaciones.Cod_Art,")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Act1,")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Por1,")
            loComandoSeleccionar.AppendLine("			Cotizaciones.Cod_Mon")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT      ")
            loComandoSeleccionar.AppendLine("        	Renglones_Facturas.Cod_Art,")
            loComandoSeleccionar.AppendLine("        	SUM(Renglones_Facturas.Can_Art1) AS Can_Art1,")
            loComandoSeleccionar.AppendLine("        	SUM(Renglones_Facturas.Mon_Net) AS Mon_Net,")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Mon")
            loComandoSeleccionar.AppendLine("INTO	#tmpFacturas")
            loComandoSeleccionar.AppendLine("FROM	Facturas")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            loComandoSeleccionar.AppendLine(" 			                AND Renglones_Facturas.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine("JOIN	Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE  	Facturas.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.cod_Suc BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Rev BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Renglones_Facturas.Cod_Art,")
            loComandoSeleccionar.AppendLine("			Facturas.Cod_Mon")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("SELECT		Renglones_Compras.Cod_Art,")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Compras.Can_Art1) As Can_Art1,")
            loComandoSeleccionar.AppendLine("			Renglones_Compras.Precio1,")
            loComandoSeleccionar.AppendLine("			Compras.Cod_Mon,")
            loComandoSeleccionar.AppendLine("			MAX(Compras.Fec_Ini) AS Fecha_Ultima_Compra")
            loComandoSeleccionar.AppendLine("INTO	#tmpCompras")
            loComandoSeleccionar.AppendLine("FROM	Compras")
            'loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
            loComandoSeleccionar.AppendLine("                           AND Renglones_Compras.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Compras.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("WHERE  Compras.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine("           AND Compras.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Compras.Tip_Ori	<>	'Recepciones' ")
            loComandoSeleccionar.AppendLine("GROUP BY Renglones_Compras.Cod_Art,Renglones_Compras.Precio1,Compras.Cod_Mon")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("UNION ALL")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Renglones_Recepciones.Cod_Art,")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Recepciones.Can_Art1) As Can_Art1,")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Precio1,")
            loComandoSeleccionar.AppendLine("           Recepciones.Cod_Mon,")
            loComandoSeleccionar.AppendLine("           MAX(Recepciones.Fec_Ini) AS Fecha_Ultima_Compra")
            loComandoSeleccionar.AppendLine("FROM	Recepciones")
            ' loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Recepciones.Cod_Pro")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("                           AND Renglones_Recepciones.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Recepciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("WHERE  Recepciones.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine("           AND Recepciones.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Renglones_Recepciones.Cod_Art,Renglones_Recepciones.Precio1,Recepciones.Cod_Mon")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("UNION ALL")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Renglones_Ajustes.Cod_Art,")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Ajustes.Can_Art1) As Can_Art1,")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes.Cos_Ult1 As Precio1,")
            loComandoSeleccionar.AppendLine("           Ajustes.Cod_Mon,")
            loComandoSeleccionar.AppendLine("           MAX(Ajustes.Fec_Ini) AS Fecha_Ultima_Compra")
            loComandoSeleccionar.AppendLine("FROM	Ajustes")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("                           AND Renglones_Ajustes.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		                    AND	Renglones_Ajustes.Tipo	IN	('Entrada') ")
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Ajustes.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("WHERE  Ajustes.Status IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("           AND Ajustes.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Renglones_Ajustes.Cod_Art,Renglones_Ajustes.Cos_Ult1,Ajustes.Cod_Mon")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("SELECT		  #tmpTemporal.Cod_Art,")
            loComandoSeleccionar.AppendLine("             #tmpTemporal.Can_Art1,")
            loComandoSeleccionar.AppendLine("             #tmpTemporal.Precio1,")
            loComandoSeleccionar.AppendLine("             #tmpTemporal.Cod_Mon,")
            loComandoSeleccionar.AppendLine("             #tmpTemporal.Fecha_Ultima_Compra")
            loComandoSeleccionar.AppendLine("INTO #tmpUltimasCompras")
            loComandoSeleccionar.AppendLine("FROM #tmpCompras AS #tmpTemporal")
            loComandoSeleccionar.AppendLine("	INNER JOIN (SELECT Cod_Art,MAX(Fecha_Ultima_Compra) As Fecha_Maxima FROM #tmpCompras GROUP BY Cod_Art) AS Tabla")
            loComandoSeleccionar.AppendLine("		ON #tmpTemporal.Fecha_Ultima_Compra = Tabla.Fecha_Maxima AND #tmpTemporal.Cod_Art = Tabla.Cod_Art")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Art Asc")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT      ")
            loComandoSeleccionar.AppendLine("           #tmpCotizaciones.Cod_Art,")
            loComandoSeleccionar.AppendLine("           #tmpCotizaciones.Nom_Art,")
            loComandoSeleccionar.AppendLine("           #tmpCotizaciones.Cod_Mon,")
            loComandoSeleccionar.AppendLine("           SUM(#tmpCotizaciones.Can_Art1)			AS Can_Art1_Cot,")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(#tmpFacturas.Can_Art1),0)    AS Can_Art1_Fac,")
            loComandoSeleccionar.AppendLine("           ISNULL(((SUM(#tmpFacturas.Can_Art1) / SUM(#tmpCotizaciones.Can_Art1)) * 100),0)	AS Efic1,")
            loComandoSeleccionar.AppendLine("           #tmpCotizaciones.Exi_Act1,")
            loComandoSeleccionar.AppendLine("           #tmpCotizaciones.Exi_Por1")
            loComandoSeleccionar.AppendLine("INTO		#tmpFinal")
            loComandoSeleccionar.AppendLine("FROM		#tmpCotizaciones")
            loComandoSeleccionar.AppendLine("LEFT JOIN	#tmpFacturas ON #tmpFacturas.Cod_Art = #tmpCotizaciones.Cod_Art")
            loComandoSeleccionar.AppendLine("GROUP BY	#tmpCotizaciones.Cod_Art, #tmpCotizaciones.Nom_Art,#tmpCotizaciones.Cod_Mon, #tmpCotizaciones.Exi_Act1,	#tmpCotizaciones.Exi_Por1")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT		#tmpFinal.*,")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmpUltimasCompras.Can_Art1,0)		        AS Cantidad_Comprada,")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmpUltimasCompras.Precio1,0)		        AS Costo_Ultimo,")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmpUltimasCompras.Fecha_Ultima_Compra,0)	AS Fecha_Ultima_Compra ")
            loComandoSeleccionar.AppendLine("FROM		#tmpFinal            ")
            loComandoSeleccionar.AppendLine("LEFT JOIN  #tmpUltimasCompras  ON #tmpUltimasCompras.Cod_Art = #tmpFinal.Cod_Art   ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCompras")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFinal")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpUltimasCompras")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCotizaciones ")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFacturas ")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEfectividadArticulos2", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrEfectividadArticulos2.ReportSource = loObjetoReporte


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
' MAT: 14/09/11: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 24/08/12: Se agregó agrupación a las dos primeras tablas (cotizaciones y Facturas)	'
'				 para corregir valores repetidos.											'
'-------------------------------------------------------------------------------------------'
