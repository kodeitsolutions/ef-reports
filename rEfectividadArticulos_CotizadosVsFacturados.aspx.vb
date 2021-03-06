﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEfectividadArticulos_CotizadosVsFacturados"
'-------------------------------------------------------------------------------------------'
Partial Class rEfectividadArticulos_CotizadosVsFacturados
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
            'Dim lcParametro9Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(14))
            Dim lcParametro14Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(14))



            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("SELECT												")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Art,						")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art,						")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Cotizaciones.Can_Art1)	AS Can_Art1,")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Cotizaciones.Mon_Net)		AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Act1      ")
            loComandoSeleccionar.AppendLine("INTO		#tmpCOTIZACIONES      ")
            loComandoSeleccionar.AppendLine("FROM		Cotizaciones      ")
            loComandoSeleccionar.AppendLine("   JOIN	Renglones_Cotizaciones ON Renglones_Cotizaciones.Documento = Cotizaciones.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN	Articulos ON Renglones_Cotizaciones.Cod_Art = Articulos.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   JOIN	Almacenes ON Renglones_Cotizaciones.Cod_Alm = Almacenes.Cod_Alm      ")
            loComandoSeleccionar.AppendLine("   JOIN	Clientes ON Clientes.Cod_Cli = Cotizaciones.Cod_CLi      ")
            loComandoSeleccionar.AppendLine("   JOIN	Vendedores ON Vendedores.Cod_Ven = Cotizaciones.Cod_Ven      ")

            loComandoSeleccionar.AppendLine("WHERE      ")
            loComandoSeleccionar.AppendLine(" 			Cotizaciones.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Cotizaciones.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Cotizaciones.Cod_Alm BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Status IN (" & lcParametro9Desde & ")")
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Zon BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Cla BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Tip BETWEEN " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Suc BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Rev BETWEEN " & lcParametro14Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro14Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Articulos.Cod_Art, Articulos.Nom_Art, Articulos.Exi_Act1  ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("SELECT      ")
            loComandoSeleccionar.AppendLine("        	Renglones_Facturas.Cod_Art,					")
            loComandoSeleccionar.AppendLine("        	SUM(Renglones_Facturas.Can_Art1) AS Can_Art1,")
            loComandoSeleccionar.AppendLine("        	SUM(Renglones_Facturas.Mon_Net) AS Mon_Net")
            loComandoSeleccionar.AppendLine("INTO		#tmpFACTURAS	")
            loComandoSeleccionar.AppendLine("FROM		Facturas		")
            loComandoSeleccionar.AppendLine("   JOIN	Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN	Articulos ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   JOIN	Almacenes ON Renglones_Facturas.Cod_Alm = Almacenes.Cod_Alm      ")
            loComandoSeleccionar.AppendLine("   JOIN	Clientes ON Clientes.Cod_Cli = Facturas.Cod_CLi      ")
            loComandoSeleccionar.AppendLine("   JOIN	Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven      ")

            loComandoSeleccionar.AppendLine("WHERE      ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.cod_tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.cod_Alm BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Zon BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Cla BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Tip BETWEEN " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Suc BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Rev BETWEEN " & lcParametro14Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro14Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT #tmpCOTIZACIONES.Cod_Art 														AS Cod_Art,  	")
            loComandoSeleccionar.AppendLine("   	#tmpCOTIZACIONES.Nom_Art 														AS Nom_Art,  	")
            loComandoSeleccionar.AppendLine("   	SUM(#tmpCOTIZACIONES.Can_Art1)													AS Can_Art1_Cot,")
            loComandoSeleccionar.AppendLine("   	SUM(COALESCE(#tmpFACTURAS.Can_Art1, 0))											AS Can_Art1_Fac,")
            loComandoSeleccionar.AppendLine("		CASE WHEN SUM(#tmpCOTIZACIONES.Can_Art1) > 0													")
            loComandoSeleccionar.AppendLine("			THEN SUM(COALESCE(#tmpFACTURAS.Can_Art1, 0))/ SUM(#tmpCOTIZACIONES.Can_Art1)*100			")
            loComandoSeleccionar.AppendLine("			ELSE 100																					")
            loComandoSeleccionar.AppendLine("		END																				AS Efic1,		")
            loComandoSeleccionar.AppendLine("		SUM(#tmpCOTIZACIONES.Mon_Net)													AS Mon_Net_Cot, ")
            loComandoSeleccionar.AppendLine("		SUM(COALESCE(#tmpFACTURAS.Mon_Net,0))											AS Mon_Net_Fac, ")
            loComandoSeleccionar.AppendLine("		CASE WHEN SUM(#tmpCOTIZACIONES.Mon_Net) > 0														")
            loComandoSeleccionar.AppendLine("			THEN SUM(COALESCE(#tmpFACTURAS.Mon_Net, 0)) / SUM(#tmpCOTIZACIONES.Mon_Net) * 100			")
            loComandoSeleccionar.AppendLine("			ELSE 100																					")
            loComandoSeleccionar.AppendLine("		END																				AS Efic2,		")
            loComandoSeleccionar.AppendLine("		#tmpCOTIZACIONES.Exi_Act1 														AS Exi_Act1		")
            loComandoSeleccionar.AppendLine("FROM	#tmpCOTIZACIONES")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpFACTURAS ON #tmpFACTURAS.Cod_Art = #tmpCOTIZACIONES.Cod_Art")
            loComandoSeleccionar.AppendLine("GROUP BY	#tmpCOTIZACIONES.Cod_Art, #tmpCOTIZACIONES.Nom_Art, #tmpCOTIZACIONES.Exi_Act1")
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCOTIZACIONES")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFACTURAS")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
			
				
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())


			Dim loServicios As New cusDatos.goDatos
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEfectividadArticulos_CotizadosVsFacturados", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)
			Me.mFormatearCamposReporte(loObjetoReporte)
			Me.crvrEfectividadArticulos_CotizadosVsFacturados.ReportSource = loObjetoReporte


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
' CMS: 08/06/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' AAP: 29/06/09: Filtro "Sucursal"
'-------------------------------------------------------------------------------------------'
' RJG: 09/12/10: Ajustado el estatus de las facturas de venta en el filtro. (el estatus ya	'
'				 no es un parámetro)														' 
'-------------------------------------------------------------------------------------------'
' RJG: 09/10/12: Se corrigió error de divición por cero en el último SELECT.				' 
'-------------------------------------------------------------------------------------------'
