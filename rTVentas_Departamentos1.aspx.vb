'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_Departamentos1"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_Departamentos1
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
            Dim lcParametro11Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lnParametro12Desde As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(14))
            Dim lcParametro14Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(14))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			
            Dim lcSeleccionarPrimeros As String

            Dim loComandoSeleccionar As New StringBuilder()	 

            If lnParametro12Desde > 0 Then
                lcSeleccionarPrimeros = "SELECT TOP		" + lnParametro12Desde.ToString()
            Else
                lcSeleccionarPrimeros = "SELECT			"
            End If

            loComandoSeleccionar.AppendLine("SELECT			")
            loComandoSeleccionar.AppendLine("         		Articulos.Cod_Dep											AS Cod_Dep, ")
            loComandoSeleccionar.AppendLine("         		Departamentos.Nom_Dep										AS Nom_Dep, ")
            loComandoSeleccionar.AppendLine("         		Renglones_Facturas.Can_Art1									AS Can_Art, ")
            loComandoSeleccionar.AppendLine("         		Renglones_Facturas.Mon_Net									AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("         		(Renglones_Facturas.Cos_Ult1 * Renglones_Facturas.Can_Art1)	AS Mon_Cos ")
            loComandoSeleccionar.AppendLine("INTO			#curTemporal ")
            loComandoSeleccionar.AppendLine("FROM			Facturas ")
            loComandoSeleccionar.AppendLine("	JOIN		Renglones_Facturas")
            loComandoSeleccionar.AppendLine("			ON	Renglones_Facturas.Documento = Facturas.Documento")
			loComandoSeleccionar.AppendLine("			AND Renglones_Facturas.Cod_Art	BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			AND Renglones_Facturas.Cod_Alm	BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("	JOIN		Articulos  ")
            loComandoSeleccionar.AppendLine("			ON	Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Dep			BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("	 		AND Articulos.Cod_Sec			BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Cla			BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Tip			BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Pro			BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro9Hasta)
         	loComandoSeleccionar.AppendLine("	JOIN		Departamentos ON (Departamentos.Cod_Dep  = Articulos.Cod_Dep)")
            loComandoSeleccionar.AppendLine("WHERE			Facturas.Status <> 'Anulado' ")
         	loComandoSeleccionar.AppendLine("           AND Facturas.Fec_Ini			BETWEEN " & lcParametro0Desde)
         	loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
         	loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Cli			BETWEEN " & lcParametro1Desde)
         	loComandoSeleccionar.AppendLine("				AND " & lcParametro1Hasta)
         	loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven			BETWEEN " & lcParametro2Desde)
         	loComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Mon			BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Status				IN ( " & lcParametro11Desde & " ) ")
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Rev			BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine("			AND Facturas.Cod_Suc			BETWEEN " & lcParametro14Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro14Hasta)

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE	@lnTotalNeto DECIMAL(28, 10)")
            loComandoSeleccionar.AppendLine("SET		@lnTotalNeto  = (SELECT SUM(Mon_Net) FROM #curTemporal)")		    
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
							
            'loComandoSeleccionar.AppendLine("SELECT			")
            loComandoSeleccionar.AppendLine(lcSeleccionarPrimeros)
            loComandoSeleccionar.AppendLine(" 			Cod_Dep									AS  Cod_Dep,")
            loComandoSeleccionar.AppendLine(" 			Nom_Dep									AS  Nom_Dep,")
            loComandoSeleccionar.AppendLine(" 			SUM(Can_Art)							AS  Can_Art,")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Net)            				AS  Mon_Net,")			   
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Cos)            				AS  Mon_Cos,")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Net-Mon_Cos)    				AS  Mon_Gan,")
            loComandoSeleccionar.AppendLine(" 			SUM(Mon_Net-Mon_Cos)/SUM(Mon_Net)*100	AS  Por_Gan,")
            loComandoSeleccionar.AppendLine(" 			@lnTotalNeto               				AS  Tot_Net,")
            loComandoSeleccionar.AppendLine("			(SUM(Mon_Net)/@lnTotalNeto)*100 		AS  Por_Ven ")
            loComandoSeleccionar.AppendLine("FROM		#curTemporal ")
            loComandoSeleccionar.AppendLine("GROUP BY	Cod_Dep, Nom_Dep ")	
            loComandoSeleccionar.AppendLine("ORDER  BY	 " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #curTemporal")
            loComandoSeleccionar.AppendLine("")



            Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_Departamentos1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTVentas_Departamentos1.ReportSource = loObjetoReporte

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
' JFP: 07/10/08: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' CMS:  08/05/09: Estandarización del codigo, Ordenamiento y se agregaron las restricciones	'
'       Seccion, Artículo, Almacén, Clases de Artículos, Tipos de Artículos, Proveedor,		'
'       Moneda, Estatus y el top (los mejores)												'
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”														'
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”														'
'-------------------------------------------------------------------------------------------'
' MAT: 01/02/11: Programación y Ajuste del reporte (No mostraba información Alguna)			'
'-------------------------------------------------------------------------------------------'
' RJG: 26/01/12: Ajuste en uniones y filtros: eliminadas tablas innecesarias en uniones,	'	
'				 ajute en filtro de monedas (aplica a Compras, no a Clientes).				'
'-------------------------------------------------------------------------------------------'
