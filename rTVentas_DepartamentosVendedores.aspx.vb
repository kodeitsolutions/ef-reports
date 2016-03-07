'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_DepartamentosVendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_DepartamentosVendedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = lcParametro2Desde
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
            Dim lnParametro12Desde As Integer = cusAplicacion.goReportes.paParametrosIniciales(12)
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lnParametro14Desde As Integer = cusAplicacion.goReportes.paParametrosIniciales(14)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Articulos.Cod_Dep				AS Cod_Dep, ")
            loComandoSeleccionar.AppendLine("    		Departamentos.Nom_Dep			AS Nom_Dep, ")
            loComandoSeleccionar.AppendLine("    		Renglones_Facturas.Can_Art1     AS Can_Art, ")
            loComandoSeleccionar.AppendLine("    		Vendedores.Cod_Ven				AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("    		Vendedores.Nom_Ven				AS Nom_Ven	")
            loComandoSeleccionar.AppendLine("INTO		#curTemporal								")
            loComandoSeleccionar.AppendLine("FROM    	Facturas									")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas	ON Renglones_Facturas.Documento	= Facturas.Documento        ")
            loComandoSeleccionar.AppendLine("	JOIN	Articulos			ON Articulos.Cod_Art			= Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN	Departamentos   	ON Departamentos.Cod_Dep		= Articulos.Cod_Dep         ")
            loComandoSeleccionar.AppendLine("	JOIN	Vendedores			ON Facturas.Cod_Ven				= Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("WHERE		Facturas.Fec_Ini                BETWEEN " & lcParametro0Desde & "")
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta )
            loComandoSeleccionar.AppendLine("       AND Facturas.Cod_Cli            BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta )
            loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Dep           BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta )
            loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Sec           BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta )
            loComandoSeleccionar.AppendLine("       AND Renglones_Facturas.Cod_Art  BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro5Hasta )
            loComandoSeleccionar.AppendLine("       AND Renglones_Facturas.Cod_Alm  BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta )
            loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Cla           BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro7Hasta )
            loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Tip           BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro8Hasta )
            loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Pro			BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta )
            loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Mon           BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("       AND Facturas.Cod_Rev             BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("    		AND " & lcParametro13Hasta & "")
            loComandoSeleccionar.AppendLine("       AND Facturas.Status IN (" & lcParametro11Desde & ")")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            If lnParametro12Desde > 0 Then
				loComandoSeleccionar.AppendLine("SELECT		TOP "  &  lnParametro12Desde.ToString() &  " Cod_Dep,	")
            Else								 
				loComandoSeleccionar.AppendLine("SELECT		Cod_Dep,	")
            End If
            loComandoSeleccionar.AppendLine("			SUM(Can_Art) AS Can_Dep")
            loComandoSeleccionar.AppendLine("INTO		#tmpTopDepartamentos")
            loComandoSeleccionar.AppendLine("FROM		#curTemporal")
            loComandoSeleccionar.AppendLine("GROUP BY	Cod_Dep")
            loComandoSeleccionar.AppendLine("")
            If lnParametro14Desde > 0 Then
				loComandoSeleccionar.AppendLine("SELECT		TOP "  &  lnParametro14Desde.ToString() &  " Cod_Ven,	")            
            Else
				loComandoSeleccionar.AppendLine("SELECT		Cod_Ven,	")
            End If
            loComandoSeleccionar.AppendLine("			SUM(Can_Art) AS Can_Ven")
            loComandoSeleccionar.AppendLine("INTO		#tmpTopVendedores")
            loComandoSeleccionar.AppendLine("FROM		#curTemporal")
            loComandoSeleccionar.AppendLine("GROUP BY	Cod_Ven")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#curTemporal.Cod_dep, #curTemporal.Nom_Dep, ")
            loComandoSeleccionar.AppendLine("		#curTemporal.Can_Art, #curTemporal.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("		#curTemporal.Nom_Ven  ,Can_Dep, Can_Ven")
            loComandoSeleccionar.AppendLine("FROM	#curTemporal")
            loComandoSeleccionar.AppendLine("	JOIN #tmpTopDepartamentos ON #tmpTopDepartamentos.Cod_Dep = #curTemporal.Cod_Dep")
            loComandoSeleccionar.AppendLine("	JOIN #tmpTopVendedores ON #tmpTopVendedores.Cod_Ven = #curTemporal.Cod_Ven")
            'loComandoSeleccionar.AppendLine("ORDER BY Can_Dep DESC, Can_Ven DESC")
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento & ", Can_Ven DESC" )
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #curTemporal")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpTopDepartamentos")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpTopVendedores")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())		  			
			
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_DepartamentosVendedores", laDatosReporte)


            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTVentas_DepartamentosVendedores.ReportSource = loObjetoReporte

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
' JJD: 05/12/09: Programacion inicial.														'
'-------------------------------------------------------------------------------------------'
' RJG: 19/01/12: Ajuste en las uniones y filtros. Se eliminaron algunas tablas que no eran	'
'				 necesarias. Se eliminó el pivot y cambió el reporte por un agrupamiento por'
'				 departamento.																'
'-------------------------------------------------------------------------------------------'
