'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rFacturas_UltimoCostosProveedores"
'-------------------------------------------------------------------------------------------'

Partial Class rFacturas_UltimoCostosProveedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro10Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            
            
			loComandoSeleccionar.AppendLine("SELECT		Compras.Cod_Pro,  ")
			loComandoSeleccionar.AppendLine("  			Renglones_Compras.Cod_Art,  ")
			loComandoSeleccionar.AppendLine("  			Compras.Fec_Ini,  ")
			loComandoSeleccionar.AppendLine("  			Renglones_Compras.Precio1, ")
			loComandoSeleccionar.AppendLine("  			Proveedores.Nom_Pro,  ")
			loComandoSeleccionar.AppendLine("  			Articulos.Nom_Art  ")
			loComandoSeleccionar.AppendLine("INTO 		#tmpDatos  ")
			loComandoSeleccionar.AppendLine("FROM		Compras  ")
			loComandoSeleccionar.AppendLine("	JOIN 	Renglones_Compras ON Compras.Documento = Renglones_Compras.Documento  ")
			loComandoSeleccionar.AppendLine("	JOIN 	Proveedores ON Compras.Cod_Pro = Proveedores.Cod_Pro   ")
			loComandoSeleccionar.AppendLine("	JOIN 	Articulos ON Articulos.Cod_Art = Renglones_Compras.Cod_Art  ")
			loComandoSeleccionar.AppendLine(" WHERE		")			
            loComandoSeleccionar.AppendLine("           Renglones_Compras.Cod_Art   BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Fec_Ini             BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Pro             BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Ven             BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep           BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Mar           BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Status              IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Renglones_Compras.Cod_Alm   BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Mon             BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Tra             BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Proveedores.Status              IN (" & lcParametro10Desde & ")") 
			loComandoSeleccionar.AppendLine("SELECT		")
			loComandoSeleccionar.AppendLine("			Renglones_Compras.cod_Art, ")
			loComandoSeleccionar.AppendLine("			Compras.Cod_Pro, ")
			loComandoSeleccionar.AppendLine("			MAX(Compras.Fec_Ini) AS Fec_Ini")
			loComandoSeleccionar.AppendLine("INTO 		#tmpFechaUltimoPrecio")
			loComandoSeleccionar.AppendLine("FROM 		Compras")
			loComandoSeleccionar.AppendLine("	JOIN	Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
			loComandoSeleccionar.AppendLine("WHERE ")	
			loComandoSeleccionar.AppendLine("           Renglones_Compras.Cod_Art   BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)		
            loComandoSeleccionar.AppendLine("           AND Compras.Fec_Ini	BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Pro	BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Ven	BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Status	IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Renglones_Compras.Cod_Alm   BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Mon	BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Compras.Cod_Tra	BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
			loComandoSeleccionar.AppendLine("GROUP BY  Renglones_Compras.cod_Art, Compras.Cod_Pro")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		")
			loComandoSeleccionar.AppendLine(" 			#tmpDatos.Cod_Pro,  ")
			loComandoSeleccionar.AppendLine(" 			#tmpDatos.Cod_Art,")
			loComandoSeleccionar.AppendLine(" 			#tmpDatos.Fec_Ini,  ")
			loComandoSeleccionar.AppendLine(" 			#tmpDatos.Precio1,")
			loComandoSeleccionar.AppendLine(" 			#tmpDatos.Nom_Pro,  ")
			loComandoSeleccionar.AppendLine(" 			#tmpDatos.Nom_Art  ")
			loComandoSeleccionar.AppendLine("FROM		#tmpDatos")
			loComandoSeleccionar.AppendLine("	JOIN	#tmpFechaUltimoPrecio ON (#tmpFechaUltimoPrecio.Fec_Ini = #tmpDatos.Fec_Ini) AND (#tmpFechaUltimoPrecio.Cod_Pro = #tmpDatos.Cod_Pro) AND (#tmpFechaUltimoPrecio.Cod_Art = #tmpDatos.Cod_Art) ")
            loComandoSeleccionar.AppendLine("ORDER BY   #tmpDatos.Cod_Art, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFacturas_UltimoCostosProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCompras_UltimoCostosProveedores.ReportSource = loObjetoReporte

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
' MAT: 01/04/11: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
' RJG: 26/01/12: Correccion menor en layout (etiquetas).									'
'-------------------------------------------------------------------------------------------'
