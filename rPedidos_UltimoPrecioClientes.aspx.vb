'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPedidos_UltimoPrecioClientes"
'-------------------------------------------------------------------------------------------'

Partial Class rPedidos_UltimoPrecioClientes

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
            
            loComandoSeleccionar.AppendLine(" SELECT Pedidos.Cod_Cli,    ")
            loComandoSeleccionar.AppendLine("  		Renglones_Pedidos.Cod_Art,    ")
            loComandoSeleccionar.AppendLine("  		Pedidos.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("  		Renglones_Pedidos.Precio1,   ")
            loComandoSeleccionar.AppendLine("  		Clientes.Nom_Cli,    ")
            loComandoSeleccionar.AppendLine("  		Articulos.Nom_Art    ")
            loComandoSeleccionar.AppendLine("  INTO #tmpDatos    ")
            loComandoSeleccionar.AppendLine("  FROM	Pedidos    ")
            loComandoSeleccionar.AppendLine("  		JOIN Renglones_Pedidos ON Pedidos.Documento = Renglones_Pedidos.Documento   ")
            loComandoSeleccionar.AppendLine("  		JOIN Clientes ON Pedidos.Cod_Cli = Clientes.Cod_Cli     ")
            loComandoSeleccionar.AppendLine("  		JOIN Articulos ON Articulos.Cod_Art = Renglones_Pedidos.Cod_Art    ")
            loComandoSeleccionar.AppendLine("  WHERE ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Art   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Fec_Ini             Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Cli             Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Ven             Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep           Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar           Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Status              IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           And Renglones_Pedidos.Cod_Alm   Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Mon             Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Tra             Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Status              IN (" & lcParametro10Desde & ")")
            
            loComandoSeleccionar.AppendLine(" SELECT   ")
            loComandoSeleccionar.AppendLine(" 		Renglones_Pedidos.cod_Art,   ")
            loComandoSeleccionar.AppendLine(" 		Pedidos.Cod_Cli,   ")
            loComandoSeleccionar.AppendLine(" 		MAX(Pedidos.Fec_Ini) AS Fec_Ini  ")
            loComandoSeleccionar.AppendLine(" INTO #tmpFechaUltimoPrecio  ")
            loComandoSeleccionar.AppendLine(" FROM Pedidos  ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Pedidos ON Renglones_Pedidos.Documento = Pedidos.Documento  ")
            loComandoSeleccionar.AppendLine(" WHERE   ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Art   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Fec_Ini             Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Cli             Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Ven             Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Status              IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           And Renglones_Pedidos.Cod_Alm   Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Mon             Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Tra             Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
			loComandoSeleccionar.AppendLine(" GROUP BY  Renglones_Pedidos.cod_Art, Pedidos.Cod_Cli ")
			
			loComandoSeleccionar.AppendLine(" SELECT   ")
            loComandoSeleccionar.AppendLine(" 		#tmpDatos.Cod_Cli,    ")
            loComandoSeleccionar.AppendLine(" 		#tmpDatos.Cod_Art,  ")
            loComandoSeleccionar.AppendLine(" 		#tmpDatos.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine(" 		#tmpDatos.Precio1,  ")
            loComandoSeleccionar.AppendLine(" 		#tmpDatos.Nom_Cli,    ")
            loComandoSeleccionar.AppendLine(" 		#tmpDatos.Nom_Art    ")
            loComandoSeleccionar.AppendLine(" FROM #tmpDatos  ")
            loComandoSeleccionar.AppendLine(" JOIN #tmpFechaUltimoPrecio ON (#tmpFechaUltimoPrecio.Fec_Ini = #tmpDatos.Fec_Ini) AND (#tmpFechaUltimoPrecio.Cod_Cli = #tmpDatos.Cod_Cli) AND (#tmpFechaUltimoPrecio.Cod_Art = #tmpDatos.Cod_Art)   ")
            loComandoSeleccionar.AppendLine("ORDER BY  #tmpDatos.Cod_Art, " & lcOrdenamiento)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_UltimoPrecioClientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_UltimoPrecioClientes.ReportSource = loObjetoReporte

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
' CMS: 25/02/10: Codigo inicial.
'-------------------------------------------------------------------------------------------'