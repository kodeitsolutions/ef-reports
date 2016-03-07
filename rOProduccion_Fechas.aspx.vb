'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOProduccion_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rOProduccion_Fechas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("   SELECT      Ordenes_Produccion.Documento    AS Doc_Pro,     ")
            loComandoSeleccionar.AppendLine("	            Ordenes_Produccion.Fec_Ini      AS Fec_Pro,     ")
            loComandoSeleccionar.AppendLine("	            Ordenes_Produccion.Cod_Art      AS Cod_Art_Pro, ")
            loComandoSeleccionar.AppendLine("	            Ordenes_Produccion.Nom_Art      AS Nom_Art_Pro, ")
            loComandoSeleccionar.AppendLine("	            Ordenes_Produccion.Status       AS Status_Pro,  ")
            loComandoSeleccionar.AppendLine("	            Ordenes_Produccion.Modelo       AS Modelo_Pro,  ")
            loComandoSeleccionar.AppendLine("	            Ordenes_Produccion.Cod_Uni      AS Cod_Uni_Pro, ")
            loComandoSeleccionar.AppendLine("	            Ordenes_Produccion.Can_Art1     AS Can_Art1_Pro, " & lcParametro7Desde & " AS Firma ")
            loComandoSeleccionar.AppendLine("    FROM       Ordenes_Produccion, Articulos                   ")
            loComandoSeleccionar.AppendLine("    WHERE      Ordenes_Produccion.Cod_Art          =   Articulos.Cod_Art   ")
            loComandoSeleccionar.AppendLine("               AND Ordenes_Produccion.Documento    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			    AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			    AND Ordenes_Produccion.Fec_Ini      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			    AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			    AND Ordenes_Produccion.Cod_Art      Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			    AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			    AND Articulos.Cod_Dep               Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			    AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			    AND Articulos.Modelo                Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			    AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			    AND Ordenes_Produccion.Status       IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine(" 			    AND Ordenes_Produccion.Cod_Suc      Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			    AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("    ORDER BY   Ordenes_Produccion.Documento, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOProduccion_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOProduccion_Fechas.ReportSource = loObjetoReporte

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
' JJD: 24/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
