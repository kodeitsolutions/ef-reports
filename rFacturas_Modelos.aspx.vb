'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rFacturas_Modelos"
'-------------------------------------------------------------------------------------------'
Partial Class rFacturas_Modelos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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
            Dim lcParametro10Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As String = cusAplicacion.goReportes.paParametrosIniciales(14)
            Dim lcParametro15Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(15))
            Dim lcParametro15Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(15))
            Dim lcParametro16Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(16))
            Dim lcParametro16Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(16))

            If (lcParametro16Desde = "'Sin Modelo'") Then
                lcParametro16Desde = "''"
            End If

            If (lcParametro16Hasta = "'Sin Modelo'") Then
                lcParametro16Hasta = "''"
            End If

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 			CASE ")
            loComandoSeleccionar.AppendLine(" 				WHEN Articulos.Modelo = '' THEN 'Sin Modelo' ")
            loComandoSeleccionar.AppendLine(" 			ELSE  ")
            loComandoSeleccionar.AppendLine(" 				Articulos.Modelo ")
            loComandoSeleccionar.AppendLine(" 			END AS Modelo,  ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Documento,			 ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 			SUM(Renglones_Facturas.Can_Art1) AS Can_Art1, ")
            loComandoSeleccionar.AppendLine(" 			SUBSTRING(Cast (Facturas.Comentario As VARCHAR),1,30) AS Comentario")
            loComandoSeleccionar.AppendLine(" FROM Facturas ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento ")
            loComandoSeleccionar.AppendLine(" JOIN Articulos ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" WHERE	 ")
            loComandoSeleccionar.AppendLine("           Facturas.Documento                  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.Cod_Art		BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Fec_Ini				BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Cli				BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Ven				BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep       		BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Cla				BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Mon				BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Tra				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar				BETWEEN" & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Status					IN (" & lcParametro10Desde & ")")
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.Cod_Alm		BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_For				BETWEEN" & lcParametro12Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro12Hasta)

            If lcParametro14Desde = "Igual" Then
                loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Rev                BETWEEN " & lcParametro13Desde)
            Else
                loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Rev           NOT  BETWEEN " & lcParametro13Desde)
            End If

            loComandoSeleccionar.AppendLine("    		AND " & lcParametro13Hasta)

            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Suc				BETWEEN" & lcParametro15Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro15Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Modelo 				BETWEEN" & lcParametro16Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro16Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Articulos.Modelo, Facturas.Documento, Facturas.Fec_Ini, Facturas.Cod_Cli, Clientes.Nom_Cli, SUBSTRING(Cast (Facturas.Comentario As VARCHAR),1,30)")
            loComandoSeleccionar.AppendLine(" ORDER BY Articulos.Modelo, " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFacturas_Modelos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrFacturas_Modelos.ReportSource = loObjetoReporte

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
' CMS: 02/09/08: Codigo inicial.
'-------------------------------------------------------------------------------------------'