'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTPedidos_ArticulosFechasVendedores"
'-------------------------------------------------------------------------------------------'
Partial Class rTPedidos_ArticulosFechasVendedores
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
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosIniciales(6) 'goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosIniciales(8)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" EXEC sp_TPedidos_AriculosFechasVendedores")
            loComandoSeleccionar.AppendLine(" 	@sp_lcFecIni = " & lcParametro0Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcFecFin = " & lcParametro0Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodArtIni = " & lcParametro1Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodArtFin = " & lcParametro1Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodCliIni = " & lcParametro2Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodCliFin = " & lcParametro2Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodVenIni = " & lcParametro3Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodVenFin = " & lcParametro3Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodDepIni = " & lcParametro4Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodDepFin = " & lcParametro4Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodTipIni = " & lcParametro5Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodTipFin = " & lcParametro5Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcStatus  = '" & lcParametro6Desde & "',")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodRevIni = " & lcParametro7Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodRevFin = " & lcParametro7Hasta & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcTipRev  = '" & lcParametro8Desde & "',")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodSucIni = " & lcParametro9Desde & ",")
            loComandoSeleccionar.AppendLine(" 	@sp_lcCodSucFin = " & lcParametro9Hasta & ",")
            loComandoSeleccionar.AppendLine("   @sp_NumDecCant = " & cusAplicacion.goOpciones.pnDecimalesParaCantidad)

            'Response.Clear()
            'Response.Write("<html><body><pre>" & vbNewLine)
            'Response.Write(loComandoSeleccionar.ToString)
            'Response.Write("</pre></body></html>")
            'Response.Flush()
            'Response.End()

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTPedidos_ArticulosFechasVendedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTPedidos_ArticulosFechasVendedores.ReportSource = loObjetoReporte

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
' Douglas Cortez:  27/05/2010 : Codigo inicial
'-------------------------------------------------------------------------------------------'
