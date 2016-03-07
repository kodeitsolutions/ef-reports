Imports System.Data
Partial Class rPagos_Vendedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("			Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Pagos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Proveedores.Nom_Pro,1,50)    AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("			Pagos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("			Pagos.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("			Pagos.Status, ")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM		Proveedores, ")
            loComandoSeleccionar.AppendLine("			Pagos, ")
            loComandoSeleccionar.AppendLine("			Vendedores ")
            loComandoSeleccionar.AppendLine(" WHERE		Pagos.Cod_Pro          =	Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("			And Pagos.Cod_Ven      =	Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("			And Pagos.Documento	Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Pagos.Fec_Ini	    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Pagos.Cod_Pro      Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Pagos.Cod_Ven      Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			And Pagos.Status       IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("			And Pagos.Cod_Mon      Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			And Pagos.Cod_rev      Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			And Pagos.Cod_Suc      Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro7Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Pagos.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPagos_Vendedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPagos_Vendedores.ReportSource = loObjetoReporte

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
' JJD: 10/01/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  13/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' RJG:  10/04/12: Se agregó el total de registros.											'
'-------------------------------------------------------------------------------------------'
