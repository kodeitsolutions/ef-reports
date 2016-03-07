Imports System.Data
Imports cusAplicacion

Partial Class rOPagos_Conceptos2
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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Ordenes_Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Status, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Motivo, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Tasa, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Deb,")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Mon_Hab,")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos.Cod_Con,")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Pagos, ")
            loComandoSeleccionar.AppendLine("           Renglones_oPagos, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Conceptos")
            loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Pagos.Cod_Pro		=   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Documento = Renglones_oPagos.Documento")
            loComandoSeleccionar.AppendLine("           And Conceptos.Cod_Con		= Renglones_oPagos.Cod_Con")
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Documento Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Fec_Ini   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Cod_Pro   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Status    IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Renglones_oPagos.Cod_Con  Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Cod_Mon  Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			AND Ordenes_Pagos.Cod_rev  Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)


            'loComandoSeleccionar.AppendLine(" ORDER BY  Ordenes_Pagos.Documento, Ordenes_Pagos.Cod_Pro ")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_Conceptos2", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_Conceptos2.ReportSource = loObjetoReporte

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
' GCR: 12/03/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
