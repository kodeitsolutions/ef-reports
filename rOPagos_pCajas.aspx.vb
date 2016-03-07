'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOPagos_pCajas"
'-------------------------------------------------------------------------------------------'
Partial Class rOPagos_pCajas

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
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Ordenes_Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Status, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Motivo, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Detalles_OPagos.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Detalles_OPagos.Comentario As Comentario_Renglon, ")
            loComandoSeleccionar.AppendLine("           Detalles_OPagos.Tasa, ")
            loComandoSeleccionar.AppendLine("           Detalles_OPagos.Tip_Ope, ")
            loComandoSeleccionar.AppendLine("           Detalles_OPagos.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("           Cajas.Nom_Caj, ")
            loComandoSeleccionar.AppendLine("           Detalles_OPagos.Fec_Ini As Fec_Ini_Renglon, ")
            loComandoSeleccionar.AppendLine("           Detalles_OPagos.Mon_Net         AS  Mon_Caj ")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Pagos, ")
            loComandoSeleccionar.AppendLine("           Detalles_OPagos, ")
            loComandoSeleccionar.AppendLine("           Cajas, ")
            loComandoSeleccionar.AppendLine("           Proveedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Pagos.Cod_Pro       =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Documento =   Detalles_OPagos.Documento ")
            loComandoSeleccionar.AppendLine("           And Detalles_OPagos.Cod_Caj =   Cajas.Cod_Caj ")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Documento Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Fec_Ini   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Pro   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Status    IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Mon   Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Rev   Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Suc   Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Detalles_OPagos.Cod_Caj Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Detalles_OPagos.Cod_Caj <> '' " )
            loComandoSeleccionar.AppendLine(" ORDER BY  Detalles_OPagos.Cod_Caj, " & lcOrdenamiento & ", Detalles_OPagos.Fec_Ini Desc")

            'Me.Response.Clear()
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_pCajas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_pCajas.ReportSource = loObjetoReporte

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
' JJD: 07/05/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
