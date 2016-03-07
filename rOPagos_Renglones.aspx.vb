Imports System.Data
Partial Class rOPagos_Renglones

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
          
         
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Ordenes_Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Status, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Motivo, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Tasa, ")
            loComandoSeleccionar.AppendLine("           Renglones_OPagos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_OPagos.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con, ")
            loComandoSeleccionar.AppendLine("           Renglones_OPagos.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_OPagos.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           (Renglones_OPagos.Mon_Imp1 + Renglones_OPagos.Mon_Imp2 + Renglones_OPagos.Mon_Imp3) AS  Mon_Imp ")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Pagos, ")
            loComandoSeleccionar.AppendLine("           Renglones_OPagos, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Conceptos ")
            loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Pagos.Documento			=   Renglones_OPagos.Documento ")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Pro		=   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Renglones_OPagos.Cod_Con    =   Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Documento Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Fec_Ini   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Pro   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Status    IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           And Renglones_OPagos.Cod_Con   Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_rev   Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Suc   Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  Ordenes_Pagos.Documento, Renglones_OPagos.Renglon ")

            'Me.Response.Clear()
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return  

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_Renglones.ReportSource = loObjetoReporte

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
' JJD: 24/01/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' GCR: 27/03/09: Ajustes al Diseño
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro Revisiones
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
