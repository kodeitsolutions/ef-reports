Imports System.Data
Partial Class rComprobantes_Presupuesto
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT     Presupuesto.Documento, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.resumen,  ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Tipo,  ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Origen,  ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Notas, ")
            loComandoSeleccionar.AppendLine("		    Renglones_Presupuesto.Renglon,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Cue,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Cen,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Gas,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Mon_Deb,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Mon_Hab,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Comentario,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Tip_Ori,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Doc_Ori,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Reg, ")
            loComandoSeleccionar.AppendLine("		    COALESCE(Cuentas_Contables.Nom_Cue,'') AS Nom_Cue, ")
            loComandoSeleccionar.AppendLine("		    COALESCE(Centros_Costos.Nom_Cen,'') AS Nom_Cen, ")
            loComandoSeleccionar.AppendLine("		    COALESCE(Cuentas_Gastos.Nom_Gas,'') AS Nom_Gas ")
            loComandoSeleccionar.AppendLine("FROM Presupuesto ")
            loComandoSeleccionar.AppendLine("JOIN Renglones_Presupuesto ON Renglones_Presupuesto.documento = Presupuesto.Documento ")
            loComandoSeleccionar.AppendLine("LEFT JOIN Cuentas_Contables ON Cuentas_Contables.Cod_Cue = Renglones_Presupuesto.Cod_Cue ")
            loComandoSeleccionar.AppendLine("LEFT JOIN Centros_Costos ON Centros_Costos.cod_cen = Renglones_Presupuesto.cod_cen ")
            loComandoSeleccionar.AppendLine("LEFT JOIN Cuentas_Gastos ON Cuentas_Gastos.Cod_Gas = Renglones_Presupuesto.cod_gas ")
            loComandoSeleccionar.AppendLine(" WHERE     Presupuesto.Documento                  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Presupuesto.Fec_Ini                    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_Presupuesto.Cod_Mon          Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rComprobantes_Presupuesto", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrComprobantes_Presupuesto.ReportSource = loObjetoReporte

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
' MAT: 16/05/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
