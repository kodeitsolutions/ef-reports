'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCDiario_CentroCostos"
'-------------------------------------------------------------------------------------------'
Partial Class rCDiario_CentroCostos
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

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      SUM(Renglones_Comprobantes.Mon_Deb) AS Mon_Deb,")
            loConsulta.AppendLine("			SUM(Renglones_Comprobantes.Mon_Hab) AS Mon_Hab,")
            loConsulta.AppendLine("			Renglones_Comprobantes.Cod_Cen,")
            loConsulta.AppendLine("			Centros_Costos.Nom_Cen")
            loConsulta.AppendLine("FROM        Renglones_Comprobantes WITH(INDEX(IX_Fec_Ini_VariosCampos))")
            loConsulta.AppendLine("    JOIN    Comprobantes WITH(INDEX(PK_Comprobantes))")
            loConsulta.AppendLine("        ON  Comprobantes.Documento = Renglones_Comprobantes.Documento")
            loConsulta.AppendLine("        AND Comprobantes.Adicional = Renglones_Comprobantes.Adicional")
            loConsulta.AppendLine("    JOIN    Centros_Costos")
            loConsulta.AppendLine("       ON   Centros_Costos.Cod_Cen = Renglones_Comprobantes.Cod_Cen")
            loConsulta.AppendLine("WHERE       Renglones_Comprobantes.Documento  BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("                AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Renglones_Comprobantes.Fec_Ini    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("                AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Renglones_Comprobantes.Cod_Mon   BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("                AND " & lcParametro2Hasta)
            loConsulta.AppendLine("GROUP BY  Renglones_Comprobantes.Cod_Cen, Centros_Costos.Nom_Cen")
            loConsulta.AppendLine("ORDER BY  " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes", 900)

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCDiario_CentroCostos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCDiario_CentroCostos.ReportSource = loObjetoReporte

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
' CMS:  15/09/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  04/12/13: Codigo inicial
'-------------------------------------------------------------------------------------------'
