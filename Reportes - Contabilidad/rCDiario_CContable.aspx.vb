Imports System.Data
Partial Class rCDiario_CContable
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

            loComandoSeleccionar.AppendLine(" SELECT	Comprobantes.Documento, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           DATEPART(yyyy,Renglones_Comprobantes.Fec_Ini)   AS Anno, ")
            loComandoSeleccionar.AppendLine("           DATEPART(mm,Renglones_Comprobantes.Fec_Ini)     AS Mes ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes01 ")
            loComandoSeleccionar.AppendLine(" FROM      Comprobantes, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine(" WHERE     Comprobantes.Documento                      =   Renglones_Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("           And Comprobantes.Documento                  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Comprobantes.Fec_Ini                    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_Comprobantes.Cod_Mon          Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And YEAR(Renglones_Comprobantes.Fec_Ini)    =   YEAR(Comprobantes.Fec_Ini) ")
            loComandoSeleccionar.AppendLine("           And MONTH(Renglones_Comprobantes.Fec_Ini)   =   MONTH(Comprobantes.Fec_Ini) ")

            loComandoSeleccionar.AppendLine(" SELECT    Documento, ")
            loComandoSeleccionar.AppendLine("           Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Deb) AS Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Hab) AS Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Anno, ")
            loComandoSeleccionar.AppendLine("           Mes ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes02 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes01 ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Anno, ")
            loComandoSeleccionar.AppendLine("           Mes, ")
            loComandoSeleccionar.AppendLine("           Documento, ")
            loComandoSeleccionar.AppendLine("           Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT    #tmpComprobantes02.*, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Resumen, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Tipo, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Origen, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Integracion, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Status, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Notas ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes03 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes02, Comprobantes ")
            loComandoSeleccionar.AppendLine(" WHERE     #tmpComprobantes02.Documento    =   Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("           And #tmpComprobantes02.Anno     =   YEAR(Comprobantes.Fec_Ini) ")
            loComandoSeleccionar.AppendLine("           And #tmpComprobantes02.Mes      =   MONTH(Comprobantes.Fec_Ini) ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes03.*, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Contables.Nom_Cue AS Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes04 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes03 LEFT JOIN Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes03.Cod_Cue   =   Cuentas_Contables.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes04.* ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes04 ")
            'loComandoSeleccionar.AppendLine(" ORDER BY  Fec_Ini, Documento, Cod_Cue ")
            loComandoSeleccionar.AppendLine("ORDER BY   Fec_Ini, Documento,  " & lcOrdenamiento)

            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCDiario_CContable", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCDiario_CContable.ReportSource = loObjetoReporte

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
' JJD: 23/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  17/08/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'