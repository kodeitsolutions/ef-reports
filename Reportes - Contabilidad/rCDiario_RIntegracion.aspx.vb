Imports System.Data
Partial Class rCDiario_RIntegracion
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

            loComandoSeleccionar.AppendLine(" SELECT	Renglones_Comprobantes.Documento, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Reg, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes01 ")
            loComandoSeleccionar.AppendLine(" FROM      Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine(" WHERE     Renglones_Comprobantes.Documento    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_Comprobantes.Fec_Ini  Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_Comprobantes.Cod_Mon  Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)

            loComandoSeleccionar.AppendLine(" SELECT    Documento, ")
            loComandoSeleccionar.AppendLine("           Cod_Reg, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Deb) AS Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Hab) AS Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Adicional ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes02 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes01 ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Adicional, ")
            loComandoSeleccionar.AppendLine("           Documento, ")
            loComandoSeleccionar.AppendLine("           Cod_Reg ")

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
            loComandoSeleccionar.AppendLine(" WHERE     #tmpComprobantes02.Documento        =   Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("           And #tmpComprobantes02.Adicional    =   Comprobantes.Adicional ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes03.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes03.Cod_Reg = '' THEN 'REGLA NO ASIGNADA' ELSE SUBSTRING(Reglas_Integracion.Nom_Reg,1,50) END) AS Nom_Reg ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes04 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes03 LEFT JOIN Reglas_Integracion ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes03.Cod_Reg   =   Reglas_Integracion.Cod_Reg ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes04.* ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes04 ")
            'loComandoSeleccionar.AppendLine(" ORDER BY  Adicional, Documento, Cod_Reg ")
            loComandoSeleccionar.AppendLine("ORDER BY    Adicional, Documento, " & lcOrdenamiento)

            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCDiario_RIntegracion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCDiario_RIntegracion.ReportSource = loObjetoReporte

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
' CMS:  17/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'