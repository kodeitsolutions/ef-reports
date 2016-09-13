Imports System.Data
Partial Class rComprobantes_Modelo
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

            loComandoSeleccionar.AppendLine(" SELECT	Modelos.Documento, ")
            loComandoSeleccionar.AppendLine("           YEAR(Modelos.Fec_Ini)  AS  Anno, ")
            loComandoSeleccionar.AppendLine("           MONTH(Modelos.Fec_Ini) AS  Mes, ")
            loComandoSeleccionar.AppendLine("           Modelos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Modelos.Resumen, ")
            loComandoSeleccionar.AppendLine("           Modelos.Tipo, ")
            loComandoSeleccionar.AppendLine("           Modelos.Origen, ")
            loComandoSeleccionar.AppendLine("           Modelos.Integracion, ")
            loComandoSeleccionar.AppendLine("           Modelos.Status, ")
            loComandoSeleccionar.AppendLine("           Modelos.Notas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Cod_Cen, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Cod_Gas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Cod_Act, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Tasa, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos.Cod_Reg ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos01 ")
            loComandoSeleccionar.AppendLine(" FROM      Modelos, ")
            loComandoSeleccionar.AppendLine("           Renglones_Modelos ")
            loComandoSeleccionar.AppendLine(" WHERE     Modelos.Documento                      =   Renglones_Modelos.Documento ")
            loComandoSeleccionar.AppendLine("           And Modelos.Documento                  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Modelos.Fec_Ini                    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_Modelos.Cod_Mon          Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And YEAR(Renglones_Modelos.Fec_Ini)    =   YEAR(Modelos.Fec_Ini) ")
            loComandoSeleccionar.AppendLine("           And MONTH(Renglones_Modelos.Fec_Ini)   =   MONTH(Modelos.Fec_Ini) ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos01.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Nom_Cue,1,30) END) AS Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos02 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos01 LEFT JOIN Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos01.Cod_Cue   =   Cuentas_Contables.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos02.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos02.Cod_Cen = '' THEN '' ELSE SUBSTRING(Centros_Costos.Nom_Cen,1,30) END) AS Nom_Cen ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos03 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos02 LEFT JOIN Centros_Costos ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos02.Cod_Cen   =   Centros_Costos.Cod_Cen ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos03.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos03.Cod_Gas = '' THEN '' ELSE SUBSTRING(Cuentas_Gastos.Nom_Gas,1,30) END) AS Nom_Gas ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos04 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos03 LEFT JOIN Cuentas_Gastos ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos03.Cod_Gas   =   Cuentas_Gastos.Cod_Gas ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos04.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos04.Cod_Act = '' THEN '' ELSE SUBSTRING(Activos_Fijos.Nom_Act,1,30) END) AS Nom_Act ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos05 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos04 LEFT JOIN Activos_Fijos ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos04.Cod_Act   =   Activos_Fijos.Cod_Act ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos05.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos05.Cod_Tip = '' THEN '' ELSE SUBSTRING(Tipos_Documentos.Nom_Tip,1,30) END) AS Nom_Tip ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos06 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos05 LEFT JOIN Tipos_Documentos ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos05.Cod_Tip   =   Tipos_Documentos.Cod_Tip ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos06.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos06.Cod_Cla = '' THEN '' ELSE SUBSTRING(Clasificadores.Nom_Cla,1,30) END) AS Nom_Cla ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos07 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos06 LEFT JOIN Clasificadores ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos06.Cod_Cla   =   Clasificadores.Cod_Cla ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos07.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos07.Cod_Mon = '' THEN '' ELSE SUBSTRING(Monedas.Nom_Mon,1,30) END) AS Nom_Mon ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos08 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos07 LEFT JOIN Monedas ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos07.Cod_Mon   =   Monedas.Cod_Mon ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos08.* ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos08 ")
            loComandoSeleccionar.AppendLine("ORDER BY    Documento, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rComprobantes_Modelo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrComprobantes_Modelo.ReportSource = loObjetoReporte

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
