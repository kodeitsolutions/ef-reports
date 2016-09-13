Imports System.Data
Partial Class rComprobantes_Diario
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
            loComandoSeleccionar.AppendLine("           YEAR(Comprobantes.Fec_Ini)  AS  Anno, ")
            loComandoSeleccionar.AppendLine("           MONTH(Comprobantes.Fec_Ini) AS  Mes, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Resumen, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Tipo, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Origen, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Integracion, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Status, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Notas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Cen, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Gas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Act, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Tasa, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Reg ")
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

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes01.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Nom_Cue,1,30) END) AS Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes02 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes01 LEFT JOIN Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes01.Cod_Cue   =   Cuentas_Contables.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes02.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes02.Cod_Cen = '' THEN '' ELSE SUBSTRING(Centros_Costos.Nom_Cen,1,30) END) AS Nom_Cen ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes03 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes02 LEFT JOIN Centros_Costos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes02.Cod_Cen   =   Centros_Costos.Cod_Cen ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes03.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes03.Cod_Gas = '' THEN '' ELSE SUBSTRING(Cuentas_Gastos.Nom_Gas,1,30) END) AS Nom_Gas ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes04 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes03 LEFT JOIN Cuentas_Gastos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes03.Cod_Gas   =   Cuentas_Gastos.Cod_Gas ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes04.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes04.Cod_Act = '' THEN '' ELSE SUBSTRING(Activos_Fijos.Nom_Act,1,30) END) AS Nom_Act ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes05 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes04 LEFT JOIN Activos_Fijos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes04.Cod_Act   =   Activos_Fijos.Cod_Act ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes05.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes05.Cod_Tip = '' THEN '' ELSE SUBSTRING(Tipos_Documentos.Nom_Tip,1,30) END) AS Nom_Tip ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes06 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes05 LEFT JOIN Tipos_Documentos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes05.Cod_Tip   =   Tipos_Documentos.Cod_Tip ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes06.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes06.Cod_Cla = '' THEN '' ELSE SUBSTRING(Clasificadores.Nom_Cla,1,30) END) AS Nom_Cla ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes07 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes06 LEFT JOIN Clasificadores ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes06.Cod_Cla   =   Clasificadores.Cod_Cla ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes07.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes07.Cod_Mon = '' THEN '' ELSE SUBSTRING(Monedas.Nom_Mon,1,30) END) AS Nom_Mon ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes08 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes07 LEFT JOIN Monedas ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes07.Cod_Mon   =   Monedas.Cod_Mon ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes08.* ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes08 ")
            'loComandoSeleccionar.AppendLine(" ORDER BY  Documento, Renglon ")
            loComandoSeleccionar.AppendLine("ORDER BY    Documento, " & lcOrdenamiento)

         

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rComprobantes_Diario", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrComprobantes_Diario.ReportSource = loObjetoReporte

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
' MAT:  16/05/11: Mejora de la vista de Diseño
'-------------------------------------------------------------------------------------------'