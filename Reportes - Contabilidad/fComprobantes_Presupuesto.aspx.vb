Imports System.Data
Partial Class fComprobantes_Presupuesto

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Presupuesto.Documento, ")
            loComandoSeleccionar.AppendLine("           YEAR(Presupuesto.Fec_Ini)  AS  Anno, ")
            loComandoSeleccionar.AppendLine("           MONTH(Presupuesto.Fec_Ini) AS  Mes, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Resumen, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Tipo, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Origen, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Integracion, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Status, ")
            loComandoSeleccionar.AppendLine("           Presupuesto.Notas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Cen, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Gas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Act, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Tasa, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto.Cod_Reg ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpPresupuesto01 ")
            loComandoSeleccionar.AppendLine(" FROM      Presupuesto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Presupuesto ")
            loComandoSeleccionar.AppendLine(" WHERE     Presupuesto.Documento  =  Renglones_Presupuesto.Documento ")
            loComandoSeleccionar.AppendLine("           AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loComandoSeleccionar.AppendLine(" SELECT	#tmpPresupuesto01.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpPresupuesto01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Nom_Cue,1,35) END) AS Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpPresupuesto02 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpPresupuesto01 LEFT JOIN Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("           ON #tmpPresupuesto01.Cod_Cue   =   Cuentas_Contables.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpPresupuesto02.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpPresupuesto02.Cod_Cen = '' THEN '' ELSE Centros_Costos.Nom_Cen END) AS Nom_Cen ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpPresupuesto03 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpPresupuesto02 LEFT JOIN Centros_Costos ")
            loComandoSeleccionar.AppendLine("           ON #tmpPresupuesto02.Cod_Cen   =   Centros_Costos.Cod_Cen ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpPresupuesto03.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpPresupuesto03.Cod_Gas = '' THEN '' ELSE Cuentas_Gastos.Nom_Gas END) AS Nom_Gas ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpPresupuesto04 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpPresupuesto03 LEFT JOIN Cuentas_Gastos ")
            loComandoSeleccionar.AppendLine("           ON #tmpPresupuesto03.Cod_Gas   =   Cuentas_Gastos.Cod_Gas ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpPresupuesto04.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpPresupuesto04.Cod_Act = '' THEN '' ELSE Activos_Fijos.Nom_Act END) AS Nom_Act ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpPresupuesto05 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpPresupuesto04 LEFT JOIN Activos_Fijos ")
            loComandoSeleccionar.AppendLine("           ON #tmpPresupuesto04.Cod_Act   =   Activos_Fijos.Cod_Act ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpPresupuesto05.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpPresupuesto05.Cod_Tip = '' THEN '' ELSE Tipos_Documentos.Nom_Tip END) AS Nom_Tip ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpPresupuesto06 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpPresupuesto05 LEFT JOIN Tipos_Documentos ")
            loComandoSeleccionar.AppendLine("           ON #tmpPresupuesto05.Cod_Tip   =   Tipos_Documentos.Cod_Tip ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpPresupuesto06.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpPresupuesto06.Cod_Cla = '' THEN '' ELSE Clasificadores.Nom_Cla END) AS Nom_Cla ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpPresupuesto07 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpPresupuesto06 LEFT JOIN Clasificadores ")
            loComandoSeleccionar.AppendLine("           ON #tmpPresupuesto06.Cod_Cla   =   Clasificadores.Cod_Cla ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpPresupuesto07.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpPresupuesto07.Cod_Mon = '' THEN '' ELSE Monedas.Nom_Mon END) AS Nom_Mon ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpPresupuesto08 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpPresupuesto07 LEFT JOIN Monedas ")
            loComandoSeleccionar.AppendLine("           ON #tmpPresupuesto07.Cod_Mon   =   Monedas.Cod_Mon ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpPresupuesto08.* ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpPresupuesto08 ")
            loComandoSeleccionar.AppendLine(" ORDER BY  Documento, Renglon ")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			
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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPresupuesto_Presupuesto", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfComprobantes_Presupuesto.ReportSource = loObjetoReporte

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
' MAT: 29/04/11: Codigo inicial
'-------------------------------------------------------------------------------------------'


