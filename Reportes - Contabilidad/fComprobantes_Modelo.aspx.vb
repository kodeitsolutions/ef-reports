Imports System.Data
Partial Class fComprobantes_Modelo

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Modelos.Documento, ")
            loComandoSeleccionar.AppendLine("           YEAR(Modelos.Fec_Ini)  AS  Anno, ")
            loComandoSeleccionar.AppendLine("           MONTH(Modelos.Fec_Ini) AS  Mes, ")
            loComandoSeleccionar.AppendLine("           Modelos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Modelos.Fec_Fin, ")
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
            loComandoSeleccionar.AppendLine(" WHERE     Modelos.Documento  =  Renglones_Modelos.Documento ")
            loComandoSeleccionar.AppendLine("           AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos01.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Nom_Cue,1,35) END) AS Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos02 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos01 LEFT JOIN Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos01.Cod_Cue   =   Cuentas_Contables.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos02.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos02.Cod_Cen = '' THEN '' ELSE Centros_Costos.Nom_Cen END) AS Nom_Cen ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos03 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos02 LEFT JOIN Centros_Costos ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos02.Cod_Cen   =   Centros_Costos.Cod_Cen ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos03.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos03.Cod_Gas = '' THEN '' ELSE Cuentas_Gastos.Nom_Gas END) AS Nom_Gas ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos04 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos03 LEFT JOIN Cuentas_Gastos ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos03.Cod_Gas   =   Cuentas_Gastos.Cod_Gas ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos04.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos04.Cod_Act = '' THEN '' ELSE Activos_Fijos.Nom_Act END) AS Nom_Act ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos05 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos04 LEFT JOIN Activos_Fijos ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos04.Cod_Act   =   Activos_Fijos.Cod_Act ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos05.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos05.Cod_Tip = '' THEN '' ELSE Tipos_Documentos.Nom_Tip END) AS Nom_Tip ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos06 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos05 LEFT JOIN Tipos_Documentos ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos05.Cod_Tip   =   Tipos_Documentos.Cod_Tip ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos06.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos06.Cod_Cla = '' THEN '' ELSE Clasificadores.Nom_Cla END) AS Nom_Cla ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos07 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos06 LEFT JOIN Clasificadores ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos06.Cod_Cla   =   Clasificadores.Cod_Cla ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos07.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpModelos07.Cod_Mon = '' THEN '' ELSE Monedas.Nom_Mon END) AS Nom_Mon ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpModelos08 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos07 LEFT JOIN Monedas ")
            loComandoSeleccionar.AppendLine("           ON #tmpModelos07.Cod_Mon   =   Monedas.Cod_Mon ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpModelos08.* ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpModelos08 ")
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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fModelos_Modelo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfComprobantes_Modelo.ReportSource = loObjetoReporte

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


