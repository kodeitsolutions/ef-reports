Imports System.Data
Partial Class fComprobantes_Diario

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Comprobantes.Documento, ")
            loComandoSeleccionar.AppendLine("           YEAR(Comprobantes.Fec_Ini)  AS  Anno, ")
            loComandoSeleccionar.AppendLine("           MONTH(Comprobantes.Fec_Ini) AS  Mes, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Fec_Fin, ")
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
            loComandoSeleccionar.AppendLine(" WHERE     Comprobantes.Documento  =  Renglones_Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("           AND Comprobantes.Adicional  =  Renglones_Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("           AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes01.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Nom_Cue,1,35) END) AS Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes02 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes01 LEFT JOIN Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes01.Cod_Cue   =   Cuentas_Contables.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes02.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes02.Cod_Cen = '' THEN '' ELSE Centros_Costos.Nom_Cen END) AS Nom_Cen ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes03 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes02 LEFT JOIN Centros_Costos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes02.Cod_Cen   =   Centros_Costos.Cod_Cen ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes03.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes03.Cod_Gas = '' THEN '' ELSE Cuentas_Gastos.Nom_Gas END) AS Nom_Gas ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes04 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes03 LEFT JOIN Cuentas_Gastos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes03.Cod_Gas   =   Cuentas_Gastos.Cod_Gas ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes04.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes04.Cod_Act = '' THEN '' ELSE Activos_Fijos.Nom_Act END) AS Nom_Act ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes05 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes04 LEFT JOIN Activos_Fijos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes04.Cod_Act   =   Activos_Fijos.Cod_Act ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes05.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes05.Cod_Tip = '' THEN '' ELSE Tipos_Documentos.Nom_Tip END) AS Nom_Tip ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes06 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes05 LEFT JOIN Tipos_Documentos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes05.Cod_Tip   =   Tipos_Documentos.Cod_Tip ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes06.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes06.Cod_Cla = '' THEN '' ELSE Clasificadores.Nom_Cla END) AS Nom_Cla ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes07 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes06 LEFT JOIN Clasificadores ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes06.Cod_Cla   =   Clasificadores.Cod_Cla ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes07.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes07.Cod_Mon = '' THEN '' ELSE Monedas.Nom_Mon END) AS Nom_Mon ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes08 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes07 LEFT JOIN Monedas ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes07.Cod_Mon   =   Monedas.Cod_Mon ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes08.* ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes08 ")
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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fComprobantes_Diario", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfComprobantes_Diario.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' JJD: 24/02/09: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' MAT: 03/03/11: Mejora en la vista de diseño												'
'-------------------------------------------------------------------------------------------'
' MAT: 04/03/11: Se aplicaron los metodos carga de imagen y validacion de registro cero.	'
'-------------------------------------------------------------------------------------------'
' RJG: 19/01/12: Se agregó el campo Adicional a la unión entre el encabezado y los renglones'
'-------------------------------------------------------------------------------------------'
