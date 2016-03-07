'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fVariables"
'-------------------------------------------------------------------------------------------'
Partial Class fVariables
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Variables.cod_var,")
            loComandoSeleccionar.AppendLine("		(CASE Variables.status ")
            loComandoSeleccionar.AppendLine("                WHEN 'A' THEN 'ACTIVO'")
            loComandoSeleccionar.AppendLine("                WHEN 'I' THEN 'INACTIVO'")
            loComandoSeleccionar.AppendLine("                WHEN 'S' THEN 'SUSPENDIDO'")
            loComandoSeleccionar.AppendLine("       END) AS status,")
            loComandoSeleccionar.AppendLine("		Variables.nom_var,")
            loComandoSeleccionar.AppendLine("		Variables.comentario,")
            loComandoSeleccionar.AppendLine("		Variables.tip_Var,")
            loComandoSeleccionar.AppendLine("		Variables.val_max_esp,")
            loComandoSeleccionar.AppendLine("		Variables.val_min_esp,")
            loComandoSeleccionar.AppendLine("		Variables.Prioridad,")
            loComandoSeleccionar.AppendLine("		Variables.Tip_Com,")
            loComandoSeleccionar.AppendLine("		Factory_Global.dbo.Combos.nom_com,")
            loComandoSeleccionar.AppendLine("		Variables.opcional,")
            loComandoSeleccionar.AppendLine("		Variables.Orden,")
            loComandoSeleccionar.AppendLine("		Variables.Nivel,")
            loComandoSeleccionar.AppendLine("		Variables.Peso,")
            loComandoSeleccionar.AppendLine("		Unidades.Nom_Uni,")
            loComandoSeleccionar.AppendLine("		Variables.Cod_Uni")
            loComandoSeleccionar.AppendLine("FROM Variables")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Unidades ON Unidades.cod_uni = Variables.cod_uni")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Factory_Global.dbo.Combos ON Factory_Global.dbo.Combos.cod_com collate database_default = Variables.cod_com collate database_default")
            loComandoSeleccionar.AppendLine("WHERE    " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fVariables", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfVariables.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' EAG: 28/08/2015 : Codigo inicial
'-------------------------------------------------------------------------------------------'
