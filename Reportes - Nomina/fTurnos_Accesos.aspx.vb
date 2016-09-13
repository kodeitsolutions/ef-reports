'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fTurnos_Accesos"
'-------------------------------------------------------------------------------------------'
Partial Class fTurnos_Accesos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("    SELECT Cod_Tur, ")
            loComandoSeleccionar.AppendLine("    		Nom_Tur, ")
            loComandoSeleccionar.AppendLine("    		Tip_Tur, ")
            loComandoSeleccionar.AppendLine("    		CASE WHEN Status = 'A' THEN 'Activo' ")
            loComandoSeleccionar.AppendLine("    			 WHEN Status = 'I' THEN 'Inactivo' ELSE 'Otro' END AS Status, ")
            loComandoSeleccionar.AppendLine("    		Hor_Ini, ")
            loComandoSeleccionar.AppendLine("    		Hor_Fin, ")
            loComandoSeleccionar.AppendLine("    		Com_Ini, ")
            loComandoSeleccionar.AppendLine("    		Com_Fin, ")
            loComandoSeleccionar.AppendLine("    		Min_Ret, ")
            loComandoSeleccionar.AppendLine("    		Lun_Ini1, ")
            loComandoSeleccionar.AppendLine("    		Lun_Fin1, ")
            loComandoSeleccionar.AppendLine("    		Lun_Ini2, ")
            loComandoSeleccionar.AppendLine("    		Lun_Fin2, ")
            loComandoSeleccionar.AppendLine("    		Mar_Ini1, ")
            loComandoSeleccionar.AppendLine("    		Mar_Fin1, ")
            loComandoSeleccionar.AppendLine("    		Mar_Ini2, ")
            loComandoSeleccionar.AppendLine("    		Mar_Fin2, ")
            loComandoSeleccionar.AppendLine("    		Mie_Ini1, ")
            loComandoSeleccionar.AppendLine("    		Mie_Fin1, ")
            loComandoSeleccionar.AppendLine("    		Mie_Ini2, ")
            loComandoSeleccionar.AppendLine("    		Mie_Fin2, ")
            loComandoSeleccionar.AppendLine("    		Jue_Ini1, ")
            loComandoSeleccionar.AppendLine("    		Jue_Fin1, ")
            loComandoSeleccionar.AppendLine("    		Jue_Ini2, ")
            loComandoSeleccionar.AppendLine("    		Jue_Fin2, ")
            loComandoSeleccionar.AppendLine("    		Vie_Ini1, ")
            loComandoSeleccionar.AppendLine("    		Vie_Fin1, ")
            loComandoSeleccionar.AppendLine("    		Vie_Ini2, ")
            loComandoSeleccionar.AppendLine("    		Vie_Fin2, ")
            loComandoSeleccionar.AppendLine("    		Sab_Ini1, ")
            loComandoSeleccionar.AppendLine("    		Sab_Fin1, ")
            loComandoSeleccionar.AppendLine("    		Sab_Ini2, ")
            loComandoSeleccionar.AppendLine("    		Sab_Fin2, ")
            loComandoSeleccionar.AppendLine("    		Dom_Ini1, ")
            loComandoSeleccionar.AppendLine("    		Dom_Fin1, ")
            loComandoSeleccionar.AppendLine("    		Dom_Ini2, ")
            loComandoSeleccionar.AppendLine("    		Dom_Fin2, ")
            loComandoSeleccionar.AppendLine("    		Fer_Ini1, ")
            loComandoSeleccionar.AppendLine("    		Fer_Fin1, ")
            loComandoSeleccionar.AppendLine("    		Fer_Ini2, ")
            loComandoSeleccionar.AppendLine("    		Fer_Fin2, ")
            loComandoSeleccionar.AppendLine("    		Comentario ")
            loComandoSeleccionar.AppendLine(" FROM      Turnos ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fTurnos_Accesos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfTurnos_Accesos.ReportSource = loObjetoReporte

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
' JJD: 09/01/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 06/02/10: Se le incluyo el Comentario
'-------------------------------------------------------------------------------------------'
