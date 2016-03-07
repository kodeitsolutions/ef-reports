'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fTipos_Containers"
'-------------------------------------------------------------------------------------------'
Partial Class fTipos_Containers
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Cod_Tip, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Nom_Tip, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Tip_Tra, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Material, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Tipo, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Status, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Ancho, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Alto, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Largo, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Tara, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Car_Max, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Pes_Bru, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Uni_Anc, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Uni_Alt, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Uni_Lar, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Uni_Tar, ")
            loComandoSeleccionar.AppendLine(" 	        Tipos_Containers.Uni_Car, ")
            loComandoSeleccionar.AppendLine(" 			Tipos_Containers.Uni_Pes, ")
            loComandoSeleccionar.AppendLine(" 	        Tipos_Containers.Comentario")
            loComandoSeleccionar.AppendLine(" FROM Tipos_Containers")
            loComandoSeleccionar.AppendLine(" WHERE")
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim cadena As String = loComandoSeleccionar.ToString()

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fTipos_Containers", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfTipos_Containers.ReportSource = loObjetoReporte

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
' RAC:28/03/11 : Codigo inicial
'-------------------------------------------------------------------------------------------'
