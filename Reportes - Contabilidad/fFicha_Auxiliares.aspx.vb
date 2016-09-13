'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFicha_Auxiliares"
'-------------------------------------------------------------------------------------------'
Partial Class fFicha_Auxiliares
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            Dim lcTipo As String = goServicios.mObtenerCampoFormatoSQL(goEmpresa.pcCodigo & "Auxiliares")

            loComandoSeleccionar.AppendLine("SELECT	Auxiliares.cod_aux,")
            loComandoSeleccionar.AppendLine("		Auxiliares.nom_aux,")
            loComandoSeleccionar.AppendLine("		Auxiliares.tipo,")
            loComandoSeleccionar.AppendLine("		(CASE Auxiliares.status")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO'")
            loComandoSeleccionar.AppendLine("		END) AS status,")
            loComandoSeleccionar.AppendLine("		Auxiliares.grupo,")
            loComandoSeleccionar.AppendLine("		Auxiliares.comentario,")
            loComandoSeleccionar.AppendLine("		Auxiliares.Atributo_A,")
            loComandoSeleccionar.AppendLine("		Auxiliares.Atributo_B,")
            loComandoSeleccionar.AppendLine("		Auxiliares.Atributo_C,")
            loComandoSeleccionar.AppendLine("		Auxiliares.Atributo_D,")
            loComandoSeleccionar.AppendLine("		Auxiliares.Atributo_E,")
            loComandoSeleccionar.AppendLine("       Auxiliares.Atributo_F")
            loComandoSeleccionar.AppendLine("FROM Auxiliares")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)


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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFicha_Auxiliares", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFicha_Auxiliares.ReportSource = loObjetoReporte

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
' EAG: 20/08/2015 : Codigo Inicial
'-------------------------------------------------------------------------------------------'