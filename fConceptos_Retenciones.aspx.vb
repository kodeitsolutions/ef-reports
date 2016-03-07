﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fConceptos_Retenciones"
'-------------------------------------------------------------------------------------------'
Partial Class fConceptos_Retenciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	retenciones.cod_ret, ")
            loComandoSeleccionar.AppendLine("		(CASE retenciones.status ")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO' ")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO' ")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO' ")
            loComandoSeleccionar.AppendLine("		END) AS status, ")
            loComandoSeleccionar.AppendLine("		retenciones.nom_ret, ")
            loComandoSeleccionar.AppendLine("		retenciones.cod_per, ")
            loComandoSeleccionar.AppendLine("		Personas.nom_per, ")
            loComandoSeleccionar.AppendLine("		retenciones.por_bas, ")
            loComandoSeleccionar.AppendLine("		retenciones.mon_des, ")
            loComandoSeleccionar.AppendLine("		retenciones.por_ret, ")
            loComandoSeleccionar.AppendLine("		retenciones.mon_has, ")
            loComandoSeleccionar.AppendLine("		retenciones.mon_sus, ")
            loComandoSeleccionar.AppendLine("		retenciones.comentario ")
            loComandoSeleccionar.AppendLine("FROM retenciones ")
            loComandoSeleccionar.AppendLine("	JOIN Personas On Personas.cod_per = retenciones.cod_per ")
            loComandoSeleccionar.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fConceptos_Retenciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfConceptos_Retenciones.ReportSource = loObjetoReporte

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
' EAG: 02/09/2015 : Codigo inicial
'-------------------------------------------------------------------------------------------'
