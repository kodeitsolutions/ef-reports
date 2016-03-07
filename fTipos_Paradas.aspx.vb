'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fTipos_Paradas"
'-------------------------------------------------------------------------------------------'
Partial Class fTipos_Paradas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            loComandoSeleccionar.AppendLine("SELECT	Tipos_Paradas.cod_tip, ")
            loComandoSeleccionar.AppendLine("		(CASE Tipos_Paradas.status  ")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO'  ")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO'  ")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO'  ")
            loComandoSeleccionar.AppendLine("		END) AS status,  ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.nom_tip, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.cod_dep, ")
            loComandoSeleccionar.AppendLine("		Departamentos.nom_dep, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.tiempo, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.motivo, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.solucion, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.fec_ini, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.fec_fin, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.tipo, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.clase, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.grupo, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.Valor1, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.Valor2, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.Valor3, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.Valor4, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.Valor5, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.comentario, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.Nivel, ")
            loComandoSeleccionar.AppendLine("		Tipos_Paradas.Prioridad ")
            loComandoSeleccionar.AppendLine("FROM Tipos_Paradas ")
            loComandoSeleccionar.AppendLine("	JOIN Departamentos ON Departamentos.cod_dep = Tipos_Paradas.cod_dep ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fTipos_Paradas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfTipos_Paradas.ReportSource = loObjetoReporte

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
' EAG: 03/09/2015 : Codigo inicial
'-------------------------------------------------------------------------------------------'
