'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSecciones"
'-------------------------------------------------------------------------------------------'
Partial Class fSecciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT Secciones.cod_sec,")
            loComandoSeleccionar.AppendLine("		(CASE Secciones.status ")
            loComandoSeleccionar.AppendLine("                WHEN 'A' THEN 'ACTIVO'")
            loComandoSeleccionar.AppendLine("                WHEN 'I' THEN 'INACTIVO'")
            loComandoSeleccionar.AppendLine("                WHEN 'S' THEN 'SUSPENDIDO'")
            loComandoSeleccionar.AppendLine("       END) AS status,")
            loComandoSeleccionar.AppendLine("		Secciones.nom_sec,")
            loComandoSeleccionar.AppendLine("		Secciones.ABC,")
            loComandoSeleccionar.AppendLine("		Secciones.prioridad,")
            loComandoSeleccionar.AppendLine("		Secciones.Cod_Ret,")
            loComandoSeleccionar.AppendLine("		COALESCE(Retenciones.Nom_ret,'') AS Nom_Ret,")
            loComandoSeleccionar.AppendLine("		Secciones.Por_Ven,")
            loComandoSeleccionar.AppendLine("		Secciones.Por_Cob,")
            loComandoSeleccionar.AppendLine("		Secciones.Por_Des,")
            loComandoSeleccionar.AppendLine("		Secciones.Mon_des,")
            loComandoSeleccionar.AppendLine("		Secciones.comentario,")
            loComandoSeleccionar.AppendLine("		Secciones.portal,")
            loComandoSeleccionar.AppendLine("		Departamentos.Nom_Dep")
            loComandoSeleccionar.AppendLine("FROM Secciones")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones On Retenciones.cod_ret = Secciones.cod_ret")
            loComandoSeleccionar.AppendLine("	JOIN		Departamentos ON Departamentos.Cod_Dep = secciones.cod_dep")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSecciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfSecciones.ReportSource = loObjetoReporte

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
