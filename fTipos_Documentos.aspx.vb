'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fTipos_Documentos"
'-------------------------------------------------------------------------------------------'
Partial Class fTipos_Documentos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            loComandoSeleccionar.AppendLine("SELECT	Tipos_Documentos.Cod_Tip,")
            loComandoSeleccionar.AppendLine("		(CASE Tipos_Documentos.status ")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO'")
            loComandoSeleccionar.AppendLine("		END) AS status,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.nom_tip,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.clase,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.tipo,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.grupo,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.Comentario,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.Usa_Pan,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.Usa_ret,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.Usa_des,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.Doc_Leg,")
            loComandoSeleccionar.AppendLine("		Tipos_Documentos.Cod_Ven,")
            loComandoSeleccionar.AppendLine("		Ventas.nom_con AS nom_ven,")
            loComandoSeleccionar.AppendLine("		tipos_documentos.Cod_com,")
            loComandoSeleccionar.AppendLine("		Compras.nom_con AS nom_com")
            loComandoSeleccionar.AppendLine("FROM Tipos_Documentos")
            loComandoSeleccionar.AppendLine("	JOIN Contadores AS Ventas ON Ventas.cod_con = Tipos_Documentos.Cod_Ven")
            loComandoSeleccionar.AppendLine("		AND Ventas.Cod_Suc = 'Factory'")
            loComandoSeleccionar.AppendLine("	JOIN Contadores AS Compras On Compras.cod_con = Tipos_Documentos.cod_com")
            loComandoSeleccionar.AppendLine("		AND Compras.Cod_Suc = 'Factory'")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fTipos_Documentos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfTipos_Documentos.ReportSource = loObjetoReporte

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
' EAG: 31/08/2015 : Codigo inicial
'-------------------------------------------------------------------------------------------'
