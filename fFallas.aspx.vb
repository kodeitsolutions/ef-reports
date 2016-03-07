'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFallas"
'-------------------------------------------------------------------------------------------'
Partial Class fFallas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            loComandoSeleccionar.AppendLine("SELECT	Fallas.cod_fal, ")
            loComandoSeleccionar.AppendLine("		(CASE Fallas.status  ")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO'  ")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO'  ")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO'  ")
            loComandoSeleccionar.AppendLine("		END) AS status,  ")
            loComandoSeleccionar.AppendLine("		Fallas.nom_fal, ")
            loComandoSeleccionar.AppendLine("		Fallas.modulo, ")
            loComandoSeleccionar.AppendLine("		Fallas.cod_dep, ")
            loComandoSeleccionar.AppendLine("		Departamentos.nom_dep, ")
            loComandoSeleccionar.AppendLine("		Fallas.cod_sec, ")
            loComandoSeleccionar.AppendLine("		secciones.nom_sec, ")
            loComandoSeleccionar.AppendLine("		Fallas.cod_mar, ")
            loComandoSeleccionar.AppendLine("		Marcas.nom_mar, ")
            loComandoSeleccionar.AppendLine("		Fallas.cod_tip, ")
            loComandoSeleccionar.AppendLine("		Tipos_Articulos.nom_tip, ")
            loComandoSeleccionar.AppendLine("		Fallas.cod_cla, ")
            loComandoSeleccionar.AppendLine("		clases_articulos.nom_cla, ")
            loComandoSeleccionar.AppendLine("		Fallas.tiempo, ")
            loComandoSeleccionar.AppendLine("		Fallas.motivo, ")
            loComandoSeleccionar.AppendLine("		Fallas.solucion, ")
            loComandoSeleccionar.AppendLine("		Fallas.fec_ini, ")
            loComandoSeleccionar.AppendLine("		Fallas.fec_fin, ")
            loComandoSeleccionar.AppendLine("		Fallas.tipo, ")
            loComandoSeleccionar.AppendLine("		Fallas.clase, ")
            loComandoSeleccionar.AppendLine("		Fallas.grupo, ")
            loComandoSeleccionar.AppendLine("		Fallas.Valor1, ")
            loComandoSeleccionar.AppendLine("		Fallas.Valor2, ")
            loComandoSeleccionar.AppendLine("		Fallas.Valor3, ")
            loComandoSeleccionar.AppendLine("		Fallas.Valor4, ")
            loComandoSeleccionar.AppendLine("		Fallas.Valor5, ")
            loComandoSeleccionar.AppendLine("		Fallas.comentario, ")
            loComandoSeleccionar.AppendLine("		Fallas.Nivel, ")
            loComandoSeleccionar.AppendLine("		Fallas.Prioridad ")
            loComandoSeleccionar.AppendLine("FROM Fallas ")
            loComandoSeleccionar.AppendLine("	JOIN Departamentos ON Departamentos.cod_dep = Fallas.cod_dep ")
            loComandoSeleccionar.AppendLine("	JOIN Secciones ON Secciones.cod_sec = Fallas.cod_sec ")
            loComandoSeleccionar.AppendLine("		AND Secciones.cod_dep = Departamentos.cod_dep ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Marcas ON Marcas.cod_mar = Fallas.cod_mar ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN tipos_articulos ON tipos_articulos.cod_tip = fallas.cod_tip ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN clases_articulos ON clases_articulos.cod_cla = Fallas.cod_cla ")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFallas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfFallas.ReportSource = loObjetoReporte

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
