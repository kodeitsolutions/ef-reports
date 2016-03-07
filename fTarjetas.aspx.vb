'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fTarjetas"
'-------------------------------------------------------------------------------------------'
Partial Class fTarjetas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()
            loComandoSeleccionar.AppendLine("SELECT	Tarjetas.Cod_Tar,")
            loComandoSeleccionar.AppendLine("		(CASE Tarjetas.status")
            loComandoSeleccionar.AppendLine("			WHEN 'A' THEN 'ACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'I' THEN 'INACTIVO'")
            loComandoSeleccionar.AppendLine("			WHEN 'S' THEN 'SUSPENDIDO'")
            loComandoSeleccionar.AppendLine("		END)		AS status,")
            loComandoSeleccionar.AppendLine("		Tarjetas.nom_tar,")
            loComandoSeleccionar.AppendLine("		Tarjetas.telefonos,")
            loComandoSeleccionar.AppendLine("		Tarjetas.cod_tip AS tip_cod,")
            loComandoSeleccionar.AppendLine("		Tipos_Tarjetas.nom_tip AS tip_nom,")
            loComandoSeleccionar.AppendLine("		(CASE Tarjetas.tip_tar")
            loComandoSeleccionar.AppendLine("			WHEN 'C' THEN 'CREDITO'")
            loComandoSeleccionar.AppendLine("			WHEN 'D' THEN 'DEBITO'")
            loComandoSeleccionar.AppendLine("		END)			AS tip_tar,")
            loComandoSeleccionar.AppendLine("		Tarjetas.cod_ban,")
            loComandoSeleccionar.AppendLine("		Bancos.nom_ban,")
            loComandoSeleccionar.AppendLine("		Tarjetas.cod_cue,")
            loComandoSeleccionar.AppendLine("		Cuentas_Bancarias.nom_cue,")
            loComandoSeleccionar.AppendLine("		Tarjetas.Clave,")
            loComandoSeleccionar.AppendLine("		Tarjetas.dep_dir,")
            loComandoSeleccionar.AppendLine("		Tarjetas.por_rec,")
            loComandoSeleccionar.AppendLine("		Tarjetas.mon_rec,")
            loComandoSeleccionar.AppendLine("		Tarjetas.por_com,")
            loComandoSeleccionar.AppendLine("		Tarjetas.mon_com,")
            loComandoSeleccionar.AppendLine("		Tarjetas.por_imp,")
            loComandoSeleccionar.AppendLine("		Tarjetas.por_ret,")
            loComandoSeleccionar.AppendLine("		Tarjetas.comentario")
            loComandoSeleccionar.AppendLine("FROM Tarjetas")
            loComandoSeleccionar.AppendLine("	JOIN Tipos_Tarjetas ON Tipos_Tarjetas.cod_tip = Tarjetas.cod_tip")
            loComandoSeleccionar.AppendLine("	JOIN bancos ON Bancos.cod_ban = Tarjetas.cod_ban")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Bancarias ON Cuentas_Bancarias.cod_cue = Tarjetas.cod_cue ")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fTarjetas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfTarjetas.ReportSource = loObjetoReporte

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
