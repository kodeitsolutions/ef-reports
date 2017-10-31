'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "KDE_fRequerimientos"
'-------------------------------------------------------------------------------------------'
Partial Class KDE_fRequerimientos
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT CAST(Solicitudes.Seguimientos as XML) AS Seguimiento INTO #xmlData FROM Solicitudes")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Solicitudes.Documento                       AS Documento,")
            loComandoSeleccionar.AppendLine("       Solicitudes.Status                          AS Estatus,")
            loComandoSeleccionar.AppendLine("       Solicitudes.Asunto                          AS Asunto,")
            loComandoSeleccionar.AppendLine("       Solicitudes.Requerimiento                   AS Requerimiento,")
            loComandoSeleccionar.AppendLine("       Solicitudes.Tip_Sol                         AS Tip_Sol,")
            loComandoSeleccionar.AppendLine("       Solicitudes.Fec_Ini                         AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("       Solicitudes.Fec_Fin                         AS Fec_Fin,")
            loComandoSeleccionar.AppendLine("       Solicitudes.Tipo                            AS Tipo,")
            loComandoSeleccionar.AppendLine("		Solicitudes.Cod_Reg                         AS Cod_Reg,")
            loComandoSeleccionar.AppendLine("       Clientes.Nom_Cli                            AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("       Solicitudes.Comentario                      AS Comentario,")
            loComandoSeleccionar.AppendLine("		ROW_NUMBER() OVER (ORDER BY D.C.value('@status', 'VARCHAR(15)') DESC) AS se_renglon,")
            loComandoSeleccionar.AppendLine(" 		D.C.value('@fecha', 'DATETIME')             AS se_fecha,")
            loComandoSeleccionar.AppendLine(" 		D.C.value('@contacto', 'VARCHAR(300)')      AS se_contacto,")
            loComandoSeleccionar.AppendLine(" 		D.C.value('@accion', 'VARCHAR(300)')        AS se_accion,")
            loComandoSeleccionar.AppendLine(" 		D.C.value('@comentario', 'VARCHAR(300)')    AS se_comentario,")
            loComandoSeleccionar.AppendLine(" 		D.C.value('@medio', 'VARCHAR(300)')         AS se_medio")
            loComandoSeleccionar.AppendLine("FROM Solicitudes, Clientes,#xmlData")
            loComandoSeleccionar.AppendLine("	CROSS APPLY Seguimiento.nodes('elementos/elemento') D(c)")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("   AND Solicitudes.Cod_Reg = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine(" DROP TABLE #xmlData")
            loComandoSeleccionar.AppendLine("")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
					 
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("KDE_fRequerimientos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvKDE_fRequerimientos.ReportSource = loObjetoReporte

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

        End Try

    End Sub
End Class

'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' MAT: 11/08/11: Programacion inicial
'-------------------------------------------------------------------------------------------'