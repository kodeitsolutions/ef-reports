'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_fMovimientos_Cuentas"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_fMovimientos_Cuentas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("          Tipos_Movimientos.Nom_Tip, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Comentario, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("          Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("          Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("          Bancos.Nom_Ban       AS Banco, ")
            loComandoSeleccionar.AppendLine("          Conceptos.Nom_Con ")
            loComandoSeleccionar.AppendLine("FROM Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("   JOIN Bancos ON  Movimientos_Cuentas.Cod_Ban = Bancos.Cod_Ban")
            loComandoSeleccionar.AppendLine("   JOIN Cuentas_Bancarias ON  Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue")
            loComandoSeleccionar.AppendLine("   JOIN Conceptos ON Movimientos_Cuentas.Cod_Con = Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine("   JOIN Tipos_Movimientos ON Movimientos_Cuentas.Cod_Tip = Tipos_Movimientos.Cod_Tip")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fMovimientos_Cuentas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fMovimientos_Cuentas.ReportSource = loObjetoReporte

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