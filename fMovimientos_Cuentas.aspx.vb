'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fMovimientos_Cuentas"
'-------------------------------------------------------------------------------------------'
Partial Class fMovimientos_Cuentas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tip_Doc, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Status, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tasa, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias.Num_Cue, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban, ")
            loComandoSeleccionar.AppendLine("           Conceptos.Nom_Con ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpMovimientosCuentas ")
            loComandoSeleccionar.AppendLine(" FROM      Movimientos_Cuentas LEFT JOIN Bancos ON  Movimientos_Cuentas.Cod_Ban = Bancos.Cod_Ban, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Bancarias, ")
            loComandoSeleccionar.AppendLine("           Conceptos ")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cuentas.Cod_Cue         =   Cuentas_Bancarias.Cod_Cue ")
            loComandoSeleccionar.AppendLine("           And Movimientos_Cuentas.Cod_Con     =   Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine("           And " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            loComandoSeleccionar.AppendLine(" SELECT    #tmpMovimientosCuentas.*, ")
            loComandoSeleccionar.AppendLine("           Bancos.Nom_Ban          AS  Banco ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpMovimientosCuentas LEFT JOIN Bancos ON #tmpMovimientosCuentas.Cod_Ban = Bancos.Cod_Ban")


            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fMovimientos_Cuentas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfMovimientos_Cuentas.ReportSource = loObjetoReporte

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
' JJD: 24/01/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' JJD: 23/01/10: Ajuste a la busqueda del banco
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT: 18/04/11: Mejora en la vista de Diseño
'-------------------------------------------------------------------------------------------'