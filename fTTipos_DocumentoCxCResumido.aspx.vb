'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fTTipos_DocumentoCxCResumido"
'-------------------------------------------------------------------------------------------'
Partial Class fTTipos_DocumentoCxCResumido
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("  SELECT		Cuentas_Cobrar.cod_tip, ")
            loComandoSeleccionar.AppendLine(" 				Tipos_Documentos.nom_tip, ")
            loComandoSeleccionar.AppendLine(" 				Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 				cast(count (Cuentas_Cobrar.cod_tip) as decimal) as Cant_Doc, ")
            loComandoSeleccionar.AppendLine(" 				SUM(Cuentas_Cobrar.mon_bas1) AS mon_bas1, ")
            loComandoSeleccionar.AppendLine(" 				SUM(Cuentas_Cobrar.mon_imp1) AS mon_imp1, ")
            loComandoSeleccionar.AppendLine(" 				SUM(Cuentas_Cobrar.mon_net) AS mon_net, ")
            loComandoSeleccionar.AppendLine("               SUM(Case when Cuentas_Cobrar.Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Sal *(-1) Else Cuentas_Cobrar.Mon_Sal End) As Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM			Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine(" 				Tipos_Documentos, ")
            loComandoSeleccionar.AppendLine(" 				Clientes ")
            loComandoSeleccionar.AppendLine(" WHERE 		Cuentas_Cobrar.Cod_tip = Tipos_Documentos.Cod_tip")
            loComandoSeleccionar.AppendLine("               AND     Cuentas_Cobrar.Mon_Sal <> 0")
            loComandoSeleccionar.AppendLine("               AND     Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("               AND 	Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")

			If cusAplicacion.goFormatos.pcCondicionPrincipal.Substring(1,6).ToUpper = "COBROS" Then
				loComandoSeleccionar.AppendLine("               AND  Clientes.Cod_Cli = (SELECT Cod_Cli FROM Cobros WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")
			Else
				loComandoSeleccionar.AppendLine("               AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			End If
            
            loComandoSeleccionar.AppendLine(" GROUP BY		Cuentas_Cobrar.Cod_tip, Tipos_Documentos.nom_tip, Clientes.Nom_Cli")


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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fTTipos_DocumentoCxCResumido", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfTTipos_DocumentoCxCResumido.ReportSource = loObjetoReporte

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
' CMS :  18/05/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS :  19/02/10 : verificación de registros 
'-------------------------------------------------------------------------------------------'
' CMS:	12/06/10: metodo de carga de imagen 
'-------------------------------------------------------------------------------------------'
' MAT:	11/04/11: Ajuste del Select y de la vista de diseño. 
'-------------------------------------------------------------------------------------------'