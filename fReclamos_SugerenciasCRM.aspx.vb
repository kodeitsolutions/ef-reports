'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fReclamos_SugerenciasCRM"
'-------------------------------------------------------------------------------------------'
Partial Class fReclamos_SugerenciasCRM
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine(" SELECT")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Documento, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Status, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Fec_Ini, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Fec_Fin, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Fec_Reg, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Fec_Cie, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Tip_Doc, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Tip_Pro, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cod_Tip, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cod_Cli, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cod_Ven, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cod_Ope, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Res_Cli, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Res_Emp, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Cuerpo, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Motivo, ")
			loComandoSeleccionar.AppendLine(" 			Reclamos.Comentario, ")
			loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli, ")
			loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven, ")
			loComandoSeleccionar.AppendLine(" 			Operador.Nom_Ven AS Nom_Ope, ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Documento AS Renglon,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Cod_Art,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Cod_Uni,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Can_Art1,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Nom_Art,   ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Comentario AS Com_Ren,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Fec_Ini AS Emision ,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Fec_Fin As Vence,  ")
			loComandoSeleccionar.AppendLine(" 			Renglones_Reclamos.Cod_Ven AS Cod_Ven_Renglon,  ")
			loComandoSeleccionar.AppendLine(" 			Vendedor_Renglon.Nom_Ven AS Nom_Ven_Renglon, ")
			loComandoSeleccionar.AppendLine(" 			Tipos_Reclamos.Nom_Tip")
			loComandoSeleccionar.AppendLine(" FROM Reclamos")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Renglones_Reclamos ON Renglones_Reclamos.Documento = Reclamos.Documento")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Tipos_Reclamos ON Tipos_Reclamos.Cod_Tip = Reclamos.Cod_Tip")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Clientes ON Clientes.Cod_Cli = Reclamos.Cod_Cli")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Vendedores ON Vendedores.Cod_Ven = Reclamos.Cod_Ven")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Vendedores AS Operador ON Operador.Cod_Ven = Reclamos.Cod_Ope")
			loComandoSeleccionar.AppendLine(" LEFT JOIN Vendedores AS Vendedor_Renglon ON Vendedor_Renglon.Cod_Ven = Reclamos.Cod_Ven")
			loComandoSeleccionar.AppendLine(" WHERE")	   			
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
           	'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
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

            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fReclamos_SugerenciasCRM", laDatosReporte)
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfReclamos_SugerenciasCRM.ReportSource = loObjetoReporte

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
' MAT: 04/12/10 : Codigo inicial
'-------------------------------------------------------------------------------------------'
