'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fHistorico_uAños"
'-------------------------------------------------------------------------------------------'
Partial Class fHistorico_uAños

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim loComandoSeleccionar As New StringBuilder()
									   

			'Lista de Parametros del srored procedure
			'@Ultimo_Año	AS INT			: Obligatorio, Año apartir del cual se calculan hacia atras el rango de años del reporte
			'@Cant_Años		AS INT			: Opcional - Numero de años hacia atras que devolvera el procedimiento 
			'@Cod_Cli		AS VARCHAR(10)	: Excluyente - Codigo de cliente
			'@Cod_Ven		AS VARCHAR(10)	: Excluyente - Codigo de vendedor
			'@Cod_Art		AS VARCHAR(10)	: Excluyente - codigo de articulo


			loComandoSeleccionar.AppendLine(" EXEC sp_Historico_uAños ")
			loComandoSeleccionar.AppendLine(" @Ultimo_Año = '" & Now.Year & "', ")
			

			If cusAplicacion.goFormatos.pcCondicionPrincipal.ToLower.Contains("vendedores") Then
				loComandoSeleccionar.AppendLine(" @Cod_Ven = " & cusAplicacion.goFormatos.pcCondicionPrincipal.Remove(0,20).Replace(")","") & "  ")
			End If
			
			If cusAplicacion.goFormatos.pcCondicionPrincipal.ToLower.Contains("clientes") Then
				loComandoSeleccionar.AppendLine(" @Cod_Cli = " & cusAplicacion.goFormatos.pcCondicionPrincipal.Remove(0,18).Replace(")","") & "  ")
			End If
			
			If cusAplicacion.goFormatos.pcCondicionPrincipal.ToLower.Contains("articulos") Then
				loComandoSeleccionar.AppendLine(" @Cod_Art = " & cusAplicacion.goFormatos.pcCondicionPrincipal.Remove(0,19).Replace(")","") & "  ")
			End If
			
'me.mEscribirConsulta(loComandoSeleccionar.ToString)			
			


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fHistorico_uAños", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfHistorico_uAños.ReportSource = loObjetoReporte

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
' CMS: 26/07/2010: Codigo inicial.
'-------------------------------------------------------------------------------------------'