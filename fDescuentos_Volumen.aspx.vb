'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fDescuentos_Volumen"
'-------------------------------------------------------------------------------------------'
Partial Class fDescuentos_Volumen

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Descuentos_Articulos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Cod_Des, ")
            loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Nom_Des, ")
            loComandoSeleccionar.AppendLine("           Departamentos.Nom_Dep, ")
            loComandoSeleccionar.AppendLine("           Secciones.Nom_Sec, ")
            loComandoSeleccionar.AppendLine("           Marcas.Nom_Mar, ")
            loComandoSeleccionar.AppendLine("           Clases_Articulos.Nom_Cla, ")
            loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Can_Des, ")
            loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Can_Has, ")
			loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Por_Des1, ")
			loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Por_Des2, ")
			loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Por_Des3, ")
			loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Por_Des4, ")
			loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Por_Des5, ")
            loComandoSeleccionar.AppendLine("           Descuentos_Articulos.Comentario, ")            
            loComandoSeleccionar.AppendLine("           CASE") 
            loCOmandoSeleccionar.AppendLine("				 WHEN Descuentos_Articulos.Status  = 'A' Then 'Activo' ")
            loComandoSeleccionar.AppendLine("				 WHEN Descuentos_Articulos.Status  = 'I' Then 'Inactivo' ")
            loComandoSeleccionar.AppendLine("				 WHEN Descuentos_Articulos.Status  = 'S' Then 'Suspendido' ")
            loComandoSeleccionar.AppendLine("           END As Status")
            loComandoSeleccionar.AppendLine(" FROM      Descuentos_Articulos,clientes,Departamentos,Secciones,Marcas,Clases_Articulos")
            loComandoSeleccionar.AppendLine(" WHERE     Descuentos_Articulos.Clase  =   'Volumen'")
            loComandoSeleccionar.AppendLine("			AND Clientes.Cod_Cli = Descuentos_Articulos.Cod_Cli")
            loComandoSeleccionar.AppendLine("			AND Departamentos.Cod_Dep = Descuentos_Articulos.Cod_Dep")
            loComandoSeleccionar.AppendLine("			AND Secciones.Cod_Sec = Descuentos_Articulos.Cod_Sec")
            loComandoSeleccionar.AppendLine("			AND Secciones.Cod_Dep = Departamentos.Cod_Dep")
            loComandoSeleccionar.AppendLine("			AND Marcas.Cod_Mar = Descuentos_Articulos.Cod_Mar")
            loComandoSeleccionar.AppendLine("			AND Clases_Articulos.Cod_Cla = Descuentos_Articulos.Cod_Cla")
            loComandoSeleccionar.AppendLine("           And Descuentos_Articulos.Cod_Cli = (SELECT Cod_cli FROM " & cusAplicacion.goFormatos.pcCondicionPrincipal.Split(".")(0).Replace("(", "") & " WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal & ")")
            loComandoSeleccionar.AppendLine("ORDER BY	Descuentos_Articulos.Cod_Cli ASC")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
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
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDescuentos_Volumen", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfDescuentos_Volumen.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' MAT: 21/02/11: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 25/06/12: Se corrigió el filtro de estatus inactivo: tenía "Í" en lugar de "I".		'
'-------------------------------------------------------------------------------------------'
