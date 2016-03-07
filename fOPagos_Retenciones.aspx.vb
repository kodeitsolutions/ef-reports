'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fOPagos_Retenciones"
'-------------------------------------------------------------------------------------------'
Partial Class fOPagos_Retenciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("  SELECT	Ordenes_Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("  			Ordenes_Pagos.Fec_Ini,   ")
            loComandoSeleccionar.AppendLine("  			Ordenes_Pagos.Fec_Fin,   ")
            loComandoSeleccionar.AppendLine("  			Ordenes_Pagos.Status,  ")
            loComandoSeleccionar.AppendLine("  			Ordenes_Pagos.Cod_Pro,") 
            loComandoSeleccionar.AppendLine("  			Proveedores.Nom_Pro,  ")
            loComandoSeleccionar.AppendLine("  			Proveedores.Rif,	  ")
            loComandoSeleccionar.AppendLine("  			Proveedores.Dir_Fis,  ")
            loComandoSeleccionar.AppendLine("  			Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("  			Ordenes_Pagos.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("  			Ordenes_Pagos.Tasa,	   ")
            loComandoSeleccionar.AppendLine("  			Ordenes_Pagos.Motivo,  ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Clase        AS  Clase, ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Tip_Ori      AS  Tip_Ori, ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Renglon      AS  Renglon, ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Cod_Ret      AS  Cod_Ret, ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Notas		AS  Notas, ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Num_Com      AS  Num_Com, ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Comentario	AS  Comentario, ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Mon_Bas      AS  Mon_Bas, ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Por_Ret      AS  Por_Ret, ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Mon_Ret      AS  Mon_Ret,  ")
            loComandoSeleccionar.AppendLine("  			Retenciones_Documentos.Mon_Sus      AS  Mon_Sus	 ")
            loComandoSeleccionar.AppendLine("  FROM		Ordenes_Pagos  ")
            loComandoSeleccionar.AppendLine("  JOIN Proveedores ON (Ordenes_Pagos.Cod_Pro = Proveedores.Cod_Pro)")
            loComandoSeleccionar.AppendLine("  JOIN Retenciones_Documentos ON (Ordenes_Pagos.Documento    =   Retenciones_Documentos.Documento AND Origen='Ordenes_Pagos')")
            loComandoSeleccionar.AppendLine("  WHERE   " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("  ORDER BY Renglon ASC	")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" ")          



          


            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            'Me.mEscribirConsulta(loCOmandoSeleccionar.ToString())
            
            
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOPagos_Retenciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfOPagos_Retenciones.ReportSource = loObjetoReporte

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
' MAT: 30/05/11: Programacion inicial														'
'-------------------------------------------------------------------------------------------'

