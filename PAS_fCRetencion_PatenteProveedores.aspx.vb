'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "PAS_fCRetencion_PatenteProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class PAS_fCRetencion_PatenteProveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Doc_Ori				AS Numero_Pago,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Control				AS Control_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Factura				AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Fec_Ini			    AS Fecha_Factura,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Mon_Abo				AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Num_Com		AS Comprobante,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Proveedores.Dir_Fis					AS Direccion")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		JOIN	Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.clase = 'PATENTE'")
            loComandoSeleccionar.AppendLine("		JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine(" WHERE  " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº 0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("PAS_fCRetencion_PatenteProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPAS_fCRetencion_PatenteProveedores.ReportSource = loObjetoReporte


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
' RJG: 09/04/15: Codigo inicial, a partir de rCRetencion_ISLRProveedores.           		'
'-------------------------------------------------------------------------------------------'
