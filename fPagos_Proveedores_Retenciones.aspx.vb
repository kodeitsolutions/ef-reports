'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPagos_Proveedores_Retenciones"
'-------------------------------------------------------------------------------------------'
Partial Class fPagos_Proveedores_Retenciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Pagos.Fec_Ini, 103)		AS	Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Pagos.Recibo, ")
            loComandoSeleccionar.AppendLine("			Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Pagos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Pagos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpPagos ")
            loComandoSeleccionar.AppendLine(" FROM      Pagos, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Vendedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Pagos.Cod_Pro          =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Pagos.Cod_Ven          =   Vendedores.Cod_Ven  AND ")
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)



            loComandoSeleccionar.AppendLine(" SELECT	#tmpPagos.Documento, ")
            loComandoSeleccionar.AppendLine("           #tmpPagos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			#tmpPagos.Recibo, ")
            loComandoSeleccionar.AppendLine("			#tmpPagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           #tmpPagos.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           #tmpPagos.Rif, ")
            loComandoSeleccionar.AppendLine("           #tmpPagos.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           #tmpPagos.Telefonos, ")
            loComandoSeleccionar.AppendLine("           #tmpPagos.Comentario, ")
            loComandoSeleccionar.AppendLine("           #tmpPagos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           #tmpPagos.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Documento    AS  Num_Doc, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Net      AS  Mon_Net ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpDocumentos ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpPagos, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Pagar.Doc_Ori  =   #tmpPagos.Documento    AND ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Tip_Ori  =   'Pagos'                ")



            loComandoSeleccionar.AppendLine(" SELECT	#tmpDocumentos.Documento, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Recibo, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Rif, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Telefonos, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Comentario, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Num_Doc, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Cod_Tip      AS  Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Doc_Ori      AS  Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Cla_Des      AS  Tip_Des, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Clase        AS  Tip_Ret, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Cod_Con      AS  Cod_Con, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Num_Com      AS  Num_Com, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Mon_Bas      AS  Mon_Bas, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Por_Ret      AS  Por_Ret, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Mon_Ret      AS  Mon_Ret, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Mon_Sus      AS  Mon_Sus, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Mon_Net ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpDocumentos, ")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos ")
            loComandoSeleccionar.AppendLine(" WHERE     #tmpDocumentos.Documento    =   Retenciones_Documentos.Documento    AND ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Cod_Tip      =   Retenciones_Documentos.Cla_Des      AND ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Num_Doc      =   Retenciones_Documentos.Doc_Des      ")
            loComandoSeleccionar.AppendLine(" ORDER BY  #tmpDocumentos.Cod_Tip ")


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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPagos_Proveedores_Retenciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPagos_Proveedores_Retenciones.ReportSource = loObjetoReporte

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
' JJD: 28/04/11: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' JJD: 29/04/11: Continuacion de la programacion											'
'-------------------------------------------------------------------------------------------'
