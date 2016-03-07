'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCobros_Clientes_Retenciones"
'-------------------------------------------------------------------------------------------'
Partial Class fCobros_Clientes_Retenciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cobros.Documento, ")
            loComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Cobros.Fec_Ini, 103)		AS	Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cobros.Recibo, ")
            loComandoSeleccionar.AppendLine("			Cobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Cobros.Comentario, ")
            loComandoSeleccionar.AppendLine("           Cobros.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpCobros ")
            loComandoSeleccionar.AppendLine(" FROM      Cobros, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Vendedores ")
            loComandoSeleccionar.AppendLine(" WHERE     Cobros.Cod_Cli          =   Clientes.Cod_Cli    AND ")
            loComandoSeleccionar.AppendLine("           Cobros.Cod_Ven          =   Vendedores.Cod_Ven  AND ")
            loComandoSeleccionar.AppendLine("           " & cusAplicacion.goFormatos.pcCondicionPrincipal)



            loComandoSeleccionar.AppendLine(" SELECT	#tmpCobros.Documento, ")
            loComandoSeleccionar.AppendLine("           #tmpCobros.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			#tmpCobros.Recibo, ")
            loComandoSeleccionar.AppendLine("			#tmpCobros.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           #tmpCobros.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           #tmpCobros.Rif, ")
            loComandoSeleccionar.AppendLine("           #tmpCobros.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           #tmpCobros.Telefonos, ")
            loComandoSeleccionar.AppendLine("           #tmpCobros.Comentario, ")
            loComandoSeleccionar.AppendLine("           #tmpCobros.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           #tmpCobros.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Documento    AS  Num_Doc, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Mon_Net      AS  Mon_Net ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpDocumentos ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpCobros, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Doc_Ori  =   #tmpCobros.Documento    AND ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Tip_Ori  =   'Cobros'                ")



            loComandoSeleccionar.AppendLine(" SELECT	#tmpDocumentos.Documento, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Recibo, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           #tmpDocumentos.Nom_Cli, ")
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCobros_Clientes_Retenciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCobros_Clientes_Retenciones.ReportSource = loObjetoReporte

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
