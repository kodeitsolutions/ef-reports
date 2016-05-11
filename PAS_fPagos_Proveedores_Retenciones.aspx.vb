'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "PAS_fPagos_Proveedores_Retenciones"
'-------------------------------------------------------------------------------------------'
Partial Class PAS_fPagos_Proveedores_Retenciones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	Pagos.Documento, ")
            loComandoSeleccionar.AppendLine("       CONVERT(NCHAR(10), Pagos.Fec_Ini, 103)		AS	Fec_Ini, ")
            loComandoSeleccionar.AppendLine("	    Pagos.Recibo, ")
            loComandoSeleccionar.AppendLine("		Pagos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Telefonos, ")
            loComandoSeleccionar.AppendLine("       Pagos.Comentario,")
            loComandoSeleccionar.AppendLine("       Pagos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("       Vendedores.Nom_Ven,")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Mon_Net               AS  Mon_Net,")
            loComandoSeleccionar.AppendLine("       Documentos.Factura                  AS Factura,")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Cod_Tip      AS  Tip_Ori, ")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Clase        AS  Tip_Ret, ")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Num_Com      AS  Num_Com, ")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Mon_Bas      AS  Mon_Bas, ")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Por_Ret      AS  Por_Ret, ")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Mon_Ret      AS  Mon_Ret, ")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Mon_Sus      AS  Mon_Sus  ")
            loComandoSeleccionar.AppendLine(" FROM Pagos  ")
            loComandoSeleccionar.AppendLine("       JOIN Proveedores ON Pagos.Cod_Pro = Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("       JOIN Vendedores ON Pagos.Cod_Ven  =   Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("       JOIN Cuentas_Pagar ON Cuentas_Pagar.Doc_Ori  = Pagos.Documento")
            loComandoSeleccionar.AppendLine("         AND Cuentas_Pagar.Tip_Ori  =   'Pagos'")
            loComandoSeleccionar.AppendLine("       JOIN Retenciones_Documentos ON Pagos.Documento = Retenciones_Documentos.Documento")
            loComandoSeleccionar.AppendLine("         AND Cuentas_Pagar.Cod_Tip = Retenciones_Documentos.Cla_Des")
            loComandoSeleccionar.AppendLine("         AND Cuentas_Pagar.Documento = Retenciones_Documentos.Doc_Des")
            loComandoSeleccionar.AppendLine("       JOIN Cuentas_Pagar AS Documentos ON Documentos.Documento = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine(" WHERE  " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ORDER BY Cuentas_Pagar.Cod_Tip")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("PAS_fPagos_Proveedores_Retenciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvPAS_fPagos_Proveedores_Retenciones.ReportSource = loObjetoReporte

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
