'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCuentas_Pagar_IHP"
'-------------------------------------------------------------------------------------------'
Partial Class fCuentas_Pagar_IHP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Pagar.Documento, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Factura, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Control, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Status, ")
            loComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Cuentas_Pagar.Fec_Ini, 103)                                 AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Cuentas_Pagar.Fec_Fin, 103)                                 AS  Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Tra                                                          AS  Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_For                                                          AS  Cod_For, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Cod_Mon                                                          AS  Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Tasa                                                             AS  Tasa, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Bru                                                          AS  Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Por_Imp1                                                         AS  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Imp1                                                         AS  Mon_Imp, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Net                                                          AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Sal                                                          AS  Saldo, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Por_Rec                                                          AS  Por_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Rec                                                          AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Por_Des                                                          AS  Por_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Des                                                          AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           (Cuentas_Pagar.Mon_Otr1 + Cuentas_Pagar.Mon_Otr2 + Cuentas_Pagar.Mon_Otr3)     AS  Mon_Otr, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Comentario                                                       AS  Comentario, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Doc_Ori, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            'loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           SUBSTRING(Proveedores.Telefonos,1,15)                                          AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Cuentas_Pagar.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Pagar.Rif = '') THEN Proveedores.Rif ELSE Cuentas_Pagar.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Cuentas_Pagar.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Cuentas_Pagar.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Pagar.Telefonos = '') THEN Proveedores.Telefonos ELSE Cuentas_Pagar.Telefonos END) END) AS  Telefonos, ")
            
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Transportes.Nom_Tra ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Pagar, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Pagar.Cod_Pro      =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Ven  =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_For  =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Tra  =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCuentas_Pagar_IHP", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCuentas_Pagar_IHP.ReportSource = loObjetoReporte

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
' MAT: 22/08/11: Programacion inicial
'-------------------------------------------------------------------------------------------'
