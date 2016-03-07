Imports System.Data
Partial Class fOrdenes_Compras_Ingles
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Compras.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")

            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Ent,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Ent,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Ent,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Ent,1, 200) END) END) AS  Dir_Ent, ")

            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Compras.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Nom_Pro         As Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Rif             As Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Nit             As Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Dir_Fis         As Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Telefonos       As Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Documento, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Por_Des1 AS Por_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Des1 AS Mon_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Por_Rec1 AS Por_Rec1_Enc, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Rec1 AS Mon_Rec1_Enc, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Comentario, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Art, ")
            loComandoSeleccionar.AppendLine("		CASE")
            loComandoSeleccionar.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
            loComandoSeleccionar.AppendLine("			ELSE Renglones_OCompras.Notas")
            loComandoSeleccionar.AppendLine("		END														AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Por_Des      As Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Precio1      As Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Precio1      As Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Comentario   As Comentario_Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Mon_Net      As Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Imp      As Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Por_Imp1     As Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Mon_Imp1     As Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Compras, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Compras.Documento   =   Renglones_OCompras.Documento AND ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Pro     =   Proveedores.Cod_Pro AND ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_For     =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art           =   Renglones_OCompras.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)



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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Compras_Ingles", laDatosReporte)
           
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrdenes_Compras_Ingles.ReportSource = loObjetoReporte

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
' MAT: 04/10/2011: Código Inicial
'-------------------------------------------------------------------------------------------'