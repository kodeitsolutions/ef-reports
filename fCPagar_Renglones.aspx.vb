'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCPagar_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class fCPagar_Renglones

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
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Proveedores.Telefonos,1,15)                                          AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Transportes.Nom_Tra, ")
			loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")            
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Renglon                AS  Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Art                AS  Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Precio1                AS  Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Can_Art                AS  Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Cod_Uni                AS  Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Documentos.Mon_Net                AS  Mon_Net_Renglon ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Renglones_Documentos ON Cuentas_Pagar.Documento = Renglones_Documentos.Documento")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Vendedores ON Cuentas_Pagar.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Formas_Pagos ON Formas_Pagos.Cod_For = Cuentas_Pagar.Cod_For ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Articulos ON Articulos.Cod_Art = Renglones_Documentos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           LEFT JOIN Transportes ON Transportes.Cod_Tra = Cuentas_Pagar.Cod_Tra")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
            
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCPagar_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCPagar_Renglones.ReportSource = loObjetoReporte

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
' CMS: 04/03/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
' MAT: 18/04/11: Ajuste de la Vista de Diseño
'-------------------------------------------------------------------------------------------'
