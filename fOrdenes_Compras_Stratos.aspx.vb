'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
										 
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fOrdenes_Compras_Stratos"
'-------------------------------------------------------------------------------------------'
Partial Class fOrdenes_Compras_Stratos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Compras.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0) THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Compras.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Fax, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Nom_Pro         As Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Rif             As Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Nit             As Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Dir_Fis         As Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Telefonos       As Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Documento, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Control, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Comentario, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("		CASE")
			loComandoSeleccionar.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
			loComandoSeleccionar.AppendLine("			ELSE Renglones_OCompras.Notas")
			loComandoSeleccionar.AppendLine("		END														AS Nom_Art,  ")            
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Por_Des      As Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Precio1      As Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Precio1      As Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Mon_Net      As Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Cod_Imp      As Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Por_Imp1     As Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_OCompras.Mon_Imp1     As Impuesto, ")
            loComandoSeleccionar.AppendLine("           ISNULL(Unidades_Articulos.Can_Uni,Articulos.Can_Uni) AS Can_Uni ")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Compras")
            loComandoSeleccionar.AppendLine(" JOIN  Renglones_OCompras ON (Ordenes_Compras.Documento   =   Renglones_OCompras.Documento)")
            loComandoSeleccionar.AppendLine(" JOIN  Proveedores ON (Ordenes_Compras.Cod_Pro     =   Proveedores.Cod_Pro)")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Formas_Pagos ON ( Ordenes_Compras.Cod_For     =   Formas_Pagos.Cod_For) ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Articulos ON (Articulos.Cod_Art     =   Renglones_OCompras.Cod_Art) ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN Vendedores ON (Ordenes_Compras.Cod_Ven =   Vendedores.Cod_Ven) ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN	Unidades_Articulos ON (Renglones_OCompras.Cod_Art	=	Unidades_Articulos.Cod_Art)")
            loComandoSeleccionar.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

			
			 
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Compras_Stratos", laDatosReporte)

            'CType(loObjetoReporte.ReportDefinition.ReportObjects("Text25"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrdenes_Compras_Stratos.ReportSource = loObjetoReporte

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
' CMS: 08/11/08: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' RJG: 23/04/10: Corrección ortográfica. Eliminada dirección de la empresa bajo el logo.	'	
'				 Cambiada columna "Discount" por "Units/Case".								'
'-------------------------------------------------------------------------------------------'
' MAT: 08/04/11: Ajuste del Select y de la vista de diseño									'
'-------------------------------------------------------------------------------------------'


