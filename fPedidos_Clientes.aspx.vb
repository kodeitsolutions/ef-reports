'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPedidos_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class fPedidos_Clientes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Pedidos.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Pedidos.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Pedidos.Rif = '') THEN Clientes.Rif ELSE Pedidos.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Pedidos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Pedidos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Pedidos.Telefonos = '') THEN Clientes.Telefonos ELSE Pedidos.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Documento, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Articulos.UPC, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Art, ")
            'loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Pedidos.Notas END AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Pedidos, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Pedidos.Documento  =   Renglones_Pedidos.Documento AND ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art  =   Renglones_Pedidos.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("ORDER BY Articulos.UPC ASC,Renglones_Pedidos.Cod_Art")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
			Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpuesto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpuesto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
						If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
								If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)<> 0 Then
									lcPorcentajesImpuesto = lcPorcentajesImpuesto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
								End If
						End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpuesto = lcPorcentajesImpuesto & ")"
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace("(,", "(")
            
            
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

			
			
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPedidos_Clientes", laDatosReporte)
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace(".",",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpuesto.ToString

            
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPedidos_Clientes.ReportSource = loObjetoReporte

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
' JJD: 13/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 24/12/09: Se le incluyo los montos de descuentos y recargos.
'-------------------------------------------------------------------------------------------'
' JJD: 09/01/10: Se cambio para que leyera los datos genericos del Pedido cuando aplique
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT: 01/03/11: Se programo la distribución de impuestos para mostrarlo en el formato.		'
'-------------------------------------------------------------------------------------------'