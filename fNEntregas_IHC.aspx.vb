'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fNEntregas_IHC"
'-------------------------------------------------------------------------------------------'
Partial Class fNEntregas_IHC

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Entregas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Entregas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Entregas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Entregas.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Entregas.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Entregas.Rif = '') THEN Clientes.Rif ELSE Entregas.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Entregas.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Entregas.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Entregas.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Entregas.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Entregas.Telefonos = '') THEN Clientes.Telefonos ELSE Entregas.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Entregas.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Entregas.Documento, ")
            loComandoSeleccionar.AppendLine("           Entregas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Entregas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_For, ")            
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("           Entregas.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Entregas.Por_Des1, ")            
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Entregas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Cod_Art, ")
            'loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("		CASE")
			loComandoSeleccionar.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
			loComandoSeleccionar.AppendLine("			ELSE Renglones_Entregas.Notas")
			loComandoSeleccionar.AppendLine("		END														AS Nom_Art,  ")            
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Renglon, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Entregas.Can_Art1, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Entregas.Cod_Uni, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Entregas.Precio1, ")
            'loComandoSeleccionar.AppendLine("       (CASE WHEN (Renglones_Entregas.Cod_Uni2='')")
            'loComandoSeleccionar.AppendLine("			THEN Renglones_Entregas.Can_Art1")
            'loComandoSeleccionar.AppendLine("			ELSE Renglones_Entregas.Can_Art2")
            'loComandoSeleccionar.AppendLine("		END)													AS Can_Art1, ")
            'loComandoSeleccionar.AppendLine("       (CASE WHEN (Renglones_Entregas.Cod_Uni2='')")
            'loComandoSeleccionar.AppendLine("			THEN Renglones_Entregas.Cod_Uni")
            'loComandoSeleccionar.AppendLine("			ELSE Renglones_Entregas.Cod_Uni2")
            'loComandoSeleccionar.AppendLine("		END)													AS Cod_Uni, ")
            'loComandoSeleccionar.AppendLine("       (CASE WHEN (Renglones_Entregas.Cod_Uni2='')")
            'loComandoSeleccionar.AppendLine("			THEN Renglones_Entregas.Precio1")
            'loComandoSeleccionar.AppendLine("			ELSE Renglones_Entregas.Precio1*Renglones_Entregas.Can_Uni2")
            'loComandoSeleccionar.AppendLine("		END)													AS Precio1, ")
            
			loComandoSeleccionar.AppendLine("      (CASE WHEN (Renglones_Entregas.Cod_Uni = Renglones_Entregas.Cod_Uni2) ")
			loComandoSeleccionar.AppendLine("		  THEN Renglones_Entregas.Can_Art1 ")
			loComandoSeleccionar.AppendLine("		  ELSE Renglones_Entregas.Can_Art2 ")
			loComandoSeleccionar.AppendLine("		END)                                                                        AS Can_Art1,")
			loComandoSeleccionar.AppendLine("      (CASE WHEN (Renglones_Entregas.Cod_Uni = Renglones_Entregas.Cod_Uni2) ")
			loComandoSeleccionar.AppendLine("		  THEN Renglones_Entregas.Cod_Uni ")
			loComandoSeleccionar.AppendLine("		  ELSE Renglones_Entregas.Cod_Uni2 ")
			loComandoSeleccionar.AppendLine("		END)                                                                        AS Cod_Uni, ")
			loComandoSeleccionar.AppendLine("      (CASE WHEN (Renglones_Entregas.Cod_Uni = Renglones_Entregas.Cod_Uni2) ")
			loComandoSeleccionar.AppendLine("		  THEN Renglones_Entregas.Precio1 ")
			loComandoSeleccionar.AppendLine("		  ELSE Renglones_Entregas.Precio1*Renglones_Entregas.Can_Uni2 ")
			loComandoSeleccionar.AppendLine("		END)                                                                        AS Precio1, ")

            
			loComandoSeleccionar.AppendLine("           Renglones_Entregas.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Entregas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Entregas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Entregas.Documento  =   Renglones_Entregas.Documento AND ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Entregas.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art   =   Renglones_Entregas.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

			
			Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpueto = "("

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
                        lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,", "(")
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".", ",")
			
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

            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fNEntregas_IHC", laDatosReporte)
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text27"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)
			Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfNEntregas_Clientes.ReportSource = loObjetoReporte

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
' MAT: 08/06/11: Codigo inicial
'-------------------------------------------------------------------------------------------'