Imports System.Data
Partial Class PAS_fOrdenes_Compras
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenCompra As String = cusAplicacion.goFormatos.pcCondicionPrincipal
        lcOrdenCompra = lcOrdenCompra.Substring(lcOrdenCompra.IndexOf("'"), (lcOrdenCompra.LastIndexOf("'") - lcOrdenCompra.IndexOf("'")) + 1)

        Try

            Dim loComandoSeleccionar1 As New StringBuilder()

            loComandoSeleccionar1.AppendLine(" SELECT  Ordenes_Compras.Usu_Cre ")
            loComandoSeleccionar1.AppendLine("FROM      Ordenes_Compras")
            loComandoSeleccionar1.AppendLine("	JOIN Renglones_OCompras")
            loComandoSeleccionar1.AppendLine("		ON Ordenes_Compras.Documento = Renglones_OCompras.Documento")
            loComandoSeleccionar1.AppendLine("    JOIN Proveedores")
            loComandoSeleccionar1.AppendLine("		ON  Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar1.AppendLine("    JOIN Formas_Pagos")
            loComandoSeleccionar1.AppendLine("		ON Ordenes_Compras.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar1.AppendLine("    JOIN Articulos ")
            loComandoSeleccionar1.AppendLine("		ON Articulos.Cod_Art = Renglones_OCompras.Cod_Art")
            loComandoSeleccionar1.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios1 As New cusDatos.goDatos

            Dim laDatosReporte1 As DataSet = loServicios1.mObtenerTodosSinEsquema(loComandoSeleccionar1.ToString, "usuarioCreador")

            Dim aString As String = laDatosReporte1.Tables(0).Rows(0).Item(0)
            aString = Trim(aString)

            Dim loComandoSeleccionar2 As New StringBuilder()

            loComandoSeleccionar2.AppendLine(" SELECT    Nom_Usu ")
            loComandoSeleccionar2.AppendLine(" FROM      Usuarios ")
            loComandoSeleccionar2.AppendLine(" WHERE     Cod_Usu = '" & aString & "'")
            loComandoSeleccionar2.AppendLine("		    AND Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))

            Dim loServicios2 As New cusDatos.goDatos

            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte2 As DataSet = loServicios2.mObtenerTodosSinEsquema(loComandoSeleccionar2.ToString, "nombreUsuarioCreador")
            Dim aString2 As String = laDatosReporte2.Tables("nombreUsuarioCreador").Rows(0).Item(0)
            aString2 = RTrim(aString2)

            Dim loComandoSeleccionar3 As New StringBuilder()

            loComandoSeleccionar3.AppendLine("SELECT TOP 1 Cod_Usu ")
            loComandoSeleccionar3.AppendLine("FROM  Auditorias")
            loComandoSeleccionar3.AppendLine("WHERE Documento = " & lcOrdenCompra)
            loComandoSeleccionar3.AppendLine("  AND Tabla = 'Ordenes_Compras'")
            loComandoSeleccionar3.AppendLine("  AND Accion = 'Confirmar'")
            loComandoSeleccionar3.AppendLine("ORDER BY Registro DESC  ")

            Dim loServicios3 As New cusDatos.goDatos

            Dim laDatosReporte3 As DataSet = loServicios3.mObtenerTodosSinEsquema(loComandoSeleccionar3.ToString, "usuarioConfirma")

            Dim aString3 As String = laDatosReporte3.Tables(0).Rows(0).Item(0)
            aString3 = Trim(aString3)

            Dim loComandoSeleccionar4 As New StringBuilder()

            loComandoSeleccionar4.AppendLine(" SELECT    Nom_Usu ")
            loComandoSeleccionar4.AppendLine(" FROM      Usuarios ")
            loComandoSeleccionar4.AppendLine(" WHERE     Cod_Usu = '" & aString3 & "'")
            loComandoSeleccionar4.AppendLine("		    AND Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))

            Dim loServicios4 As New cusDatos.goDatos

            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte4 As DataSet = loServicios4.mObtenerTodosSinEsquema(loComandoSeleccionar4.ToString, "nombreUsuarioConfirma")
            Dim aString4 As String = laDatosReporte4.Tables("nombreUsuarioConfirma").Rows(0).Item(0)
            aString4 = RTrim(aString4)

            Dim loComandoSeleccionar5 As New StringBuilder()

            loComandoSeleccionar5.AppendLine(" SELECT    Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar5.AppendLine("        (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar5.AppendLine("            (CASE WHEN (Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar5.AppendLine("        (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar5.AppendLine("            (CASE WHEN (Ordenes_Compras.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Compras.Rif END) END) AS  Rif, ")
            loComandoSeleccionar5.AppendLine("        Proveedores.Nit, ")
            loComandoSeleccionar5.AppendLine("        (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar5.AppendLine("            (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar5.AppendLine("        (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar5.AppendLine("            (CASE WHEN (Ordenes_Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Compras.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar5.AppendLine("        Proveedores.Correo, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Documento, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Por_Des1 AS Por_Des1_Enc, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Mon_Des1 AS Mon_Des1_Enc, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Por_Rec1 AS Por_Rec1_Enc, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Mon_Rec1 AS Mon_Rec1_Enc, ")
            loComandoSeleccionar5.AppendLine("        Renglones_OCompras.Cod_Uni, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Fec_Fin, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Mon_Bru, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Por_Imp1, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Dis_Imp, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Mon_Imp1, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Mon_Net, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Comentario, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Status,")
            loComandoSeleccionar5.AppendLine("		  '" & aString2 & "' AS Usuario_Crea,")
            loComandoSeleccionar5.AppendLine("		  '" & aString4 & "' AS Usuario_Confirma,")
            loComandoSeleccionar5.AppendLine("        Formas_Pagos.Nom_For, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar5.AppendLine("        Renglones_OCompras.Cod_Art, ")
            loComandoSeleccionar5.AppendLine("        Articulos.Nom_Art               AS Nom_Art, ")
            loComandoSeleccionar5.AppendLine("        Articulos.Generico              AS Generico,")
            loComandoSeleccionar5.AppendLine("        Renglones_OCompras.Notas        AS Notas,")
            loComandoSeleccionar5.AppendLine("        Renglones_OCompras.Can_Art1, ")
            loComandoSeleccionar5.AppendLine("        Renglones_OCompras.Precio1      As Precio1, ")
            loComandoSeleccionar5.AppendLine("        Renglones_OCompras.Mon_Net      As Neto, ")
            loComandoSeleccionar5.AppendLine("        Renglones_OCompras.Doc_Ori, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Registro        As Fec_Cre, ")
            loComandoSeleccionar5.AppendLine("        Ordenes_Compras.Fec_Aut1        As Fec_Aut ")
            loComandoSeleccionar5.AppendLine("FROM      Ordenes_Compras")
            loComandoSeleccionar5.AppendLine("	JOIN Renglones_OCompras")
            loComandoSeleccionar5.AppendLine("		ON Ordenes_Compras.Documento = Renglones_OCompras.Documento")
            loComandoSeleccionar5.AppendLine("    JOIN Proveedores")
            loComandoSeleccionar5.AppendLine("		ON  Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar5.AppendLine("    JOIN Formas_Pagos")
            loComandoSeleccionar5.AppendLine("		ON Ordenes_Compras.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar5.AppendLine("    JOIN Articulos ")
            loComandoSeleccionar5.AppendLine("		ON Articulos.Cod_Art = Renglones_OCompras.Cod_Art")
            loComandoSeleccionar5.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar5.ToString())

            Dim loServicios5 As New cusDatos.goDatos

            Dim laDatosReporte5 As DataSet = loServicios5.mObtenerTodosSinEsquema(loComandoSeleccionar5.ToString, "curReportes")

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpueto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte5.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte5.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte5.Tables(0).Rows.Count - 1 Then
                        If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) <> 0 Then
                            lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
                        End If
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,", "(")


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte5.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte5.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("PAS_fOrdenes_Compras", laDatosReporte5)
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".", ",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text25"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvPAS_fOrdenes_Compras.ReportSource = loObjetoReporte

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
' JJD: 08/11/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos gen.
'-------------------------------------------------------------------------------------------'
' JJD: 09/01/10: Se cambio para que leyera datos de genericos de la Cotizacion cuando aplique
'-------------------------------------------------------------------------------------------'
' CMS: 17/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' MAT: 02/09/11: Adición de Comentario en Renglones
'-------------------------------------------------------------------------------------------'
