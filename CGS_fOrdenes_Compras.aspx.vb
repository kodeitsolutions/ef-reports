Imports System.Data
Partial Class CGS_fOrdenes_Compras
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT Ordenes_Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Rif, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("       Proveedores.Telefonos,")
            loComandoSeleccionar.AppendLine("       Proveedores.Correo, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Documento, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Por_Des1 AS Por_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Des1 AS Mon_Des1_Enc, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Por_Rec1 AS Por_Rec1_Enc, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Rec1 AS Mon_Rec1_Enc, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Dis_Imp, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Comentario, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Status,")
            loComandoSeleccionar.AppendLine("       COALESCE((SELECT Factory_Global.dbo.Usuarios.Nom_Usu ")
            loComandoSeleccionar.AppendLine("       FROM Factory_Global.dbo.Usuarios ")
            loComandoSeleccionar.AppendLine("       WHERE Factory_Global.dbo.Usuarios.Cod_Usu COLLATE SQL_Latin1_General_CP1_CI_AS = Ordenes_Compras.Usu_Cre COLLATE SQL_Latin1_General_CP1_CI_AS),'') AS Usuario,")
            loComandoSeleccionar.AppendLine("       Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Cod_Art, ")
            loComandoSeleccionar.AppendLine("       Articulos.Nom_Art               AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("       Articulos.Generico              AS Generico,")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Notas        AS Notas,")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Precio1      AS Precio1, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Mon_Net      AS Neto, ")
            loComandoSeleccionar.AppendLine("       Renglones_OCompras.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Registro        AS Fec_Cre, ")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fec_Aut1        AS Fec_Aut, ")
            loComandoSeleccionar.AppendLine("       COALESCE((SELECT TOP 1 Auditorias.Cod_Usu")
            loComandoSeleccionar.AppendLine("                FROM Auditorias")
            loComandoSeleccionar.AppendLine("                WHERE	Ordenes_Compras.Documento = Auditorias.Documento")
            loComandoSeleccionar.AppendLine("                  AND Auditorias.Tabla = 'Ordenes_Compras'")
            loComandoSeleccionar.AppendLine("                  AND Auditorias.Accion = 'Confirmar'")
            loComandoSeleccionar.AppendLine("                ORDER BY Auditorias.Registro DESC)")
            loComandoSeleccionar.AppendLine("       ,'')	AS Usu_Aut,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Prioridad       AS Tipo,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico1         AS mgentili,")
            'loComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico2          AS ssimanca,")
            loComandoSeleccionar.AppendLine("	   CASE WHEN COALESCE((SELECT TOP 1 Auditorias.Cod_Usu")
            loComandoSeleccionar.AppendLine("							FROM Auditorias")
            loComandoSeleccionar.AppendLine("							WHERE	Ordenes_Compras.Documento = Auditorias.Documento")
            loComandoSeleccionar.AppendLine("							  AND Auditorias.Tabla = 'Ordenes_Compras'")
            loComandoSeleccionar.AppendLine("							  AND Auditorias.Accion = 'Confirmar'")
            loComandoSeleccionar.AppendLine("							ORDER BY Auditorias.Registro DESC),'') = 'ssimanca'")
            loComandoSeleccionar.AppendLine("			THEN CAST(1 AS BIT)")
            loComandoSeleccionar.AppendLine("			ELSE CAST(0 AS BIT) END		AS ssimanca,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico3         AS lcarrizal,")
            'loComandoSeleccionar.AppendLine("       Ordenes_Compras.Logico4          AS yreina,")
            loComandoSeleccionar.AppendLine("	   CASE WHEN COALESCE((SELECT TOP 1 Auditorias.Cod_Usu")
            loComandoSeleccionar.AppendLine("							FROM Auditorias")
            loComandoSeleccionar.AppendLine("							WHERE	Ordenes_Compras.Documento = Auditorias.Documento")
            loComandoSeleccionar.AppendLine("							  AND Auditorias.Tabla = 'Ordenes_Compras'")
            loComandoSeleccionar.AppendLine("							  AND Auditorias.Accion = 'Confirmar'")
            loComandoSeleccionar.AppendLine("							ORDER BY Auditorias.Registro DESC),'') = 'yreina'")
            loComandoSeleccionar.AppendLine("			THEN CAST(1 AS BIT)")
            loComandoSeleccionar.AppendLine("			ELSE CAST(0 AS BIT) END     AS yreina,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fecha1          AS Faut_mgentili,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fecha2          AS Faut_ssimanca,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fecha3          AS Faut_lcarrizal,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Fecha4          AS Faut_yreina,   ")
            loComandoSeleccionar.AppendLine("       COALESCE(CONCAT(RTRIM(Requisiciones.Caracter1),CHAR(13),RTRIM(Requisiciones.Caracter2),CHAR(13),")
            loComandoSeleccionar.AppendLine("	    RTRIM(Requisiciones.Caracter3),CHAR(13),RTRIM(Requisiciones.Caracter4)),'') AS Solicitante ")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Compras")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Requisiciones ON Renglones_OCompras.Doc_Ori = Requisiciones.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_OCompras.Tip_Ori = 'Requisiciones'")
            loComandoSeleccionar.AppendLine("   JOIN Proveedores ON  Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("   JOIN Formas_Pagos ON Ordenes_Compras.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ON Articulos.Cod_Art = Renglones_OCompras.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios3 As New cusDatos.goDatos

            Dim laDatosReporte3 As DataSet = loServicios3.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpueto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte3.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte3.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte3.Tables(0).Rows.Count - 1 Then
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
            Me.mCargarLogoEmpresa(laDatosReporte3.Tables(0), "LogoEmpresa")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte3.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fOrdenes_Compras", laDatosReporte3)
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".", ",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text25"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fOrdenes_Compras.ReportSource = loObjetoReporte

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
