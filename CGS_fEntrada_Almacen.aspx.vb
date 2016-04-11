Imports System.Data
Partial Class CGS_fEntrada_Almacen

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar1 As New StringBuilder()

            loComandoSeleccionar1.AppendLine(" SELECT  Recepciones.Usu_Cre ")
            loComandoSeleccionar1.AppendLine("FROM      Recepciones")
            loComandoSeleccionar1.AppendLine("	JOIN Renglones_Recepciones")
            loComandoSeleccionar1.AppendLine("		ON Recepciones.Documento = Renglones_Recepciones.Documento")
            loComandoSeleccionar1.AppendLine("    JOIN Proveedores")
            loComandoSeleccionar1.AppendLine("		ON  Recepciones.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar1.AppendLine("    JOIN Articulos ")
            loComandoSeleccionar1.AppendLine("		ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art")
            loComandoSeleccionar1.AppendLine("WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios1 As New cusDatos.goDatos

            Dim laDatosReporte1 As DataSet = loServicios1.mObtenerTodosSinEsquema(loComandoSeleccionar1.ToString, "usuario")

            Dim aString As String = laDatosReporte1.Tables(0).Rows(0).Item(0)
            aString = Trim(aString)

            Dim loComandoSeleccionar2 As New StringBuilder()

            loComandoSeleccionar2.AppendLine(" SELECT    Nom_Usu ")
            loComandoSeleccionar2.AppendLine(" FROM      Usuarios ")
            loComandoSeleccionar2.AppendLine(" WHERE     Cod_Usu = '" & aString & "'")
            loComandoSeleccionar2.AppendLine("		    AND Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))

            Dim loServicios2 As New cusDatos.goDatos

            goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte2 As DataSet = loServicios2.mObtenerTodosSinEsquema(loComandoSeleccionar2.ToString, "nombreUsuario")
            Dim aString2 As String = laDatosReporte2.Tables("nombreUsuario").Rows(0).Item(0)
            aString2 = RTrim(aString2)

            Dim loComandoSeleccionar3 As New StringBuilder()

            loComandoSeleccionar3.AppendLine("SELECT	ROW_NUMBER() OVER (ORDER BY Registros.Fec_Ini) AS Num,*")
            loComandoSeleccionar3.AppendLine("FROM(SELECT	Recepciones.Cod_Pro                    As  Cod_Cli, ")
            loComandoSeleccionar3.AppendLine("            (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar3.AppendLine("                (CASE WHEN (Recepciones.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Recepciones.Nom_Pro END) END) AS  Nom_Cli, ")
            loComandoSeleccionar3.AppendLine("            (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar3.AppendLine("                (CASE WHEN (Recepciones.Rif = '') THEN Proveedores.Rif ELSE Recepciones.Rif END) END) AS  Rif, ")
            loComandoSeleccionar3.AppendLine("            Proveedores.Nit, ")
            loComandoSeleccionar3.AppendLine("            (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar3.AppendLine("                (CASE WHEN (SUBSTRING(Recepciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Recepciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar3.AppendLine("            (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar3.AppendLine("                (CASE WHEN (Recepciones.Telefonos = '') THEN Proveedores.Telefonos ELSE Recepciones.Telefonos END) END) AS  Telefonos, ")

            loComandoSeleccionar3.AppendLine("            Proveedores.Fax						AS Fax, ")
            loComandoSeleccionar3.AppendLine("            Recepciones.Documento					AS Documento,")
            loComandoSeleccionar3.AppendLine("            Recepciones.Fec_Ini					AS Fec_Ini,")
            loComandoSeleccionar3.AppendLine("            Recepciones.Cod_Ven					AS Cod_Ven, ")
            loComandoSeleccionar3.AppendLine("            Recepciones.Comentario				AS Comentario,")
            loComandoSeleccionar3.AppendLine("            SUBSTRING(Vendedores.Nom_Ven,1,25)	AS Nom_Ven, ")
            loComandoSeleccionar3.AppendLine("            RR.Cod_Art							AS Cod_Art,")
            loComandoSeleccionar3.AppendLine("            Articulos.Nom_Art						AS Nom_Art,")
            loComandoSeleccionar3.AppendLine("            Articulos.Generico					AS Generico,")
            loComandoSeleccionar3.AppendLine("            RR.Notas								AS Notas, ")
            loComandoSeleccionar3.AppendLine("            RR.Can_Art1						    AS Cantidad,")
            loComandoSeleccionar3.AppendLine("            RR.Cod_Uni							AS Cod_Uni,")
            loComandoSeleccionar3.AppendLine("			  RR.Doc_Ori							AS OC_Origen,")
            loComandoSeleccionar3.AppendLine("            CASE WHEN ((SELECT Renglones_OCompras.Can_Art1")
            loComandoSeleccionar3.AppendLine("						FROM Renglones_OCompras")
            loComandoSeleccionar3.AppendLine("							JOIN Renglones_Recepciones ON Renglones_OCompras.Documento =  Renglones_Recepciones.Doc_Ori")
            loComandoSeleccionar3.AppendLine("								AND Renglones_OCompras.Renglon = Renglones_Recepciones.Ren_Ori")
            loComandoSeleccionar3.AppendLine("						 WHERE Renglones_OCompras.Documento = RR.Doc_Ori")
            loComandoSeleccionar3.AppendLine("							AND Renglones_OCompras.Renglon = RR.Ren_Ori")
            loComandoSeleccionar3.AppendLine("			) = RR.Can_Art1) THEN 'TOTAL' ELSE 'PARCIAL' END AS Entrega,")
            loComandoSeleccionar3.AppendLine("		    '" & aString2 & "'                        AS Almacenista,")
            loComandoSeleccionar3.AppendLine("		    Recepciones.Fec_Reg                       AS Fec_Rec")
            loComandoSeleccionar3.AppendLine("FROM       Recepciones")
            loComandoSeleccionar3.AppendLine("            JOIN Renglones_Recepciones AS RR")
            loComandoSeleccionar3.AppendLine("            	ON Recepciones.Documento = RR.Documento")
            loComandoSeleccionar3.AppendLine("            JOIN Proveedores")
            loComandoSeleccionar3.AppendLine("            	ON Recepciones.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar3.AppendLine("            JOIN Formas_Pagos")
            loComandoSeleccionar3.AppendLine("            	ON Recepciones.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar3.AppendLine("            JOIN Vendedores")
            loComandoSeleccionar3.AppendLine("            	ON Recepciones.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar3.AppendLine("            JOIN Articulos")
            loComandoSeleccionar3.AppendLine("            	ON Articulos.Cod_Art = RR.Cod_Art ")
            loComandoSeleccionar3.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal & ") Registros ")

            'Me.mEscribirConsulta(loComandoSeleccionar3.ToString)

            Dim loServicios3 As New cusDatos.goDatos
            Dim laDatosReporte3 As DataSet = loServicios3.mObtenerTodosSinEsquema(loComandoSeleccionar3.ToString, "curReportes")



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


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte3.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fEntrada_Almacen", laDatosReporte3)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_fRecepciones_Proveedores.ReportSource = loObjetoReporte

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
' JJD: 27/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos genericos
'-------------------------------------------------------------------------------------------'
' CMS: 16/09/09: Se Agrego la distribucion de impuesto
'-------------------------------------------------------------------------------------------'
' MAT: 01/03/11: Ajuste de la vista de diseño												'
'-------------------------------------------------------------------------------------------'