Imports System.Data
Partial Class CGS_fRecepciones_Proveedores

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Recepciones.Cod_Pro                    As  Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Recepciones.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Recepciones.Nom_Pro END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Recepciones.Rif = '') THEN Proveedores.Rif ELSE Recepciones.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Recepciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Recepciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Recepciones.Telefonos = '') THEN Proveedores.Telefonos ELSE Recepciones.Telefonos END) END) AS  Telefonos, ")

            loComandoSeleccionar.AppendLine("	        Proveedores.Fax							AS Fax, ")
            loComandoSeleccionar.AppendLine("	        Recepciones.Documento					AS Documento,")
            loComandoSeleccionar.AppendLine("	        Recepciones.Fec_Ini						AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("	        Recepciones.Cod_Ven						AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("	        Recepciones.Comentario					AS Comentario,")
            loComandoSeleccionar.AppendLine("	        SUBSTRING(Vendedores.Nom_Ven,1,25)		AS Nom_Ven, ")
            loComandoSeleccionar.AppendLine("	        Renglones_Recepciones.Cod_Art			AS Cod_Art,")
            loComandoSeleccionar.AppendLine("	        Articulos.Nom_Art						AS Nom_Art,")
            loComandoSeleccionar.AppendLine("           Articulos.Generico						AS Generico,")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Notas				AS Notas, ")
            loComandoSeleccionar.AppendLine("	        Renglones_Recepciones.Can_Art1          AS Cantidad,")
            loComandoSeleccionar.AppendLine("	        Renglones_Recepciones.Cod_Uni			AS Cod_Uni")
            loComandoSeleccionar.AppendLine("FROM       Recepciones")
            loComandoSeleccionar.AppendLine("	        JOIN Renglones_Recepciones")
            loComandoSeleccionar.AppendLine("		        ON Recepciones.Documento = Renglones_Recepciones.Documento")
            loComandoSeleccionar.AppendLine("	        JOIN Proveedores")
            loComandoSeleccionar.AppendLine("		        ON Recepciones.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("	        JOIN Formas_Pagos")
            loComandoSeleccionar.AppendLine("		        ON Recepciones.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("	        JOIN Vendedores")
            loComandoSeleccionar.AppendLine("		        ON Recepciones.Cod_Ven = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("	        JOIN Articulos")
            loComandoSeleccionar.AppendLine("		        ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art ")
            loComandoSeleccionar.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            

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


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("CGS_fRecepciones_Proveedores", laDatosReporte)

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