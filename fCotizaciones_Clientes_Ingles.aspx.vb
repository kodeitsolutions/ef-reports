'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCotizaciones_Clientes_Ingles"
'-------------------------------------------------------------------------------------------'
Partial Class fCotizaciones_Clientes_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cotizaciones.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cotizaciones.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cotizaciones.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cotizaciones.Rif = '') THEN Clientes.Rif ELSE Cotizaciones.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Cotizaciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Cotizaciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cotizaciones.Telefonos = '') THEN Clientes.Telefonos ELSE Cotizaciones.Telefonos END) END) AS  Telefonos, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nom_Cli                    As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Rif                        As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nit                        As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Dir_Fis                    As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Telefonos                  As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Des1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Por_Rec1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Des1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Rec1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Cotizaciones.Notas END AS Nom_Art,  ")
            'loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Net          As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Por_Imp1         As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Imp1         As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Cotizaciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Cotizaciones.Documento  =   Renglones_Cotizaciones.Documento AND ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Cli    =   Clientes.Cod_Cli AND ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_For    =   Formas_Pagos.Cod_For AND ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Ven    =   Vendedores.Cod_Ven AND ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Art       =   Renglones_Cotizaciones.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCotizaciones_Clientes_Ingles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCotizaciones_Clientes_Ingles.ReportSource = loObjetoReporte

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
' Douglas Cortez: 05/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RAC: 23/03/2011: Se modifico la etiqueta QUOTE por QUOTATION en el archivo rpt.
'-------------------------------------------------------------------------------------------'
' RAC: 24/03/2011: Se modifico la etiqueta QUOTE NUMBER por QUOTATION NUMBER en el archivo rpt.
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Ajuste de la vista de Diseño
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Eliminación del Pie de Página de eFactory según Requerimientos
'-------------------------------------------------------------------------------------------'

