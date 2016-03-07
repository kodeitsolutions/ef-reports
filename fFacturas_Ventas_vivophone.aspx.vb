Imports System.Data
Partial Class fFacturas_Ventas_Vivophone
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END) END) AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Rif = '') THEN Clientes.Rif ELSE Facturas.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (SUBSTRING(Facturas.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Facturas.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Facturas.Telefonos = '') THEN Clientes.Telefonos ELSE Facturas.Telefonos END) END) AS  Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Clientes.Generico, ")
            loComandoSeleccionar.AppendLine("           Facturas.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Facturas.Documento, ")
            loComandoSeleccionar.AppendLine("           Facturas.Status,	")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Ini,	")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Fin,	")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Mon, 	")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Bru,	")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Imp1,	")
            loComandoSeleccionar.AppendLine("           Facturas.Por_Imp1,	")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Net,	")
            loComandoSeleccionar.AppendLine("           Facturas.Por_Des1,	")
            loComandoSeleccionar.AppendLine("           Facturas.Dis_Imp,	")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Des1       As  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Facturas.Por_Rec1,	")
            loComandoSeleccionar.AppendLine("           Facturas.Mon_Rec1                       AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_For                        AS  Cod_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Notas END AS Nom_Art,  ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Renglon, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Can_Art1")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Can_Art2 END) AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Cod_Uni")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Cod_Uni2 END) AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Precio1")
            loComandoSeleccionar.AppendLine("			    ELSE Renglones_Facturas.Precio1*Renglones_Facturas.Can_Uni2 END) AS Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Imp1 As  Impuesto ")
            loComandoSeleccionar.AppendLine(" FROM      Facturas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Facturas.Documento      =   Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Cli    =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_For    =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven    =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_Ventas_Vivophone", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFacturas_Ventas_Vivophone.ReportSource = loObjetoReporte

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
' GMO: 16/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 08/11/08: Ajustes al select
'-------------------------------------------------------------------------------------------'
' RJG: 01/09/09: Agregado código para mostrar unidad segundaria.							'
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos gen.
'-------------------------------------------------------------------------------------------'
' JJD: 09/01/10: Se cambio para que leyera los datos genericos de la Factura cuando aplique
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 19/03/10: Se a justo la logica para determinar el nombre del cliente
'		(Clientes.Generico = 0 ) a (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '')
'-------------------------------------------------------------------------------------------'
' RAC: 23/03/11: Se cambio la etiqueta Client por Customer en el archivo rpt.
'-------------------------------------------------------------------------------------------'
' MAT :  24/08/11 : Ajuste del Select, Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
' MAT :  15/09/11: Eliminación del Pie de Página del eFactory según requerimiento
'-------------------------------------------------------------------------------------------'
' RJG: 13/12/12: Se agregó el campo Moneda (Cod_Mon).										'
'-------------------------------------------------------------------------------------------'
