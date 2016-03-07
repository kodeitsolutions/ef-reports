Imports System.Data
Partial Class fCotizaciones_Ventas_FSV1
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cotizaciones.Cod_Cli, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Nit, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Dir_Fis, ")
            'loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Cotizaciones.Nom_Cli>'' THEN Cotizaciones.Nom_Cli ELSE Clientes.Nom_Cli END) AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Cotizaciones.Rif>'' THEN Cotizaciones.Rif ELSE Clientes.Rif END) AS Rif,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Cotizaciones.Nit>'' THEN Cotizaciones.Nit ELSE Clientes.Nit END) AS Nit,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Cotizaciones.Dir_Fis>'' THEN REPLACE(Cotizaciones.Dir_Fis,CHAR(10),'') ELSE REPLACE(Clientes.Dir_Fis,CHAR(10),'') END) AS Dir_Fis,")
            loComandoSeleccionar.AppendLine("           (CASE WHEN Cotizaciones.Telefonos>'' THEN Cotizaciones.Telefonos ELSE Clientes.Telefonos END) AS Telefonos,")
            loComandoSeleccionar.AppendLine("           Clientes.Fax, ")
            loComandoSeleccionar.AppendLine("           Clientes.movil, ")
            loComandoSeleccionar.AppendLine("           Clientes.contacto, ")
            loComandoSeleccionar.AppendLine("           Clientes.correo, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nom_Cli        As  Nom_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Rif            As  Rif_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Nit            As  Nit_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Dir_Fis        As  Dir_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Telefonos      As  Tel_Gen, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Documento, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Tasa, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cotizaciones.Comentario, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art + Renglones_Cotizaciones.Comentario    As  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Net  As  Neto, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Por_Imp1 As  Por_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Cod_Imp, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Mon_Imp1 As  Impuesto ")
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

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
		    
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCotizaciones_Ventas_FSV1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCotizaciones_Ventas_FSV1.ReportSource = loObjetoReporte

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
' MAT: 19/04/11: Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'
' JFP: 09/04/13: Se elimino el substring del comentario del renglon
'-------------------------------------------------------------------------------------------'
' RJG: 06/02/15: Ahora se muestran los datos genericos del cliente; Si están en blanco, se 
'                muestran los datos reales del cliente.                                     
'-------------------------------------------------------------------------------------------'
' EAG: 04/09/15: Se agregaron los campos contactos, movil y correo de Clientes                                 
'-------------------------------------------------------------------------------------------'
