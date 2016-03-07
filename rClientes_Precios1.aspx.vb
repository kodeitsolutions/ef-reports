Imports System.Data
Partial Class rClientes_Precios1

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Clientes.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Clientes.Rif, ")
            loComandoSeleccionar.AppendLine("           Clientes.Telefonos, ")
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Zon, ")
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("           Clientes.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Precios_Clientes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Precios_Clientes.Precio1, ")
            loComandoSeleccionar.AppendLine("           Precios_Clientes.Precio2, ")
            loComandoSeleccionar.AppendLine("           Precios_Clientes.Precio3, ")
            loComandoSeleccionar.AppendLine("           Precios_Clientes.Precio4, ")
            loComandoSeleccionar.AppendLine("           Precios_Clientes.Precio5, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine(" FROM      Clientes, ")
            loComandoSeleccionar.AppendLine("           Tipos_Clientes, ")
            loComandoSeleccionar.AppendLine("           Zonas, ")
            loComandoSeleccionar.AppendLine("           Clases_Clientes, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Precios_Clientes, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     Clientes.Cod_Tip                =   Tipos_Clientes.Cod_Tip ")
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Zon            =   Zonas.Cod_Zon ")
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Cla            =   Clases_Clientes.Cod_Cla ")
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Ven            =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Cli            =   Precios_Clientes.Cod_Reg ")
            loComandoSeleccionar.AppendLine("           And Precios_Clientes.Cod_Art    =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Precios_Clientes.Origen     =   'Clientes' ")
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Cli            Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Status             IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Tip            Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Zon            Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Cla            Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Ven            Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Clientes.Cod_Cli, Precios_Clientes.Cod_Art ")

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rClientes_Precios1", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrClientes_Precios1.ReportSource = loObjetoReporte

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
' JFP: 08/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 23/02/09: Revision del Codigo inicial. Normalizacion del codigo
'-------------------------------------------------------------------------------------------'
