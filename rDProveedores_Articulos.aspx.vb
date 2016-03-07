Imports System.Data

Partial Class rDProveedores_Articulos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Cod_Mar, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Status, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Documento, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Renglones_dProveedores.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_dProveedores.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_dProveedores.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_dProveedores.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_dProveedores.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_dProveedores.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_dProveedores.Mon_Net ")
            loComandoSeleccionar.AppendLine(" FROM      Articulos, ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Proveedores, ")
            loComandoSeleccionar.AppendLine("           Renglones_dProveedores, ")
            loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Almacenes, ")
            loComandoSeleccionar.AppendLine("           Marcas ")
            loComandoSeleccionar.AppendLine(" WHERE     Renglones_dProveedores.Cod_Art          =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Renglones_dProveedores.Documento    =   Devoluciones_Proveedores.Documento ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar                   =   Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("           And Devoluciones_Proveedores.Cod_Pro    =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Devoluciones_Proveedores.Cod_Ven    =   Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine("           And Renglones_dProveedores.Cod_Alm      =   Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("           And Renglones_dProveedores.Cod_Art      Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Devoluciones_Proveedores.Fec_Ini    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Devoluciones_Proveedores.Cod_Pro    Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Devoluciones_Proveedores.Cod_Ven    Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar                   Between" & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Devoluciones_Proveedores.Status     IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("           And Renglones_dProveedores.Cod_Alm      Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Devoluciones_Proveedores.Cod_rev    Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Devoluciones_Proveedores.Cod_Suc    Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("ORDER BY   Articulos.Cod_Art, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDProveedores_Articulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDProveedores_Articulos.ReportSource = loObjetoReporte

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
' JJD: 10/01/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 21/03/09: Ajustes al seleccionar el Status
'-------------------------------------------------------------------------------------------'
' GCR: 21/03/09: Ajustes al seleccionar el Status
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'
' CMS:  04/08/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  11/08/09: Verificacion de registros
'-------------------------------------------------------------------------------------------'