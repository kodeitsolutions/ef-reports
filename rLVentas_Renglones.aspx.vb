'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLVentas_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rLVentas_Renglones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Libres_Ventas.Documento, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Renglon, ")
            loComandoSeleccionar.AppendLine("           Libres_Ventas.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Articulos.Nom_Art,1,37)       AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Vendedores.Nom_Ven,1,20)      AS  Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Transportes.Nom_Tra,1,20)     AS  Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,45)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine(" FROM      Libres_Ventas, ")
            loComandoSeleccionar.AppendLine("           Renglones_LVentas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE     Libres_Ventas.Documento         =   Renglones_LVentas.Documento ")
            loComandoSeleccionar.AppendLine("           And Libres_Ventas.Cod_Cli       =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art           =   Renglones_LVentas.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Libres_Ventas.Cod_Ven       =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Libres_Ventas.Cod_For       =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           And Libres_Ventas.Cod_Tra       =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Libres_Ventas.Documento     Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Ventas.Fec_Ini       Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Ventas.Cod_Cli       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Ventas.Status        IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Libres_Ventas.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro4Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY  Libres_Ventas.Documento")
            loComandoSeleccionar.AppendLine("ORDER BY    Libres_Ventas.Documento,   " & lcOrdenamiento)

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLVentas_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLVentas_Renglones.ReportSource = loObjetoReporte

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
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  06/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'