'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLProduccion_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rLProduccion_Renglones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Libres_Produccion.Documento, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Renglon, ")
            loComandoSeleccionar.AppendLine("           Libres_Produccion.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Transportes.Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine(" FROM      Libres_Produccion, ")
            loComandoSeleccionar.AppendLine("           Renglones_LProduccion, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE     Libres_Produccion.Documento         =   Renglones_LProduccion.Documento ")
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_Cli       =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art               =   Renglones_LProduccion.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_Ven       =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_For       =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Cod_Tra       =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Documento     Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Fec_Ini       Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_LProduccion.Cod_Art   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep               Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Modelo                Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Libres_Produccion.Status        IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Libres_Produccion.Cod_Suc       Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  Libres_Produccion.Documento," & lcOrdenamiento)

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLProduccion_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLProduccion_Renglones.ReportSource = loObjetoReporte

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
' JJD: 09/01/10: Codigo inicial
'-------------------------------------------------------------------------------------------'