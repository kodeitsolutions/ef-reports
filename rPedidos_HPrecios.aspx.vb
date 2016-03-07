Imports System.Data
Partial Class rFacturas_HPrecios

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
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro10Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            'loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Facturas.Fec_Ini, 103)		AS	Fec_Ini, ")
            'loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Facturas.Fec_Ini, 103)		AS	Fec_Ini1, ")
            'loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Facturas.Fec_Ini, 112)     AS	Fec_Ini2, ")

            loComandoSeleccionar.AppendLine(" SELECT	Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Facturas.Fec_Ini, 103)    AS	Fec_Ini1, ")
            loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(10), Facturas.Fec_Ini, 112)    AS	Fec_Ini2, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Precio1 ")
            loComandoSeleccionar.AppendLine(" INTO      #curTemporal ")
            loComandoSeleccionar.AppendLine(" FROM      Facturas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Almacenes ")
            loComandoSeleccionar.AppendLine(" WHERE     Facturas.Documento               =   Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Cli             =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Ven             =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_For             =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Tra             =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art           =   Renglones_Facturas.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Almacenes.Cod_Alm           =   Renglones_Facturas.Cod_Alm ")
            loComandoSeleccionar.AppendLine("           And Renglones_Facturas.Cod_Art   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Fec_Ini             Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Cli             Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Ven             Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Dep           Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Mar           Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Status              IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           And Renglones_Facturas.Cod_Alm   Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Mon             Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Tra             Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Status              IN (" & lcParametro10Desde & ")")
            loComandoSeleccionar.AppendLine(" ORDER BY  Facturas.Cod_Cli, Renglones_Facturas.Cod_Art, Facturas.Fec_Ini ")


            loComandoSeleccionar.AppendLine(" SELECT	TOP 1 Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Precio1, ")
            loComandoSeleccionar.AppendLine("           Fec_Ini1, ")
            loComandoSeleccionar.AppendLine("           Fec_Ini2, ")
            loComandoSeleccionar.AppendLine("           'Borrame'	AS	Cam_Eli ")
            loComandoSeleccionar.AppendLine(" INTO		#curTemporal3 ")
            loComandoSeleccionar.AppendLine(" FROM		#curTemporal ")

            loComandoSeleccionar.AppendLine(" DECLARE   @Pre_Ant	FLOAT ")
            loComandoSeleccionar.AppendLine(" SET       @Pre_Ant    =   0.00 ")

            loComandoSeleccionar.AppendLine(" DECLARE curPrueba CURSOR SCROLL KEYSET FOR ")
            loComandoSeleccionar.AppendLine(" SELECT	* FROM #curTemporal ")
            loComandoSeleccionar.AppendLine(" OPEN curPrueba ")

            loComandoSeleccionar.AppendLine(" DECLARE @Cod_Cli	CHAR(10) ")
            loComandoSeleccionar.AppendLine(" DECLARE @Cod_Art	CHAR(30) ")
            loComandoSeleccionar.AppendLine(" DECLARE @Fec_Ini1	CHAR(10) ")
            loComandoSeleccionar.AppendLine(" DECLARE @Fec_Ini2	INTEGER ")
            loComandoSeleccionar.AppendLine(" DECLARE @Precio1	FLOAT ")
            loComandoSeleccionar.AppendLine(" FETCH NEXT FROM curPrueba INTO @Cod_Cli, @Cod_Art, @Fec_Ini1, @Fec_Ini2, @Precio1 ")
            loComandoSeleccionar.AppendLine(" WHILE @@FETCH_STATUS = 0 ")
            loComandoSeleccionar.AppendLine(" BEGIN ")
            loComandoSeleccionar.AppendLine(" IF @Precio1 <> @Pre_Ant ")
            loComandoSeleccionar.AppendLine("	BEGIN ")
            loComandoSeleccionar.AppendLine("       INSERT INTO #curTemporal3	(Cod_Cli, Cod_Art, Fec_Ini1, Fec_Ini2, Precio1, Cam_Eli) ")
            loComandoSeleccionar.AppendLine("       VALUES						(@Cod_Cli, @Cod_Art, @Fec_Ini1, @Fec_Ini2, @Precio1, 'Dejame') ")
            loComandoSeleccionar.AppendLine("		SET @Pre_Ant = @Precio1 ")
            loComandoSeleccionar.AppendLine("	END  ")
            loComandoSeleccionar.AppendLine("	FETCH NEXT FROM curPrueba INTO @Cod_Cli, @Cod_Art, @Fec_Ini1, @Fec_Ini2, @Precio1 ")
            loComandoSeleccionar.AppendLine(" END ")
            loComandoSeleccionar.AppendLine(" CLOSE curPrueba ")
            loComandoSeleccionar.AppendLine(" DEALLOCATE curPrueba ")

            loComandoSeleccionar.AppendLine(" SELECT	#curTemporal3.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           #curTemporal3.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           #curTemporal3.Fec_Ini1  AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           #curTemporal3.Fec_Ini2  AS  Fec_Ini2, ")
            loComandoSeleccionar.AppendLine("           #curTemporal3.Precio1, ")
            loComandoSeleccionar.AppendLine("           Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine(" INTO      #curTemporal4 ")
            loComandoSeleccionar.AppendLine(" FROM      #curTemporal3, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE     #curTemporal3.Cod_Cli       =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And #curTemporal3.Cod_Art   =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And #curTemporal3.Cam_Eli   <>  'Borrame' ")

            loComandoSeleccionar.AppendLine(" SELECT	* ")
            loComandoSeleccionar.AppendLine(" FROM      #curTemporal4 ")
            'loComandoSeleccionar.AppendLine(" ORDER BY  Cod_Cli, Cod_Art, Fec_Ini2 ")
            loComandoSeleccionar.AppendLine("ORDER BY    Cod_Cli, Cod_Art, " & lcOrdenamiento)

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFacturas_HPrecios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrFacturas_HPrecios.ReportSource = loObjetoReporte

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
' JJD: 24/02/09: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' CMS:  18/08/09: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' CMS:  10/09/09: Se agrego el filtro estatus de cliente
'-------------------------------------------------------------------------------------------'