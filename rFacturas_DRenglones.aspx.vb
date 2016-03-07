Imports System.Data
Partial Class rFacturas_DRenglones

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

            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))

            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	")
            loComandoSeleccionar.AppendLine("           CASE	")
            loComandoSeleccionar.AppendLine("           		WHEN DATEPART(DW, Facturas.Fec_Ini) = 2 THEN 'Lunes'	")
            loComandoSeleccionar.AppendLine("           		WHEN DATEPART(DW, Facturas.Fec_Ini) = 3 THEN 'Martes'	")
            loComandoSeleccionar.AppendLine("           		WHEN DATEPART(DW, Facturas.Fec_Ini) = 4 THEN 'Miercoles'	")
            loComandoSeleccionar.AppendLine("           		WHEN DATEPART(DW, Facturas.Fec_Ini) = 5 THEN 'Jueves'	")
            loComandoSeleccionar.AppendLine("           		WHEN DATEPART(DW, Facturas.Fec_Ini) = 6 THEN 'Viernes'	")
            loComandoSeleccionar.AppendLine("           		WHEN DATEPART(DW, Facturas.Fec_Ini) = 7 THEN 'Sabado'	")
            loComandoSeleccionar.AppendLine("           		WHEN DATEPART(DW, Facturas.Fec_Ini) = 1 THEN 'Domingo'	")
            loComandoSeleccionar.AppendLine("           END AS	NombreDia,	")
            loComandoSeleccionar.AppendLine("           DATEPART(WEEKDAY,Facturas.Fec_Ini)   AS  NumeroDia, ")
            loComandoSeleccionar.AppendLine("           Facturas.Documento, ")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Facturas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Facturas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Facturas.Tasa, ")
            loComandoSeleccionar.AppendLine("           Facturas.Status, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Articulos.Nom_Art,1,50)       AS  Nom_Art, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Vendedores.Nom_Ven,1,30)      AS  Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Transportes.Nom_Tra,1,30)     AS  Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Clientes.Nom_Cli,1,50)        AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Almacenes.Nom_Alm,1,30)       AS  Nom_Alm ")
            loComandoSeleccionar.AppendLine(" INTO      #curTemporal ")
            loComandoSeleccionar.AppendLine(" FROM      Facturas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Almacenes ")
            loComandoSeleccionar.AppendLine(" WHERE     Facturas.Documento      =   Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Cli    =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Ven    =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_For    =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Tra    =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Almacenes.Cod_Alm   =   Renglones_Facturas.Cod_Alm ")
            loComandoSeleccionar.AppendLine("           And Facturas.Documento  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Fec_Ini    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Cli    Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Ven    Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Status     IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("           And Almacenes.Cod_Alm   Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Mon    Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Facturas.Cod_Suc    Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)

            loComandoSeleccionar.AppendLine(" ORDER BY  Facturas.Documento")


            loComandoSeleccionar.AppendLine(" SELECT	NombreDia, ")
            loComandoSeleccionar.AppendLine("           NumeroDia, ")
            loComandoSeleccionar.AppendLine("           Documento, ")
            loComandoSeleccionar.AppendLine("           Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Cod_For, ")
            loComandoSeleccionar.AppendLine("           Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Tasa, ")
            loComandoSeleccionar.AppendLine("           Status, ")
            loComandoSeleccionar.AppendLine("           Renglon, ")
            loComandoSeleccionar.AppendLine("           Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Precio1, ")
            loComandoSeleccionar.AppendLine("           Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Por_Des, ")
            loComandoSeleccionar.AppendLine("           Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Comentario, ")
            loComandoSeleccionar.AppendLine("           Nom_Art, ")
            loComandoSeleccionar.AppendLine("           Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           Nom_For, ")
            loComandoSeleccionar.AppendLine("           Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           Nom_Alm ")
            loComandoSeleccionar.AppendLine(" FROM      #curTemporal ")
            'loComandoSeleccionar.AppendLine(" ORDER BY  Fec_Ini, Renglon ")
            loComandoSeleccionar.AppendLine("ORDER BY     Documento, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFacturas_DRenglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrFacturas_DRenglones.ReportSource = loObjetoReporte

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
' JJD: 23/02/09: Codigo inicial.
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  14/07/09: Metodo de Ordenamiento, Se tradujeron los dias de la semana
'-------------------------------------------------------------------------------------------'
' CMS: 26/03/10: Se cambio la funcion DATENAME por DATEPART
'-------------------------------------------------------------------------------------------'