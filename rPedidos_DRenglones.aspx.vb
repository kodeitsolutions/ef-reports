Imports System.Data
Partial Class rPedidos_DRenglones

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

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	DATENAME(DW, Pedidos.Fec_Ini)       AS	NombreDia, ")
            loComandoSeleccionar.AppendLine("           DATEPART(WEEKDAY,Pedidos.Fec_Ini)   AS  NumeroDia, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Documento, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_For, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Tasa, ")
            loComandoSeleccionar.AppendLine("           Pedidos.Status, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Precio1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Por_Des, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos.Comentario, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Vendedores.Nom_Ven,1,30)      AS  Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Transportes.Nom_Tra,1,30)     AS  Nom_Tra, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)    AS  Nom_For, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Clientes.Nom_Cli,1,50)        AS  Nom_Cli, ")
            loComandoSeleccionar.AppendLine("           SUBSTRING(Almacenes.Nom_Alm,1,30)       AS  Nom_Alm ")
            loComandoSeleccionar.AppendLine(" INTO      #curTemporal ")
            loComandoSeleccionar.AppendLine(" FROM      Pedidos, ")
            loComandoSeleccionar.AppendLine("           Renglones_Pedidos, ")
            loComandoSeleccionar.AppendLine("           Clientes, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            loComandoSeleccionar.AppendLine("           Vendedores, ")
            loComandoSeleccionar.AppendLine("           Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("           Transportes, ")
            loComandoSeleccionar.AppendLine("           Almacenes ")
            loComandoSeleccionar.AppendLine(" WHERE     Pedidos.Documento      =   Renglones_Pedidos.Documento ")
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Cli    =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Ven    =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_For    =   Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Tra    =   Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art   =   Renglones_Pedidos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           And Almacenes.Cod_Alm   =   Renglones_Pedidos.Cod_Alm ")
            loComandoSeleccionar.AppendLine("           And Pedidos.Documento  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Fec_Ini    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Articulos.Cod_Art   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Cli    Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Ven    Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Status     IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("           And Almacenes.Cod_Alm   Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("           And Pedidos.Cod_Mon    Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  Pedidos.Documento")




            loComandoSeleccionar.AppendLine(" SELECT	 ")
            loComandoSeleccionar.AppendLine("           case NombreDia ")
            loComandoSeleccionar.AppendLine("                    when 'Monday'   then 'Lunes' ")
            loComandoSeleccionar.AppendLine("                    when 'Tuesday'  then 'Martes' ")
            loComandoSeleccionar.AppendLine("                    when 'Wednesday' then 'Miercoles' ")
            loComandoSeleccionar.AppendLine("                    when 'Thursday' then 'Jueves' ")
            loComandoSeleccionar.AppendLine("                    when 'Friday'  then 'Viernes' ")
            loComandoSeleccionar.AppendLine("                    when 'Saturday' then 'Sabado' ")
            loComandoSeleccionar.AppendLine("                    when 'Sunday'   then 'Domingo' ")
            loComandoSeleccionar.AppendLine("           end as NombreDia,   ")
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
             loComandoSeleccionar.AppendLine("ORDER BY    Documento, " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_DRenglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_DRenglones.ReportSource = loObjetoReporte

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