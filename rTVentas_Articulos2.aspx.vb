'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_Articulos2"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_Articulos2
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" WITH curTemporal AS ( ")

            loComandoSeleccionar.AppendLine("SELECT			Facturas.Cod_Cli				      AS Cod_Cli, ")
            loComandoSeleccionar.AppendLine("               Clientes.Nom_Cli                      AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine("               Vendedores.Cod_Ven                    AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("               Vendedores.Nom_Ven                    AS Nom_Ven, ")
            loComandoSeleccionar.AppendLine("               Renglones_Facturas.Cod_Art            AS Cod_Art, ")
            loComandoSeleccionar.AppendLine("               Articulos.Nom_Art,								  ")
            loComandoSeleccionar.AppendLine("               Articulos.Cod_Uni1                    AS Cod_Uni, ")
            loComandoSeleccionar.AppendLine("               Renglones_Facturas.Can_Art1           AS Can_Art, ")
            loComandoSeleccionar.AppendLine("               Renglones_Facturas.Mon_Net            AS Mon_Net ")
            loComandoSeleccionar.AppendLine("FROM			Facturas, ")
            loComandoSeleccionar.AppendLine("               Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine("               Articulos, ")
            loComandoSeleccionar.AppendLine("               Vendedores, ")
            loComandoSeleccionar.AppendLine("               Clientes ")
            loComandoSeleccionar.AppendLine("WHERE			Facturas.Documento			=   Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("		AND 	Facturas.Cod_Cli            =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("		AND 	Facturas.Cod_Ven            =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("		AND 	Renglones_Facturas.Cod_Art  =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("		AND 	Facturas.Status             IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine("		AND		Facturas.Fec_Ini		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND		Facturas.Cod_Cli		BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND		Facturas.Cod_Ven		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND		Articulos.Cod_Art		BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND		Facturas.Cod_Rev		BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("   		AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("   	AND		Facturas.Cod_Suc		BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("   		AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(") ")


            loComandoSeleccionar.AppendLine(" SELECT Cod_Cli, ")
            loComandoSeleccionar.AppendLine("                 Nom_Cli, ")
            loComandoSeleccionar.AppendLine("                 Cod_Art, ")
            loComandoSeleccionar.AppendLine("                 Nom_Art, ")
            loComandoSeleccionar.AppendLine("                 Cod_Uni, ")
            loComandoSeleccionar.AppendLine("                 SUM(Can_Art) AS  Can_Art, ")
            loComandoSeleccionar.AppendLine("                 SUM(Mon_Net) AS  Mon_Net ")
            loComandoSeleccionar.AppendLine(" FROM curTemporal ")
            loComandoSeleccionar.AppendLine(" GROUP BY Cod_Cli, Nom_Cli, Cod_Art, Nom_Art, Cod_Uni ")
            loComandoSeleccionar.AppendLine("ORDER BY  Cod_Cli, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_Articulos2", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTVentas_Articulos2.ReportSource = loObjetoReporte

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
' JFP: 07/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Estandarizacion de codigo y Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' RJG: 09/12/10: Ajustado el estatus de las facturas de venta en el filtro.					' 
'-------------------------------------------------------------------------------------------'
