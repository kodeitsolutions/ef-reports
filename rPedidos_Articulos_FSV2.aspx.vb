Imports System.Data
Partial Class rPedidos_Articulos_FSV2
    Inherits vis2Formularios.frmReporte

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

		Dim loComandoSeleccionar As New StringBuilder()

				loComandoSeleccionar.Appendline("SELECT  Articulos.Cod_Art, ")
				loComandoSeleccionar.Appendline("		Articulos.Nom_art, ")
				loComandoSeleccionar.Appendline("		Articulos.Cod_Mar, ")
				loComandoSeleccionar.Appendline("		Articulos.Status, ")
				loComandoSeleccionar.Appendline("		Pedidos.Documento, ")
				loComandoSeleccionar.Appendline("		Pedidos.Cod_Cli, ")
				loComandoSeleccionar.Appendline("		Clientes.Nom_Cli, ")
				loComandoSeleccionar.Appendline("		Pedidos.Fec_Ini, ")
				loComandoSeleccionar.Appendline("		Pedidos.Cod_Ven, ")
				loComandoSeleccionar.Appendline("		Renglones_Pedidos.Renglon, ")
				loComandoSeleccionar.Appendline("		Renglones_Pedidos.Cod_Alm, ")
				loComandoSeleccionar.Appendline("		Renglones_Pedidos.Can_Art1, ")
				loComandoSeleccionar.Appendline("		Renglones_Pedidos.Cod_Uni, ")
				loComandoSeleccionar.Appendline("		Renglones_Pedidos.Precio1, ")
				loComandoSeleccionar.Appendline("		Renglones_Pedidos.Por_Des, ")
				loComandoSeleccionar.Appendline("		Renglones_Pedidos.Mon_Net ")
				loComandoSeleccionar.Appendline("From Articulos, ")
				loComandoSeleccionar.Appendline("		Pedidos, ")
				loComandoSeleccionar.Appendline("		Renglones_Pedidos, ")
				loComandoSeleccionar.Appendline("		Clientes, ")
				loComandoSeleccionar.Appendline("		Vendedores, ")
				loComandoSeleccionar.Appendline("		Almacenes, ")
				loComandoSeleccionar.Appendline("		Marcas ")
				loComandoSeleccionar.Appendline("WHERE Articulos.Cod_Art = Renglones_Pedidos.Cod_Art ")
				loComandoSeleccionar.Appendline("		AND Renglones_Pedidos.Documento = Pedidos.Documento ")
				loComandoSeleccionar.Appendline("		AND Articulos.Cod_Mar = Marcas.Cod_Mar ")
				loComandoSeleccionar.Appendline("		AND Pedidos.Cod_Cli = Clientes.Cod_Cli ")
				loComandoSeleccionar.Appendline("		AND Pedidos.Cod_Ven = Vendedores.Cod_Ven")
				loComandoSeleccionar.Appendline("		AND Renglones_Pedidos.Cod_Alm = Almacenes.Cod_Alm ")
				loComandoSeleccionar.Appendline("		AND Clientes.Status ='A' ")
				loComandoSeleccionar.Appendline("		AND Articulos.Cod_Art between " & lcParametro0Desde )
				loComandoSeleccionar.Appendline("		AND " & lcParametro0Hasta )
				loComandoSeleccionar.Appendline("		AND Pedidos.Fec_Ini between " & lcParametro1Desde )
				loComandoSeleccionar.Appendline("		AND " & lcParametro1Hasta )
				loComandoSeleccionar.Appendline("		AND Clientes.Cod_Cli between " & lcParametro2Desde )
				loComandoSeleccionar.Appendline("		AND " & lcParametro2Hasta )
				loComandoSeleccionar.Appendline("		AND Vendedores.Cod_Ven between " & lcParametro3Desde )
				loComandoSeleccionar.Appendline("		AND " & lcParametro3Hasta )
				loComandoSeleccionar.Appendline("		AND Marcas.Cod_Mar  between" & lcParametro4Desde )
				loComandoSeleccionar.Appendline("		AND " & lcParametro4Hasta )
				loComandoSeleccionar.Appendline("		AND Pedidos.Status IN (" & lcParametro5Desde & ")" )
				loComandoSeleccionar.Appendline("		AND Almacenes.Cod_Alm between " & lcParametro6Desde )
				loComandoSeleccionar.Appendline("		AND " & lcParametro6Hasta )
				loComandoSeleccionar.Appendline("GROUP BY Pedidos.Cod_Cli, Articulos.Cod_Art, Renglones_Pedidos.Precio1, Articulos.Nom_art, Articulos.Cod_Mar, Articulos.Status, Pedidos.Documento, Clientes.Nom_Cli, Pedidos.Fec_Ini, Pedidos.Cod_Ven, Renglones_Pedidos.Renglon, Renglones_Pedidos.Cod_Alm, Renglones_Pedidos.Can_Art1, Renglones_Pedidos.Cod_Uni, Renglones_Pedidos.Por_Des, Renglones_Pedidos.Mon_Net, Renglones_Pedidos.Cod_Uni, Renglones_Pedidos.Por_Des, Renglones_Pedidos.Mon_Net " )
				loComandoSeleccionar.Appendline("ORDER BY Pedidos.Cod_Cli, Articulos.Cod_Art, Pedidos.Fec_Ini ASC" )
				

				
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString , "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_Articulos_FSV2", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_Articulos_FSV2.ReportSource = loObjetoReporte


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
' JFP: 01/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT:  17/02/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
