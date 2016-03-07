Imports System.Data
Partial Class rFacturas_Paises
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Paises.Cod_Pai, ")
            loComandoSeleccionar.AppendLine("			Paises.Nom_Pai, ")
            loComandoSeleccionar.AppendLine("			Paises.Status, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Documento, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.fec_ini, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.fec_fin, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Tra, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Control, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Comentario, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 			Facturas.Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM		Clientes, ")
            loComandoSeleccionar.AppendLine(" 			Facturas, ")
            loComandoSeleccionar.AppendLine(" 			Vendedores, ")
            loComandoSeleccionar.AppendLine(" 			Transportes, ")
            loComandoSeleccionar.AppendLine(" 			Paises, ")
            loComandoSeleccionar.AppendLine(" 			Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE		Clientes.Cod_Cli = Facturas.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" AND 		Paises.Cod_Pai = Clientes.Cod_Pai ")
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Cod_Tra = Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Cod_Mon = Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" AND 		" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" AND 		" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" AND 		" & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" AND 		" & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Cod_Tra between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" AND 		" & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Cod_Mon between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" AND 		" & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" AND 		Facturas.Cod_Suc between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" AND 		" & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY    Paises.Cod_Pai, " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine(" ORDER BY	Facturas.Cod_Cli, ")
            'loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini, ")
            'loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Fin, ")
            'loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Ven ")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFacturas_Paises", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrFacturas_Paises.ReportSource = loObjetoReporte


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
' MJP: 17/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
