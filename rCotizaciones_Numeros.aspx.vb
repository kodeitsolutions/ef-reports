'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCotizaciones_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rCotizaciones_Numeros
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("  SELECT		Cotizaciones.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Fec_Fin, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 				Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Cod_Tra, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores.Status, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Control, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones.Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" FROM			Clientes, ")
            loComandoSeleccionar.AppendLine(" 				Cotizaciones, ")
            loComandoSeleccionar.AppendLine(" 				Vendedores, ")
            loComandoSeleccionar.AppendLine(" 				Transportes, ")
            loComandoSeleccionar.AppendLine(" 				Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE			Cotizaciones.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Tra = Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Mon = Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Ven between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Status IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Tra between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Cotizaciones.Cod_Mon between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("               AND Cotizaciones.Cod_Rev between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    		    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("               AND Cotizaciones.Cod_Suc between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    		    AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY Cotizaciones." & lcOrdenamiento)


            'loComandoSeleccionar.AppendLine(" ORDER BY		Cotizaciones.Documento ")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCotizaciones_Numeros", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCotizaciones_Numeros.ReportSource = loObjetoReporte


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
' JJD: 24/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' GCR: 16/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS: 14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP: 29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' JJD: 15/08/09: Se incluyo el orden de los registros
'-------------------------------------------------------------------------------------------'
