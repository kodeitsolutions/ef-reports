'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CSJ_rFacturas_Pendientes_Aval"
'-------------------------------------------------------------------------------------------'
Partial Class CSJ_rFacturas_Pendientes_Aval
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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("  SELECT			Facturas.Documento, ")
            loComandoSeleccionar.AppendLine(" 					Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 					Facturas.Fec_Fin, ")
            loComandoSeleccionar.AppendLine(" 					Facturas.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 					Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 					Facturas.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 					Facturas.Cod_Tra, ")
            loComandoSeleccionar.AppendLine(" 					Facturas.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine(" 					Vendedores.Status, ")
            loComandoSeleccionar.AppendLine(" 					Facturas.Control, ")
            loComandoSeleccionar.AppendLine(" 					Facturas.Mon_Net, ")
            loComandoSeleccionar.AppendLine(" 					Facturas.Mon_Sal  ")
            loComandoSeleccionar.AppendLine(" 					From Clientes, ")
            loComandoSeleccionar.AppendLine(" 					Facturas, ")
            loComandoSeleccionar.AppendLine(" 					Vendedores, ")
            loComandoSeleccionar.AppendLine(" 					Transportes, ")
            loComandoSeleccionar.AppendLine(" 					Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE			Facturas.Cod_Cli		= Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" 					And Facturas.Cod_Ven	= Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 					And Facturas.Cod_Tra	= Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine(" 					And Facturas.Cod_Mon	= Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" 					And Facturas.Documento	between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 					And Facturas.Fec_Ini	between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 					And Facturas.Cod_Cli	between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 					And Facturas.Cod_Ven	between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 					And Facturas.Status		IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" 					And Facturas.Cod_Tra between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 					And Facturas.Cod_Mon		between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			    	AND Facturas.Cod_For				BETWEEN" & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			    	AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("                   AND Facturas.Cod_Rev between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    		        AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine(" 			    	AND Facturas.Cod_Suc				BETWEEN" & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			    	AND " & lcParametro9Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY			Facturas.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CSJ_rFacturas_Pendientes_Aval", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCSJ_rFacturas_Pendientes_Aval.ReportSource = loObjetoReporte


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
' MJP: 18/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS:  21/07/09: Se agregaron los siguientes filtros: Forma de pago, sucursal,
'                 Verificacion de registro, Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'