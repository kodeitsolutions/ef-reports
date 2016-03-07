'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPedidos_PFacturar"
'-------------------------------------------------------------------------------------------'
Partial Class rPedidos_PFacturar
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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))

            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcInstruccion As String = "Renglones_Pedidos.Can_Art1 "
            Select Case lcParametro10Desde
                Case Is = "Todos"
                    lcInstruccion = "Renglones_Pedidos.Can_Art1 "
                Case Is = "Bacorder"
                    lcInstruccion = "Renglones_Pedidos.Can_Pen1 "
                Case Is = "Procesado"
                    lcInstruccion = "(Renglones_Pedidos.Can_Art1 - Renglones_Pedidos.Can_Pen1)"
            End Select
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Pedidos.Documento, ")
            loComandoSeleccionar.AppendLine("    		Pedidos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("	    	Pedidos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("		    Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine("	    	Clientes.Cod_Zon, ")
            loComandoSeleccionar.AppendLine("	    	Zonas.Nom_Zon, ")
            loComandoSeleccionar.AppendLine("	    	Clientes.Dir_Fis, ")
            loComandoSeleccionar.AppendLine("	    	Pedidos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("	    	Renglones_Pedidos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("	    	" & lcInstruccion & " AS Can_Art1, ")
            loComandoSeleccionar.AppendLine("		    Articulos.Modelo, ")
            loComandoSeleccionar.AppendLine("		    Articulos.Exi_Act1, ")
            loComandoSeleccionar.AppendLine("		    Vendedores.Nom_Ven ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTemporal ")
            loComandoSeleccionar.AppendLine(" FROM	    Pedidos, ")
            loComandoSeleccionar.AppendLine("		    Renglones_Pedidos, ")
            loComandoSeleccionar.AppendLine("		    Clientes, ")
            loComandoSeleccionar.AppendLine("		    Zonas, ")
            loComandoSeleccionar.AppendLine("		    Articulos, ")
            loComandoSeleccionar.AppendLine("		    Clases_Articulos, ")
            loComandoSeleccionar.AppendLine("		    Vendedores ")
            loComandoSeleccionar.AppendLine(" WHERE	    Pedidos.Documento               =   Renglones_Pedidos.Documento ")
            loComandoSeleccionar.AppendLine(" 		    AND Pedidos.Cod_Cli             =   Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" 		    AND Clientes.Cod_Zon            =   Zonas.Cod_Zon ")
            loComandoSeleccionar.AppendLine(" 		    AND Renglones_Pedidos.Cod_Art   =   Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine(" 		    AND Articulos.Cod_Cla           =   Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine(" 		    AND Pedidos.Cod_Ven             =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 		    AND Renglones_Pedidos.Cod_Art   BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Pedidos.Fec_Ini             BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Pedidos.Cod_Cli             BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Pedidos.Cod_Ven             BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Articulos.Cod_Dep           BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Articulos.Cod_Cla           BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Pedidos.Status IN ( " & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine(" 		    AND Renglones_Pedidos.Cod_Alm   BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Pedidos.Cod_Mon             BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Clientes.Cod_Est            BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Pedidos.Cod_Tra             BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Clientes.Cod_Zon            BETWEEN " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento & "")

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Ven, ")
            loComandoSeleccionar.AppendLine("    		Nom_Ven, ")
            loComandoSeleccionar.AppendLine("    		Documento, ")
            loComandoSeleccionar.AppendLine("    		Dir_Fis, ")
            loComandoSeleccionar.AppendLine("    		Nom_Zon, ")
            loComandoSeleccionar.AppendLine("    		CONVERT(NCHAR(10), Fec_Ini, 103)	AS	Fec_Ini, ")
            loComandoSeleccionar.AppendLine("	    	Cod_Cli, ")
            loComandoSeleccionar.AppendLine("	    	Nom_Cli, ")
            loComandoSeleccionar.AppendLine("	    	Cod_Art, ")
            loComandoSeleccionar.AppendLine("		    Modelo, ")
            loComandoSeleccionar.AppendLine("		    Can_Art1, ")
            loComandoSeleccionar.AppendLine("           Exi_Act1 ")
            loComandoSeleccionar.AppendLine(" FROM	    #tmpTemporal ")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_PFacturar", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_PFacturar.ReportSource = loObjetoReporte

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
' JJD: 01/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 12/10/09: Inclusion del filtro de la zona
'-------------------------------------------------------------------------------------------'
