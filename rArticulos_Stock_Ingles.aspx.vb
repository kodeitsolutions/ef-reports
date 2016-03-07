'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_Stock_Ingles"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_Stock_Ingles
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)

            Dim lcExisiencia As String = ""

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("    		Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("    		Articulos.Cod_Uni1, ")
            'loComandoSeleccionar.AppendLine("    		Articulos.Exi_Act1, ")

            Select Case lcParametro10Desde
                Case "Actual"
                    loComandoSeleccionar.AppendLine("	Articulos.Exi_Act1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Act1"
                Case "Comprometida"
                    loComandoSeleccionar.AppendLine("	Articulos.Exi_Ped1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Ped1"
                Case "Cotizada"
                    loComandoSeleccionar.AppendLine("	Articulos.Exi_Cot1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Cot1"
                Case "En_Produccion"
                    loComandoSeleccionar.AppendLine("	Articulos.Exi_Pro1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Pro1"
                Case "Por_Llegar"
                    loComandoSeleccionar.AppendLine("	Articulos.Exi_Por1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Por1"
                Case "Por_Despachar"
                    loComandoSeleccionar.AppendLine("	Articulos.Exi_Des1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Des1"
                Case "Por_Distribuir"
                    loComandoSeleccionar.AppendLine("	Articulos.Exi_Dis1 AS Exi_Act1,")
                    lcExisiencia = "Articulos.Exi_Dis1"
            End Select

            loComandoSeleccionar.AppendLine("    		Articulos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine("    		Articulos.Cod_Sec, ")
            loComandoSeleccionar.AppendLine("    		Articulos.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("    		Articulos.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("    		Articulos.Cod_Mar ")
            loComandoSeleccionar.AppendLine("FROM		Articulos")
            loComandoSeleccionar.AppendLine("WHERE		Articulos.Cod_Art       BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Dep       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Sec       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Mar       BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Tip       BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Cla       BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			AND Articulos.Cod_Ubi		BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("      		And Articulos.Cod_Pro between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			And " & lcParametro8Hasta)

            Select Case lcParametro9Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("      ")
                Case "Igual"
                    loComandoSeleccionar.AppendLine("     AND Articulos.Exi_Act1          =   0  ")
                Case "Mayor"
                    loComandoSeleccionar.AppendLine("     AND " & lcExisiencia & "          >   0  ")
                Case "Menor"
                    loComandoSeleccionar.AppendLine("     AND Articulos.Exi_Act1          <   0  ")
                Case "Maximo"
                    loComandoSeleccionar.AppendLine("     AND Articulos.Exi_Max           =   " & lcExisiencia & "  ")
                Case "Minimo"
                    loComandoSeleccionar.AppendLine("     And Articulos.Exi_Min           =   " & lcExisiencia & "  ")
                Case "Pedido"
                    loComandoSeleccionar.AppendLine("     And Articulos.Exi_pto           =   " & lcExisiencia & "  ")
            End Select

            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_Stock_Ingles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_Stock_Ingles.ReportSource = loObjetoReporte

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message & loExcepcion.StackTrace, _
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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' Douglas Cortez: 20/05/2010: Codigo inicial (Copia de rArticulos_Stock)
'-------------------------------------------------------------------------------------------'
