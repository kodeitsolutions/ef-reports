Imports System.Data
Partial Class rPresupuestos_Renglones
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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            '-------------------------------------------------------------------------------------------'
            ' 1 - Select Principal
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT	Presupuestos.Documento, ")
            loComandoSeleccionar.AppendLine("			Presupuestos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Presupuestos.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Presupuestos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("			Presupuestos.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("			Presupuestos.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("			Presupuestos.Status,  ")
            loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Renglon, ")
            loComandoSeleccionar.AppendLine("			Presupuestos.Cod_For, ")
            loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Precio1, ")
            'loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Can_Art1, ")

            Select Case lcParametro5Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("             Renglones_Presupuestos.Can_Art1, ")
                Case "Backorder"
                    loComandoSeleccionar.AppendLine("             Renglones_Presupuestos.Can_Pen1 AS Can_Art1, ")
                Case "Procesado"
                    loComandoSeleccionar.AppendLine("             (Renglones_Presupuestos.Can_Art1 - Renglones_Presupuestos.Can_Pen1) AS Can_Art1, ")
            End Select

            loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Por_Des, ")
            loComandoSeleccionar.AppendLine("			Renglones_Presupuestos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven, ")
            loComandoSeleccionar.AppendLine("			Transportes.Nom_Tra, ")
            loComandoSeleccionar.AppendLine("			Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Proveedores.Nom_Pro,1,60) AS  Nom_Pro ")
            loComandoSeleccionar.AppendLine(" FROM		Presupuestos, ")
            loComandoSeleccionar.AppendLine("			Renglones_Presupuestos, ")
            loComandoSeleccionar.AppendLine("			Proveedores, ")
            loComandoSeleccionar.AppendLine("			Articulos, ")
            loComandoSeleccionar.AppendLine("			Vendedores, ")
            loComandoSeleccionar.AppendLine("			Formas_Pagos, ")
            loComandoSeleccionar.AppendLine("			Transportes ")
            loComandoSeleccionar.AppendLine(" WHERE		Presupuestos.Documento				=	Renglones_Presupuestos.Documento ")
            loComandoSeleccionar.AppendLine("			And Presupuestos.Cod_Pro			=	Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("			And Renglones_Presupuestos.Cod_Art	=	Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("			And Presupuestos.Cod_Ven			=	Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("			And Presupuestos.Cod_For			=	Formas_Pagos.Cod_For ")
            loComandoSeleccionar.AppendLine("			And Presupuestos.Cod_Tra			=	Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine("			And Presupuestos.Documento			Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Presupuestos.Fec_Ini			Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Proveedores.Cod_Pro				Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Presupuestos.Status				IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("			And Presupuestos.Cod_rev				Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            
            'loComandoSeleccionar.AppendLine(" ORDER BY	Presupuestos.Documento, Renglones_Presupuestos.Renglon ")


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPresupuestos_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPresupuestos_Renglones.ReportSource = loObjetoReporte

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
' JJD: 01/11/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' TJP: 14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' CMS: 22/07/09: Filtro BackOrder, lo conllevo al anexo del campo Can_Pen1,
'                 verificacion de registros
'-------------------------------------------------------------------------------------------'