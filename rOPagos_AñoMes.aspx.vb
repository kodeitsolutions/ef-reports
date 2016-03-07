'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOPagos_AñoMes"
'-------------------------------------------------------------------------------------------'
Partial Class rOPagos_AñoMes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine("            DATEPART(YEAR, Ordenes_Pagos.fec_ini) as Anno,")
            loComandoSeleccionar.AppendLine("            DATEPART(MONTH, Ordenes_Pagos.fec_ini) as Mes,")
            loComandoSeleccionar.AppendLine("            (Ordenes_Pagos.Documento) as Cantidad,")
            loComandoSeleccionar.AppendLine("            (Ordenes_Pagos.Mon_Imp) as Mon_Imp,")
            loComandoSeleccionar.AppendLine("            (Renglones_OPagos.Mon_Deb) as Mon_Deb,")
            loComandoSeleccionar.AppendLine("            (Renglones_OPagos.Mon_Hab) as Mon_Hab")
            loComandoSeleccionar.AppendLine(" INTO       #tmp")
            loComandoSeleccionar.AppendLine(" FROM       Ordenes_Pagos, Renglones_OPagos")
            loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Pagos.Documento     =   Renglones_OPagos.Documento ")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Fec_Ini Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Pro   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Status    IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Mon   Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_OPagos.Cod_Con   Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("SELECT ")
            loComandoSeleccionar.AppendLine("            Anno,")
            loComandoSeleccionar.AppendLine("            Mes,")
            loComandoSeleccionar.AppendLine("            COUNT(Cantidad) AS Cantidad,")
            loComandoSeleccionar.AppendLine("            SUM(Mon_Imp) AS Mon_Imp,")
            loComandoSeleccionar.AppendLine("            SUM (Mon_Deb) AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("            SUM(Mon_Hab) AS Mon_Hab")
            loComandoSeleccionar.AppendLine("FROM #tmp")
            loComandoSeleccionar.AppendLine("GROUP BY Anno,mes")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_AñoMes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_AñoMes.ReportSource = loObjetoReporte

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
' CMS: 16/05/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' MAT:  06/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'