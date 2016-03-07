'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOPagos_Fechas_Resumido"
'-------------------------------------------------------------------------------------------'
Partial Class rOPagos_Fechas_Resumido

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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = cusAplicacion.goReportes.paParametrosFinales(8)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Ordenes_Pagos.Documento					AS Documento,")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Status					AS Status,")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Fec_Ini					AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Cod_Pro					AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro						AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Net					AS Mon_Net,")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Motivo					AS Motivo,")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Comentario				AS Comentario,")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Cod_Mon					AS Cod_Mon,")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Tasa						AS Tasa,")
            loComandoSeleccionar.AppendLine("			Detalles_oPagos.Cod_Cue					AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("			ISNULL(Cajas.Nom_Caj,'')				AS Cod_Caj,")
            loComandoSeleccionar.AppendLine("			Detalles_oPagos.Num_Doc					AS Num_Doc,")
            loComandoSeleccionar.AppendLine("			Detalles_oPagos.Mon_Net					AS Mon_Det,")
            loComandoSeleccionar.AppendLine("			Detalles_oPagos.Renglon					AS Renglon_Det")
            loComandoSeleccionar.AppendLine("FROM		Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine("   JOIN	Detalles_oPagos ON Detalles_oPagos.Documento = Ordenes_Pagos.Documento ")
            loComandoSeleccionar.AppendLine("   JOIN	Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN	Cajas ON Cajas.Cod_Caj = Detalles_oPagos.Cod_Caj ")
            loComandoSeleccionar.AppendLine("   LEFT JOIN	Cuentas_Bancarias ON Cuentas_Bancarias.Cod_Cue = Detalles_oPagos.Cod_Cue ")
            loComandoSeleccionar.AppendLine("WHERE     ")
            loComandoSeleccionar.AppendLine("           Ordenes_Pagos.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Fec_Ini   BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Cod_Pro   BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Status    IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Cod_Mon    BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("           AND Detalles_oPagos.Cod_Cue    BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Cod_Suc    BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            
            If lcParametro8Desde = "Igual" Then
                loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Cod_rev   BETWEEN " & lcParametro7Desde)
            Else
                loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Cod_rev  NOT BETWEEN " & lcParametro7Desde)
            End If
            
            loComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
                 
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento & ", Ordenes_Pagos.Documento ASC")

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

           'me.mEscribirConsulta(loComandoSeleccionar.ToString)
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_Fechas_Resumido", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_Fechas_Resumido.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' JJD: 24/01/09: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' GCR: 27/03/09: Ajustes al diseño															'
'-------------------------------------------------------------------------------------------'
' YJP: 14/05/09: Agregar filtro Revisión													'
'-------------------------------------------------------------------------------------------'
' AAP:  01/07/09: Filtro "Sucursal:"														'
'-------------------------------------------------------------------------------------------'
' CMS:  04/07/09: Metodo de ordenamiento													'
'-------------------------------------------------------------------------------------------'
' CMS:  02/07/10: Filtro Cuenta																'
'-------------------------------------------------------------------------------------------'
' CMS:  03/07/10: Filtro Tipo de revision													'
'-------------------------------------------------------------------------------------------'
' MAT:  25/10/10: Mejora visual del reporte													'
'-------------------------------------------------------------------------------------------'
' RJG:  18/04/12: Se colocó el total de registros, y se corrigió el monto neto (lo tomaba	'
'				  del encabezado en lugar de tomarlo de los renglones.						'
'-------------------------------------------------------------------------------------------'
