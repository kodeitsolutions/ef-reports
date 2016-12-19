'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rAsistencia_Trabajadores"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rAsistencia_Trabajadores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFechaDesde      AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaHasta      AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodTra_Desde	AS VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodTra_Hasta	AS VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Recibos.Cod_Tra						AS Cod_Tra,")
            loComandoSeleccionar.AppendLine("		Trabajadores.Nom_Tra				AS Nom_Tra,")
            loComandoSeleccionar.AppendLine("		Departamentos_Nomina.Nom_Dep		AS Departamento,")
            loComandoSeleccionar.AppendLine("		Cargos.Nom_Car						AS Cargo,")
            loComandoSeleccionar.AppendLine("		Contratos.Nom_Con					AS Contrato,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Recibos.Val_Num)		AS Total_Horas,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Recibos.Val_Num) / 8	AS Total_Dias, ")
            loComandoSeleccionar.AppendLine("       @ldFechaDesde                       AS Desde,")
            loComandoSeleccionar.AppendLine("       @ldFechaHasta                       AS Hasta")
            loComandoSeleccionar.AppendLine("FROM Renglones_Recibos ")
            loComandoSeleccionar.AppendLine("	JOIN Recibos ON Renglones_Recibos.Documento = Recibos.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Trabajadores ON Trabajadores.Cod_Tra = Recibos.Cod_Tra")
            loComandoSeleccionar.AppendLine("	JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
            loComandoSeleccionar.AppendLine("	JOIN Cargos ON Cargos.Cod_Car = Trabajadores.Cod_Car")
            loComandoSeleccionar.AppendLine("	JOIN Contratos ON Contratos.Cod_Con = Trabajadores.Cod_Con")
            loComandoSeleccionar.AppendLine("		AND Contratos.Cod_Con = Recibos.Cod_Con")
            loComandoSeleccionar.AppendLine("WHERE Renglones_Recibos.Cod_Con = 'E002'")
            loComandoSeleccionar.AppendLine("	AND Recibos.Fecha BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Recibos.Cod_Tra BETWEEN @lcCodTra_Desde AND @lcCodTra_Hasta")
            loComandoSeleccionar.AppendLine("GROUP BY Recibos.Cod_Tra, Trabajadores.Nom_Tra, Cargos.Nom_Car, Departamentos_Nomina.Nom_Dep, Contratos.Nom_Con")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rAsistencia_Trabajadores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rAsistencia_Trabajadores.ReportSource = loObjetoReporte

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
' DLC: 02/09/2010: Programacion inicial (Replica del reporte rLEstadoCuenta_HistoricoVentas)'
'                   - Cambio de la consulta a procedimiento almacenado.						'
'-------------------------------------------------------------------------------------------'
' DLC: 15/09/2010: Ajuste en la forma de obtener los detalles de Pagos, asi como también,	'
'                ajustar en el RPT, la forma de mostrar los detalles de Pagos.				'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Reprogramación del Reporte y su respectivo Store Procedure					'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Ajuste de la vista de Diseño.												'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Se elimino el filtro Detalle												'
'-------------------------------------------------------------------------------------------'
' RJG: 05/12/11: Eliminado el SP: ahora la consulta se hace desde un Query en línea para	'
'				 corregir cálculo de saldo y optimizar.										'
'-------------------------------------------------------------------------------------------'
