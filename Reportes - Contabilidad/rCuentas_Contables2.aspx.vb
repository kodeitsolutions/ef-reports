'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCuentas_Contables2"
'-------------------------------------------------------------------------------------------'
Partial Class rCuentas_Contables2
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))


        Dim loComandoSeleccionar As New StringBuilder()
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Contables.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Status, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Movimiento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Cod_Cen, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Cod_Gas, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Categoria, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Contables.Movimiento = 1 ")
            loComandoSeleccionar.AppendLine("		        THEN LOWER(Cuentas_Contables.Nom_Cue) ")
            loComandoSeleccionar.AppendLine("			    ELSE UPPER(Cuentas_Contables.Nom_Cue) ")
            loComandoSeleccionar.AppendLine("			END) AS Nombre, ")
            loComandoSeleccionar.AppendLine("			(CASE Cuentas_Contables.Status ")
            loComandoSeleccionar.AppendLine("				WHEN 'A' THEN	'Activo' ")
            loComandoSeleccionar.AppendLine("				WHEN 'S' THEN	'Suspendido' ")
            loComandoSeleccionar.AppendLine("				ELSE			'Inactivo' ")
            loComandoSeleccionar.AppendLine("			END) AS Status_Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Contables.Cod_Cue BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Contables.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		    AND Cuentas_Contables.Cod_Cen BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		    AND Cuentas_Contables.Cod_Gas BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Contables.Categoria IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY	Cuentas_Contables." & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCuentas_Contables2", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCuentas_Contables2.ReportSource = loObjetoReporte


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
' MJP: 15/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 05/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------' 
' MAT: 16/05/11: Mejora de la vista de diseño, Ajuste del Select
'-------------------------------------------------------------------------------------------'
' PMV: 17/06/15: Creacion del Reporte Con Salida de Datos en MAYUSCULAS y minusculas
'-------------------------------------------------------------------------------------------'
