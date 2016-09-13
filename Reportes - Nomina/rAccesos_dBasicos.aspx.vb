'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAccesos_dBasicos"
'-------------------------------------------------------------------------------------------'
Partial Class rAccesos_dBasicos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))

            loComandoSeleccionar.AppendLine(" SELECT	Accesos.Documento, ")
            loComandoSeleccionar.AppendLine(" Accesos.Registro, ")
            loComandoSeleccionar.AppendLine(" Accesos.Cod_Tra, ")
            loComandoSeleccionar.AppendLine(" Accesos.Cod_Pue, ")
            loComandoSeleccionar.AppendLine(" Accesos.Cod_Tar, ")
            loComandoSeleccionar.AppendLine(" Accesos.Cod_Tur, ")
            loComandoSeleccionar.AppendLine(" Accesos.Tipo, ")
            loComandoSeleccionar.AppendLine(" Accesos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine(" Accesos.Cod_Tar, ")
            loComandoSeleccionar.AppendLine(" Trabajadores.Nom_Tra ")
            loComandoSeleccionar.AppendLine(" FROM	 Accesos, ")
            loComandoSeleccionar.AppendLine(" Trabajadores ")
            loComandoSeleccionar.AppendLine(" WHERE	Accesos.Cod_Tra = Trabajadores.Cod_Tra ")
            loComandoSeleccionar.AppendLine(" And Accesos.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" And Accesos.Registro between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" And Trabajadores.Cod_Tra between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY Accesos.Documento")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAccesos_dBasicos", laDatosReporte)

            Me.crvrAccesos_dBasicos.ReportSource = loObjetoReporte


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
' MVP:  14/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP:  24/07/08: Se suprimión el filtro de la puertas puesto que inicialmente no se trabaja_
'                 con este dato necesario al ingresar o salir de la empresa.
'-------------------------------------------------------------------------------------------'
