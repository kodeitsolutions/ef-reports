'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rParametros_Comisiones"
'-------------------------------------------------------------------------------------------'
Partial Class rParametros_Comisiones
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.cod_par,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.nom_par,     ")
            loComandoSeleccionar.AppendLine("		CASE     ")
            loComandoSeleccionar.AppendLine("			WHEN Parametros_Comisiones.status = 'A' THEN 'Activo'     ")
            loComandoSeleccionar.AppendLine("			WHEN Parametros_Comisiones.status = 'I' THEN 'Inactivo'     ")
            loComandoSeleccionar.AppendLine("			WHEN Parametros_Comisiones.status = 'S' THEN 'Suspendido'     ")
            loComandoSeleccionar.AppendLine("		END AS Status,       ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.cod_ven,     ")
            loComandoSeleccionar.AppendLine("		Vendedores.Nom_ven,                ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.dia_des,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.dia_has,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.mon_des,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.mon_has,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.can_des,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.can_has,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.tip_com,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.por_com1,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.por_com2,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.por_com3,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.mon_com1,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.mon_com2,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.mon_com3,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.por_otr1,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.por_otr2,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.por_otr3,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.mon_otr1,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.mon_otr2,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.mon_otr3,    ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.logico1,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.logico2,     ")
            loComandoSeleccionar.AppendLine("		Parametros_Comisiones.logico3      ")
            loComandoSeleccionar.AppendLine("FROM	Parametros_Comisiones    ")
            loComandoSeleccionar.AppendLine("JOIN Vendedores ON Vendedores.Cod_Ven = Parametros_Comisiones.Cod_ven")
            loComandoSeleccionar.AppendLine("WHERE   ")
            loComandoSeleccionar.AppendLine(" 			 Parametros_Comisiones.Cod_Par between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			 AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			 AND Parametros_Comisiones.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY    " & lcOrdenamiento)
            

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rParametros_Comisiones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrParametros_Comisiones.ReportSource = loObjetoReporte


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
' CMS:  07/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
