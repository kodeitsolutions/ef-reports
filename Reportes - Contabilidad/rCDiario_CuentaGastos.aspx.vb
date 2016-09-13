'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCDiario_CuentaGastos"
'-------------------------------------------------------------------------------------------'
Partial Class rCDiario_CuentaGastos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 			Renglones_Comprobantes.Mon_Deb,")
            loComandoSeleccionar.AppendLine(" 			Renglones_Comprobantes.Mon_Hab,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Gastos.Cod_Gas,")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Gastos.Nom_Gas")
			loComandoSeleccionar.AppendLine(" INTO #Temp")
            loComandoSeleccionar.AppendLine(" FROM Renglones_Comprobantes")
            loComandoSeleccionar.AppendLine(" JOIN Comprobantes ON Comprobantes.Documento = Renglones_Comprobantes.Documento")
            loComandoSeleccionar.AppendLine(" JOIN Cuentas_Gastos ON Cuentas_Gastos.Cod_Gas = Renglones_Comprobantes.Cod_Gas")
            loComandoSeleccionar.AppendLine(" WHERE     Comprobantes.Documento                      =   Renglones_Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("           And Comprobantes.Documento                  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Comprobantes.Fec_Ini                    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_Comprobantes.Cod_Mon          Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And YEAR(Renglones_Comprobantes.Fec_Ini)    =   YEAR(Comprobantes.Fec_Ini) ")
            loComandoSeleccionar.AppendLine("           And MONTH(Renglones_Comprobantes.Fec_Ini)   =   MONTH(Comprobantes.Fec_Ini) ")
            
			loComandoSeleccionar.AppendLine(" SELECT ")
			loComandoSeleccionar.AppendLine(" 		Sum(Mon_Deb) AS Mon_Deb,")
			loComandoSeleccionar.AppendLine(" 		Sum(Mon_Hab) AS Mon_Hab,")
			loComandoSeleccionar.AppendLine(" 		Nom_Gas,")
			loComandoSeleccionar.AppendLine(" 		Cod_Gas")
			loComandoSeleccionar.AppendLine(" FROM #Temp")
			loComandoSeleccionar.AppendLine(" GROUP BY Cod_Gas, Nom_Gas")
            loComandoSeleccionar.AppendLine("ORDER BY  " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCDiario_CuentaGastos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCDiario_CuentaGastos.ReportSource = loObjetoReporte

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
' CMS:  16/09/09: Codigo inicial
'-------------------------------------------------------------------------------------------'