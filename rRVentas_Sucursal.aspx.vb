'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRVentas_Sucursal"
'-------------------------------------------------------------------------------------------'
Partial Class rRVentas_Sucursal
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

             Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            
             loComandoSeleccionar.AppendLine("   SELECT ")
            loComandoSeleccionar.AppendLine("           Cuentas_Cobrar.Cod_Suc,")
            loComandoSeleccionar.AppendLine("           Sucursales.Nom_Suc,")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(Cuentas_Cobrar.Mon_Bru), 0) as Mon_Bru,")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(Cuentas_Cobrar.Mon_Imp1), 0) AS Mon_Imp1,")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(Cuentas_Cobrar.Mon_Bru)/30, 0)AS Mon_Bru_Prom,")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(Cuentas_Cobrar.Mon_Imp1)/30, 0) AS Mon_Imp1_Prom")
            loComandoSeleccionar.AppendLine("   INTO	#tmpTemporal1")
            loComandoSeleccionar.AppendLine("   FROM	Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("	JOIN	Sucursales ON Sucursales.Cod_Suc = Cuentas_Cobrar.Cod_Suc ")
            loComandoSeleccionar.AppendLine("   WHERE	Cuentas_Cobrar.Cod_Tip = 'Fact' ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Status <> 'Anulado' ")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Suc BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_rev BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("   GROUP BY Cuentas_Cobrar.Cod_Suc,Sucursales.Nom_Suc")
			loComandoSeleccionar.AppendLine("     ")
			loComandoSeleccionar.AppendLine("     ")
			
			
            loComandoSeleccionar.AppendLine("   SELECT  ")
            loComandoSeleccionar.AppendLine("           Devoluciones_Clientes.Cod_Suc,")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(Devoluciones_Clientes.Mon_Bru), 0) as Mon_Dev, ")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(Devoluciones_Clientes.Mon_Imp1), 0) as Dev_Imp1,")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(Devoluciones_Clientes.Mon_Bru)/30, 0) AS Mon_Dev_Prom, ")
            loComandoSeleccionar.AppendLine("           ISNULL(sum(Devoluciones_Clientes.Mon_Imp1)/30, 0) AS Dev_Imp1_Prom")
            loComandoSeleccionar.AppendLine("   INTO	#tmpTemporal2 ")
            loComandoSeleccionar.AppendLine("   FROM	Devoluciones_Clientes  ")
            loComandoSeleccionar.AppendLine("	JOIN	Sucursales ON Sucursales.Cod_Suc = Devoluciones_Clientes.Cod_Suc ")
            loComandoSeleccionar.AppendLine("   WHERE	Devoluciones_Clientes.Status = 'Confirmado' ")
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Suc BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.cod_ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Devoluciones_Clientes.cod_mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("   GROUP BY Devoluciones_Clientes.Cod_Suc")
			loComandoSeleccionar.AppendLine("     ")
			loComandoSeleccionar.AppendLine("     ")

            loComandoSeleccionar.AppendLine("   SELECT  ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal1.Cod_Suc,")
            loComandoSeleccionar.AppendLine("           #tmpTemporal1.Nom_Suc,")
            loComandoSeleccionar.AppendLine("           #tmpTemporal1.Mon_Bru,")
            loComandoSeleccionar.AppendLine("           #tmpTemporal1.Mon_Imp1,")
            loComandoSeleccionar.AppendLine("           #tmpTemporal1.Mon_Bru_Prom, ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal1.Mon_Imp1_Prom,")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmpTemporal2.Mon_Dev,0) AS Mon_Dev, ")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmpTemporal2.Dev_Imp1,0) AS Dev_Imp1,")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmpTemporal2.Mon_Dev_Prom,0) AS Mon_Dev_Prom ,")
            loComandoSeleccionar.AppendLine("           ISNULL( #tmpTemporal2.Dev_Imp1_Prom,0) AS Dev_Imp1_Prom")
            loComandoSeleccionar.AppendLine("   FROM	#tmpTemporal1")
            loComandoSeleccionar.AppendLine("   LEFT JOIN #tmpTemporal2 ON #tmpTemporal1.cod_suc = #tmpTemporal2.cod_Suc")
		    loComandoSeleccionar.AppendLine("   ORDER BY " & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("     ")
			
			
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
           
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRVentas_Sucursal", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRVentas_Sucursal.ReportSource = loObjetoReporte

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
' CMS: 18/06/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' MAT: 17/05/11: Ajuste del Select, mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
