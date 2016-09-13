'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCContables_Presupuestos_CentrosCostos"
'-------------------------------------------------------------------------------------------'
Partial Class rCContables_Presupuestos_CentrosCostos
     Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		
		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
		Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
		Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
		Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
		Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
		Dim lcParametro4Desde As String = cusAplicacion.goReportes.paParametrosIniciales(4)
		Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)
	    
        If (lcParametro0Desde = "'0'") Then 
            lcParametro0Desde = goServicios.mObtenerCampoFormatoSQL(Date.Now.Year)
        End If

		Dim loConsulta As New StringBuilder()
		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
        
        Dim llSoloCuentasDeMovimiento As Boolean 
        Dim llSoloCuentasConPresupuestoAsignado As Boolean 

        llSoloCuentasDeMovimiento = (lcParametro4Desde.Trim().ToUpper() = "SI")
        llSoloCuentasConPresupuestoAsignado = (lcParametro5Desde.Trim().ToUpper() = "SI")

		Try	
			
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT      Cuentas_Contables.Cod_Cue       AS Cod_Cue,")
			loConsulta.AppendLine("            Cuentas_Contables.Nom_Cue       AS Nom_Cue,")
			loConsulta.AppendLine("            Cuentas_Contables.Status        AS Status,")
			loConsulta.AppendLine("            (CASE Cuentas_Contables.Status   ")
			loConsulta.AppendLine("                WHEN 'A' THEN 'Activo'")
			loConsulta.AppendLine("                WHEN 'S' THEN 'Suspendido'")
			loConsulta.AppendLine("                ELSE 'Inactivo'")
			loConsulta.AppendLine("            END)                            AS Estatus, ")
			loConsulta.AppendLine("            Metas.Adicional                 AS Cod_Cen,")
			loConsulta.AppendLine("            Centros_Costos.Nom_Cen          AS Nom_Cen,")
			loConsulta.AppendLine("            Metas.Mes                       AS Mes,")
			loConsulta.AppendLine("            Metas.Año                       AS Año,")
			loConsulta.AppendLine("            Metas.Monto1                    AS Presupuesto, ")
			loConsulta.AppendLine("            Metas.Monto2                    AS Ejecutado, ")
			loConsulta.AppendLine("            Metas.Monto3                    AS Saldo")
			loConsulta.AppendLine("FROM        Cuentas_Contables")
            If llSoloCuentasConPresupuestoAsignado Then
			    loConsulta.AppendLine("    JOIN    Metas ")
            Else
			    loConsulta.AppendLine("    LEFT JOIN Metas ")
            End If
			loConsulta.AppendLine("        ON  Metas.Cod_Reg = Cuentas_Contables.Cod_Cue")
			loConsulta.AppendLine("        AND Metas.Origen = 'Cuentas_Contables'")
			loConsulta.AppendLine("        AND Metas.Clase = 'Presupuesto'")
			loConsulta.AppendLine("		   AND Metas.Tip_Met = 'Cuentas_Contables_Centros_Costos_Mensualmente'")
			loConsulta.AppendLine("        AND Metas.Año = " & lcParametro0Desde)
            If llSoloCuentasConPresupuestoAsignado Then
			    loConsulta.AppendLine("    JOIN    Centros_Costos ")
            Else
			    loConsulta.AppendLine("    LEFT JOIN Centros_Costos ")
            End If
			loConsulta.AppendLine("        ON  Centros_Costos.Cod_Cen = Metas.Adicional")
			loConsulta.AppendLine("        AND Centros_Costos.Cod_Cen BETWEEN " & lcParametro3Desde)
			loConsulta.AppendLine("		   AND " & lcParametro3Hasta)
			loConsulta.AppendLine("WHERE		Cuentas_Contables.Cod_Cue BETWEEN " & lcParametro1Desde)
			loConsulta.AppendLine("			AND " & lcParametro1Hasta)
			loConsulta.AppendLine("			AND Cuentas_Contables.Status IN (" & lcParametro2Desde & ")")
            IF llSoloCuentasDeMovimiento Then
			    loConsulta.AppendLine("			AND Cuentas_Contables.Movimiento = 1")
            End If
			loConsulta.AppendLine("ORDER BY    " & lcOrdenamiento & ", Metas.Adicional, Metas.Año, Metas.Mes")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			
            'Me.mEscribirConsulta(loConsulta.ToString())

			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCContables_Presupuestos_CentrosCostos", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrCContables_Presupuestos_CentrosCostos.ReportSource = loObjetoReporte
	   

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
' RJG: 10/01/15: Codigo inicial.
'-------------------------------------------------------------------------------------------'
