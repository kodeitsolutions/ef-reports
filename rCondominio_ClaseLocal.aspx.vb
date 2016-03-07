'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCondominio_ClaseLocal"
'-------------------------------------------------------------------------------------------'
Partial Class rCondominio_ClaseLocal
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		Try
	
		
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @lnMontoPropietarios DECIMAL(28,10);")
            loConsulta.AppendLine("DECLARE @lnMontoOpcionVenta DECIMAL(28,10);")
            loConsulta.AppendLine("SET @lnMontoPropietarios = " & lcParametro4Desde & ";")
            loConsulta.AppendLine("SET @lnMontoOpcionVenta = " & lcParametro5Desde & ";")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  CxC.documento               AS documento,")
            loConsulta.AppendLine("        CxC.Cod_Cli                 AS Cod_Cli,")
            loConsulta.AppendLine("        C.nom_cli                   AS nom_cli,")
            loConsulta.AppendLine("        CxC.mon_bru                 AS mon_bru,")
            loConsulta.AppendLine("        CxC.mon_imp1                AS mon_imp1,")
            loConsulta.AppendLine("        CxC.mon_net                 AS mon_net,")
            loConsulta.AppendLine("        A.Cod_Cla                   AS Cod_Cla,")
            loConsulta.AppendLine("        CA.nom_cla                  AS nom_cla,")
            loConsulta.AppendLine("        CxC.comentario              AS comentario,")
            loConsulta.AppendLine("        A.Cod_Art                   AS Local,")
            loConsulta.AppendLine("        A.Precio1                   AS Canon,")
            loConsulta.AppendLine("        ROUND(A.precio1*.25,2)      AS Canon_25,")
            loConsulta.AppendLine("        A.Por_Ali                   AS Alicuota, ")
            loConsulta.AppendLine("        ROUND(A.por_ali*(")
            loConsulta.AppendLine("            CASE A.Cod_Cla ")
            loConsulta.AppendLine("            WHEN 'PROP' THEN @lnMontoPropietarios ")
            loConsulta.AppendLine("            WHEN 'OPVE' THEN @lnMontoOpcionVenta ")
            loConsulta.AppendLine("            ELSE 0 END)/100,2)      AS Condominio_Alicuota")
            loConsulta.AppendLine("FROM    cuentas_cobrar AS CxC")
            loConsulta.AppendLine("    JOIN clientes AS C ON C.Cod_Cli = CxC.Cod_Cli")
            loConsulta.AppendLine("    JOIN precios_clientes PC ON PC.cod_reg = CxC.Cod_Cli and PC.origen = 'clientes'")
            loConsulta.AppendLine("    JOIN articulos A ON A.cod_art = PC.Cod_Art ")
            loConsulta.AppendLine("    JOIN clases_articulos CA ON CA.Cod_Cla = A.Cod_Cla ")
            loConsulta.AppendLine("WHERE   CxC.caracter3 = 'CONDOMINIO'")
            loConsulta.AppendLine("    AND CxC.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("    AND CxC.fec_ini BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loConsulta.AppendLine("    AND CxC.cod_cli BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loConsulta.AppendLine("    AND A.Cod_Cla BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("    AND RTRIM(CAST(CxC.Comentario AS VARCHAR(MAX))) LIKE  '%' +RTRIM(A.cod_art)")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loConsulta.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCondominio_ClaseLocal", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
	
            Me.crvrCondominio_ClaseLocal.ReportSource = loObjetoReporte	


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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 19/02/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
