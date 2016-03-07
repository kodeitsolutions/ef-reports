'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rProveedores_Paises"
'-------------------------------------------------------------------------------------------'
Partial Class rProveedores_Paises 
    Inherits vis2Formularios.frmReporte
    
	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
		    Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()


			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT      Paises.Cod_Pai,")
			loConsulta.AppendLine("            Paises.Nom_Pai,")
			loConsulta.AppendLine("            Proveedores.Cod_Pro, ")
			loConsulta.AppendLine("            Proveedores.Nom_Pro,")
			loConsulta.AppendLine("            REPLACE(REPLACE(Proveedores.Dir_Fis, CHAR(10), ' '), CHAR(13), ' ') Dir_Fis,")
			loConsulta.AppendLine("            Proveedores.Web,")
			loConsulta.AppendLine("            Proveedores.Contacto,")
			loConsulta.AppendLine("            Proveedores.Correo,")
			loConsulta.AppendLine("            Proveedores.Telefonos,")
			loConsulta.AppendLine("            Proveedores.Movil,")
			loConsulta.AppendLine("            Proveedores.Mon_Sal,")
			loConsulta.AppendLine("            Vendedores.Cod_Ven,")
			loConsulta.AppendLine("            Vendedores.Nom_Ven,")
			loConsulta.AppendLine("            Estados.Cod_Est,")
			loConsulta.AppendLine("            Estados.Nom_Est,")
			loConsulta.AppendLine("            Tipos_Proveedores.Cod_Tip,")
			loConsulta.AppendLine("            Tipos_Proveedores.Nom_Tip,")
			loConsulta.AppendLine("            Clases_Proveedores.Cod_Cla,")
			loConsulta.AppendLine("            Clases_Proveedores.Nom_Cla,")
			loConsulta.AppendLine("            RTRIM(Tipos_Proveedores.Nom_Tip) + ")
			loConsulta.AppendLine("                '/' + RTRIM(Clases_Proveedores.Nom_Cla) AS Tipo_y_Clase")
			loConsulta.AppendLine("FROM        Proveedores")
			loConsulta.AppendLine("    JOIN    Paises ON Paises.Cod_Pai = Proveedores.Cod_Pai")
			loConsulta.AppendLine("    JOIN    Estados ON Estados.Cod_Est = Proveedores.Cod_Est")
			loConsulta.AppendLine("    JOIN    Vendedores ON Vendedores.Cod_Ven = Proveedores.Cod_Ven")
			loConsulta.AppendLine("    JOIN    Tipos_Proveedores ON Tipos_Proveedores.Cod_Tip = Proveedores.Cod_Tip")
			loConsulta.AppendLine("    JOIN    Clases_Proveedores ON Clases_Proveedores.Cod_Cla = Proveedores.Cod_Cla")
			loConsulta.AppendLine("WHERE	    Proveedores.Cod_Pro BETWEEN " & lcParametro0Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro0Hasta)
			loConsulta.AppendLine(" 	    AND Proveedores.Status IN (" & lcParametro1Desde &  ")")
			loConsulta.AppendLine(" 	    AND Vendedores.Cod_Ven BETWEEN " & lcParametro2Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro2Hasta)
			loConsulta.AppendLine(" 	    AND Proveedores.Cod_Zon BETWEEN " & lcParametro3Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro3Hasta)
			loConsulta.AppendLine(" 	    AND Paises.Cod_Pai BETWEEN " & lcParametro4Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro4Hasta)
			loConsulta.AppendLine(" 	    AND Proveedores.Cod_Suc BETWEEN " & lcParametro5Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro5Hasta)
			loConsulta.AppendLine(" 	    AND Clases_Proveedores.Cod_Cla BETWEEN " & lcParametro6Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro6Hasta)
			loConsulta.AppendLine(" 	    AND Tipos_Proveedores.Cod_Tip BETWEEN " & lcParametro7Desde)
			loConsulta.AppendLine(" 	    AND " & lcParametro7Hasta)
			loConsulta.AppendLine("ORDER BY     Paises.Nom_Pai ASC, " & lcOrdenamiento)
			loConsulta.AppendLine("")
     
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProveedores_Paises", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProveedores_Paises.ReportSource = loObjetoReporte
            

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
' RJG: 07/05/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
