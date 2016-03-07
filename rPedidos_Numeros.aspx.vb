'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPedidos_Numeros"
'-------------------------------------------------------------------------------------------'
Partial Class rPedidos_Numeros
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()


            loConsulta.AppendLine("SELECT		 Pedidos.Documento, ")
            loConsulta.AppendLine(" 			 Pedidos.Status, ")
            loConsulta.AppendLine(" 			 Pedidos.Fec_Ini, ")
            loConsulta.AppendLine(" 			 Pedidos.Fec_Fin, ")
            loConsulta.AppendLine(" 			 Pedidos.Cod_Cli, ")
            loConsulta.AppendLine(" 			 Clientes.Nom_Cli, ")
            loConsulta.AppendLine(" 			 Pedidos.Cod_Ven, ")
            loConsulta.AppendLine(" 			 Pedidos.Cod_Tra, ")
            loConsulta.AppendLine(" 			 Pedidos.Mon_Imp1, ")
            loConsulta.AppendLine(" 			 Pedidos.Control, ")
            loConsulta.AppendLine(" 			 Pedidos.Mon_Net, ")
            loConsulta.AppendLine(" 			 Pedidos.Mon_Bru  ")
            loConsulta.AppendLine("FROM		 Pedidos  ")
            loConsulta.AppendLine("    JOIN    Clientes ON  Clientes.Cod_Cli = Pedidos.Cod_Cli")
            loConsulta.AppendLine("WHERE		     Pedidos.Documento BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("            AND " & lcParametro0Hasta)
            loConsulta.AppendLine("            AND Pedidos.Fec_Ini BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("            AND " & lcParametro1Hasta)
            loConsulta.AppendLine("            AND Pedidos.Cod_Cli BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("            AND " & lcParametro2Hasta)
            loConsulta.AppendLine("            AND Pedidos.Cod_Ven BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("            AND " & lcParametro3Hasta)
            loConsulta.AppendLine("            AND Pedidos.Status IN (" & lcParametro4Desde & ")")
            loConsulta.AppendLine("            AND Pedidos.Cod_Tra BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("            AND " & lcParametro5Hasta)
            loConsulta.AppendLine("            AND Pedidos.Cod_Mon BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("            AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 		     AND Pedidos.Cod_mon BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine(" 		     AND " & lcParametro6Hasta)
            loConsulta.AppendLine("            AND Pedidos.Cod_Rev between " & lcParametro7Desde)
            loConsulta.AppendLine("    	     AND " & lcParametro7Hasta)
            loConsulta.AppendLine("            AND Pedidos.Cod_Suc between " & lcParametro8Desde)
            loConsulta.AppendLine("    	     AND " & lcParametro8Hasta)
            loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_Numeros", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_Numeros.ReportSource = loObjetoReporte


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
' JJD: 24/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.                 '
'-------------------------------------------------------------------------------------------'
' CMS: 14/05/09: Filtro “Revisión:” y se le agrego el filtro moneda.                        '
'-------------------------------------------------------------------------------------------'
' AAP: 30/06/09: Filtro “Sucursal:”.                                                        '
'-------------------------------------------------------------------------------------------'
' CMS: 10/08/09: Metodo de ordenamiento, verificacionde registros.                          '
'-------------------------------------------------------------------------------------------'
' RJG: 23/07/14: Ajuste en consulta (se eliminaron tablas innecesarias). Ajustes en interfaz'
'-------------------------------------------------------------------------------------------'