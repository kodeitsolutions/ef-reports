'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_sAlmacen"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_sAlmacen 
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
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)
            Dim lcExistencia As String = ""

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT  Articulos.Cod_Art, " )
			loConsulta.AppendLine("        Articulos.Nom_Art, " )
			loConsulta.AppendLine("        Articulos.Cod_Uni1, " )

			Select Case lcParametro11Desde
                Case "Actual"
                    loConsulta.AppendLine("        Renglones_Almacenes.Exi_Act1     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Act1"
                Case "Comprometida"
                    loConsulta.AppendLine("        Renglones_Almacenes.Exi_Ped1     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Ped1"
                Case "Cotizada"
                    loConsulta.AppendLine("        Renglones_Almacenes.Exi_Cot1     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Cot1"
                Case "En_Produccion"
                    loConsulta.AppendLine("        Renglones_Almacenes.Exi_Pro1     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Pro1"
                Case "Por_Llegar"
                    loConsulta.AppendLine("        Renglones_Almacenes.Exi_Por1     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Por1"
                Case "Por_Despachar"
                    loConsulta.AppendLine("        Renglones_Almacenes.Exi_Des1     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Des1"
                Case "Por_Distribuir"
                    loConsulta.AppendLine("        Renglones_Almacenes.Exi_Dis1     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Dis1"
                Case Else
                    loConsulta.AppendLine("        0")
            End Select					 			
			
			loConsulta.AppendLine("        Renglones_Almacenes.Cod_Alm, " )
			loConsulta.AppendLine("        Articulos.Cod_Dep, " )
			loConsulta.AppendLine("        Articulos.Cod_Sec, " )
			loConsulta.AppendLine("        Articulos.Cod_Tip, " )
			loConsulta.AppendLine("        Articulos.Cod_Cla, " )
            loConsulta.AppendLine("        Articulos.Cod_Mar, ")
            loConsulta.AppendLine("        Almacenes.Nom_Alm  ")
			loConsulta.AppendLine("FROM    Articulos")
            loConsulta.AppendLine("    JOIN Renglones_Almacenes ")
            loConsulta.AppendLine("        ON Renglones_Almacenes.Cod_Art = Articulos.Cod_Art")
			loConsulta.AppendLine("    JOIN Almacenes  ")
            loConsulta.AppendLine("        ON Almacenes.Cod_Alm = Renglones_Almacenes.Cod_Alm")
			loConsulta.AppendLine("WHERE		Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
			loConsulta.AppendLine(" 		AND " & lcParametro0Hasta)
			loConsulta.AppendLine(" 		AND Articulos.status IN (" & lcParametro1Desde & ")")
			loConsulta.AppendLine(" 		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde)
			loConsulta.AppendLine(" 		AND " & lcParametro2Hasta)
			loConsulta.AppendLine(" 		AND Articulos.Cod_Dep BETWEEN " & lcParametro3Desde)
			loConsulta.AppendLine(" 		AND " & lcParametro3Hasta)
			loConsulta.AppendLine(" 		AND Articulos.Cod_Sec BETWEEN " & lcParametro4Desde)
			loConsulta.AppendLine(" 		AND " & lcParametro4Hasta)
			loConsulta.AppendLine(" 		AND Articulos.Cod_Mar BETWEEN " & lcParametro5Desde)
			loConsulta.AppendLine(" 		AND " & lcParametro5Hasta)
			loConsulta.AppendLine(" 		AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
			loConsulta.AppendLine(" 		AND " & lcParametro6Hasta)
			loConsulta.AppendLine(" 		AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro7Hasta)
            loConsulta.AppendLine("      	AND Articulos.Cod_Ubi between " & lcParametro8Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro8Hasta)             
			loConsulta.AppendLine("      	AND Articulos.Cod_Pro between " & lcParametro9Desde)
            loConsulta.AppendLine(" 		AND " & lcParametro9Hasta)
            
            Select Case lcParametro10Desde
                Case "Todos"
                    loConsulta.AppendLine("      ")
                Case "Igual"
                    loConsulta.AppendLine("     AND " & lcExistencia & "          =   0  ")
                Case "Mayor"							  
                    loConsulta.AppendLine("     AND " & lcExistencia & "          >   0  ")
                Case "Menor"							  
                    loConsulta.AppendLine("     AND " & lcExistencia & "          <   0  ")
                Case "Maximo"							  
                    loConsulta.AppendLine("     AND Articulos.Exi_Max           =   " & lcExistencia & "  ")
                Case "Minimo"							  
                    loConsulta.AppendLine("     AND Articulos.Exi_Min           =   " & lcExistencia & "  ")
                Case "Pedido"
                    loConsulta.AppendLine("     AND Articulos.Exi_pto           =   " & lcExistencia & "  ")
            End Select
            
            loConsulta.AppendLine("ORDER BY   Articulos.Cod_Art, " & lcOrdenamiento)
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

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

			loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rArticulos_sAlmacen", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrArticulos_sAlmacen.ReportSource =   loObjetoReporte	



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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' MVP: 09/07/08 : Codigo inicial.                                                           '
'-------------------------------------------------------------------------------------------'
' MVP: 11/07/08 : Adición de loObjetoReporte para eliminar los archivos temp en Uranus.     '
'-------------------------------------------------------------------------------------------'
' MVP: 01/08/08: Cambios para multi idioma, mensaje de error y clase padre.                 '
'-------------------------------------------------------------------------------------------'
' CMS: 05/05/09: Ordenamiento.                                                              '
'-------------------------------------------------------------------------------------------'
' CMS: 11/08/09: Verificacion de registros y se agregaro el filtro_ Ubicación.              '
'-------------------------------------------------------------------------------------------'
' CMS: 21/09/09: Se agregaro los filtros tipo de existencia, nivel de existencia y proveedor
'-------------------------------------------------------------------------------------------'
' CMS: 21/09/09: Se cambio el origen de las existencias de articulos a renglones.           '
'-------------------------------------------------------------------------------------------'
' RJG: 25/09/13: Se ajustó el SELECT: se eliminaron tablas redundantes.                     '
'-------------------------------------------------------------------------------------------'
