'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_sProveedor"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_sProveedor 
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

			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("SELECT	Articulos.Cod_Art		AS Cod_Art, " )
			loComandoSeleccionar.AppendLine("		Articulos.Nom_Art		AS Nom_Art, " )
			loComandoSeleccionar.AppendLine("		Articulos.Cod_Uni1		AS Cod_Uni1," )
			
			Select Case lcParametro11Desde
                Case "Actual"
                    loComandoSeleccionar.AppendLine("		SUM(Renglones_Almacenes.Exi_Act1) AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Act1"
                Case "Comprometida"
                    loComandoSeleccionar.AppendLine("		SUM(Renglones_Almacenes.Exi_Ped1) AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Ped1"
                Case "Cotizada"
                    loComandoSeleccionar.AppendLine("		SUM(Renglones_Almacenes.Exi_Cot1) AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Cot1"
                Case "En_Produccion"
                    loComandoSeleccionar.AppendLine("		SUM(Renglones_Almacenes.Exi_Pro1) AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Pro1"
                Case "Por_Llegar"
                    loComandoSeleccionar.AppendLine("		SUM(Renglones_Almacenes.Exi_Por1) AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Por1"
                Case "Por_Despachar"
                    loComandoSeleccionar.AppendLine("		SUM(Renglones_Almacenes.Exi_Des1) AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Des1"
                Case "Por_Distribuir"
                    loComandoSeleccionar.AppendLine("		SUM(Renglones_Almacenes.Exi_Dis1) AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Dis1"
            End Select					 			
			
			loComandoSeleccionar.AppendLine("		Proveedores.Cod_Pro		AS Cod_Pro, " )
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro		AS Nom_Pro")
			loComandoSeleccionar.AppendLine("FROM	Articulos" )
			loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Articulos.Cod_Pro" )
			loComandoSeleccionar.AppendLine("	JOIN Renglones_Almacenes ON Renglones_Almacenes.Cod_Art = Articulos.Cod_Art" )
			loComandoSeleccionar.AppendLine("WHERE		Articulos.Cod_Pro = Proveedores.Cod_Pro ")
			loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 		AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 		AND Articulos.status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" 		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 		AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Dep BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 		AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Sec BETWEEN " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 		AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Mar BETWEEN " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 		AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
			loComandoSeleccionar.AppendLine(" 		AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("      	AND Articulos.Cod_Ubi BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro8Hasta)             
			loComandoSeleccionar.AppendLine("      	AND Articulos.Cod_Pro BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro9Hasta)
            
			loComandoSeleccionar.AppendLine("GROUP BY	Articulos.Cod_Art, Articulos.Nom_Art, Articulos.Cod_Uni1,")
			loComandoSeleccionar.AppendLine("			Proveedores.Cod_Pro, Proveedores.Nom_Pro")
            
            Select Case lcParametro10Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("      ")
                Case "Igual"
                    loComandoSeleccionar.AppendLine("HAVING  SUM(" & lcExistencia & ")	=   0  ")
                Case "Mayor"							  
                    loComandoSeleccionar.AppendLine("HAVING  SUM(" & lcExistencia & ")	>   0  ")
                Case "Menor"							  
                    loComandoSeleccionar.AppendLine("HAVING  SUM(" & lcExistencia & ")	<   0  ")
                Case "Maximo"							  
                    loComandoSeleccionar.AppendLine("HAVING  MAX(Articulos.Exi_Max)		=   SUM(" & lcExistencia & ")  ")
                Case "Minimo"																
                    loComandoSeleccionar.AppendLine("HAVING  MAX(Articulos.Exi_Min)		=   SUM(" & lcExistencia & ")  ")
                Case "Pedido"																
                    loComandoSeleccionar.AppendLine("HAVING  MAX(Articulos.Exi_pto)		=   SUM(" & lcExistencia & ")  ")
            End Select
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento & ", Articulos.Cod_Art")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
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

			loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rArticulos_sProveedor", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrArticulos_sProveedor.ReportSource =   loObjetoReporte	



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
' RJG: 10/02/12: Codigo inicial
'-------------------------------------------------------------------------------------------'
