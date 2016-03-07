'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCPagar_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rCPagar_Fechas
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
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)
            Dim lcParametro12Desde As String = cusAplicacion.goReportes.paParametrosIniciales(12)
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.Appendline(" SELECT Cuentas_Pagar.Documento, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Cod_Tip, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Fec_Ini, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Fec_Fin, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Cod_Pro, "  )
            loComandoSeleccionar.AppendLine(" 			Proveedores.Nom_Pro,   ")
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Cod_Ven, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Cod_Tra, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Cod_Mon, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Control, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Comentario, "  )
			If lcParametro12Desde.ToString = "Si" Then
				loComandoSeleccionar.Appendline(" 		'Si' As cComentario, "  )
			Else
				loComandoSeleccionar.Appendline(" 		'No' As cComentario, "  )
			End If
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Pagar.Mon_Bru *(-1) Else Cuentas_Pagar.Mon_Bru End) As Mon_Bru,  ")
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Mon_Imp1, "  )
            'loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Mon_Net, "  )
            'loComandoSeleccionar.Appendline(" 			Cuentas_Pagar.Mon_Sal  "  )
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Pagar.Mon_Net *(-1) Else Cuentas_Pagar.Mon_Net End) As Mon_Net,  ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Pagar.Mon_Sal *(-1) Else Cuentas_Pagar.Mon_Sal End) As Mon_Sal  ")
			loComandoSeleccionar.Appendline(" From Proveedores, "  )
			loComandoSeleccionar.Appendline(" 			Cuentas_Pagar, "  )
			loComandoSeleccionar.Appendline(" 			Vendedores, "  )
			loComandoSeleccionar.Appendline(" 			Transportes, "  )
			loComandoSeleccionar.Appendline(" 			Monedas "  )
			loComandoSeleccionar.Appendline(" WHERE Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro "  )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Pagar.Cod_Ven = Vendedores.Cod_Ven "  )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Pagar.Cod_Tra = Transportes.Cod_Tra "  )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Pagar.Cod_Mon = Monedas.Cod_Mon "  )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Pagar.Documento BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.Appendline(" 			AND " & lcParametro0Hasta )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.Appendline(" 			AND " & lcParametro1Hasta )
			loComandoSeleccionar.Appendline(" 			AND Proveedores.Cod_Pro BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.Appendline(" 			AND " & lcParametro2Hasta )
			loComandoSeleccionar.Appendline(" 			AND Vendedores.Cod_Ven BETWEEN " & lcParametro3Desde )
			loComandoSeleccionar.Appendline(" 			AND " & lcParametro3Hasta )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Pagar.Status IN ( " & lcParametro4Desde & " ) " )
			loComandoSeleccionar.Appendline(" 			AND Transportes.Cod_Tra BETWEEN " & lcParametro5Desde )
			loComandoSeleccionar.Appendline(" 			AND " & lcParametro5Hasta )
			loComandoSeleccionar.Appendline(" 			AND Monedas.Cod_Mon BETWEEN " & lcParametro6Desde )
			loComandoSeleccionar.Appendline(" 			AND " & lcParametro6Hasta )
			loComandoSeleccionar.Appendline(" 			AND Cuentas_Pagar.Cod_Tip BETWEEN " & lcParametro7Desde )
			loComandoSeleccionar.Appendline(" 			AND " & lcParametro7Hasta )
			
			loComandoSeleccionar.AppendLine("		  AND ((" & lcParametro8Desde & " = 'Si' AND Cuentas_Pagar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("		  OR (" & lcParametro8Desde & " <> 'Si' AND (Cuentas_Pagar.Mon_Sal >= 0 or Cuentas_Pagar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("         AND Cuentas_Pagar.Cod_Suc between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("         AND " & lcParametro9Hasta)
           
            If lcParametro11Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev between " & lcParametro10Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev NOT between " & lcParametro10Desde)
            End If
            loComandoSeleccionar.AppendLine("         AND " & lcParametro10Hasta)
			
            'loComandoSeleccionar.AppendLine(" ORDER BY  Cuentas_Pagar.Fec_Ini, Cuentas_Pagar.Cod_Pro,  Cuentas_Pagar.Cod_Tip, Cuentas_Pagar.Documento ")
            loComandoSeleccionar.AppendLine("ORDER BY   CONVERT(nchar(30), Cuentas_Pagar.Fec_Ini,112), " & lcOrdenamiento)
'me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString , "curReportes")

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


            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCPagar_Fechas", laDatosReporte)
            
           'If lcParametro12Desde.ToString = "Si" Then
           '     loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 1
           '     CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Right = 0
           '     CType(loObjetoReporte.ReportDefinition.ReportObjects("Line1"), CrystalDecisions.CrystalReports.Engine.LineObject).Left = 0
           ' Else
           '     loObjetoReporte.DataDefinition.FormulaFields("Comentario_Notas").Text = 0
           '     loObjetoReporte.ReportDefinition.ReportObjects("text1").Height = 0
           '     loObjetoReporte.ReportDefinition.ReportObjects("text16").Height = 0
           '     loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Height = 0
           '     loObjetoReporte.ReportDefinition.ReportObjects("text1").Top = 0
           '     loObjetoReporte.ReportDefinition.ReportObjects("text16").Top = 0
           '     loObjetoReporte.ReportDefinition.ReportObjects("Comentario1").Top = 0
           '     loObjetoReporte.ReportDefinition.Sections(3).Height = 310
           '     loObjetoReporte.ReportDefinition.Sections("DetailSection2").Height = 3
           ' End If           
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCPagar_Fechas.ReportSource =	 loObjetoReporte

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
' JJD: 22/09/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 21/04/09: Estandarización del código y Corrección del estatus
'-------------------------------------------------------------------------------------------'
' YJP : 14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' CMS: 15/07/09: Metodo de Ordenamiento.
'                Verificación de Registros.
'                Multiplicación (*-1) al campo Mon_Net, Mon_Sal, Mon_Bru
'-------------------------------------------------------------------------------------------'
'-------------------------------------------------------------------------------------------'
' CMS: 21/09/09: Se agregaron los filtros ¿Solo Con Saldo?, Sucursal, Tipo de Revisión,
'				 Comentario
'-------------------------------------------------------------------------------------------'