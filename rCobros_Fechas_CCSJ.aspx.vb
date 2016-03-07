'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCobros_Fechas_CCSJ"
'-------------------------------------------------------------------------------------------'
Partial Class rCobros_Fechas_CCSJ
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
			Dim lcParametro0Desde As String = cusAplicacion.goReportes.paParametrosIniciales(0)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            

            'Escapa los caracteres especiales para el LIKE
            Dim llBuscarPlanilla As Boolean = False
            If Not String.IsNullOrEmpty(Strings.Trim(lcParametro0Desde)) Then 
                llBuscarPlanilla = True
                lcParametro0Desde = lcParametro0Desde.Replace("%", "[%]").Replace("_", "[_]")
                lcParametro0Desde = goServicios.mObtenerCampoFormatoSQL("%" & lcParametro0Desde & "%")
            End If

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            


			Dim loConsulta As New StringBuilder()

			loConsulta.AppendLine("SELECT			Cobros.Documento			AS Documento,	")
			loConsulta.AppendLine(" 				Cobros.Fec_Ini				AS Fec_Ini,		")
			loConsulta.AppendLine(" 				Cobros.Cod_Cli				AS Cod_Cli,")
			loConsulta.AppendLine(" 				CASE WHEN (COALESCE(cuentas_cobrar.nom_cli, '')>'')")
			loConsulta.AppendLine(" 				    THEN COALESCE(cuentas_cobrar.nom_cli, '')")
			loConsulta.AppendLine(" 				    ELSE Clientes.Nom_Cli	")
			loConsulta.AppendLine(" 				END                         AS Nom_Cli,		")
			loConsulta.AppendLine(" 				cuentas_cobrar.Documento	AS Factura,		")
			loConsulta.AppendLine(" 				cuentas_cobrar.Control		AS Control,		")
			loConsulta.AppendLine(" 				Cobros.Mon_Net				AS Mon_Net,		")
			loConsulta.AppendLine(" 				Cobros.Cod_Mon				AS Cod_Mon,		")
			loConsulta.AppendLine(" 				Cobros.Cod_Ven				AS Cod_Ven,		")
			loConsulta.AppendLine(" 				Vendedores.Nom_Ven			AS Nom_Ven,		")
			loConsulta.AppendLine(" 				Detalles_Cobros.Tip_Ope		AS Tip_Ope,		")
			loConsulta.AppendLine(" 				Detalles_Cobros.Renglon		AS Renglon,		")
			loConsulta.AppendLine(" 				Detalles_Cobros.Mon_Net		AS Cob_Ope,		")
			loConsulta.AppendLine(" 				Detalles_Cobros.Doc_Des		AS Doc_Des,		")
			loConsulta.AppendLine(" 				Detalles_Cobros.Num_Doc		AS Num_Doc,		")
			loConsulta.AppendLine(" 				CASE WHEN(Detalles_Cobros.Cod_Cue = '')		")
			loConsulta.AppendLine(" 					THEN	Detalles_Cobros.Cod_Ban			")
			loConsulta.AppendLine(" 					ELSE	Cuentas_Bancarias.Cod_Ban		")
			loConsulta.AppendLine(" 				END							AS Cod_Ban,		")
			loConsulta.AppendLine(" 				Detalles_Cobros.Cod_Caj		AS Cod_Caj,		")
			loConsulta.AppendLine(" 				Detalles_Cobros.Cod_Cue		AS Cod_Cue,		")
			loConsulta.AppendLine(" 				ROW_NUMBER()")
			loConsulta.AppendLine(" 				    OVER(PARTITION BY Cobros.Documento")
			loConsulta.AppendLine(" 				         ORDER BY Cobros.Fec_Ini," & lcOrdenamiento & ") AS Posicion_Detalle")
			loConsulta.AppendLine("FROM			    Cobros ")
			loConsulta.AppendLine("	    LEFT JOIN 	renglones_cobros ")
			loConsulta.AppendLine("	            ON  renglones_cobros.Documento = Cobros.Documento ")
			loConsulta.AppendLine("	            AND renglones_cobros.renglon = 1")
			loConsulta.AppendLine("	    LEFT JOIN 	cuentas_cobrar ")
			loConsulta.AppendLine("	            ON  cuentas_cobrar.Documento = renglones_cobros.doc_ori")
			loConsulta.AppendLine("	            AND cuentas_cobrar.cod_tip = renglones_cobros.cod_tip")
			loConsulta.AppendLine("	    JOIN 	    Detalles_Cobros ")
			loConsulta.AppendLine("	            ON  Cobros.Documento = Detalles_Cobros.Documento")
			loConsulta.AppendLine("	    JOIN 	    Clientes ")
			loConsulta.AppendLine("	            ON  Cobros.Cod_Cli = Clientes.Cod_Cli")
			loConsulta.AppendLine("	    JOIN 	    Vendedores ")
			loConsulta.AppendLine("	            ON  Cobros.Cod_Ven = Vendedores.Cod_Ven ")
			loConsulta.AppendLine("	    LEFT JOIN   Cuentas_Bancarias ")
			loConsulta.AppendLine("	            ON  Cuentas_Bancarias.Cod_Cue = Detalles_Cobros.Cod_Cue")
			loConsulta.AppendLine("WHERE		Cobros.Documento BETWEEN " & lcParametro1Desde)
			loConsulta.AppendLine(" 				AND " & lcParametro1Hasta)
			loConsulta.AppendLine(" 			AND Cobros.Fec_Ini BETWEEN " & lcParametro2Desde)
			loConsulta.AppendLine(" 				AND " & lcParametro2Hasta)
			loConsulta.AppendLine(" 			AND Cobros.Cod_Cli BETWEEN " & lcParametro3Desde)
			loConsulta.AppendLine(" 				AND " & lcParametro3Hasta)
			loConsulta.AppendLine(" 			AND Cobros.Cod_Mon BETWEEN " & lcParametro4Desde)
			loConsulta.AppendLine(" 				AND " & lcParametro4Hasta)
			loConsulta.AppendLine(" 			AND Cobros.Cod_Ven BETWEEN " & lcParametro5Desde)
			loConsulta.AppendLine(" 				AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 			AND Cobros.Status IN (" & lcParametro6Desde & ")")
            loConsulta.AppendLine("       	    AND Cobros.Cod_Rev BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine("    		        AND " & lcParametro7Hasta)
            loConsulta.AppendLine("       	    AND Cobros.Cod_Suc BETWEEN " & lcParametro8Desde)
            loConsulta.AppendLine("    		        AND " & lcParametro8Hasta)
            loConsulta.AppendLine("       	    AND Cobros.Cod_Usu BETWEEN " & lcParametro9Desde)
            loConsulta.AppendLine("    		        AND " & lcParametro9Hasta)
            If llBuscarPlanilla Then 
                loConsulta.AppendLine("       	 AND Cobros.Comentario LIKE " & lcParametro0Desde)
            End if
            loConsulta.AppendLine("ORDER BY   Cobros.Fec_Ini, " & lcOrdenamiento)
            
            Dim loServicios As New cusDatos.goDatos

		    'Me.mEscribirConsulta(loConsulta.ToString)

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

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCobros_Fechas_CCSJ", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCobros_Fechas_CCSJ.ReportSource =	 loObjetoReporte

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
' Fin del codigo                                                    			            '
'-------------------------------------------------------------------------------------------'
' RJG:  13/01/14: Programacion inicial, a partir de rCobros_Fechas.				            '
'-------------------------------------------------------------------------------------------'
' RJG:  27/01/14: Se corrigieron los total de registros: estaban invertidos.				'
'-------------------------------------------------------------------------------------------'
' RJG:  18/02/14: Se cambió el campo Usu_Cre por Cod_Usu para filtrar.				        '
'-------------------------------------------------------------------------------------------'
' RJG:  08/07/14: Se agregaron los campos Documento y Control (de la factura) y se bajó el  '
'                 nombre a la segunda línea.				                                '
'-------------------------------------------------------------------------------------------'
