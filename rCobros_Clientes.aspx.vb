'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCobros_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class rCobros_Clientes
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
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("SELECT ")
            loComandoSeleccionar.AppendLine("                Renglones_Cobros.Documento, ")
            loComandoSeleccionar.AppendLine("                Renglones_Cobros.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("                Renglones_Cobros.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("                Renglones_Cobros.Fec_ini, ")
            loComandoSeleccionar.AppendLine("                Renglones_Cobros.Renglon, ")
            loComandoSeleccionar.AppendLine("                Renglones_Cobros.mon_abo, ")

            loComandoSeleccionar.AppendLine("  			    Cobros.mon_Net AS Cobro_Mon_Net,")
            loComandoSeleccionar.AppendLine("  			    Cobros.Cod_Cli,")
            loComandoSeleccionar.AppendLine("  			    Clientes.Nom_Cli,")

            loComandoSeleccionar.AppendLine("  			    Cobros.Fec_Ini AS Fec_Ini_Cobro,")
            loComandoSeleccionar.AppendLine("  			    Cobros.Recibo,")
            loComandoSeleccionar.AppendLine("  			    Cobros.Cod_Ven AS Cod_Ven_Cobro,")
            loComandoSeleccionar.AppendLine("  			    Cobros.Status AS Status_Cobro,")
            loComandoSeleccionar.AppendLine("  			    CAST(Cobros.Comentario AS VARCHAR) AS Comentario_Cobro,")

            loComandoSeleccionar.AppendLine("  			    1 AS Tabla")
            loComandoSeleccionar.AppendLine("INTO			#tmpRenglones_Cobros")
            loComandoSeleccionar.AppendLine("FROM			Cobros")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Cobros ON Cobros.Documento = Renglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Cobros.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE ")
            loComandoSeleccionar.AppendLine(" 			    Cobros.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			    And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			    And Cobros.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			    And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			    And Cobros.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			    And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			    And Cobros.Cod_Mon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			    And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			    And Cobros.Cod_Ven between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			    And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			    And Cobros.Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("               AND Cobros.Cod_Rev between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("    	        AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("               AND Cobros.Cod_Suc between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	        AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 	ORDER BY Cobros.Documento, Cobros.Cod_Cli, Clientes.Nom_Cli")

			
            loComandoSeleccionar.AppendLine(" 	SELECT ")

            loComandoSeleccionar.AppendLine("                Detalles_Cobros.Documento, ")
            loComandoSeleccionar.AppendLine("                Detalles_Cobros.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("                Detalles_Cobros.Fec_ini, ")
            loComandoSeleccionar.AppendLine("                Detalles_Cobros.Renglon, ")
            loComandoSeleccionar.AppendLine("                Detalles_Cobros.Tip_ope, ")
            loComandoSeleccionar.AppendLine("                Detalles_Cobros.Num_Doc, ")
            loComandoSeleccionar.AppendLine("                Detalles_Cobros.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("                Detalles_Cobros.Mon_Net, ")
            loComandoSeleccionar.AppendLine("                CASE ")
            loComandoSeleccionar.AppendLine("                	WHEN Detalles_Cobros.Tip_Ope = 'Tarjeta' THEN Cuentas_Bancarias.Cod_Ban ")
            loComandoSeleccionar.AppendLine("                	WHEN Detalles_Cobros.Tip_Ope = 'Transferencia' THEN Cuentas_Bancarias.Cod_Ban ")
            loComandoSeleccionar.AppendLine("                	WHEN Detalles_Cobros.Tip_Ope = 'Deposito' THEN Cuentas_Bancarias.Cod_Ban ")
            loComandoSeleccionar.AppendLine("                	ELSE Detalles_Cobros.Cod_Ban ")
            loComandoSeleccionar.AppendLine("                END AS Cod_Ban,  ")

            loComandoSeleccionar.AppendLine("  			    Cobros.Cod_Cli,")
            loComandoSeleccionar.AppendLine("  			    Clientes.Nom_Cli,")
            loComandoSeleccionar.AppendLine(" 				2 AS Tabla ")
            loComandoSeleccionar.AppendLine(" 	INTO #tmpDetalles_Cobros")
            loComandoSeleccionar.AppendLine(" 	FROM Cobros")
            loComandoSeleccionar.AppendLine(" 	JOIN Detalles_Cobros ON Cobros.Documento = Detalles_Cobros.Documento")
            loComandoSeleccionar.AppendLine("   JOIN Clientes ON  Cobros.Cod_Cli = Clientes.Cod_Cli")

            loComandoSeleccionar.AppendLine("   LEFT JOIN Cuentas_Bancarias ON Detalles_Cobros.Cod_Cue = Cuentas_Bancarias.Cod_Cue")

            loComandoSeleccionar.AppendLine(" 	    WHERE			")
            loComandoSeleccionar.AppendLine(" 				Cobros.Documento between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				And Cobros.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				And Cobros.Cod_Cli between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				And Cobros.Cod_Mon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				And Cobros.Cod_Ven between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				And Cobros.Status IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("               AND Cobros.Cod_Rev between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("    	        AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("               AND Cobros.Cod_Suc between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	        AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 	ORDER BY Cobros.Documento")

            loComandoSeleccionar.AppendLine(" 	SELECT")
            loComandoSeleccionar.AppendLine(" 				ISNULL(#tmpRenglones_Cobros.Cod_Cli, #tmpDetalles_Cobros.Cod_Cli) AS Cod_Cli,")
            loComandoSeleccionar.AppendLine(" 				ISNULL(#tmpRenglones_Cobros.Nom_Cli, #tmpDetalles_Cobros.Nom_Cli) AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("   			ISNULL(#tmpRenglones_Cobros.Documento,#tmpDetalles_Cobros.Documento) AS Documento,")
            loComandoSeleccionar.AppendLine(" 				ISNULL(#tmpRenglones_Cobros.Tabla, #tmpDetalles_Cobros.Tabla) AS Tabla,")
            loComandoSeleccionar.AppendLine(" 				#tmpRenglones_Cobros.Fec_ini AS FEC_Cob,")
            loComandoSeleccionar.AppendLine(" 				#tmpDetalles_Cobros.Fec_ini AS Fec_Pag,")
            loComandoSeleccionar.AppendLine(" 				#tmpRenglones_Cobros.Cod_Tip,")
            loComandoSeleccionar.AppendLine(" 				#tmpDetalles_Cobros.Tip_ope,")
            loComandoSeleccionar.AppendLine(" 				#tmpDetalles_Cobros.Num_Doc,")
            loComandoSeleccionar.AppendLine(" 				#tmpRenglones_Cobros.Doc_Ori,")
            loComandoSeleccionar.AppendLine("				 (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN #tmpDetalles_Cobros.mon_net *(-1) ELSE #tmpDetalles_Cobros.mon_net END) AS Pago, ")
            loComandoSeleccionar.AppendLine("				 (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN #tmpRenglones_Cobros.Mon_Abo *(-1) ELSE #tmpRenglones_Cobros.Mon_Abo END) AS Cargo, ")
            loComandoSeleccionar.AppendLine("				 (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Net *(-1) ELSE Cuentas_Cobrar.Mon_Net END) AS Mon_Doc,")
            loComandoSeleccionar.AppendLine("				 (CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN #tmpRenglones_Cobros.Cobro_Mon_Net *(-1) ELSE #tmpRenglones_Cobros.Cobro_Mon_Net END) AS Cobro_Mon_Net,")

            loComandoSeleccionar.AppendLine("  			    #tmpRenglones_Cobros.Fec_Ini_Cobro,")
            loComandoSeleccionar.AppendLine("  			    #tmpRenglones_Cobros.Recibo,")
            loComandoSeleccionar.AppendLine("  			    RTRIM(#tmpRenglones_Cobros.Cod_Ven_Cobro) +'  '+ RTRIM(Vendedores.Nom_Ven) AS Cod_Ven_Cobro,")
            loComandoSeleccionar.AppendLine("  			    #tmpRenglones_Cobros.Status_Cobro,")
            loComandoSeleccionar.AppendLine("  			    #tmpRenglones_Cobros.Comentario_Cobro,")
            loComandoSeleccionar.AppendLine("  			    #tmpDetalles_Cobros.Cod_Cue,")
            loComandoSeleccionar.AppendLine("  			    RTRIM(#tmpDetalles_Cobros.Cod_Caj) +'  '+  RTRIM(Cajas.Nom_Caj) AS Cod_Caj,")
            loComandoSeleccionar.AppendLine("  			    RTRIM(#tmpDetalles_Cobros.Cod_Ban) +'  '+RTRIM(Bancos.Nom_Ban) AS Cod_ban,")
            loComandoSeleccionar.AppendLine("  			    RTRIM(Cajas.Nom_Caj) AS Caja")


            loComandoSeleccionar.AppendLine(" 	FROM #tmpRenglones_Cobros")
            loComandoSeleccionar.AppendLine(" 	FULL JOIN #tmpDetalles_Cobros ON #tmpDetalles_Cobros.Documento = #tmpRenglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine(" 			AND #tmpRenglones_Cobros.Tabla = #tmpDetalles_Cobros.Tabla")
            loComandoSeleccionar.AppendLine(" 	LEFT JOIN Cuentas_Cobrar ON  Cuentas_Cobrar.Documento = #tmpRenglones_Cobros.Doc_Ori")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Tip = #tmpRenglones_Cobros.Cod_Tip ")

            loComandoSeleccionar.AppendLine(" 	LEFT JOIN Vendedores ON #tmpRenglones_Cobros.Cod_Ven_Cobro = Vendedores.Cod_Ven")
            loComandoSeleccionar.AppendLine(" 	LEFT JOIN Cajas ON Cajas.Cod_Caj = #tmpDetalles_Cobros.Cod_Caj")
            loComandoSeleccionar.AppendLine(" 	LEFT JOIN Bancos ON Bancos.Cod_Ban = #tmpDetalles_Cobros.Cod_Ban")

            loComandoSeleccionar.AppendLine(" 	GROUP BY    ISNULL(#tmpRenglones_Cobros.Cod_Cli, #tmpDetalles_Cobros.Cod_Cli), ")
            loComandoSeleccionar.AppendLine(" 			    ISNULL(#tmpRenglones_Cobros.Nom_Cli, #tmpDetalles_Cobros.Nom_Cli), ")
            loComandoSeleccionar.AppendLine(" 			    ISNULL(#tmpRenglones_Cobros.Tabla, #tmpDetalles_Cobros.Tabla),")
            loComandoSeleccionar.AppendLine(" 			    ISNULL(#tmpRenglones_Cobros.Documento,#tmpDetalles_Cobros.Documento),")
			loComandoSeleccionar.AppendLine(" 			    Cuentas_Cobrar.Tip_Doc,")
            loComandoSeleccionar.AppendLine(" 			    #tmpRenglones_Cobros.Fec_ini,")
            loComandoSeleccionar.AppendLine("   	        #tmpDetalles_Cobros.Fec_ini,")
            loComandoSeleccionar.AppendLine("   	        #tmpRenglones_Cobros.Fec_Ini_Cobro,")
            loComandoSeleccionar.AppendLine("   	        #tmpRenglones_Cobros.Recibo,")
            loComandoSeleccionar.AppendLine("   	        #tmpRenglones_Cobros.Cod_Ven_Cobro,")
            loComandoSeleccionar.AppendLine("   	        Vendedores.nom_ven,")
            loComandoSeleccionar.AppendLine("   	        #tmpRenglones_Cobros.Status_Cobro,")
            loComandoSeleccionar.AppendLine("   	        #tmpRenglones_Cobros.Comentario_Cobro,")

            loComandoSeleccionar.AppendLine(" 			    #tmpRenglones_Cobros.Cod_Tip,")
            loComandoSeleccionar.AppendLine(" 			    #tmpRenglones_Cobros.Renglon,")
            loComandoSeleccionar.AppendLine(" 			    #tmpDetalles_Cobros.Renglon,")
            loComandoSeleccionar.AppendLine(" 			    #tmpDetalles_Cobros.Tip_ope,")
            loComandoSeleccionar.AppendLine(" 				#tmpDetalles_Cobros.Num_Doc,")
            loComandoSeleccionar.AppendLine(" 			    #tmpRenglones_Cobros.Doc_Ori,")

            loComandoSeleccionar.AppendLine(" 			    #tmpDetalles_Cobros.Cod_Cue,")
            loComandoSeleccionar.AppendLine(" 			    #tmpDetalles_Cobros.Cod_Caj,")
            loComandoSeleccionar.AppendLine(" 			    Cajas.Nom_Caj,")
            loComandoSeleccionar.AppendLine(" 			    #tmpDetalles_Cobros.Cod_Ban,")
            loComandoSeleccionar.AppendLine(" 			    Bancos.Nom_Ban,")

            loComandoSeleccionar.AppendLine(" 			    #tmpDetalles_Cobros.mon_net,")
            loComandoSeleccionar.AppendLine(" 			    #tmpRenglones_Cobros.mon_abo,")
            loComandoSeleccionar.AppendLine(" 			    Cuentas_Cobrar.Mon_Net,")
            loComandoSeleccionar.AppendLine("   	        #tmpRenglones_Cobros.Cobro_Mon_Net ")
            loComandoSeleccionar.AppendLine(" 	ORDER BY   ")
            loComandoSeleccionar.AppendLine(" 				ISNULL(#tmpRenglones_Cobros.Cod_Cli, #tmpDetalles_Cobros.Cod_Cli), ")
            loComandoSeleccionar.AppendLine(" 				ISNULL(#tmpRenglones_Cobros.Nom_Cli, #tmpDetalles_Cobros.Nom_Cli), ")
            loComandoSeleccionar.AppendLine(" 			    ISNULL(#tmpRenglones_Cobros.Documento,#tmpDetalles_Cobros.Documento)" & lcOrdenamiento.Replace("?","") & ", ")
            loComandoSeleccionar.AppendLine(" 				ISNULL(#tmpRenglones_Cobros.Tabla, #tmpDetalles_Cobros.Tabla),")
            loComandoSeleccionar.AppendLine(" 			    #tmpRenglones_Cobros.Renglon,")
            loComandoSeleccionar.AppendLine(" 			    #tmpDetalles_Cobros.Renglon,")
            loComandoSeleccionar.AppendLine(" 			    #tmpRenglones_Cobros.Fec_ini, ")
            loComandoSeleccionar.AppendLine("   		    #tmpDetalles_Cobros.Fec_ini,")
            loComandoSeleccionar.AppendLine(" 		        #tmpRenglones_Cobros.Cod_Tip")

            'loComandoSeleccionar.AppendLine("ORDER BY           Cobros.Cod_Cli," & lcOrdenamiento & ", Convert(nchar(30), Cobros.Fec_Ini, 112) Desc")
'me.mEscribirConsulta (loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCobros_Clientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCobros_Clientes.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' JJD: 22/09/08: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”														'
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”														'
'-------------------------------------------------------------------------------------------'
' CMS:  16/07/09: Metodo de ordenamieto, verificacion de registros, reprogramacion			'
'-------------------------------------------------------------------------------------------'
' CMS:  09/09/09: Se ajusto el signo de los montos segun su naturaleza						'
'-------------------------------------------------------------------------------------------'
' CMS:  04/05/10: Ajuste al Metodo de ordenamieto											'
'-------------------------------------------------------------------------------------------'
' RJG:  10/04/12: Se agregó el total de documentos.											'
'-------------------------------------------------------------------------------------------'
