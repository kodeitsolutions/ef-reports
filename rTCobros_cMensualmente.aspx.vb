'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTCobros_cMensualmente"
'-------------------------------------------------------------------------------------------'
Partial Class rTCobros_cMensualmente
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            loComandoSeleccionar.AppendLine("  SELECT	")
            loComandoSeleccionar.AppendLine("  		 	  Clientes.Cod_Cli,	")
            loComandoSeleccionar.AppendLine("  			  Clientes.Nom_Cli,	")
            loComandoSeleccionar.AppendLine("  			  CASE  	")
            loComandoSeleccionar.AppendLine("  				  WHEN DATEPART(MONTH, Cobros.Fec_ini) = 1 THEN  Cobros.Mon_net 	")
            loComandoSeleccionar.AppendLine("  				  else 0  	")
            loComandoSeleccionar.AppendLine("  			  END as Ene, 	")
            loComandoSeleccionar.AppendLine("  			  CASE  	")
            loComandoSeleccionar.AppendLine("  				  WHEN DATEPART(MONTH, Cobros.Fec_ini) = 2 THEN Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				  else 0 	")
            loComandoSeleccionar.AppendLine("  			  END as Feb, 	")
            loComandoSeleccionar.AppendLine("  			  CASE  	")
            loComandoSeleccionar.AppendLine("  				  WHEN DATEPART(MONTH, Cobros.Fec_ini) = 3 THEN Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				  else 0 	")
            loComandoSeleccionar.AppendLine("  			  END as Mar, 	")
            loComandoSeleccionar.AppendLine("  			  CASE  	")
            loComandoSeleccionar.AppendLine("  				  WHEN DATEPART(MONTH, Cobros.Fec_ini) = 4 THEN  Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				  else 0 	")
            loComandoSeleccionar.AppendLine("  			  END as Abr,  	")
            loComandoSeleccionar.AppendLine("  			  CASE  	")
            loComandoSeleccionar.AppendLine("  				  WHEN DATEPART(MONTH, Cobros.Fec_ini) = 5 THEN  Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				  else 0 	")
            loComandoSeleccionar.AppendLine("  			  END as May, 	")
            loComandoSeleccionar.AppendLine("  			  CASE  	")
            loComandoSeleccionar.AppendLine("  				  WHEN DATEPART(MONTH, Cobros.Fec_ini) = 6 THEN  Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				  else 0 	")
            loComandoSeleccionar.AppendLine("  			  END as Jun, 	")
            loComandoSeleccionar.AppendLine("  			  CASE  	")
            loComandoSeleccionar.AppendLine("  				  WHEN DATEPART(MONTH, Cobros.Fec_ini) = 7 THEN  Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				  else 0 	")
            loComandoSeleccionar.AppendLine("  			END as Jul, 	")
            loComandoSeleccionar.AppendLine("  			CASE  	")
            loComandoSeleccionar.AppendLine("  				  WHEN DATEPART(MONTH, Cobros.Fec_ini) = 8 THEN  Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				  else 0 	")
            loComandoSeleccionar.AppendLine("  			END as Ago, 	")
            loComandoSeleccionar.AppendLine("  			CASE  	")
            loComandoSeleccionar.AppendLine("  				  WHEN DATEPART(MONTH, Cobros.Fec_ini) = 9 THEN  Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				  else 0 	")
            loComandoSeleccionar.AppendLine("  			 END as Sep, 	")
            loComandoSeleccionar.AppendLine("  			 CASE  	")
            loComandoSeleccionar.AppendLine("  				 WHEN DATEPART(MONTH, Cobros.Fec_ini) = 10 THEN  Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				 else 0 	")
            loComandoSeleccionar.AppendLine("  			 END as Oct, 	")
            loComandoSeleccionar.AppendLine("  			 CASE  	")
            loComandoSeleccionar.AppendLine("  				 WHEN DATEPART(MONTH, Cobros.Fec_ini) = 11 THEN  Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				 else 0 	")
            loComandoSeleccionar.AppendLine("  			 END as Nov, 	")
            loComandoSeleccionar.AppendLine("  			 CASE  	")
            loComandoSeleccionar.AppendLine("  				 WHEN DATEPART(MONTH, Cobros.Fec_ini) = 12 THEN  Cobros.Mon_net  	")
            loComandoSeleccionar.AppendLine("  				 else 0 	")
            loComandoSeleccionar.AppendLine("  			 END as Dic, 	")
            loComandoSeleccionar.AppendLine("               Cobros.Mon_net AS Total	")
            loComandoSeleccionar.AppendLine("  			INTO        #tmpTemporal	")
            loComandoSeleccionar.AppendLine("              from	")
            loComandoSeleccionar.AppendLine("  			Clientes, 	")
            loComandoSeleccionar.AppendLine("  			Cobros 	")
            'loComandoSeleccionar.AppendLine("  			Cuentas_Cobrar, 	")
            'loComandoSeleccionar.AppendLine("  			Vendedores, 	")
            'loComandoSeleccionar.AppendLine("  			Renglones_Cobros,	")
            'loComandoSeleccionar.AppendLine("  			Detalles_Cobros, 	")
            'loComandoSeleccionar.AppendLine("           Monedas	")
            loComandoSeleccionar.AppendLine("  where	")
            'loComandoSeleccionar.AppendLine("              Cuentas_Cobrar.Documento = Renglones_Cobros.Doc_Ori	")
            'loComandoSeleccionar.AppendLine("  			AND Cuentas_Cobrar.Cod_Tip = Renglones_Cobros.Cod_Tip 	")
            'loComandoSeleccionar.AppendLine("  			AND Cobros.Documento = Renglones_Cobros.Documento 	")
            'loComandoSeleccionar.AppendLine("  			AND Cobros.Documento = Detalles_Cobros.Documento 	")
            loComandoSeleccionar.AppendLine("  			 Cobros.Cod_Cli = Clientes.Cod_Cli 	")
            'loComandoSeleccionar.AppendLine("  			AND Cobros.Cod_Ven = Vendedores.Cod_Ven 	")
            'loComandoSeleccionar.AppendLine("  			AND Cobros.Cod_Mon = Monedas.Cod_Mon 	")
            loComandoSeleccionar.AppendLine(" 					And Cobros.Fec_Ini between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 					And Cobros.Cod_Cli between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 					And Clientes.Cod_Zon between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 					And Clientes.Cod_Cla between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 					And Clientes.Cod_Tip between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 					And Cobros.Cod_Mon between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 					And Cobros.Cod_Ven between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 					And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 					And Cobros.Status IN (" & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("                   AND Cobros.Cod_Rev between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    	            AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("  SELECT  	")
            loComandoSeleccionar.AppendLine("                #tmpTemporal.Cod_Cli, 	")
            loComandoSeleccionar.AppendLine("                #tmpTemporal.Nom_Cli, 	")
            loComandoSeleccionar.AppendLine("                sum(ene) as Ene, 	")
            loComandoSeleccionar.AppendLine("                sum(feb) as Feb, 	")
            loComandoSeleccionar.AppendLine("                sum(mar) as Mar, 	")
            loComandoSeleccionar.AppendLine("                sum(abr) as Abr, 	")
            loComandoSeleccionar.AppendLine("                sum(may) as May, 	")
            loComandoSeleccionar.AppendLine("                sum(jun) as Jun, 	")
            loComandoSeleccionar.AppendLine("                sum(jul) as Jul, 	")
            loComandoSeleccionar.AppendLine("                sum(ago) as Ago, 	")
            loComandoSeleccionar.AppendLine("                sum(sep) as Sep, 	")
            loComandoSeleccionar.AppendLine("                sum(oct) as Oct, 	")
            loComandoSeleccionar.AppendLine("                sum(nov) as Nov, 	")
            loComandoSeleccionar.AppendLine("                sum(dic) as Dic, 	")
            loComandoSeleccionar.AppendLine("                sum(total) as Total	")
            loComandoSeleccionar.AppendLine("  FROM #tmpTemporal 	")
            loComandoSeleccionar.AppendLine("  WHERE Total > 0    	")
            loComandoSeleccionar.AppendLine("              Group by	")
            loComandoSeleccionar.AppendLine("        Cod_Cli, 	")
            loComandoSeleccionar.AppendLine("        Nom_Cli	")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTCobros_cMensualmente", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTCobros_cMensualmente.ReportSource = loObjetoReporte


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
' CMS: 07/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' CMS: 26/03/10: Se cambio la funcion DATEPART por DATEPART
'-------------------------------------------------------------------------------------------'