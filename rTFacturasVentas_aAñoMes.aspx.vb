'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTFacturasVentas_aAñoMes"
'-------------------------------------------------------------------------------------------'
Partial Class rTFacturasVentas_aAñoMes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcComandoSeleccionar As New StringBuilder()
												  
			lcComandoSeleccionar.AppendLine("SELECT ")
			lcComandoSeleccionar.AppendLine("   		Articulos.cod_art,")
			lcComandoSeleccionar.AppendLine("   		SUBSTRING(Articulos.nom_art,0,70)		AS nom_art, ") 
			lcComandoSeleccionar.AppendLine("			DATEPART(MONTH,Facturas.Fec_Ini)		AS num_mes," ) 
			lcComandoSeleccionar.AppendLine("   		DATEPART(YEAR,Facturas.Fec_Ini)			AS AÑO,   ")
			lcComandoSeleccionar.AppendLine("   		CASE DatePart(MONTH,Facturas.Fec_Ini)")
			lcComandoSeleccionar.AppendLine("   			WHEN 1 THEN 'ENE'	 ")
			lcComandoSeleccionar.AppendLine("   			WHEN 2 THEN 'FEB'	")
			lcComandoSeleccionar.AppendLine("				WHEN 3 THEN 'MAR'   ")
			lcComandoSeleccionar.AppendLine("				WHEN 4 THEN 'ABR'   ")
			lcComandoSeleccionar.AppendLine("				WHEN 5 THEN 'MAY'   ")
			lcComandoSeleccionar.AppendLine("				WHEN 6 THEN 'JUN'   ")
			lcComandoSeleccionar.AppendLine("				WHEN 7 THEN 'JUL'   ")
			lcComandoSeleccionar.AppendLine("				WHEN 8 THEN 'AGO'   ")
			lcComandoSeleccionar.AppendLine("				WHEN 9 THEN 'SEP'   ")
			lcComandoSeleccionar.AppendLine("				WHEN 10 THEN 'OCT' ")
			lcComandoSeleccionar.AppendLine("				WHEN 11 THEN 'NOV'	")
			lcComandoSeleccionar.AppendLine("				WHEN 12 THEN 'DIC' ")
			lcComandoSeleccionar.AppendLine("   		END										AS mes, ")
			lcComandoSeleccionar.AppendLine("   		Renglones_Facturas.can_art1				AS cant,")
			lcComandoSeleccionar.AppendLine("   		Renglones_Facturas.mon_net				AS monto")
			lcComandoSeleccionar.AppendLine("INTO		#curTemporal")
			lcComandoSeleccionar.AppendLine("FROM		Facturas")
			lcComandoSeleccionar.AppendLine("	JOIN Renglones_Facturas ON Renglones_Facturas.documento = Facturas.documento ")
			lcComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.cod_art = Renglones_Facturas.cod_art")
			lcComandoSeleccionar.AppendLine(" 			AND Articulos.cod_art				BETWEEN " & lcParametro1Desde)
			lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
			lcComandoSeleccionar.AppendLine("           AND Articulos.cod_dep				BETWEEN " & lcParametro4Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)	
			lcComandoSeleccionar.AppendLine("           AND Articulos.cod_tip				BETWEEN " & lcParametro5Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
			lcComandoSeleccionar.AppendLine("WHERE		")
			lcComandoSeleccionar.AppendLine("			Facturas.fec_ini				BETWEEN " & lcParametro0Desde)
			lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
			lcComandoSeleccionar.AppendLine("           AND Facturas.cod_cli			BETWEEN " & lcParametro2Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
		 	lcComandoSeleccionar.AppendLine("           AND Facturas.cod_ven			BETWEEN " & lcParametro3Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
			lcComandoSeleccionar.AppendLine("           AND Facturas.status				IN (" & lcParametro6Desde & ")")
			lcComandoSeleccionar.AppendLine("           AND Facturas.cod_rev			BETWEEN " & lcParametro7Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
			lcComandoSeleccionar.AppendLine("")
   
			lcComandoSeleccionar.AppendLine(" SELECT  ")  
			lcComandoSeleccionar.AppendLine(" 			ROW_NUMBER() OVER(PARTITION BY cod_Art ORDER BY cod_Art, nom_art, AÑO, num_mes ASC) AS 'Renglon',  ")                
			lcComandoSeleccionar.AppendLine(" 			cod_art,	")  
			lcComandoSeleccionar.AppendLine(" 			num_mes,	") 
			lcComandoSeleccionar.AppendLine(" 			nom_art,	")
			lcComandoSeleccionar.AppendLine(" 			AÑO,		")
			lcComandoSeleccionar.AppendLine(" 			MES,		")
			lcComandoSeleccionar.AppendLine(" 			SUM(cant)		AS cant_art1, ")
            lcComandoSeleccionar.AppendLine(" 			SUM(monto)		AS mon_net,  ")
			lcComandoSeleccionar.AppendLine("			(SUM(monto))/30	AS mon_dia, ")
			lcComandoSeleccionar.AppendLine("			(SUM(cant))/30	AS cant_dia   ")
			lcComandoSeleccionar.AppendLine("FROM		#curTemporal	 ")
			lcComandoSeleccionar.AppendLine("GROUP BY   cod_art, nom_art, AÑO, num_mes, MES")
			lcComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos
			'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")
            
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTFacturasVentas_aAñoMes", laDatosReporte)
			
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTFacturasVentas_aAñoMes.ReportSource = loObjetoReporte

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
' YJP: 22/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 01/04/11: Ajuste del Select, Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'
' RJG: 19/01/12: 
'-------------------------------------------------------------------------------------------'
