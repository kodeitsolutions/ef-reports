﻿Imports System.Data
Imports cusAplicacion

Partial Class rTdProveedores_pAñoMes
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
			lcComandoSeleccionar.AppendLine("   		proveedores.cod_pro,")
			lcComandoSeleccionar.AppendLine("   		SUBSTRING(proveedores.nom_pro,0,70) as nom_pro, ") 
			lcComandoSeleccionar.AppendLine("			DatePart(MONTH,devoluciones_proveedores.Fec_Ini) as num_mes," ) 
			lcComandoSeleccionar.AppendLine("   		DatePart(YEAR,devoluciones_proveedores.Fec_Ini)as AÑO,   ")
			lcComandoSeleccionar.AppendLine("   		CASE WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=1 THEN 'ENE'	 ")
			lcComandoSeleccionar.AppendLine("   		WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=2 THEN 'FEB'	")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=3 THEN 'MAR'   ")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=4 THEN 'ABR'   ")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=5 THEN 'MAY'   ")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=6 THEN 'JUN'   ")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=7 THEN 'JUL'   ")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=8 THEN 'AGO'   ")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=9 THEN 'SEP'   ")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=10 THEN 'OCT' ")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=11 THEN 'NOV'	")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,devoluciones_proveedores.Fec_Ini)=12 THEN 'DIC' ")
			lcComandoSeleccionar.AppendLine("   		END AS mes, ")
			lcComandoSeleccionar.AppendLine("   		renglones_dproveedores.can_art1 as cant,  ")
			lcComandoSeleccionar.AppendLine("   		renglones_dproveedores.mon_net as monto	  ")
			lcComandoSeleccionar.AppendLine("into #temporal	 ")
			lcComandoSeleccionar.AppendLine("FROM ")
			lcComandoSeleccionar.AppendLine(" devoluciones_proveedores, renglones_dproveedores, transportes, monedas, articulos, proveedores " )
			lcComandoSeleccionar.AppendLine(" WHERE	")
			lcComandoSeleccionar.AppendLine("			devoluciones_proveedores.documento		=	renglones_dproveedores.documento ")
			lcComandoSeleccionar.AppendLine("			AND  proveedores.Cod_pro				=	devoluciones_proveedores.Cod_pro ")
			'lcComandoSeleccionar.AppendLine("			AND  devoluciones_proveedores.Cod_Ven	=	Vendedores.Cod_Ven ")
			lcComandoSeleccionar.AppendLine("			AND devoluciones_proveedores.Cod_Tra	=	Transportes.Cod_Tra ")
			lcComandoSeleccionar.AppendLine("			AND devoluciones_proveedores.Cod_Mon	=	Monedas.Cod_Mon ")
			lcComandoSeleccionar.AppendLine("			AND renglones_dproveedores.Cod_art		=	articulos.Cod_art ")
			lcComandoSeleccionar.AppendLine(" 			AND    devoluciones_proveedores.fec_ini		BETWEEN " & lcParametro0Desde)
			lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
			lcComandoSeleccionar.AppendLine(" 			AND articulos.cod_art     BETWEEN " & lcParametro1Desde)
			lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
			lcComandoSeleccionar.AppendLine("           AND devoluciones_proveedores.cod_pro		BETWEEN " & lcParametro2Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
		 	lcComandoSeleccionar.AppendLine("           AND devoluciones_proveedores.cod_ven		BETWEEN " & lcParametro3Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
		   	lcComandoSeleccionar.AppendLine("           AND articulos.cod_dep						BETWEEN " & lcParametro4Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)	
			lcComandoSeleccionar.AppendLine("           AND articulos.cod_tip						BETWEEN " & lcParametro5Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
			lcComandoSeleccionar.AppendLine("           AND devoluciones_proveedores.status			IN (" & lcParametro6Desde & ")")
			lcComandoSeleccionar.AppendLine("           AND devoluciones_proveedores.cod_rev		BETWEEN " & lcParametro7Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
   
			lcComandoSeleccionar.AppendLine(" SELECT  ")  
			lcComandoSeleccionar.AppendLine(" 			ROW_NUMBER() OVER(PARTITION BY cod_pro ORDER BY  cod_pro, nom_pro, AÑO, num_mes ASC) AS 'Renglon',  ")                
			lcComandoSeleccionar.AppendLine(" 			cod_pro,  ")  
			lcComandoSeleccionar.AppendLine(" 			num_mes,  ") 
			lcComandoSeleccionar.AppendLine(" 			nom_pro,   ")
			lcComandoSeleccionar.AppendLine(" 			AÑO,	   ")
			lcComandoSeleccionar.AppendLine(" 			MES,	 ")
			lcComandoSeleccionar.AppendLine(" 			SUM(cant) as cant_art1, ")
            lcComandoSeleccionar.AppendLine(" 			SUM(monto) as mon_net,  ")
			lcComandoSeleccionar.AppendLine("			(SUM(monto))/30 as mon_dia, ")
			lcComandoSeleccionar.AppendLine("			(SUM(cant))/30 as cant_dia   ")
			lcComandoSeleccionar.AppendLine(" FROM		#temporal	 ")
			lcComandoSeleccionar.AppendLine("GROUP BY   cod_pro, nom_pro, AÑO, num_mes, MES")
			lcComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTdProveedores_pAñoMes", laDatosReporte)
			
			
			
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTdProveedores_pAñoMes.ReportSource = loObjetoReporte

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
' YJP: 20/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
