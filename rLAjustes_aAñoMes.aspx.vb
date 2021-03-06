﻿Imports System.Data
Imports cusAplicacion

Partial Class rLAjustes_aAñoMes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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


			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcComandoSeleccionar As New StringBuilder()
			
			lcComandoSeleccionar.AppendLine("SELECT		")
			lcComandoSeleccionar.AppendLine("			articulos.cod_art,")
			lcComandoSeleccionar.AppendLine("			articulos.nom_art,	")
			lcComandoSeleccionar.AppendLine("			DatePart(YEAR,ajustes.Fec_Ini)as AÑO,")   
			lcComandoSeleccionar.AppendLine("			DatePart(MONTH,ajustes.Fec_Ini) as num_mes,")
			lcComandoSeleccionar.AppendLine("			CASE WHEN DatePart(MONTH,ajustes.Fec_Ini)=1 THEN 'ENE'	") 
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=2 THEN 'FEB' ")	
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=3 THEN 'MAR' ")   
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=4 THEN 'ABR' ")  
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=5 THEN 'MAY' ")  
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=6 THEN 'JUN' ")  
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=7 THEN 'JUL' ") 
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=8 THEN 'AGO' ") 
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=9 THEN 'SEP' ")  
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=10 THEN 'OCT' ")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=11 THEN 'NOV'	")
			lcComandoSeleccionar.AppendLine("			WHEN DatePart(MONTH,ajustes.Fec_Ini)=12 THEN 'DIC' ")
			lcComandoSeleccionar.AppendLine("			END AS mes, ")
			lcComandoSeleccionar.AppendLine("			renglones_ajustes.documento, ")
			lcComandoSeleccionar.AppendLine("			renglones_ajustes.tipo, ")
			lcComandoSeleccionar.AppendLine("			renglones_ajustes.can_art1, ")
			lcComandoSeleccionar.AppendLine("			renglones_ajustes.cos_pro1 ")
			lcComandoSeleccionar.AppendLine("INTO		#tmpTemporal	")
			lcComandoSeleccionar.AppendLine("FROM		articulos, ajustes, renglones_ajustes, almacenes, tipos_ajustes, departamentos ")
			lcComandoSeleccionar.AppendLine("WHERE		")
			lcComandoSeleccionar.AppendLine("			ajustes.documento					=	renglones_ajustes.documento ")
			lcComandoSeleccionar.AppendLine("			AND  renglones_ajustes.cod_alm		=	almacenes.cod_alm")
			lcComandoSeleccionar.AppendLine("			AND renglones_ajustes.cod_tip		=	tipos_ajustes.cod_tip ")
			lcComandoSeleccionar.AppendLine("			AND renglones_ajustes.cod_Art		=	articulos.cod_Art ")
			lcComandoSeleccionar.AppendLine("			AND articulos.cod_dep				=	departamentos.cod_dep ")

			lcComandoSeleccionar.AppendLine(" 			AND    ajustes.fec_ini				BETWEEN " & lcParametro0Desde)
			lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
			lcComandoSeleccionar.AppendLine("           AND ajustes.status					IN (" & lcParametro1Desde & ")")
			lcComandoSeleccionar.AppendLine("           AND articulos.cod_art				BETWEEN " & lcParametro2Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
			lcComandoSeleccionar.AppendLine("           AND articulos.cod_dep				BETWEEN " & lcParametro3Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
			lcComandoSeleccionar.AppendLine("           AND articulos.cod_cla				BETWEEN " & lcParametro4Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
			lcComandoSeleccionar.AppendLine("           AND articulos.cod_tip				BETWEEN " & lcParametro5Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("           AND renglones_ajustes.cod_alm		BETWEEN " & lcParametro6Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine("           AND renglones_ajustes.cod_tip		BETWEEN " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("           AND " & lcParametro7Hasta)
            lcComandoSeleccionar.AppendLine("           AND ajustes.cod_suc			        BETWEEN " & lcParametro8Desde)
            lcComandoSeleccionar.AppendLine("           AND " & lcParametro8Hasta)

			lcComandoSeleccionar.AppendLine("SELECT		")
			lcComandoSeleccionar.AppendLine("			ROW_NUMBER() OVER(PARTITION BY cod_art ORDER BY " & lcOrdenamiento & ") AS 'Renglon', ")
			lcComandoSeleccionar.AppendLine("			cod_art, nom_art, AÑO, num_mes, mes, ")
			lcComandoSeleccionar.AppendLine("			count(documento) as can_doc,  ")
			lcComandoSeleccionar.AppendLine("			SUM(CASE WHEN tipo = 'Entrada' THEN can_art1 ELSE 0 END) As can_ent, ")
			lcComandoSeleccionar.AppendLine("			SUM(CASE WHEN tipo = 'Salida' THEN can_art1 ELSE 0 END) As can_sal,   ")
			lcComandoSeleccionar.AppendLine("			SUM(CASE WHEN tipo = 'Entrada' THEN can_art1 ELSE 0 END)-SUM(CASE WHEN tipo = 'Salida' THEN can_art1 ELSE 0 END)AS dif_can, ")
			lcComandoSeleccionar.AppendLine("			SUM(CASE WHEN tipo='Entrada' THEN (can_art1*cos_pro1) else 0 END) AS mon_ent,   ")
			lcComandoSeleccionar.AppendLine("			SUM(CASE WHEN tipo='Salida' THEN (can_art1*cos_pro1) else 0 END) AS mon_sal,  ")
			lcComandoSeleccionar.AppendLine("			SUM(CASE WHEN tipo='Entrada' THEN (can_art1*cos_pro1) else 0 END)-SUM(CASE WHEN tipo='Salida' THEN (can_art1*cos_pro1) else 0 END) AS dif_mon, ")
			lcComandoSeleccionar.AppendLine("			((count(documento))/(30.0)) AS doc_dia,  ")
			lcComandoSeleccionar.AppendLine("			(SUM(CASE WHEN tipo = 'Entrada' THEN can_art1 ELSE 0 END))/30 As can_ent_dia, ")			
			lcComandoSeleccionar.AppendLine("			(SUM(CASE WHEN tipo = 'Salida' THEN can_art1 ELSE 0 END)) As can_sal_dia,   ")
			lcComandoSeleccionar.AppendLine("			(SUM(CASE WHEN tipo = 'Entrada' THEN can_art1 ELSE 0 END)-SUM(CASE WHEN tipo = 'Salida' THEN can_art1 ELSE 0 END)) AS can_dif_dia, ")
			lcComandoSeleccionar.AppendLine("			(SUM(CASE WHEN tipo='Entrada' THEN (can_art1*cos_pro1) else 0 END))/30 AS mon_ent_dia,   ")
			lcComandoSeleccionar.AppendLine("			(SUM(CASE WHEN tipo='Salida' THEN (can_art1*cos_pro1) else 0 END))/30 AS mon_sal_dia,  ")
			lcComandoSeleccionar.AppendLine("			(SUM(CASE WHEN tipo='Entrada' THEN (can_art1*cos_pro1) else 0 END)-SUM(CASE WHEN tipo='Salida' THEN (can_art1*cos_pro1) else 0 END))/30 AS mon_dif_dia ")			
			lcComandoSeleccionar.AppendLine("FROM ")
			lcComandoSeleccionar.AppendLine("			#tmpTemporal ")
			lcComandoSeleccionar.AppendLine("GROUP BY	 cod_art, nom_art, AÑO, num_mes, mes ")
			lcComandoSeleccionar.AppendLine("ORDER BY	 " & lcOrdenamiento)
				       
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLAjustes_aAñoMes", laDatosReporte)
			
			
			
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLAjustes_aAñoMes.ReportSource = loObjetoReporte

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
' YJP: 25/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' AAP: 29/06/09: Filtro "Sucursal"
'-------------------------------------------------------------------------------------------'
' MAT:  18/02/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'

