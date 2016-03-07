'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLMontosAjustes_AlmacenMes"
'-------------------------------------------------------------------------------------------'
Partial Class rLMontosAjustes_AlmacenMes
    Inherits vis2formularios.frmReporte

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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcComandoSeleccionar As New StringBuilder()
            lcComandoSeleccionar.AppendLine("SELECT ")
            lcComandoSeleccionar.AppendLine("			almacenes.cod_alm, ")
            lcComandoSeleccionar.AppendLine("			almacenes.nom_alm,	 ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='1') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_ene, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='2') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_feb, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='3') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_mar, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='4') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_abr, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='5') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_may, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='6') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_jun, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='7') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_jul, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='8') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_ago, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='9') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_sep, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='10') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_oct, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='11') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_nov, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Entrada' AND DatePart(MONTH,ajustes.Fec_Ini)='12') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As ent_dic, ")

            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='1') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_ene, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='2') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_feb, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='3') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_mar, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='4') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_abr, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='5') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_may, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='6') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_jun, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='7') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_jul, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='8') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_ago, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='9') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_sep, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='10') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_oct, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='11') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_nov, ")
            lcComandoSeleccionar.AppendLine("			SUM((CASE WHEN (renglones_ajustes.tipo = 'Salida' AND DatePart(MONTH,ajustes.Fec_Ini)='12') THEN renglones_ajustes.can_art1*renglones_ajustes.cos_ult1  ELSE 0 END)) As sal_dic  ")
            lcComandoSeleccionar.AppendLine("			INTO #tmpTemporal   ")
            lcComandoSeleccionar.AppendLine("FROM articulos, ajustes, renglones_ajustes, almacenes, tipos_ajustes, departamentos ")
            lcComandoSeleccionar.AppendLine("WHERE ")
            lcComandoSeleccionar.AppendLine("			ajustes.documento				 =	renglones_ajustes.documento	")
            lcComandoSeleccionar.AppendLine("			AND	renglones_ajustes.cod_alm	 =	almacenes.cod_alm  ")
            lcComandoSeleccionar.AppendLine("			AND renglones_ajustes.cod_tip	 =	tipos_ajustes.cod_tip ")
            lcComandoSeleccionar.AppendLine("			AND renglones_ajustes.cod_Art	 =	articulos.cod_Art  ")
            lcComandoSeleccionar.AppendLine("			AND	articulos.cod_dep			 =	departamentos.cod_dep ")
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
            lcComandoSeleccionar.AppendLine("GROUP BY almacenes.cod_alm, almacenes.nom_alm ")

            lcComandoSeleccionar.AppendLine("SELECT	")
            lcComandoSeleccionar.AppendLine("			ROW_NUMBER() OVER(Order By " & lcOrdenamiento & ") as Renglon, ")
            lcComandoSeleccionar.AppendLine("			cod_alm, ")
            lcComandoSeleccionar.AppendLine("			SUBSTRING(nom_alm,1,35) AS nom_alm, ")
            lcComandoSeleccionar.AppendLine("			ent_ene, ent_feb, ent_mar, ent_abr, ent_may, ent_jun, ")
            lcComandoSeleccionar.AppendLine("			ent_jul, ent_ago, ent_sep, ent_oct, ent_nov, ent_dic, ")
            lcComandoSeleccionar.AppendLine("			sal_ene, sal_feb, sal_mar, sal_abr, sal_may, sal_jun, ")
            lcComandoSeleccionar.AppendLine("			sal_jul, sal_ago, sal_sep, sal_oct, sal_nov, sal_dic, ")
            lcComandoSeleccionar.AppendLine("			(ent_ene-sal_ene)	AS 	dif_ene, ")
            lcComandoSeleccionar.AppendLine("			(ent_feb-sal_feb)	AS 	dif_feb, ")
            lcComandoSeleccionar.AppendLine("			(ent_mar-sal_mar)	AS 	dif_mar, ")
            lcComandoSeleccionar.AppendLine("			(ent_abr-sal_abr)	AS 	dif_abr, ")
            lcComandoSeleccionar.AppendLine("			(ent_may-sal_may)	AS 	dif_may, ")
            lcComandoSeleccionar.AppendLine("			(ent_jun-sal_jun)	AS 	dif_jun, ")
            lcComandoSeleccionar.AppendLine("			(ent_jul-sal_jul)	AS 	dif_jul, ")
            lcComandoSeleccionar.AppendLine("			(ent_ago-sal_ago)	AS 	dif_ago, ")
            lcComandoSeleccionar.AppendLine("			(ent_sep-sal_sep)	AS 	dif_sep, ")
            lcComandoSeleccionar.AppendLine("			(ent_oct-sal_oct)	AS 	dif_oct, ")
            lcComandoSeleccionar.AppendLine("			(ent_nov-sal_nov)	AS 	dif_nov, ")
            lcComandoSeleccionar.AppendLine("			(ent_dic-sal_dic)	AS 	dif_dic, ")

            lcComandoSeleccionar.AppendLine("			(ent_ene + ent_feb + ent_mar + ent_abr + ent_may + ent_jun +  ")
            lcComandoSeleccionar.AppendLine("			ent_jul + ent_ago + ent_sep + ent_oct + ent_nov + ent_dic) AS tot_ent,	")

            lcComandoSeleccionar.AppendLine("			(sal_ene + sal_feb + sal_mar + sal_abr + sal_may + sal_jun + ")
            lcComandoSeleccionar.AppendLine("			sal_jul + sal_ago + sal_sep + sal_oct + sal_nov + sal_dic) AS tot_sal, ")

            lcComandoSeleccionar.AppendLine("			((ent_ene-sal_ene) + (ent_feb-sal_feb) + (ent_mar-sal_mar) + (ent_abr-sal_abr) + ")
            lcComandoSeleccionar.AppendLine("			(ent_may-sal_may)  + (ent_jun-sal_jun) + (ent_jul-sal_jul) + (ent_ago-sal_ago) +  ")
            lcComandoSeleccionar.AppendLine("			(ent_sep-sal_sep)  + (ent_oct-sal_oct) + (ent_nov-sal_nov) + (ent_dic-sal_dic)) AS tot_dif	 ")
            lcComandoSeleccionar.AppendLine("FROM		")
            lcComandoSeleccionar.AppendLine("			#tmpTemporal  ")
            lcComandoSeleccionar.AppendLine("ORDER BY	 " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLMontosAjustes_AlmacenMes", laDatosReporte)



            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLMontosAjustes_AlmacenMes.ReportSource = loObjetoReporte

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
' CMS: 29/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 08/04/10: Se cambio el campo cos_ult1 por cos_pro1 
'-------------------------------------------------------------------------------------------'
