'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPedidos_vMes"
'-------------------------------------------------------------------------------------------'
Partial Class rPedidos_vMes
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
												  
            
			lcComandoSeleccionar.AppendLine(" SELECT	SUBSTRING(vendedores.nom_ven,0,30) as nom_ven, ")
			lcComandoSeleccionar.AppendLine("			sum(case when DatePart(MONTH,pedidos.Fec_Ini) = 1 then renglones_pedidos.can_art1 else 0 end ) as ped_ene, ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 2 Then renglones_pedidos.can_art1 Else 0 End) As ped_feb, ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 3 Then renglones_pedidos.can_art1 Else 0 End) As ped_mar,   ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 4 Then renglones_pedidos.can_art1 Else 0 End) As ped_abr, ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 5 Then renglones_pedidos.can_art1 Else 0 End) As ped_may, ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 6 Then renglones_pedidos.can_art1 Else 0 End) As ped_jun, ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 7 Then renglones_pedidos.can_art1 Else 0 End) As ped_jul, ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 8 Then renglones_pedidos.can_art1 Else 0 End) As ped_ago,  ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 9 Then renglones_pedidos.can_art1 Else 0 End) As ped_sep,  ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 10 Then renglones_pedidos.can_art1 Else 0 End) As ped_oct, ")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 11 Then renglones_pedidos.can_art1 Else 0 End) As ped_nov,	")
			lcComandoSeleccionar.AppendLine("			Sum(Case when DatePart(MONTH,pedidos.Fec_Ini) = 12 Then renglones_pedidos.can_art1 Else 0 End) As ped_dic	")
			lcComandoSeleccionar.AppendLine("into #temporal	 ")
			lcComandoSeleccionar.AppendLine("from renglones_pedidos, pedidos, articulos, vendedores, departamentos ")
			lcComandoSeleccionar.AppendLine(" WHERE	")
			lcComandoSeleccionar.AppendLine(" 	pedidos.documento=renglones_pedidos.documento	")
			lcComandoSeleccionar.AppendLine(" 	AND vendedores.cod_ven=pedidos.cod_ven  ")
			lcComandoSeleccionar.AppendLine(" 	AND renglones_pedidos.cod_art = articulos.cod_art  ")
			lcComandoSeleccionar.AppendLine(" 	AND articulos.cod_dep = departamentos.cod_dep  ")
			lcComandoSeleccionar.AppendLine(" 			AND    pedidos.fec_ini        Between " & lcParametro0Desde)
			lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
			lcComandoSeleccionar.AppendLine(" 			AND articulos.cod_art     Between " & lcParametro1Desde)
			lcComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
			lcComandoSeleccionar.AppendLine("           AND pedidos.cod_cli       Between " & lcParametro2Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
			lcComandoSeleccionar.AppendLine("           AND pedidos.cod_ven      Between " & lcParametro3Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
			lcComandoSeleccionar.AppendLine("           AND articulos.cod_dep   Between " & lcParametro4Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
			lcComandoSeleccionar.AppendLine("           AND articulos.cod_tip   Between " & lcParametro5Desde)
			lcComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("           AND pedidos.status   IN (" & lcParametro6Desde & ")")
            lcComandoSeleccionar.AppendLine("           AND pedidos.Cod_Rev between " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("    	    AND " & lcParametro7Hasta)
			lcComandoSeleccionar.AppendLine("GROUP BY   vendedores.nom_ven ")

			lcComandoSeleccionar.AppendLine("select		nom_ven,	")
			lcComandoSeleccionar.AppendLine("			ped_ene, ")
			lcComandoSeleccionar.AppendLine("			ped_feb,  ")
			lcComandoSeleccionar.AppendLine("			ped_mar,  ")
			lcComandoSeleccionar.AppendLine("			ped_abr,  ")
			lcComandoSeleccionar.AppendLine("			ped_may,  ")
			lcComandoSeleccionar.AppendLine("			ped_jun,	")
			lcComandoSeleccionar.AppendLine("			ped_jul,   ")
			lcComandoSeleccionar.AppendLine("			ped_ago,   ")
			lcComandoSeleccionar.AppendLine("			ped_sep,   ")
			lcComandoSeleccionar.AppendLine("			ped_oct,	")
			lcComandoSeleccionar.AppendLine("			ped_nov,   ")
			lcComandoSeleccionar.AppendLine("			ped_dic, ")
			lcComandoSeleccionar.AppendLine("			 (ped_ene+ped_feb+ped_mar+ped_abr+ped_may+ped_jun+ped_jul+ped_ago+ped_sep+ped_oct+ped_nov+ped_dic)as total ")
			lcComandoSeleccionar.AppendLine(" from #temporal ")
			lcComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			'lcComandoSeleccionar.AppendLine("ORDER BY   nom_ven ")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPedidos_vMes", laDatosReporte)
			
			
			
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPedidos_vMes.ReportSource = loObjetoReporte

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
' YJP: 08/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 16/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' MAT: 18/02/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------'