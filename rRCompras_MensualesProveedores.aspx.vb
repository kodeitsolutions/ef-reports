'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRCompras_MensualesProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class rRCompras_MensualesProveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

			
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(YEAR, Fec_Ini)   AS  Año_CXP, ")
            loComandoSeleccionar.AppendLine("           DATEPART(MONTH, Fec_Ini)  AS  Mes_CXP, ")
            'loComandoSeleccionar.AppendLine("           SUM((mon_Bas1 + mon_exe))               AS  Mon_Bas1_CXP, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Bru)				            AS  Mon_Bas1_CXP, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Imp1)				            AS  Mon_Imp1_CXP, ")
            loComandoSeleccionar.AppendLine("           SUM((Mon_Bru/30))                      AS  Mon_Bas1_CXPP, ")
            loComandoSeleccionar.AppendLine("           SUM((Mon_Imp1/30))                      AS  Mon_Imp1_CXPP,  ")
            loComandoSeleccionar.AppendLine("           Cod_Pro                                 AS  Cod_Pro_CXP  ")
            loComandoSeleccionar.AppendLine(" INTO	    #Temporal ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine(" WHERE	    Cod_Tip         =   'Fact' ")
            loComandoSeleccionar.AppendLine("           AND Status      <>  'Anulado' ")
            loComandoSeleccionar.AppendLine("           AND Fec_Ini     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Pro     BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Ven     BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Mon     BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Rev     BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine("			Group BY  DATEPART(YEAR, Fec_Ini), DATEPART(MONTH, Fec_Ini), Cod_Pro")
			loComandoSeleccionar.AppendLine("     ")
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(YEAR, Fec_Ini)   AS  Año_DC, ")
            loComandoSeleccionar.AppendLine("			DATEPART(MONTH, Fec_Ini)  AS  Mes_DC, ")
            'loComandoSeleccionar.AppendLine("           SUM((mon_Bas1 + mon_exe))               AS  Mon_Bas1_DC, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Bru)               AS  Mon_Bas1_DC, ")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Imp1)                           AS  Mon_Imp1_DC, ")
            loComandoSeleccionar.AppendLine("           SUM((Mon_Bru/30))                      AS  Mon_Bas1_DCP, ")
            loComandoSeleccionar.AppendLine("           SUM((Mon_Imp1/30))                      AS  Mon_Imp1_DCP,  ")
            loComandoSeleccionar.AppendLine("           Cod_Pro                                 AS  Cod_Pro_DCP  ")
            loComandoSeleccionar.AppendLine(" INTO	    #Temporal2 ")
            loComandoSeleccionar.AppendLine(" FROM	    Devoluciones_Proveedores  ")
            loComandoSeleccionar.AppendLine(" WHERE	    Status         = 'Confirmado' ")
            loComandoSeleccionar.AppendLine("           AND Fec_Ini     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Pro     BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Ven     BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Mon     BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("         AND Cod_Suc BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine("			Group BY  DATEPART(YEAR, Fec_Ini), DATEPART(MONTH, Fec_Ini), Cod_Pro")
			loComandoSeleccionar.AppendLine("     ")
            loComandoSeleccionar.AppendLine(" SELECT    #Temporal.Año_CXP, ")
            loComandoSeleccionar.AppendLine("           #Temporal.Mes_CXP, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(#Temporal.Mon_Bas1_CXP), 0)                                      AS  Mon_Bas1_CXP, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(#Temporal.Mon_Imp1_CXP), 0)                                      AS  Mon_Imp1_CXP, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(#Temporal2.Mon_Bas1_DC), 0)                                      AS  Mon_Bas1_DC, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(#Temporal2.Mon_Imp1_DC), 0)                                      AS  Mon_Imp1_DC, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#Temporal.Mon_Bas1_CXP, 0) - ISNULL(#temporal2.Mon_Bas1_DC, 0))  AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#Temporal.Mon_Imp1_CXP, 0) - ISNULL(#temporal2.Mon_Imp1_DC, 0))  AS  Imp_Net, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(Mon_Bas1_CXPP), 0)                                               AS  Mon_Bas1_CXPP, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(Mon_Imp1_CXPP), 0)                                               AS  Mon_Imp1_CXPP, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(Mon_Bas1_DCP), 0)                                                AS  Mon_Bas1_DCP, ")
            loComandoSeleccionar.AppendLine("           ISNULL(SUM(Mon_Imp1_DCP), 0)                                                AS  Mon_Imp1_DCP, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(Mon_Bas1_CXPP, 0) - ISNULL(Mon_Bas1_DCP, 0))                     AS  Mon_Netp, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(Mon_Imp1_CXPP, 0) - ISNULL(Mon_Imp1_DCP, 0))                     AS  Imp_Netp, ")
            loComandoSeleccionar.AppendLine("           #Temporal.Cod_Pro_CXP                                                       AS  Cod_Pro_CXP,  ")
            loComandoSeleccionar.AppendLine("           #Temporal2.Cod_Pro_DCP                                                      AS  Cod_Pro_DCP, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro                                                         AS  Nom_Pro ")
            loComandoSeleccionar.AppendLine(" FROM      #Temporal")
            loComandoSeleccionar.AppendLine("           LEFT JOIN #Temporal2 AS #Temporal2 ON ((#Temporal.Año_CXP = #Temporal2.Año_DC) AND (#temporal.Mes_CXP = #temporal2.Mes_DC) AND (#temporal.Cod_Pro_CXP = #temporal2.Cod_Pro_DCP))  ")
            loComandoSeleccionar.AppendLine("           JOIN Proveedores ON Proveedores.Cod_Pro = #temporal.Cod_Pro_CXP ")
            loComandoSeleccionar.AppendLine(" GROUP BY  #Temporal.Año_CXP, #Temporal.Mes_CXP,  #Temporal.Cod_Pro_CXP, #Temporal2.Cod_Pro_DCP, Proveedores.Nom_Pro ")
            loComandoSeleccionar.AppendLine(" ORDER BY  #Temporal.Cod_Pro_CXP, " & lcOrdenamiento )
'me.mEscribirConsulta(loComandoSeleccionar.ToString)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRCompras_MensualesProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRCompras_MensualesProveedores.ReportSource = loObjetoReporte

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
' CMS: 28/05/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' JJD: 22/04/10: Normalizacion del codigo
'-------------------------------------------------------------------------------------------'
' CMS: 28/04/10: Se Ajusto el ordenamiento del nom_pro a la fecha
'-------------------------------------------------------------------------------------------'
' MAT: 28/02/11: Ajuste del Select. Nueva vista de Diseño
'-------------------------------------------------------------------------------------------'
' MAT: 18/05/11: Agregado Filtro "Sucursal:"
'-------------------------------------------------------------------------------------------'