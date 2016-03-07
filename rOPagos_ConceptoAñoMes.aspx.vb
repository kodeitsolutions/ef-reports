'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOPagos_ConceptoAñoMes"
'-------------------------------------------------------------------------------------------'
Partial Class rOPagos_ConceptoAñoMes

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(YEAR, Ordenes_Pagos.Fec_Ini)   AS  Anno,")
            loComandoSeleccionar.AppendLine("			DATEPART(MONTH,Ordenes_Pagos.Fec_Ini)   AS  Num_Mes,")
            loComandoSeleccionar.AppendLine("   		CASE    WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=1 THEN 'ENE' ")
            loComandoSeleccionar.AppendLine("   		        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=2 THEN 'FEB'	")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=3 THEN 'MAR' ")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=4 THEN 'ABR' ")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=5 THEN 'MAY' ")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=6 THEN 'JUN' ")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=7 THEN 'JUL' ")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=8 THEN 'AGO' ")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=9 THEN 'SEP' ")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=10 THEN 'OCT' ")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=11 THEN 'NOV' ")
            loComandoSeleccionar.AppendLine("			        WHEN DatePart(MONTH,Ordenes_Pagos.Fec_Ini)=12 THEN 'DIC' ")
            loComandoSeleccionar.AppendLine("   		END AS Mes, ")
            loComandoSeleccionar.AppendLine("           (Renglones_OPagos.Cod_Con)  AS  Cod_Con,")
            loComandoSeleccionar.AppendLine("           (Conceptos.Nom_Con)         AS  Nom_Con,")
            loComandoSeleccionar.AppendLine("           (Ordenes_Pagos.Documento)   AS  Cantidad,")
            loComandoSeleccionar.AppendLine("           (Ordenes_Pagos.Mon_Imp)     AS  Mon_Imp,")
            loComandoSeleccionar.AppendLine("           (Renglones_OPagos.Mon_Deb)  AS  Mon_Deb,")
            loComandoSeleccionar.AppendLine("           (Renglones_OPagos.Mon_Hab)  AS  Mon_Hab")
            loComandoSeleccionar.AppendLine(" INTO      #tmpDocumentos")
            loComandoSeleccionar.AppendLine(" FROM      Ordenes_Pagos, Renglones_OPagos, Conceptos")
            loComandoSeleccionar.AppendLine(" WHERE     Ordenes_Pagos.Documento         =   Renglones_OPagos.Documento ")
            loComandoSeleccionar.AppendLine("           AND Conceptos.Cod_Con		    =   Renglones_oPagos.Cod_Con  ")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Fec_Ini       Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Pro       Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Status        IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Mon       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_OPagos.Cod_Con    Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)

            loComandoSeleccionar.AppendLine(" SELECT    Anno,")
            loComandoSeleccionar.AppendLine("           Mes,")
            loComandoSeleccionar.AppendLine("           Num_Mes,")
            loComandoSeleccionar.AppendLine("           Cod_Con,")
            loComandoSeleccionar.AppendLine("           Nom_Con,")
            loComandoSeleccionar.AppendLine("           COUNT(Cantidad) AS Cantidad,")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Imp)    AS Mon_Imp,")
            loComandoSeleccionar.AppendLine("           SUM (Mon_Deb)   AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("           SUM(Mon_Hab)    AS Mon_Hab")
            loComandoSeleccionar.AppendLine(" FROM      #tmpDocumentos")
            loComandoSeleccionar.AppendLine(" GROUP BY  Cod_Con, Nom_Con, Anno, num_mes, Mes")
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_ConceptoAñoMes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_ConceptoAñoMes.ReportSource = loObjetoReporte

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
' CMS: 18/05/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' JJD: 16/01/10: Normalizacion del codigo
'-------------------------------------------------------------------------------------------'
' MAT:  06/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'