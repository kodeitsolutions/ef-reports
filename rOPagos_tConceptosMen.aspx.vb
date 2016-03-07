'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOPagos_tConceptosMen"
'-------------------------------------------------------------------------------------------'
Partial Class rOPagos_tConceptosMen
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            loComandoSeleccionar.AppendLine("SELECT      Conceptos.Cod_Con, ")
            loComandoSeleccionar.AppendLine("              Conceptos.Nom_Con,  ")
            loComandoSeleccionar.AppendLine("              CASE  ")
            loComandoSeleccionar.AppendLine("                  WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 1 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab)) ")
            loComandoSeleccionar.AppendLine("                  else 0  ")
            loComandoSeleccionar.AppendLine("              END as Ene, ")
            loComandoSeleccionar.AppendLine("              CASE  ")
            loComandoSeleccionar.AppendLine("                  WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 2 THEN ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                  else 0 ")
            loComandoSeleccionar.AppendLine("              END as Feb, ")
            loComandoSeleccionar.AppendLine("              CASE  ")
            loComandoSeleccionar.AppendLine("                  WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 3 THEN ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                  else 0 ")
            loComandoSeleccionar.AppendLine("              END as Mar, ")
            loComandoSeleccionar.AppendLine("              CASE  ")
            loComandoSeleccionar.AppendLine("                  WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 4 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                  else 0 ")
            loComandoSeleccionar.AppendLine("              END as Abr,  ")
            loComandoSeleccionar.AppendLine("              CASE  ")
            loComandoSeleccionar.AppendLine("                  WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 5 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                  else 0 ")
            loComandoSeleccionar.AppendLine("              END as May, ")
            loComandoSeleccionar.AppendLine("              CASE  ")
            loComandoSeleccionar.AppendLine("                  WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 6 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                  else 0 ")
            loComandoSeleccionar.AppendLine("              END as Jun, ")
            loComandoSeleccionar.AppendLine("              CASE  ")
            loComandoSeleccionar.AppendLine("                  WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 7 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                  else 0 ")
            loComandoSeleccionar.AppendLine("            END as Jul, ")
            loComandoSeleccionar.AppendLine("            CASE  ")
            loComandoSeleccionar.AppendLine("                  WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 8 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                  else 0 ")
            loComandoSeleccionar.AppendLine("            END as Ago, ")
            loComandoSeleccionar.AppendLine("            CASE  ")
            loComandoSeleccionar.AppendLine("                  WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 9 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                  else 0 ")
            loComandoSeleccionar.AppendLine("             END as Sep, ")
            loComandoSeleccionar.AppendLine("             CASE  ")
            loComandoSeleccionar.AppendLine("                 WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 10 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                 else 0 ")
            loComandoSeleccionar.AppendLine("             END as Oct, ")
            loComandoSeleccionar.AppendLine("             CASE  ")
            loComandoSeleccionar.AppendLine("                 WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 11 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                 else 0 ")
            loComandoSeleccionar.AppendLine("             END as Nov, ")
            loComandoSeleccionar.AppendLine("             CASE  ")
            loComandoSeleccionar.AppendLine("                 WHEN DATEPART(MONTH, Ordenes_Pagos.Fec_ini) = 12 THEN  ((Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab))  ")
            loComandoSeleccionar.AppendLine("                 else 0 ")
            loComandoSeleccionar.AppendLine("             END as Dic, ")
            loComandoSeleccionar.AppendLine("             (Renglones_oPagos.Mon_Deb - Renglones_oPagos.Mon_hab) AS Total ")
            loComandoSeleccionar.AppendLine("             INTO        #tmpTemporal ")
            loComandoSeleccionar.AppendLine("FROM        Ordenes_Pagos, ")
            loComandoSeleccionar.AppendLine("            Renglones_oPagos, ")
            loComandoSeleccionar.AppendLine("            Conceptos, ")
            loComandoSeleccionar.AppendLine("            Proveedores ")
            loComandoSeleccionar.AppendLine("            WHERE(Ordenes_Pagos.Documento = Renglones_oPagos.Documento) ")
            loComandoSeleccionar.AppendLine("            And Renglones_oPagos.Cod_Con = Conceptos.Cod_Con ")
            loComandoSeleccionar.AppendLine("            And Ordenes_Pagos.Cod_Pro = Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Documento Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Fec_Ini   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_Pro   Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Status    IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           And Conceptos.Cod_Con      Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Ordenes_Pagos.Cod_rev   Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            
            loComandoSeleccionar.AppendLine("SELECT  ")
            loComandoSeleccionar.AppendLine("              #tmpTemporal.Cod_Con, ")
            loComandoSeleccionar.AppendLine("              #tmpTemporal.Nom_Con, ")
            loComandoSeleccionar.AppendLine("              sum(ene) as Ene, ")
            loComandoSeleccionar.AppendLine("              sum(feb) as Feb, ")
            loComandoSeleccionar.AppendLine("              sum(mar) as Mar, ")
            loComandoSeleccionar.AppendLine("              sum(abr) as Abr, ")
            loComandoSeleccionar.AppendLine("              sum(may) as May, ")
            loComandoSeleccionar.AppendLine("              sum(jun) as Jun, ")
            loComandoSeleccionar.AppendLine("              sum(jul) as Jul, ")
            loComandoSeleccionar.AppendLine("              sum(ago) as Ago, ")
            loComandoSeleccionar.AppendLine("              sum(sep) as Sep, ")
            loComandoSeleccionar.AppendLine("              sum(oct) as Oct, ")
            loComandoSeleccionar.AppendLine("              sum(nov) as Nov, ")
            loComandoSeleccionar.AppendLine("              sum(dic) as Dic, ")
            loComandoSeleccionar.AppendLine("              sum(total) as Total")
            loComandoSeleccionar.AppendLine("FROM #tmpTemporal ")
            loComandoSeleccionar.AppendLine("Group By     ")
            loComandoSeleccionar.AppendLine("      Cod_Con, ")
            loComandoSeleccionar.AppendLine("      Nom_Con ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOPagos_tConceptosMen", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOPagos_tConceptosMen.ReportSource = loObjetoReporte


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
' YJP: 14/05/09: Agregar filtro Revisión
'-------------------------------------------------------------------------------------------'
' CMS: 26/03/10: Se cambio la funcion DATEPART por DATEPART
'-------------------------------------------------------------------------------------------'
' MAT: 19/05/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'