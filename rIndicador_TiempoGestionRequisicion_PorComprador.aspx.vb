'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rIndicador_TiempoGestionRequisicion_PorComprador"
'-------------------------------------------------------------------------------------------'
Partial Class rIndicador_TiempoGestionRequisicion_PorComprador
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

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
        Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosIniciales(5)
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
        Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
        Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpRequisiciones(")
            loConsulta.AppendLine("            Requisicion CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("            Estatus CHAR(15) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("            Cod_Pro CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("            Cod_Ven CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("            Req_Fecha DATE);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpRequisiciones(Requisicion, Estatus, Cod_Pro, Cod_Ven, Req_Fecha)")
            loConsulta.AppendLine("SELECT      Documento, Status, Cod_Pro, Cod_Ven, Fec_Ini")
            loConsulta.AppendLine("FROM        Requisiciones")
            loConsulta.AppendLine("WHERE       Requisiciones.Status IN ('Confirmado' , 'Afectado', 'Procesado')")
            loConsulta.AppendLine("        AND Requisiciones.Documento BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("             AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Requisiciones.Fec_Ini BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("             AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Requisiciones.Cod_Pro BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("             AND " & lcParametro2Hasta)
            loConsulta.AppendLine("        AND Requisiciones.Cod_Ven BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("             AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Requisiciones.Cod_Mon BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("             AND " & lcParametro4Hasta)
            If lcParametro5Desde.ToUpper().Trim() = "IGUAL" Then
                loConsulta.AppendLine("    AND Requisiciones.Cod_Rev BETWEEN " & lcParametro6Desde)
                loConsulta.AppendLine("         AND " & lcParametro6Hasta)
            Else
                loConsulta.AppendLine("    AND Requisiciones.Cod_Rev NOT BETWEEN " & lcParametro6Desde)
                loConsulta.AppendLine("         AND " & lcParametro6Hasta)
            End If
            loConsulta.AppendLine("    AND Requisiciones.Cod_Suc BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine("         AND " & lcParametro7Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpTraza(")
            loConsulta.AppendLine("            Requisicion CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("            Presupuesto CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("            Presupuesto_Fecha DATE,")
            loConsulta.AppendLine("            Orden_Compra CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("            OrdenCompra_Fecha DATE,")
            loConsulta.AppendLine("            Recepcion CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("            Recepcion_Fecha DATE,")
            loConsulta.AppendLine("            Compra CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("            Compra_Fecha DATE")
            loConsulta.AppendLine("            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE NONCLUSTERED INDEX Cliclo_Requisicion ON #tmpTraza(Requisicion);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- *************************************************")
            loConsulta.AppendLine("-- Ciclo #1: Req -> Pre -> Ord.C -> Not.Rec -> Com")
            loConsulta.AppendLine("-- (Ciclo completo de compra)")
            loConsulta.AppendLine("-- *************************************************")
            loConsulta.AppendLine("INSERT INTO #tmpTraza ( Requisicion, ")
            loConsulta.AppendLine("                        Presupuesto, Presupuesto_Fecha, ")
            loConsulta.AppendLine("                        Orden_Compra, OrdenCompra_Fecha, ")
            loConsulta.AppendLine("                        Recepcion, Recepcion_Fecha, ")
            loConsulta.AppendLine("                        Compra, Compra_Fecha)")
            loConsulta.AppendLine("SELECT      #tmpRequisiciones.Requisicion       AS Requisicion,")
            loConsulta.AppendLine("            Presupuestos.Documento              AS Presupuesto,")
            loConsulta.AppendLine("            Presupuestos.Fec_Ini                AS Presupuesto_Fecha,")
            loConsulta.AppendLine("            Ordenes_Compras.Documento           AS Orden_Compra,")
            loConsulta.AppendLine("            Ordenes_Compras.Fec_Ini             AS OrdenCompra_Fecha,")
            loConsulta.AppendLine("            Recepciones.Documento               AS Recepcion,")
            loConsulta.AppendLine("            Recepciones.Fec_Ini                 AS Recepcion_Fecha,")
            loConsulta.AppendLine("            Compras.Documento                   AS Compra,")
            loConsulta.AppendLine("            Compras.Fec_Ini                     AS Compra_Fecha")
            loConsulta.AppendLine("FROM        #tmpRequisiciones")
            loConsulta.AppendLine("    JOIN    Renglones_Presupuestos")
            loConsulta.AppendLine("        ON  Renglones_Presupuestos.Doc_Ori = #tmpRequisiciones.Requisicion")
            loConsulta.AppendLine("        AND Renglones_Presupuestos.Tip_Ori = 'Requisiciones'")
            loConsulta.AppendLine("    JOIN    Presupuestos")
            loConsulta.AppendLine("        ON  Presupuestos.Documento = Renglones_Presupuestos.Documento")
            loConsulta.AppendLine("        AND Presupuestos.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loConsulta.AppendLine("    LEFT JOIN Renglones_oCompras")
            loConsulta.AppendLine("        ON  Renglones_oCompras.Doc_Ori = Renglones_Presupuestos.Documento")
            loConsulta.AppendLine("        AND Renglones_oCompras.Ren_Ori = Renglones_Presupuestos.Renglon")
            loConsulta.AppendLine("        AND Renglones_oCompras.Tip_Ori = 'Presupuestos'")
            loConsulta.AppendLine("    LEFT JOIN Ordenes_Compras")
            loConsulta.AppendLine("        ON  Ordenes_Compras.Documento = Renglones_oCompras.Documento")
            loConsulta.AppendLine("        AND Ordenes_Compras.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Recepciones")
            loConsulta.AppendLine("        ON  Renglones_Recepciones.Doc_Ori = Renglones_oCompras.Documento")
            loConsulta.AppendLine("        AND Renglones_Recepciones.Ren_Ori = Renglones_oCompras.Renglon")
            loConsulta.AppendLine("        AND Renglones_Recepciones.Tip_Ori = 'Ordenes_Compras'")
            loConsulta.AppendLine("    LEFT JOIN Recepciones")
            loConsulta.AppendLine("        ON  Recepciones.Documento = Renglones_Recepciones.Documento")
            loConsulta.AppendLine("        AND Recepciones.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Compras")
            loConsulta.AppendLine("        ON  Renglones_Compras.Doc_Ori = Renglones_Recepciones.Documento")
            loConsulta.AppendLine("        AND Renglones_Compras.Ren_Ori = Renglones_Recepciones.Renglon")
            loConsulta.AppendLine("        AND Renglones_Compras.Tip_Ori = 'Recepciones'")
            loConsulta.AppendLine("    LEFT JOIN Compras")
            loConsulta.AppendLine("        ON  Compras.Documento = Renglones_Compras.Documento")
            loConsulta.AppendLine("        AND Compras.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loConsulta.AppendLine("GROUP BY    #tmpRequisiciones.Requisicion,")
            loConsulta.AppendLine("            Presupuestos.Documento,")
            loConsulta.AppendLine("            Presupuestos.Fec_Ini,")
            loConsulta.AppendLine("            Ordenes_Compras.Documento,")
            loConsulta.AppendLine("            Ordenes_Compras.Fec_Ini,")
            loConsulta.AppendLine("            Recepciones.Documento,")
            loConsulta.AppendLine("            Recepciones.Fec_Ini,")
            loConsulta.AppendLine("            Compras.Documento,")
            loConsulta.AppendLine("            Compras.Fec_Ini")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      #tmpRequisiciones.Requisicion                                               AS Requisicion,")
            loConsulta.AppendLine("            #tmpRequisiciones.Estatus                                                   AS Estatus,")
            loConsulta.AppendLine("            #tmpRequisiciones.Cod_Pro                                                   AS Cod_Pro,")
            loConsulta.AppendLine("            Vendedores.Cod_Ven                                                          AS Cod_Ven,")
            loConsulta.AppendLine("            Vendedores.Nom_Ven                                                          AS Nom_Ven,")
            loConsulta.AppendLine("            #tmpRequisiciones.Req_Fecha                                                 AS Requisicion_Fecha,")
            loConsulta.AppendLine("            #tmpTraza.Presupuesto                                                       AS Presupuesto, ")
            loConsulta.AppendLine("            #tmpTraza.Presupuesto_Fecha                                                 AS Presupuesto_Fecha, ")
            loConsulta.AppendLine("            #tmpTraza.Orden_Compra                                                      AS Orden_Compra, ")
            loConsulta.AppendLine("            #tmpTraza.OrdenCompra_Fecha                                                 AS OrdenCompra_Fecha, ")
            loConsulta.AppendLine("            #tmpTraza.Recepcion                                                         AS Recepcion, ")
            loConsulta.AppendLine("            #tmpTraza.Recepcion_Fecha                                                   AS Recepcion_Fecha, ")
            loConsulta.AppendLine("            #tmpTraza.Compra                                                            AS Compra, ")
            loConsulta.AppendLine("            #tmpTraza.Compra_Fecha                                                      AS Compra_Fecha, ")
            loConsulta.AppendLine("            CAST((CASE ")
            loConsulta.AppendLine("                WHEN COALESCE(#tmpTraza.Compra_Fecha, #tmpTraza.Recepcion_Fecha) IS NULL")
            loConsulta.AppendLine("                THEN 0 ELSE 1 END) AS BIT)                                              AS Recibido,")
            loConsulta.AppendLine("            COALESCE(#tmpTraza.Compra_Fecha, #tmpTraza.Recepcion_Fecha)                 AS Fecha_Recibido,")
            loConsulta.AppendLine("            DATEDIFF(DAY, #tmpRequisiciones.Req_Fecha,")
            loConsulta.AppendLine("                        COALESCE(#tmpTraza.Compra_Fecha, ")
            loConsulta.AppendLine("                            #tmpTraza.Recepcion_Fecha,")
            loConsulta.AppendLine("                            #tmpRequisiciones.Req_Fecha))                               AS Dias_Recibido,")
            loConsulta.AppendLine("            (CASE ")
            loConsulta.AppendLine("                WHEN COALESCE(#tmpTraza.Compra_Fecha, #tmpTraza.Recepcion_Fecha) IS NULL")
            loConsulta.AppendLine("                THEN DATEDIFF(DAY, #tmpRequisiciones.Req_Fecha, GETDATE()) ")
            loConsulta.AppendLine("                ELSE 0 END )                                                            AS Dias_Antiguedad")
            loConsulta.AppendLine("FROM        #tmpRequisiciones")
            loConsulta.AppendLine("    JOIN    Vendedores")
            loConsulta.AppendLine("        ON  Vendedores.Cod_Ven = #tmpRequisiciones.Cod_Ven")
            loConsulta.AppendLine("    LEFT JOIN #tmpTraza")
            loConsulta.AppendLine("        ON  #tmpTraza.Requisicion = #tmpRequisiciones.Requisicion")
            loConsulta.AppendLine("ORDER BY    " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rIndicador_TiempoGestionRequisicion_PorComprador", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrIndicador_TiempoGestionRequisicion_PorComprador.ReportSource = loObjetoReporte

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
' RJG: 15/11/14: Codigo inicial
'-------------------------------------------------------------------------------------------'
