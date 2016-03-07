'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOrdenes_Envio_Stratos"
'-------------------------------------------------------------------------------------------'
Partial Class rOrdenes_Envio_Stratos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT    CAST(Facturas.Tie_Log as XML) AS Tiempos,Facturas.Documento")
            loComandoSeleccionar.AppendLine(" INTO #xmlDataTiempos")
            loComandoSeleccionar.AppendLine(" FROM Facturas")
            loComandoSeleccionar.AppendLine("WHERE      Facturas.For_Env <> ''")
            loComandoSeleccionar.AppendLine(" 	        AND Facturas.Documento	BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Fec_Ini	BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Cli	BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Ven	BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Status		IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Tra	BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Rev	BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Suc	BETWEEN" & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" ")

            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT    CAST(Facturas.Embalaje as XML) AS Embalaje,Facturas.Documento")
            loComandoSeleccionar.AppendLine(" INTO #xmlDataEmbalaje")
            loComandoSeleccionar.AppendLine(" FROM Facturas")
            loComandoSeleccionar.AppendLine("WHERE      Facturas.For_Env <> ''")
            loComandoSeleccionar.AppendLine(" 	        AND Facturas.Documento	BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Fec_Ini	BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Cli	BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Ven	BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Status		IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Tra	BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Rev	BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 		    AND Facturas.Cod_Suc	BETWEEN" & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" ")

            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("SELECT	    Documento,")
            loComandoSeleccionar.AppendLine("		    D.C.value('@renglon', 'int')                as ti_renglon,")
            loComandoSeleccionar.AppendLine("           D.C.value('@inicio', 'datetime')			as ti_inicio, ")
            loComandoSeleccionar.AppendLine("           D.C.value('@hor_ini', 'Varchar(50)')		as ti_hor_ini,")
            loComandoSeleccionar.AppendLine("           D.C.value('@fin', 'datetime')			    as ti_fin, ")
            loComandoSeleccionar.AppendLine("           D.C.value('@hor_fin', 'Varchar(50)')		as ti_hor_fin,")
            loComandoSeleccionar.AppendLine("           D.C.value('@operador', 'Varchar(100)')		as ti_operador,")
            loComandoSeleccionar.AppendLine("           D.C.value('@horas', 'datetime')		        as ti_horas	")
            loComandoSeleccionar.AppendLine("INTO #tmpTiempos")
            loComandoSeleccionar.AppendLine("FROM #xmlDataTiempos")
            loComandoSeleccionar.AppendLine("CROSS APPLY Tiempos.nodes('elementos/elemento') D(c) ")
            loComandoSeleccionar.AppendLine("ORDER BY Documento ASC ,ti_renglon ASC  ")
            loComandoSeleccionar.AppendLine(" ")

            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("SELECT	 Documento,")
            loComandoSeleccionar.AppendLine("		 D.C.value('@renglon', 'int') as Em_renglon,")
            loComandoSeleccionar.AppendLine("        D.C.value('@tipo', 'Varchar(50)')		    as Em_tipo,	")
            loComandoSeleccionar.AppendLine("        D.C.value('@cantidad', 'Decimal (28,10)')	as Em_cantidad,")
            loComandoSeleccionar.AppendLine("        D.C.value('@peso', 'Decimal (28,10)')	    as Em_peso,")
            loComandoSeleccionar.AppendLine("        D.C.value('@largo', 'Decimal (28,10)')	    as Em_largo,")
            loComandoSeleccionar.AppendLine("        D.C.value('@ancho', 'Decimal (28,10)')	    as Em_ancho,")
            loComandoSeleccionar.AppendLine("        D.C.value('@alto', 'Decimal (28,10)')	    as Em_alto")
            loComandoSeleccionar.AppendLine("INTO #tmpEmbalaje")
            loComandoSeleccionar.AppendLine("FROM #xmlDataEmbalaje")
            loComandoSeleccionar.AppendLine("CROSS APPLY Embalaje.nodes('elementos/elemento') D(c) ")
            loComandoSeleccionar.AppendLine("ORDER BY Documento ASC ,Em_renglon ASC  ")
            loComandoSeleccionar.AppendLine(" ")

            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("SELECT	Documento,")
            loComandoSeleccionar.AppendLine("       SUM(CASE ")
            loComandoSeleccionar.AppendLine("	        WHEN Em_tipo = 'Paleta'	THEN Em_cantidad  ELSE 0.00")
            loComandoSeleccionar.AppendLine("       END) AS Can_Paletas,")
            loComandoSeleccionar.AppendLine("       SUM(CASE ")
            loComandoSeleccionar.AppendLine("	        WHEN Em_tipo = 'Caja'	THEN Em_cantidad ELSE 0.00")
            loComandoSeleccionar.AppendLine("       END)  AS Can_Cajas")
            loComandoSeleccionar.AppendLine("INTO #tmpPacking")
            loComandoSeleccionar.AppendLine("FROM #tmpEmbalaje")
            loComandoSeleccionar.AppendLine("GROUP BY Documento")
            loComandoSeleccionar.AppendLine(" ")


            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("  SELECT		Facturas.Documento, ")
            loComandoSeleccionar.AppendLine(" 				Facturas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 				Paises.Nom_Pai, ")
            loComandoSeleccionar.AppendLine("               Facturas.Tip_Env, ")
            loComandoSeleccionar.AppendLine("               Facturas.For_Env, ")
            loComandoSeleccionar.AppendLine("               Facturas.Fec_Rec, ")
            loComandoSeleccionar.AppendLine(" 				Facturas.Referencia, ")
            loComandoSeleccionar.AppendLine("               Facturas.Guia, ")
            loComandoSeleccionar.AppendLine("               CAST( '' AS VARCHAR (150)) AS Nombre_Chofer ")
            loComandoSeleccionar.AppendLine(" INTO #tmpEncabezado")
            loComandoSeleccionar.AppendLine(" FROM Facturas")
            loComandoSeleccionar.AppendLine(" JOIN Clientes ON Facturas.Cod_Cli	= Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine(" JOIN Paises ON Paises.Cod_Pai		= Clientes.Cod_Pai")
            loComandoSeleccionar.AppendLine("WHERE      Facturas.For_Env <> ''")
            loComandoSeleccionar.AppendLine(" 	            AND Facturas.Documento	BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Fec_Ini	BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Cli	BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Ven	BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Status		IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Tra	BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Rev	BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Facturas.Cod_Suc	BETWEEN" & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      CONVERT(nchar(30), Facturas.Fec_Ini,112), " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine(" ")


            loComandoSeleccionar.AppendLine("SELECT	     #tmpEncabezado.*,     ")
            loComandoSeleccionar.AppendLine("		     ISNULL(#tmpTiempos.ti_inicio,'')		AS ti_inicio,         ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpTiempos.ti_hor_ini,'00:00')	AS ti_hor_ini,        ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpTiempos.ti_fin,'')			AS ti_fin,            ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpTiempos.ti_hor_fin,'00:00')	AS ti_hor_fin,        ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpTiempos.ti_operador,'')		AS ti_operador,       ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpTiempos.ti_horas,'00:00')	AS ti_horas,          ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpPacking.Can_Paletas,0)		AS Can_Paletas,       ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpPacking.Can_Cajas,0)		AS Can_Cajas,     ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpEmbalaje.Em_peso,0)			AS Em_peso,           ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpEmbalaje.Em_largo,0)		AS Em_largo,      ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpEmbalaje.Em_ancho,0)		AS Em_ancho,      ")
            loComandoSeleccionar.AppendLine("            ISNULL(#tmpEmbalaje.Em_alto,0)			AS Em_alto ")
            loComandoSeleccionar.AppendLine("FROM #tmpEncabezado")
            loComandoSeleccionar.AppendLine("LEFT JOIN #tmpTiempos ON (#tmpTiempos.Documento = #tmpEncabezado.Documento AND #tmpTiempos.ti_renglon = '1')")
            loComandoSeleccionar.AppendLine("LEFT JOIN #tmpEmbalaje ON (#tmpEmbalaje.Documento = #tmpEncabezado.Documento AND #tmpEmbalaje.Em_renglon = '1')")
            loComandoSeleccionar.AppendLine("LEFT JOIN #tmpPacking ON (#tmpPacking.Documento = #tmpEncabezado.Documento)")
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine("DROP TABLE #xmlDataTiempos")
            loComandoSeleccionar.AppendLine("DROP TABLE #xmlDataEmbalaje")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpEncabezado")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpEmbalaje")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpTiempos")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpPacking")

       

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOrdenes_Envio_Stratos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOrdenes_Envio_Stratos.ReportSource = loObjetoReporte

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
'-----------------------------------------------------------------------------------------------'
' Fin del codigo
'-----------------------------------------------------------------------------------------------'
' MAT: 07/09/11: Codigo inicial
'-----------------------------------------------------------------------------------------------'
