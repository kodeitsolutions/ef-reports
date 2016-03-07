'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_sAlmacenConsolidado"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_sAlmacenConsolidado
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
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

            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)
            Dim lcExisiencia As String = ""

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" DECLARE @lcNumAlmacenes INT")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT	TOP 9")
            loComandoSeleccionar.AppendLine(" 		Almacenes.Cod_Alm,")
            loComandoSeleccionar.AppendLine(" 		Almacenes.Nom_Alm")
            loComandoSeleccionar.AppendLine(" INTO	#tempALMACENES")
            loComandoSeleccionar.AppendLine(" FROM	Almacenes")
            loComandoSeleccionar.AppendLine(" WHERE	Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @lcNumAlmacenes = (SELECT Count(*) FROM #tempALMACENES)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT	")
            loComandoSeleccionar.AppendLine(" 		Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Cod_Uni1, ")
            Select Case lcParametro11Desde
                Case "Actual"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Act1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Act1"
                Case "Comprometida"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Ped1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Ped1"
                Case "Cotizada"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Cot1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Cot1"
                Case "En_Produccion"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Pro1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Pro1"
                Case "Por_Llegar"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Por1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Por1"
                Case "Por_Despachar"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Des1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Des1"
                Case "Por_Distribuir"
                    loComandoSeleccionar.AppendLine("	Renglones_Almacenes.Exi_Dis1 AS Exi_Act1,")
                    lcExisiencia = "Renglones_Almacenes.Exi_Dis1"
            End Select
            loComandoSeleccionar.AppendLine(" 		#tempALMACENES.Cod_Alm, ")
            loComandoSeleccionar.AppendLine(" 		#tempALMACENES.Nom_Alm")
            loComandoSeleccionar.AppendLine(" INTO	#tempDATOS")
            loComandoSeleccionar.AppendLine(" FROM	Articulos, ")
            loComandoSeleccionar.AppendLine(" 	  	Renglones_Almacenes, ")
            loComandoSeleccionar.AppendLine("  		Departamentos, ")
            loComandoSeleccionar.AppendLine("  		Secciones, ")
            loComandoSeleccionar.AppendLine("  		Marcas, ")
            loComandoSeleccionar.AppendLine("  		Tipos_Articulos, ")
            loComandoSeleccionar.AppendLine("  		Clases_Articulos,")
            loComandoSeleccionar.AppendLine(" 		#tempALMACENES")
            loComandoSeleccionar.AppendLine(" WHERE		Articulos.Cod_Art = Renglones_Almacenes.Cod_Art ")
            loComandoSeleccionar.AppendLine("  		AND Renglones_Almacenes.Cod_Alm = #tempALMACENES.Cod_Alm ")
            loComandoSeleccionar.AppendLine("  		AND Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("  		AND Articulos.Cod_Sec = Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("  		AND Secciones.cod_dep = Departamentos.cod_dep ")
            loComandoSeleccionar.AppendLine("  		AND Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("  		AND Articulos.Cod_Tip = Tipos_Articulos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("  		AND Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde & "	AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Articulos.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Departamentos.Cod_Dep BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Secciones.Cod_Sec BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Marcas.Cod_Mar BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Tipos_Articulos.Cod_Tip BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Clases_Articulos.Cod_Cla BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("      	AND Articulos.Cod_Ubi between " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("      	AND Articulos.Cod_Pro between " & lcParametro9Desde & " AND " & lcParametro9Hasta)

            Select Case lcParametro10Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("      ")
                Case "Igual"
                    loComandoSeleccionar.AppendLine("     AND " & lcExisiencia & "          =   0  ")
                Case "Mayor"
                    loComandoSeleccionar.AppendLine("     AND " & lcExisiencia & "          >   0  ")
                Case "Menor"
                    loComandoSeleccionar.AppendLine("     AND " & lcExisiencia & "          <   0  ")
                Case "Maximo"
                    loComandoSeleccionar.AppendLine("     AND Articulos.Exi_Max           =   " & lcExisiencia & "  ")
                Case "Minimo"
                    loComandoSeleccionar.AppendLine("     And Articulos.Exi_Min           =   " & lcExisiencia & "  ")
                Case "Pedido"
                    loComandoSeleccionar.AppendLine("     And Articulos.Exi_pto           =   " & lcExisiencia & "  ")
            End Select
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT	")
            loComandoSeleccionar.AppendLine(" 		Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Cod_Uni1, ")
            loComandoSeleccionar.AppendLine(" 		0 AS Exi_Act1,")
            loComandoSeleccionar.AppendLine(" 		#tempALMACENES.Cod_Alm, ")
            loComandoSeleccionar.AppendLine(" 		#tempALMACENES.Nom_Alm  ")
            loComandoSeleccionar.AppendLine(" FROM	Articulos, ")
            loComandoSeleccionar.AppendLine("  		Departamentos, ")
            loComandoSeleccionar.AppendLine("  		Secciones, ")
            loComandoSeleccionar.AppendLine("  		Marcas, ")
            loComandoSeleccionar.AppendLine("  		Tipos_Articulos, ")
            If lcParametro10Desde <> "Todos" Then
                loComandoSeleccionar.AppendLine("  		Renglones_Almacenes,")
            End If
            loComandoSeleccionar.AppendLine("  		Clases_Articulos")
            loComandoSeleccionar.AppendLine(" CROSS JOIN #tempALMACENES")
            loComandoSeleccionar.AppendLine(" WHERE		Articulos.Cod_Dep = Departamentos.Cod_Dep ")
            loComandoSeleccionar.AppendLine("  		AND Articulos.Cod_Sec = Secciones.Cod_Sec ")
            loComandoSeleccionar.AppendLine("  		AND Secciones.cod_dep = Departamentos.cod_dep ")
            loComandoSeleccionar.AppendLine("  		AND Articulos.Cod_Mar = Marcas.Cod_Mar ")
            loComandoSeleccionar.AppendLine("  		AND Articulos.Cod_Tip = Tipos_Articulos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("  		AND Articulos.Cod_Cla = Clases_Articulos.Cod_Cla ")
            loComandoSeleccionar.AppendLine(" 		AND Articulos.Cod_Art BETWEEN " & lcParametro0Desde & "	AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Articulos.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 		AND Departamentos.Cod_Dep BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Secciones.Cod_Sec BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Marcas.Cod_Mar BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Tipos_Articulos.Cod_Tip BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Clases_Articulos.Cod_Cla BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("      	AND Articulos.Cod_Ubi between " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("      	AND Articulos.Cod_Pro between " & lcParametro9Desde & " AND " & lcParametro9Hasta)

            Select Case lcParametro10Desde
                Case "Todos"
                    loComandoSeleccionar.AppendLine("      ")
                Case "Igual"
                    loComandoSeleccionar.AppendLine("  		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("       AND " & lcExisiencia & "          =   0  ")
                    loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Art = Renglones_Almacenes.Cod_Art ")
                Case "Mayor"
                    loComandoSeleccionar.AppendLine("  		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("       AND " & lcExisiencia & "          >   0  ")
                    loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Art = Renglones_Almacenes.Cod_Art ")
                Case "Menor"
                    loComandoSeleccionar.AppendLine("  		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("       AND " & lcExisiencia & "          <   0  ")
                    loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Art = Renglones_Almacenes.Cod_Art ")
                Case "Maximo"
                    loComandoSeleccionar.AppendLine("  		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("       AND Articulos.Exi_Max           =   " & lcExisiencia & "  ")
                    loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Art = Renglones_Almacenes.Cod_Art ")
                Case "Minimo"
                    loComandoSeleccionar.AppendLine("  		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("       And Articulos.Exi_Min           =   " & lcExisiencia & "  ")
                    loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Art = Renglones_Almacenes.Cod_Art ")
                Case "Pedido"
                    loComandoSeleccionar.AppendLine("  		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("       And Articulos.Exi_pto           =   " & lcExisiencia & "  ")
                    loComandoSeleccionar.AppendLine("       AND Articulos.Cod_Art = Renglones_Almacenes.Cod_Art ")
            End Select

            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTablaResult TABLE(")
            loComandoSeleccionar.AppendLine(" 	Cod_Art			VARCHAR(30)		NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	Nom_Art			VARCHAR(100)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	Cod_Uni			VARCHAR(10)		NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	Arr_Cod_Alm		VARCHAR(120)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	Arr_Nom_Alm		VARCHAR(1020)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	Arr_Exi_Act		VARCHAR(220)	NOT NULL,")
            loComandoSeleccionar.AppendLine(" 	NumDecExi		INT	            NOT NULL")
            loComandoSeleccionar.AppendLine(" )")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTab_CodArt	VARCHAR(30)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTab_NomArt	VARCHAR(100)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTab_CodUni	VARCHAR(10)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTab_CodAlm	VARCHAR(10)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTab_NomAlm	VARCHAR(100)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcTab_ExiAct	FLOAT")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" DECLARE @lcArrCodAlm	VARCHAR(120)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcArrNomAlm	VARCHAR(1020)")
            loComandoSeleccionar.AppendLine(" DECLARE @lcArrExiAct	VARCHAR(220)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" DECLARE @lcCountRecord	INT")
            loComandoSeleccionar.AppendLine(" SET @lcCountRecord = 0")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" DECLARE CURSOR_RESULT CURSOR FOR")
            loComandoSeleccionar.AppendLine(" 	SELECT")
            loComandoSeleccionar.AppendLine(" 			#tempDATOS.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 			#tempDATOS.Nom_Art, ")
            loComandoSeleccionar.AppendLine(" 			#tempDATOS.Cod_Uni1, ")
            loComandoSeleccionar.AppendLine(" 			sum(#tempDATOS.Exi_Act1) AS Exi_Act1,")
            loComandoSeleccionar.AppendLine(" 			#tempDATOS.Cod_Alm, ")
            loComandoSeleccionar.AppendLine(" 			#tempDATOS.Nom_Alm")
            loComandoSeleccionar.AppendLine(" 	FROM	#tempDATOS")
            loComandoSeleccionar.AppendLine(" 	GROUP BY #tempDATOS.Cod_Art,#tempDATOS.Nom_Art,#tempDATOS.Cod_Uni1,#tempDATOS.Cod_Alm,#tempDATOS.Nom_Alm")
            loComandoSeleccionar.AppendLine("	ORDER BY #tempDATOS.Cod_Art ASC")
            loComandoSeleccionar.AppendLine(" OPEN CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" FETCH NEXT FROM CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" INTO @lcTab_CodArt,@lcTab_NomArt,@lcTab_CodUni,@lcTab_ExiAct,@lcTab_CodAlm,@lcTab_NomAlm")
            loComandoSeleccionar.AppendLine(" WHILE @@fetch_status = 0")
            loComandoSeleccionar.AppendLine(" BEGIN")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	SET @lcCountRecord = @lcCountRecord + 1")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	IF ( @lcCountRecord = 1 )")
            loComandoSeleccionar.AppendLine(" 	BEGIN")
            loComandoSeleccionar.AppendLine(" 		SET @lcArrCodAlm = ''")
            loComandoSeleccionar.AppendLine(" 		SET @lcArrNomAlm = ''")
            loComandoSeleccionar.AppendLine(" 		SET @lcArrExiAct = ''")
            loComandoSeleccionar.AppendLine(" 	END")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	SET @lcArrCodAlm = @lcArrCodAlm + RTRIM(@lcTab_CodAlm) + '#'")
            loComandoSeleccionar.AppendLine(" 	SET @lcArrNomAlm = @lcArrNomAlm + RTRIM(@lcTab_NomAlm) + '#'")
            loComandoSeleccionar.AppendLine(" 	SET @lcArrExiAct = @lcArrExiAct + CAST(@lcTab_ExiAct AS VARCHAR(20)) + '#'")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	IF ( @lcCountRecord = @lcNumAlmacenes )")
            loComandoSeleccionar.AppendLine(" 	BEGIN")
            loComandoSeleccionar.AppendLine(" 		SET @lcCountRecord = 0")
            loComandoSeleccionar.AppendLine(" 		INSERT @lcTablaResult(Cod_Art,Nom_Art,Cod_Uni,Arr_Cod_Alm,Arr_Nom_Alm,Arr_Exi_Act,NumDecExi)")
            loComandoSeleccionar.AppendLine(" 		VALUES(@lcTab_CodArt,@lcTab_NomArt,@lcTab_CodUni,@lcArrCodAlm,@lcArrNomAlm,@lcArrExiAct," & cusAplicacion.goOpciones.pnDecimalesParaCantidad & ")")
            loComandoSeleccionar.AppendLine(" 	END")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" 	FETCH NEXT FROM CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" 	INTO @lcTab_CodArt,@lcTab_NomArt,@lcTab_CodUni,@lcTab_ExiAct,@lcTab_CodAlm,@lcTab_NomAlm")
            loComandoSeleccionar.AppendLine(" END")
            loComandoSeleccionar.AppendLine(" CLOSE CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" DEALLOCATE CURSOR_RESULT")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT * FROM @lcTablaResult")
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)
           

            Dim loServicios As New cusDatos.goDatos
            
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_sAlmacenConsolidado", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_sAlmacenConsolidado.ReportSource = loObjetoReporte

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
' Douglas Cortez:  27/05/2010 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT:  03/02/11 : Ajuste del Select. No mostraba información
'-------------------------------------------------------------------------------------------'
