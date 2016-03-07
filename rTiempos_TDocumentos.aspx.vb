'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTiempos_TDocumentos"
'-------------------------------------------------------------------------------------------'
Partial Class rTiempos_TDocumentos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            '-------------------------------------------------------------------------------------------------------
            'Creando el cursor de trabajo
            '-------------------------------------------------------------------------------------------------------
           
            loComandoSeleccionar.AppendLine(" DECLARE @lnCantidad AS INT ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnMinimo  AS INT ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnMaximo AS INT")
            loComandoSeleccionar.AppendLine(" DECLARE @lnPromedio  AS INT ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnSumatoria  AS INT ")   
            loComandoSeleccionar.AppendLine(" DECLARE @lcNombre AS VARCHAR(100) ")
            loComandoSeleccionar.AppendLine(" DECLARE @lcNombreCampo AS VARCHAR(50) ")
            loComandoSeleccionar.AppendLine(" DECLARE @lcSelectExecute AS VARCHAR(5000) ")
            loComandoSeleccionar.AppendLine(" DECLARE @ldFechaInicial AS DATETIME ")
            loComandoSeleccionar.AppendLine(" DECLARE @ldFechaFinal AS DATETIME ")
            loComandoSeleccionar.AppendLine(" DECLARE @ldFechaInicial02 AS VARCHAR(50) ")
            loComandoSeleccionar.AppendLine(" DECLARE @ldFechaFinal02 AS VARCHAR(50) ")

            loComandoSeleccionar.AppendLine(" SET @lcNombreCampo = 'Reg_Ini' ")

            loComandoSeleccionar.AppendLine(" SELECT    TABLE_NAME                                  AS  Nombre, ")
            loComandoSeleccionar.AppendLine("           " & lcParametro0Desde & "					AS  Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           " & lcParametro0Hasta & "					AS  Fec_Fin, ")
            loComandoSeleccionar.AppendLine("   		CAST('19900101' AS DATETIME)                AS  Reg_Ini, ")
            loComandoSeleccionar.AppendLine("           CAST('19900101' AS DATETIME)                AS  Reg_Fin, ")
            loComandoSeleccionar.AppendLine("           CAST(0 AS INT)                              AS  Cuantos, ")
            loComandoSeleccionar.AppendLine("           CAST(0 AS INT)                              AS  Minimo_Seg, ")
            loComandoSeleccionar.AppendLine("           CAST(0 AS INT)                              AS  Maximo_Seg, ")
            loComandoSeleccionar.AppendLine("           CAST(0 AS INT)                              AS  Promedio_Seg,")
            loComandoSeleccionar.AppendLine("           CAST(0 AS INT)                              AS  Sumatoria_Seg ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTablas ")
            loComandoSeleccionar.AppendLine("           FROM INFORMATION_SCHEMA.COLUMNS ")
            loComandoSeleccionar.AppendLine(" WHERE     COLUMN_NAME LIKE @lcNombreCampo ")
            loComandoSeleccionar.AppendLine("			AND TABLE_NAME =" & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" ORDER BY  1 ")

            loComandoSeleccionar.AppendLine(" CREATE TABLE #TablaResultados (Cuantos INT, Minimo_Seg INT, Maximo_Seg INT, Promedio_Seg INT, Sumatoria_Seg INT)")

            loComandoSeleccionar.AppendLine(" DECLARE curResultante CURSOR FOR ")
            loComandoSeleccionar.AppendLine(" SELECT * FROM #tmpTablas ")
            loComandoSeleccionar.AppendLine(" OPEN curResultante ")         
            loComandoSeleccionar.AppendLine(" FETCH NEXT FROM curResultante ")
            loComandoSeleccionar.AppendLine(" INTO @lcNombre, @ldFechaInicial02, @ldFechaFinal02, @ldFechaInicial, @ldFechaFinal, @lnCantidad, @lnMinimo, @lnMaximo, @lnSumatoria, @lnPromedio ")
		    loComandoSeleccionar.AppendLine(" WHILE @@FETCH_STATUS = 0 ")
            loComandoSeleccionar.AppendLine(" BEGIN")
					loComandoSeleccionar.AppendLine(" SET @lcSelectExecute	=	'Select		COUNT(Reg_Ini)											As		Cuantos, ")
					loComandoSeleccionar.AppendLine("										Min(DATEDIFF(ss,Reg_Ini,Reg_Fin))						As		Minimo_Seg,	")
					loComandoSeleccionar.AppendLine(" 										Max(DATEDIFF(ss,Reg_Ini,Reg_Fin))						As      Maximo_Seg,	")
					loComandoSeleccionar.AppendLine("										SUM(DATEDIFF(ss,Reg_Ini,Reg_Fin))/Count(Reg_Ini)		As		Promedio_Seg,")
					loComandoSeleccionar.AppendLine("										Sum(DATEDIFF(ss,Reg_Ini,Reg_Fin))						As		SumatoriaSeg  ")
					loComandoSeleccionar.AppendLine("                           From	     ' + @lcNombre + '  ")
					loComandoSeleccionar.AppendLine("                           Where	Fec_Ini Between ''' + @ldFechaInicial02 + ''' And ''' + @ldFechaFinal02 + '''' ")
					loComandoSeleccionar.AppendLine(" INSERT  INTO #TablaResultados ")           
					loComandoSeleccionar.AppendLine(" Execute(@lcSelectExecute) ")
					loComandoSeleccionar.AppendLine(" UPDATE    #tmpTablas  ")
					loComandoSeleccionar.AppendLine(" SET		Cuantos			=	ISNULL(#tmpResultadosX.Cuantos, 0), ")
					loComandoSeleccionar.AppendLine(" 			Minimo_Seg 		=	ISNULL(#tmpResultadosX.Minimo_Seg , 0), ")
					loComandoSeleccionar.AppendLine(" 			Maximo_Seg 		=	ISNULL(#tmpResultadosX.Maximo_Seg , 0), ")
					loComandoSeleccionar.AppendLine(" 			Promedio_Seg	=	ISNULL(#tmpResultadosX.Promedio_Seg , 0), ")
					loComandoSeleccionar.AppendLine(" 			Sumatoria_Seg	=	ISNULL(#tmpResultadosX.Sumatoria_Seg , 0) ")            
					loComandoSeleccionar.AppendLine(" FROM      (SELECT #TablaResultados.Cuantos, ")
					loComandoSeleccionar.AppendLine(" 					#TablaResultados.Minimo_Seg, ")
					loComandoSeleccionar.AppendLine(" 					#TablaResultados.Maximo_Seg, ")
					loComandoSeleccionar.AppendLine(" 					#TablaResultados.Promedio_Seg, ")
					loComandoSeleccionar.AppendLine(" 					#TablaResultados.Sumatoria_Seg  ")
					loComandoSeleccionar.AppendLine(" 			FROM	#TablaResultados) AS #tmpResultadosX ")
					loComandoSeleccionar.AppendLine(" WHERE     Nombre	=	@lcNombre ")
					loComandoSeleccionar.AppendLine(" DELETE FROM #TablaResultados ") 
					loComandoSeleccionar.AppendLine(" FETCH NEXT FROM curResultante ")
					loComandoSeleccionar.AppendLine(" INTO @lcNombre, @ldFechaInicial02, @ldFechaFinal02, @ldFechaInicial, @ldFechaFinal, @lnCantidad, @lnMinimo, @lnMaximo, @lnSumatoria, @lnPromedio ")
			loComandoSeleccionar.AppendLine(" End ")

            loComandoSeleccionar.AppendLine(" Select * From #tmpTablas WHERE Sumatoria_Seg > 0")
            'loComandoSeleccionar.AppendLine(" Select * From #tmpTablas")
            
            Dim loServicios As New cusDatos.goDatos
            
           ' Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
				Return
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTiempos_TDocumentos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTiempos_TDocumentos.ReportSource = loObjetoReporte

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
' JJD: 10/11/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
' MAT: 11/11/10: Programacion final
'-------------------------------------------------------------------------------------------'
' MAT: 15/11/10: Ajuste de Horas
'-------------------------------------------------------------------------------------------'
' MAT: 30/04/11: Modificación para que tome el nombre de la tabla desde una lista desplegable
'-------------------------------------------------------------------------------------------'

