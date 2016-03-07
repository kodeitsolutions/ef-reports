'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAuditorias_uMensualmente"
'-------------------------------------------------------------------------------------------'
Partial Class rAuditorias_uMensualmente
    Inherits vis2formularios.frmReporte

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
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cod_Usu,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '1' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Ene,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '2' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Feb,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '3' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Mar,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '4' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Abr,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '5' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS May,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '6' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Jun,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '7' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Jul,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '8' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Ago,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '9' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Sep,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '10' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Oct,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '11' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Nov,   ")
            loComandoSeleccionar.AppendLine("  			CASE   ")
            loComandoSeleccionar.AppendLine("  				WHEN DATEPART(MONTH, Registro) = '12' THEN 1   ")
            loComandoSeleccionar.AppendLine("  				ELSE '0'   ")
            loComandoSeleccionar.AppendLine("  			END AS Dic	   ")
            loComandoSeleccionar.AppendLine(" INTO      #Temporal   ")
            loComandoSeleccionar.AppendLine(" FROM      Auditorias    ")
            loComandoSeleccionar.AppendLine(" WHERE		Registro between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cod_Usu between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Tabla between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Opcion between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Documento   Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Codigo      Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cod_Emp     Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("  ORDER BY Cod_Usu ")

          

            loComandoSeleccionar.AppendLine(" SELECT    Cod_Usu,		     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Ene) AS Ene,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Feb) AS Feb,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Mar) AS Mar,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Abr) AS Abr,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(May) AS May,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Jun) AS Jun,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Jul) AS Jul,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Ago) AS Ago,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Sep) AS Sep,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Oct) AS Oct,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Nov) AS Nov,     ")
            loComandoSeleccionar.AppendLine(" 			SUM(Dic) AS Dic,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Ene) + SUM(Feb) + SUM(Mar) + SUM(Abr) + SUM(May) + SUM(Jun)     ")
            loComandoSeleccionar.AppendLine(" 			+ SUM(Jul) + SUM(Ago) + SUM(Sep) + SUM(Oct) + SUM(Nov) + SUM(Dic))AS Total,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Ene)/30) AS EneD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Feb)/30) AS FebD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Mar)/30) AS MarD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Abr)/30) AS AbrD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(May)/30) AS MayD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Jun)/30) AS JunD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Jul)/30) AS JulD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Ago)/30) AS AgoD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Sep)/30) AS SepD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Oct)/30) AS OctD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Nov)/30) AS NovD,     ")
            loComandoSeleccionar.AppendLine(" 			(SUM(Dic)/30) AS DicD,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Ene) + SUM(Feb) + SUM(Mar) + SUM(Abr) + SUM(May) + SUM(Jun)     ")
            loComandoSeleccionar.AppendLine(" 			+ SUM(Jul) + SUM(Ago) + SUM(Sep) + SUM(Oct) + SUM(Nov) + SUM(Dic))/30) AS TotalD,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Ene)/30)/8) AS EneH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Feb)/30)/8) AS FebH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Mar)/30)/8) AS MarH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Abr)/30)/8) AS AbrH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(May)/30)/8) AS MayH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Jun)/30)/8) AS JunH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Jul)/30)/8) AS JulH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Ago)/30)/8) AS AgoH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Sep)/30)/8) AS SepH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Oct)/30)/8) AS OctH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Nov)/30)/8) AS NovH,     ")
            loComandoSeleccionar.AppendLine(" 			((SUM(Dic)/30)/8) AS DicH,     ")
            loComandoSeleccionar.AppendLine(" 			(((SUM(Ene) + SUM(Feb) + SUM(Mar) + SUM(Abr) + SUM(May) + SUM(Jun)     ")
            loComandoSeleccionar.AppendLine(" 			+ SUM(Jul) + SUM(Ago) + SUM(Sep) + SUM(Oct) + SUM(Nov) + SUM(Dic))/30)/8) AS TotalH   ")
            loComandoSeleccionar.AppendLine(" FROM #Temporal   ")
            loComandoSeleccionar.AppendLine(" GROUP BY Cod_Usu   ")
            loComandoSeleccionar.AppendLine(" ORDER BY   " & lcOrdenamiento)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAuditorias_uMensualmente", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrAuditorias_uMensualmente.ReportSource = loObjetoReporte

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
' CMS: 30/05/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/12: Se incluyo el filtro de la empresa
'-------------------------------------------------------------------------------------------'
' MAT: 15/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
