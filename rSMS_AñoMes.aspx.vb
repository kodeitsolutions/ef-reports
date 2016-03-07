'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rSMS_AñoMes"
'-------------------------------------------------------------------------------------------'
Partial Class rSMS_AñoMes
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
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpAuditorias(Año INT, ")
            loComandoSeleccionar.AppendLine("							Mes_Numero INT, ")
            loComandoSeleccionar.AppendLine("							Cod_Usu NCHAR(10) COLLATE Database_Default,")
            loComandoSeleccionar.AppendLine("							Accion NCHAR(10) COLLATE Database_Default, ")
            loComandoSeleccionar.AppendLine("							Cantidad INTEGER) ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT	INTO #tmpAuditorias(Año, Mes_Numero, Cod_Usu, Accion, Cantidad)")
            loComandoSeleccionar.AppendLine("SELECT	YEAR(Registro)		AS Año,	")
            loComandoSeleccionar.AppendLine("		MONTH(Registro)		AS Mes_Numero, ")
            loComandoSeleccionar.AppendLine("		Cod_Usu				AS Cod_Usu,")
            loComandoSeleccionar.AppendLine("		Accion				AS Accion,")
            loComandoSeleccionar.AppendLine("		COUNT(*)			AS Cantidad")
            loComandoSeleccionar.AppendLine("FROM 	Auditorias ")
            loComandoSeleccionar.AppendLine("WHERE		Registro    BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Cod_Usu     BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Tabla       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND Opcion      BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND Documento   BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND Codigo      BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND Cod_Emp     BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("       AND		Auditorias.Accion = 'SMS'")

            loComandoSeleccionar.AppendLine("GROUP BY	YEAR(Registro), ")
            loComandoSeleccionar.AppendLine("			MONTH(Registro),")
            loComandoSeleccionar.AppendLine("			Cod_Usu, Accion")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Año, ")
            loComandoSeleccionar.AppendLine("			Mes_Numero, ")
            loComandoSeleccionar.AppendLine(" 			CASE ")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 1 THEN 'Enero'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 2 THEN 'Febrero'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 3 THEN 'Marzo'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 4 THEN 'Abril'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 5 THEN 'Mayo'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 6 THEN 'Junio'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 7 THEN 'Julio'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 8 THEN 'Agosto'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 9 THEN 'Septiembre'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 10 THEN 'Octubre'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 11 THEN 'Noviembre'")
            loComandoSeleccionar.AppendLine("				WHEN Mes_Numero = 12 THEN 'Diciembre'")
            loComandoSeleccionar.AppendLine(" 			END AS Mes, ")
            loComandoSeleccionar.AppendLine("			COUNT(DISTINCT Cod_Usu) AS CantidadUsuarios,")
            'loComandoSeleccionar.AppendLine("			SUM((CASE WHEN Accion = 'Apertura'	THEN Cantidad ELSE 0 END)) AS Apertura, ")
            'loComandoSeleccionar.AppendLine("			SUM((CASE WHEN Accion = 'Agregar'	THEN Cantidad ELSE 0 END)) AS Agregar, ")
            'loComandoSeleccionar.AppendLine("			SUM((CASE WHEN Accion = 'Anular'	THEN Cantidad ELSE 0 END)) AS Anular, ")
            'loComandoSeleccionar.AppendLine("			SUM((CASE WHEN Accion = 'Eliminar'	THEN Cantidad ELSE 0 END)) AS Eliminar, ")
            'loComandoSeleccionar.AppendLine("			SUM((CASE WHEN Accion = 'Modificar'	THEN Cantidad ELSE 0 END)) AS Modificar,")
            'loComandoSeleccionar.AppendLine("			SUM((CASE WHEN Accion = 'Reporte'	THEN Cantidad ELSE 0 END)) AS Reporte, ")
            loComandoSeleccionar.AppendLine("			SUM((CASE WHEN Accion = 'SMS'   	THEN Cantidad ELSE 0 END)) AS SMS, ")
            'loComandoSeleccionar.AppendLine("			SUM((CASE WHEN Accion = ''			THEN Cantidad ELSE 0 END)) As Vacio, ")
            loComandoSeleccionar.AppendLine("			SUM(Cantidad) As TotalDia ")
            loComandoSeleccionar.AppendLine("FROM		#tmpAuditorias")

            loComandoSeleccionar.AppendLine("GROUP BY	Año, Mes_Numero ")
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento & ", Mes Desc")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")




            Dim loServicios As New cusDatos.goDatos

            ' Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rSMS_AñoMes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSMS_AñoMes.ReportSource = loObjetoReporte

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
' CMS: 24/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 19/12/12: Se incluyo el filtro de la empresa
'-------------------------------------------------------------------------------------------'
' MAT: 15/04/11: Ajuste de la vista de Diseño
'-------------------------------------------------------------------------------------------'
' RJG: 21/01/13: Se agregó el número de usuarios del día. Se ajustó el SELECT para hacerlo	'
'				 más rápido.																'
'-------------------------------------------------------------------------------------------'
' RJG: 11/03/13: Corrección de bug en totales de auditorias.
'-------------------------------------------------------------------------------------------'
' PMV: 11/03/13: Creacion de un Reporte de Auditorias de SMS.
'-------------------------------------------------------------------------------------------'
