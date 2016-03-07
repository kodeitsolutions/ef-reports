'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_ValorAlmacen"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_ValorAlmacen
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcComandoSeleccionar As New StringBuilder()

            lcComandoSeleccionar.AppendLine("SELECT		Articulos.Cod_Art, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Uni1, ")
            lcComandoSeleccionar.AppendLine("			SUM(Renglones_Almacenes.Exi_Act1) AS Exi_Act1, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cos_Ult1, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cos_Ult2, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cos_Pro1, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cos_Pro2, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Dep, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Sec, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Tip, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Cla, ")
            lcComandoSeleccionar.AppendLine("			Articulos.Cod_Mar, ")
            lcComandoSeleccionar.AppendLine("			Renglones_Almacenes.Cod_Alm ")
            lcComandoSeleccionar.AppendLine("INTO		#tmp ")
            lcComandoSeleccionar.AppendLine("FROM		Articulos ")
            lcComandoSeleccionar.AppendLine("	JOIN	Renglones_Almacenes   ")
            lcComandoSeleccionar.AppendLine("		ON	Renglones_Almacenes.Cod_Art = Articulos.Cod_Art ")
            lcComandoSeleccionar.AppendLine("		AND	Renglones_Almacenes.Exi_Act1 <> 0")
            lcComandoSeleccionar.AppendLine("WHERE     Articulos.Cod_Art       BETWEEN " & lcParametro0Desde)
            lcComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            lcComandoSeleccionar.AppendLine("		    AND Articulos.Status        IN (" & lcParametro1Desde & ")")
            lcComandoSeleccionar.AppendLine("		    AND Articulos.Cod_Dep       BETWEEN " & lcParametro2Desde)
            lcComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
            lcComandoSeleccionar.AppendLine("		    AND Articulos.Cod_Sec       BETWEEN " & lcParametro3Desde)
            lcComandoSeleccionar.AppendLine("				AND " & lcParametro3Hasta)
            lcComandoSeleccionar.AppendLine("		    AND Articulos.Cod_Mar       BETWEEN " & lcParametro4Desde)
            lcComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
            lcComandoSeleccionar.AppendLine("		    AND Articulos.Cod_Tip       BETWEEN " & lcParametro5Desde)
            lcComandoSeleccionar.AppendLine("				AND " & lcParametro5Hasta)
            lcComandoSeleccionar.AppendLine("		    AND Articulos.Cod_Cla       BETWEEN " & lcParametro6Desde)
            lcComandoSeleccionar.AppendLine("				AND " & lcParametro6Hasta)
            lcComandoSeleccionar.AppendLine("		    AND Articulos.Cod_Ubi    BETWEEN " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("				And " & lcParametro7Hasta)
            lcComandoSeleccionar.AppendLine("		    AND Articulos.Cod_Ubi    BETWEEN " & lcParametro7Desde)
            lcComandoSeleccionar.AppendLine("				AND " & lcParametro7Hasta)
			lcComandoSeleccionar.AppendLine("GROUP BY	Articulos.Cod_Art,       	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Nom_Art,       	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cod_Uni1,      	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cos_Ult1,      	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cos_Ult2,      	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cos_Pro1,      	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cos_Pro2,      	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cod_Dep,       	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cod_Sec,       	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cod_Tip,       	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cod_Cla,       	")
			lcComandoSeleccionar.AppendLine(" 			Articulos.Cod_Mar,			")
			lcComandoSeleccionar.AppendLine("			Renglones_Almacenes.cod_alm	")
			lcComandoSeleccionar.AppendLine("           ")
			lcComandoSeleccionar.AppendLine("           ")
			lcComandoSeleccionar.AppendLine("SELECT		")
			lcComandoSeleccionar.AppendLine("			#tmp.Cod_Art, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Nom_Art, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cod_Uni1, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Exi_Act1, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cos_Ult1, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cos_Ult2, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cos_Pro1, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cos_Pro2, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cod_Dep, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cod_Sec, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cod_Tip, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cod_Cla, ")
			lcComandoSeleccionar.AppendLine("			#tmp.Cod_Mar, ")
			lcComandoSeleccionar.AppendLine("			#tmp.cod_alm, ")
            lcComandoSeleccionar.AppendLine("			(SELECT Nom_Alm FROM Almacenes WHERE Almacenes.Cod_Alm = #tmp.Cod_Alm) As Nom_Alm")
			lcComandoSeleccionar.AppendLine("FROM		#tmp")
			lcComandoSeleccionar.AppendLine("WHERE		Cod_Alm BETWEEN " & lcParametro8Desde)
            lcComandoSeleccionar.AppendLine("				AND " & lcParametro8Hasta)
            lcComandoSeleccionar.AppendLine("ORDER BY	Cod_Alm,  " & lcOrdenamiento)
         
            

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(lcComandoSeleccionar.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_ValorAlmacen", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_ValorAlmacen.ReportSource = loObjetoReporte


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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' CMS: 12/03/10: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
' MAT: 14/09/11: Ajuste de la vista de diseño												'
'-------------------------------------------------------------------------------------------'
' RJG: 26/03/12: Se agregó un filtro fijo para solo mostrar con existencia <> 0. Se corrigió'
'				 un filtro de estatus ='A' en los almacenes (estaba de sobra).				'
'-------------------------------------------------------------------------------------------'
