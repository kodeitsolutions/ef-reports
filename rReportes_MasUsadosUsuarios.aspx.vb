'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rReportes_MasUsadosUsuarios"
'-------------------------------------------------------------------------------------------'
Partial Class rReportes_MasUsadosUsuarios
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
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            'Dim lcSql As String

            Dim loComandoSeleccionar As New StringBuilder()


            'If lcParametro4Desde > 0 Then
            '    lcSql = "Select Top " + lcParametro4Desde.ToString
            'Else
            '    lcSql = "Select "
            'End If

            'loComandoSeleccionar.AppendLine(lcSql)
            loComandoSeleccionar.AppendLine(" Select")
			loComandoSeleccionar.AppendLine("              Auditorias.Cod_Usu,    ")
            loComandoSeleccionar.AppendLine("              Usuarios.Nom_Usu,      ")
            loComandoSeleccionar.AppendLine("              Reportes.Cod_Rep,      ")
            loComandoSeleccionar.AppendLine("              Reportes.Nom_Rep,      ")
            loComandoSeleccionar.AppendLine("              Reportes.Tipo,      ")
            loComandoSeleccionar.AppendLine("              Reportes.Modulo,      ")
            loComandoSeleccionar.AppendLine("              CAST(COUNT(Reportes.Cod_Rep) AS INT) AS Cantidad  ")
            loComandoSeleccionar.AppendLine(" INTO #Temporal      ")
            loComandoSeleccionar.AppendLine("  FROM Factory_Global.dbo.Reportes AS Reportes      ")
            loComandoSeleccionar.AppendLine("  JOIN Factory_" & goAplicacion.pcNombre & "_" & goEmpresa.pcCodigo & ".dbo.auditorias AS Auditorias ON Reportes.Cod_Rep collate Modern_Spanish_CI_AS = Auditorias.Codigo collate Modern_Spanish_CI_AS      ")
            loComandoSeleccionar.AppendLine("  JOIN Factory_Global.dbo.usuarios AS usuarios ON Usuarios.Cod_usu collate Modern_Spanish_CI_AS = Auditorias.cod_usu collate Modern_Spanish_CI_AS      ")
            loComandoSeleccionar.AppendLine("  WHERE    Auditorias.Accion IN ('Reporte', 'Formato')      ")
            loComandoSeleccionar.AppendLine("           AND Auditorias.Tipo = 'Seguimiento'      ")
            loComandoSeleccionar.AppendLine("     AND			Auditorias.Registro   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("     AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("     AND			Auditorias.Cod_Usu  Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("     AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("     AND			Reportes.Opcion  Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("     AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("     AND			Reportes.Modulo  Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("     AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("     AND			Reportes.Tipo  IN (" & lcParametro5Desde & ")")
            loComandoSeleccionar.AppendLine("  GROUP BY Auditorias.Cod_Usu, Usuarios.Nom_Usu, Reportes.Modulo, Reportes.Cod_Rep, Reportes.Nom_Rep, Reportes.Tipo      ")
            loComandoSeleccionar.AppendLine("  ORDER BY Cantidad DESC ")

            'loComandoSeleccionar.AppendLine("   SELECT * FROM #Temporal      ")
            loComandoSeleccionar.AppendLine("    SELECT  ")
            loComandoSeleccionar.AppendLine("    		ROW_NUMBER() OVER (PARTITION BY Cod_Usu ORDER BY Cod_Usu, " & lcOrdenamiento &", Cod_Rep ) AS ROW_NUMBER, ")
            loComandoSeleccionar.AppendLine("    		Cod_Usu, ")
            loComandoSeleccionar.AppendLine("    		Nom_Usu, ")
            loComandoSeleccionar.AppendLine("    		Cod_Rep, ")
            loComandoSeleccionar.AppendLine("    		Nom_Rep, ")
            loComandoSeleccionar.AppendLine("    		Tipo, ")
            loComandoSeleccionar.AppendLine("    		Modulo, ")
            loComandoSeleccionar.AppendLine("    		Cantidad    ")
            loComandoSeleccionar.AppendLine("    INTO #Temporal2 ")
            loComandoSeleccionar.AppendLine("    FROM #Temporal   ")
            loComandoSeleccionar.AppendLine("   ORDER BY Cod_Usu, " & lcOrdenamiento & ", Cod_Rep")

			If lcParametro4Desde = "'0'" Then
				loComandoSeleccionar.AppendLine("    SELECT * FROM #Temporal2 ")
			Else
				loComandoSeleccionar.AppendLine("    SELECT * FROM #Temporal2 WHERE ROW_NUMBER <= " & lcParametro4Desde)
			End If


            Dim loServicios As New cusDatos.goDatos
            goDatos.pcNombreAplicativoExterno = "Framework"

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
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rReportes_MasUsadosUsuarios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrReportes_MasUsadosUsuarios.ReportSource = loObjetoReporte

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
' CMS   :  16/04/10 : Codigo inicial
'-------------------------------------------------------------------------------------------'
