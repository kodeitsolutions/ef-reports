'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rInasistencias_Acceso"
'-------------------------------------------------------------------------------------------'
Partial Class rInasistencias_Acceso
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim ldInicio As DateTime = cusAplicacion.goReportes.paParametrosIniciales(2)
        Dim ldFin As DateTime = cusAplicacion.goReportes.paParametrosFinales(2)

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(ldInicio, goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(ldFin, goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        loConsulta.AppendLine("DECLARE       @tmpAsistentes TABLE (Registro DATETIME, Cod_Tra CHAR(10));")

        loConsulta.AppendLine("INSERT INTO   @tmpAsistentes (Registro, Cod_Tra)  ")
        loConsulta.AppendLine("SELECT		(Registro) AS Registro, Cod_Tra")
        loConsulta.AppendLine("FROM		    Accesos")
        loConsulta.AppendLine("WHERE		    Registro BETWEEN " & lcParametro2Desde)
        loConsulta.AppendLine("			    AND " & lcParametro2Hasta)
        loConsulta.AppendLine("		        AND	Status <> 'Anulado'")

        loConsulta.AppendLine("SELECT    Registro, Cod_Tra FROM @tmpAsistentes")

        loConsulta.AppendLine("SELECT	Trabajadores.cod_tra,")
        loConsulta.AppendLine("			Trabajadores.nom_tra,")
        loConsulta.AppendLine("			ISNULL(Departamentos_Nomina.Cod_Dep, '*') AS cod_dep,")
        loConsulta.AppendLine("			ISNULL(Departamentos_Nomina.Nom_Dep, '*Sin Definir*') AS nom_dep,")
        loConsulta.AppendLine("			ISNULL(CAST(Tarjetas_Accesos.Cod_Tar AS CHAR(15)), '*Sin Asignar*') AS cod_tar")
        loConsulta.AppendLine("FROM		Trabajadores")
        loConsulta.AppendLine("	        LEFT JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep")
        loConsulta.AppendLine("	        LEFT JOIN Tarjetas_Accesos ON Tarjetas_Accesos.Cod_Tra = Trabajadores.Cod_Tra")
        loConsulta.AppendLine("WHERE    Trabajadores.Status = 'A'")
        loConsulta.AppendLine("		    AND	Trabajadores.Cod_Tra BETWEEN " & lcParametro0Desde)
        loConsulta.AppendLine("			AND	" & lcParametro0Hasta)
        loConsulta.AppendLine("		    AND	Trabajadores.Cod_Dep BETWEEN " & lcParametro1Desde)
        loConsulta.AppendLine("			AND	" & lcParametro1Hasta)
        'loComandoSeleccionar.AppendLine(" ORDER BY	Trabajadores.Cod_Dep, Trabajadores.Cod_Tra ")
        loConsulta.AppendLine("ORDER BY   Trabajadores.Cod_Dep,   " & lcOrdenamiento)


        Dim loServicios As New cusDatos.goDatos
        Dim laDatosReporte As DataSet

        Try

            laDatosReporte = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")
            Return

        End Try

        laDatosReporte.Tables(0).TableName = "Asistencias"
        laDatosReporte.Tables(1).TableName = "Inasistentes"

        '-------------------------------------------------------------------------------------------'
        '  Crea el contenedor para las inasistencias.												'
        '-------------------------------------------------------------------------------------------'
        Dim loTablaReporte As New DataTable("curReportes")

        loTablaReporte.Columns.Add("Fecha", GetType(String))
        loTablaReporte.Columns.Add("Dia", GetType(String))
        loTablaReporte.Columns.Add("Cod_Tra", GetType(String))
        loTablaReporte.Columns.Add("Nom_Tra", GetType(String))
        loTablaReporte.Columns.Add("Cod_Dep", GetType(String))
        loTablaReporte.Columns.Add("Nom_Dep", GetType(String))
        loTablaReporte.Columns.Add("Cod_Tar", GetType(String))


        '-------------------------------------------------------------------------------------------'
        '  Recorre la lista de trabajadores que no vinieron en el rango de fechas.					'
        '-------------------------------------------------------------------------------------------'
        Dim lnTotalDias As Integer = Math.Ceiling((ldFin - ldInicio).TotalDays)
        Dim ldDiaInicial As DateTime = New DateTime(ldInicio.Year, ldInicio.Month, ldInicio.Day)

        For Each loInasistente As DataRow In laDatosReporte.Tables("Inasistentes").Rows

            Dim lcTrabajador As String = goServicios.mObtenerCampoFormatoSQL(loInasistente("Cod_Tra"))

            '-------------------------------------------------------------------------------------------'
            '  Para cada inasistente, verifica cada día del rango de fechas y verifica si ese día vino.	'
            '-------------------------------------------------------------------------------------------'
            For lnDiaActual As Integer = 0 To lnTotalDias - 1

                Dim ldDiaActual As DateTime = ldDiaInicial.AddDays(lnDiaActual)

                Dim lcSeleccion As String = "Cod_Tra=" & lcTrabajador & " AND Registro >= '" & _
                        Microsoft.VisualBasic.Format(ldDiaActual, "MM-dd-yyyy HH:mm:ss") & "' AND Registro <= '" & _
                        Microsoft.VisualBasic.Format(ldDiaActual.AddDays(1).AddSeconds(-1), "MM-dd-yyyy HH:mm:ss") & "'"

                If laDatosReporte.Tables("Asistencias").Select(lcSeleccion).Length <= 0 Then
                    Dim loNuevoRegistro As DataRow = loTablaReporte.NewRow()

                    loNuevoRegistro("Fecha") = goServicios.mObtenerFormatoCadena(ldDiaActual, goServicios.enuOpcionesRedondeo.KN_FechaSinHoras)
                    loNuevoRegistro("Cod_Tra") = loInasistente("Cod_Tra")
                    loNuevoRegistro("Nom_Tra") = loInasistente("Nom_Tra")
                    loNuevoRegistro("Cod_Dep") = loInasistente("Cod_Dep")
                    loNuevoRegistro("Nom_Dep") = loInasistente("Nom_Dep")
                    loNuevoRegistro("Cod_Tar") = loInasistente("Cod_Tar")

                    Select Case ldDiaActual.DayOfWeek
                        Case DayOfWeek.Monday
                            loNuevoRegistro("dia") = "LUN"
                        Case DayOfWeek.Tuesday
                            loNuevoRegistro("dia") = "MAR"
                        Case DayOfWeek.Wednesday
                            loNuevoRegistro("dia") = "MIE"
                        Case DayOfWeek.Thursday
                            loNuevoRegistro("dia") = "JUE"
                        Case DayOfWeek.Friday
                            loNuevoRegistro("dia") = "VIE"
                        Case DayOfWeek.Saturday
                            loNuevoRegistro("dia") = "SAB"
                        Case DayOfWeek.Sunday
                            loNuevoRegistro("dia") = "DOM"
                    End Select

                    loTablaReporte.Rows.Add(loNuevoRegistro)

                End If

            Next lnDiaActual

        Next loInasistente

        Try

            Dim loDataSetReporte As New DataSet()
            loDataSetReporte.Tables.Add(loTablaReporte)
            'loTablaReporte.TableName = "curReportes"
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rInasistencias_Acceso", loDataSetReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrInasistencias_Acceso.ReportSource = loObjetoReporte


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
' MVP: 14/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 24/07/08: Se suprimión el filtro de la puertas puesto que inicialmente no se trabaja_
'                con este dato necesario al ingresar o salir de la empresa.
'-------------------------------------------------------------------------------------------'
' JJD: 06/02/10: Se le sustituyo el objeto mObtenerTodos() por mObtenerTodosSinEsquema() 
'-------------------------------------------------------------------------------------------'
' CMS: 15/07/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'
' RJG: 27/06/15: Actualización de tabla de Departamento (por migración de nómina a Adm.).   '
'                Ajuste de interfaz (no mostraba la vista PDF).                             '
'-------------------------------------------------------------------------------------------'
