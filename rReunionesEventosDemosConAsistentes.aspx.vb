'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rReunionesEventosDemosConAsistentes"
'-------------------------------------------------------------------------------------------'
Partial Class rReunionesEventosDemosConAsistentes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

        Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))


        Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))


        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("SELECT ")
            loConsulta.AppendLine("  Eventos_Marketing.Cod_Eve AS Codigo, ")
            loConsulta.AppendLine("  Eventos_Marketing.Tipo, ")
            loConsulta.AppendLine("  Eventos_Marketing.Status AS Estatus, ")
            loConsulta.AppendLine("  Eventos_Marketing.Nom_Eve AS Nombre,  ")
            loConsulta.AppendLine("  Eventos_Marketing.Responsable, ")
            loConsulta.AppendLine("  Eventos_Marketing.Fec_ini, ")
            loConsulta.AppendLine("  Eventos_Marketing.Fec_Fin, ")
            loConsulta.AppendLine("  Eventos_Marketing.Fec_Rev, ")
            loConsulta.AppendLine("  Eventos_Marketing.Por_Eje,")
            loConsulta.AppendLine("  COALESCE(Renglones.Renglon,0)      AS Renglon,")
            loConsulta.AppendLine("  COALESCE(Renglones.Actividad,  '') AS Actividad,")
            loConsulta.AppendLine("  COALESCE(Renglones.Responsable,'') AS Resp_Act,")
            loConsulta.AppendLine("  COALESCE(Renglones.Lugar,    '')   AS Lugar_Act,")
            loConsulta.AppendLine("  COALESCE(Renglones.Fec_Ini,  '')   AS Fec_Ini_Act,")
            loConsulta.AppendLine("  COALESCE(Renglones.Fec_Fin,  '')   AS Fec_Fin_Act,")
            loConsulta.AppendLine("  COALESCE(Renglones.Por_Eje,  0 )   AS Por_Eje_Act,")
            loConsulta.AppendLine("  COALESCE(Renglones.Estatus,  '')   AS Estatus_Act,")
            loConsulta.AppendLine("  COALESCE(Renglones.Usuario,  '')   AS Usuario_act")
            loConsulta.AppendLine("                 ")
            loConsulta.AppendLine("FROM Eventos_Marketing")
            loConsulta.AppendLine("             ")
            loConsulta.AppendLine("LEFT JOIN (")
            loConsulta.AppendLine("  SELECT ")
            loConsulta.AppendLine("   AsistentesEvento.Cod_Eve,")
            loConsulta.AppendLine("   AsistentesEvento.Origen,")
            loConsulta.AppendLine("   T.C.value('(@renglon)[1]', 'Integer')         AS Renglon,")
            loConsulta.AppendLine("   T.C.value('(@nombre)[1]', 'varchar(MAX)')     AS Nombre_Reng,")
            loConsulta.AppendLine("   T.C.value('(@cantidad)[1]', 'integer')        AS Cantidad,")
            loConsulta.AppendLine("   T.C.value('(@monto)[1]', 'decimal')           AS Monto,")
            loConsulta.AppendLine("   T.C.value('(@valor1)[1]', 'decimal')          AS valor1,")
            loConsulta.AppendLine("   T.C.value('(@valor2)[1]', 'decimal')          AS valor2,")
            loConsulta.AppendLine("   T.C.value('(@valor3)[1]', 'decimal')          AS valor3,")
            loConsulta.AppendLine("             ")
            loConsulta.AppendLine("  FROM (")
            loConsulta.AppendLine("    SELECT Cod_Eve, Origen, CAST(Asistencia AS XML) as Asistencia FROM Eventos_Marketing ")
            loConsulta.AppendLine("  )")
            loConsulta.AppendLine("  AS AsistentesEvento")
            loConsulta.AppendLine("  OUTER APPLY Asistencia.nodes('//elementos/elemento') AS T(C)")
            loConsulta.AppendLine(")")
            loConsulta.AppendLine("AS Renglones")
            loConsulta.AppendLine("ON Eventos_Marketing.Cod_Eve = Renglones.Cod_Eve AND Renglones.Origen= Eventos_Marketing.Origen")

            loConsulta.AppendLine("WHERE ")
            loConsulta.AppendLine("  Eventos_Marketing.adicional = '' ")
            loConsulta.AppendLine("  AND Eventos_Marketing.Cod_Eve BETWEEN " & _
                                 lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("  AND Eventos_Marketing.Status IN ( " & lcParametro1Desde & ")")
            loConsulta.AppendLine("  AND Eventos_Marketing.Etapa IN ( " & lcParametro2Desde & ")")

            loConsulta.AppendLine("  AND Eventos_Marketing.Fec_Ini BETWEEN " & _
                                             lcParametro3Desde & " AND " & lcParametro3Hasta)

            loConsulta.AppendLine(" AND  Eventos_Marketing.Fec_Rev BETWEEN " & _
                                 lcParametro4Desde & " AND " & lcParametro4Hasta)
            loConsulta.AppendLine("  AND Eventos_Marketing.Prioridad IN ( " & lcParametro5Desde & ")")
            loConsulta.AppendLine("  AND Eventos_Marketing.Tipo IN ( 'Reunion','Evento','Demostracion')")
            loConsulta.AppendLine("ORDER BY Eventos_Marketing.tipo ASC, " & lcOrdenamiento)

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rReunionesEventosDemosConAsistentes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrReunionesEventosDemosConAsistentes.ReportSource = loObjetoReporte

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
' JAC : 20/08/15 : Codigo inicial
'-------------------------------------------------------------------------------------------'

