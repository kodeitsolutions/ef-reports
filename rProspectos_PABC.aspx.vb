'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rProspectos_PABC"
'-------------------------------------------------------------------------------------------'
Partial Class rProspectos_PABC
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
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      vendedores.cod_ven,")
            loConsulta.AppendLine("            vendedores.nom_ven,")
            loConsulta.AppendLine("            Paises.cod_pai,")
            loConsulta.AppendLine("            Paises.nom_pai,")
            loConsulta.AppendLine("            SUM(CASE")
            loConsulta.AppendLine("                  WHEN Prospectos.abc = 'A' THEN 1 ELSE 0")
            loConsulta.AppendLine("                END) A,")
            loConsulta.AppendLine("            SUM(CASE")
            loConsulta.AppendLine("                  WHEN Prospectos.abc = 'B' THEN 1 ELSE 0")
            loConsulta.AppendLine("                END) B,")
            loConsulta.AppendLine("            SUM(CASE")
            loConsulta.AppendLine("                  WHEN Prospectos.abc = 'C' THEN 1 ELSE 0")
            loConsulta.AppendLine("                END) C,")
            loConsulta.AppendLine("            SUM(CASE")
            loConsulta.AppendLine("                  WHEN Prospectos.abc = 'A'or Prospectos.abc ='B' or Prospectos.abc='C' THEN 1 ELSE 0")
            loConsulta.AppendLine("                END) Total")
            loConsulta.AppendLine("FROM        Prospectos")
            loConsulta.AppendLine("    JOIN    Paises ON Paises.Cod_Pai = Prospectos.Cod_Pai")
            loConsulta.AppendLine("    JOIN    Vendedores ON Vendedores.Cod_Ven = Prospectos.Cod_Ven")
            loConsulta.AppendLine("WHERE	   Prospectos.Cod_Ven BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 	    AND Prospectos.status IN (" & lcParametro1Desde & ")")
            loConsulta.AppendLine(" 	    AND Prospectos.Cod_Zon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 	    AND Prospectos.Cod_Pai BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 	    AND Paises.Cod_Suc BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro4Hasta)
            loConsulta.AppendLine(" 	    AND Prospectos.Cod_Cla BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 	    AND Prospectos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro6Hasta)
            If lcParametro7Desde <> "''" Then
                loConsulta.AppendLine(" 	    AND Prospectos.Atributo_A IN (" & lcParametro7Desde & ")")
            End If
            If lcParametro8Desde <> "''" Then
                loConsulta.AppendLine(" 	    AND Prospectos.Atributo_B IN (" & lcParametro8Desde & ")")
            End If
            If lcParametro9Desde <> "''" Then
                loConsulta.AppendLine(" 	    AND Prospectos.Atributo_C IN (" & lcParametro9Desde & ")")
            End If
            loConsulta.AppendLine("GROUP BY    Vendedores.Cod_Ven, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven,")
            loConsulta.AppendLine("            Paises.Cod_Pai,")
            loConsulta.AppendLine("            Paises.Nom_Pai")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento & ", Paises.cod_pai")
            loConsulta.AppendLine("")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProspectos_PABC", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProspectos_PABC.ReportSource = loObjetoReporte


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
' EAG: 28/07/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
