﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rProspectos_PaisAtributo"
'-------------------------------------------------------------------------------------------'
Partial Class rProspectos_PaisAtributo
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
            'Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            
            'Atributo: es el nombre del campo de atributo a mostrar (sin comillas)
            Dim lcCampoAtributo As String = "[" & cusAplicacion.goReportes.paParametrosIniciales(7) & "]"
            Dim lcNombreCampo AS String = goServicios.mObtenerCampoFormatoSQL(CStr(cusAplicacion.goReportes.paParametrosIniciales(7) ).Trim().Replace("_", " "))
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Vendedores.cod_ven                                                      AS cod_ven,")
            loConsulta.AppendLine("            Vendedores.nom_ven                                                      AS nom_ven,")
            loConsulta.AppendLine("            Paises.cod_pai                                                          AS cod_pai, ")
            loConsulta.AppendLine("            Paises.nom_pai                                                          AS nom_pai,")
            loConsulta.AppendLine("            " & lcNombreCampo & "                                                   AS Campo_Atributo,")
            loConsulta.AppendLine("            (CASE WHEN Prospectos." & lcCampoAtributo & "> ''                       ")
            loConsulta.AppendLine("                THEN Prospectos." & lcCampoAtributo & "                             ")
            loConsulta.AppendLine("                ELSE '[SIN ASIGNAR]' END )                                          AS Atributo,")
            loConsulta.AppendLine("            SUM(CASE WHEN Prospectos.abc = 'A' THEN 1 ELSE 0 END)                   AS A,")
            loConsulta.AppendLine("            SUM(CASE WHEN Prospectos.abc = 'B' THEN 1 ELSE 0 END)                   AS B,")
            loConsulta.AppendLine("            SUM(CASE WHEN Prospectos.abc = 'C' THEN 1 ELSE 0 END)                   AS C,")
            loConsulta.AppendLine("            SUM(1)                                                                  AS Total,")
            loConsulta.AppendLine("            SUM( SUM(1) )")
            loConsulta.AppendLine("                 OVER(PARTITION BY Vendedores.Cod_Ven, Paises.Cod_Pai)              AS Total_VendedorPais,")
            loConsulta.AppendLine("            SUM(1)*100.0/SUM( SUM(1) ) ")
            loConsulta.AppendLine("                OVER(PARTITION BY Vendedores.Cod_Ven, Paises.Cod_Pai)               AS Porcentaje")
            loConsulta.AppendLine("FROM        Prospectos")
            loConsulta.AppendLine("    JOIN    Paises ON Paises.Cod_Pai = Prospectos.Cod_Pai")
            loConsulta.AppendLine("    JOIN    Vendedores ON Vendedores.Cod_Ven = Prospectos.Cod_Ven")
            loConsulta.AppendLine("WHERE	   Prospectos.Cod_Ven BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 	    AND Prospectos.Status IN (" & lcParametro1Desde & ")")
            loConsulta.AppendLine(" 	    AND Prospectos.Cod_Zon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 	    AND Prospectos.Cod_Pai BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 	    AND Prospectos.Cod_Suc BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro4Hasta)
            loConsulta.AppendLine(" 	    AND Prospectos.Cod_Cla BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 	    AND Prospectos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine(" 	    AND " & lcParametro6Hasta)
            loConsulta.AppendLine("GROUP BY    Vendedores.Cod_Ven,")
            loConsulta.AppendLine("            Vendedores.Nom_Ven,")
            loConsulta.AppendLine("            Paises.Cod_Pai,")
            loConsulta.AppendLine("            Paises.Nom_Pai,")
            loConsulta.AppendLine("            Prospectos." & lcCampoAtributo)
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento & ", Paises.Cod_Pai, Prospectos." & lcCampoAtributo)
            'loConsulta.AppendLine("ORDER BY    Vendedores.Cod_Ven, Paises.Cod_Pai, Prospectos.[Atributo_A]")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProspectos_PaisAtributo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProspectos_PaisAtributo.ReportSource = loObjetoReporte


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
' RJG: 29/07/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
