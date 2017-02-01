'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rCuentas_Contables"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rCuentas_Contables
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))

        Dim loComandoSeleccionar As New StringBuilder()

        Try

            loComandoSeleccionar.AppendLine("DECLARE @lcCod_Cue AS VARCHAR(12) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Cuentas_Contables.Cod_Cue	    AS Cod_Cue, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Nom_Cue	    AS Nom_Cue, ")
            loComandoSeleccionar.AppendLine("		''							    AS Cod_Aux,")
            loComandoSeleccionar.AppendLine("		''							    AS Nom_Aux")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Contables")
            If lcParametro0Desde = "''" Then
                loComandoSeleccionar.AppendLine("WHERE Cuentas_Contables.Cod_Cue NOT IN ('1.1.2.02.001','1.1.2.03.001','2.1.1.01.001','2.1.1.02.001')")
            Else
                loComandoSeleccionar.AppendLine("WHERE Cuentas_Contables.Cod_Cue = @lcCod_Cue")
            End If
            loComandoSeleccionar.AppendLine("")
            If lcParametro0Desde = "''" Or lcParametro0Desde = "'1.1.2.02.001'" Then
                loComandoSeleccionar.AppendLine("UNION")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("SELECT Cuentas_Contables.Cod_Cue   AS Cod_Cue, ")
                loComandoSeleccionar.AppendLine("		Cuentas_Contables.Nom_Cue   AS Nom_Cue, ")
                loComandoSeleccionar.AppendLine("		Auxiliares.Cod_Aux          AS Cod_Aux,")
                loComandoSeleccionar.AppendLine("		Auxiliares.Nom_Aux          AS Nom_Aux")
                loComandoSeleccionar.AppendLine("FROM Cuentas_Contables,")
                loComandoSeleccionar.AppendLine("	Auxiliares")
                loComandoSeleccionar.AppendLine("WHERE Cuentas_Contables.Cod_Cue = '1.1.2.02.001'")
                loComandoSeleccionar.AppendLine("	AND Auxiliares.Tipo = 'Cliente'")
                loComandoSeleccionar.AppendLine("	AND SUBSTRING(Auxiliares.Cod_Aux,1,2) IN ('CJ','CV','CG')")
                loComandoSeleccionar.AppendLine("")
            End If
            If lcParametro0Desde = "''" Or lcParametro0Desde = "'1.1.2.03.001'" Then
                loComandoSeleccionar.AppendLine("UNION")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("SELECT Cuentas_Contables.Cod_Cue  AS Cod_Cue, ")
                loComandoSeleccionar.AppendLine("		Cuentas_Contables.Nom_Cue   AS Nom_Cue, ")
                loComandoSeleccionar.AppendLine("		Auxiliares.Cod_Aux          AS Cod_Aux,")
                loComandoSeleccionar.AppendLine("		Auxiliares.Nom_Aux          AS Nom_Aux")
                loComandoSeleccionar.AppendLine("FROM Cuentas_Contables,")
                loComandoSeleccionar.AppendLine("	Auxiliares")
                loComandoSeleccionar.AppendLine("WHERE	Cuentas_Contables.Cod_Cue = '1.1.2.03.001'")
                loComandoSeleccionar.AppendLine("	AND Auxiliares.Tipo = 'Cliente'")
                loComandoSeleccionar.AppendLine("	AND SUBSTRING(Auxiliares.Cod_Aux,1,2) IN ('CX','C0')")
                loComandoSeleccionar.AppendLine("")
            End If
            If lcParametro0Desde = "''" Or lcParametro0Desde = "'2.1.1.01.001'" Then
                loComandoSeleccionar.AppendLine("UNION")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("SELECT Cuentas_Contables.Cod_Cue  AS Cod_Cue, ")
                loComandoSeleccionar.AppendLine("		Cuentas_Contables.Nom_Cue   AS Nom_Cue, ")
                loComandoSeleccionar.AppendLine("		Auxiliares.Cod_Aux          AS Cod_Aux,")
                loComandoSeleccionar.AppendLine("		Auxiliares.Nom_Aux          AS Nom_Aux")
                loComandoSeleccionar.AppendLine("FROM Cuentas_Contables,")
                loComandoSeleccionar.AppendLine("	Auxiliares")
                loComandoSeleccionar.AppendLine("WHERE Cuentas_Contables.Cod_Cue = '2.1.1.01.001'")
                loComandoSeleccionar.AppendLine("	AND Auxiliares.Tipo = 'Proveedor'")
                loComandoSeleccionar.AppendLine("	AND (SUBSTRING(Auxiliares.Cod_Aux,1,2) IN ('PJ','PV','PG') ")
                loComandoSeleccionar.AppendLine("		AND SUBSTRING(Auxiliares.Cod_Aux,1,3) <> 'PE0')")
                loComandoSeleccionar.AppendLine("")
            End If
            If lcParametro0Desde = "''" Or lcParametro0Desde = "'2.1.1.02.001'" Then
                loComandoSeleccionar.AppendLine("UNION")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("SELECT Cuentas_Contables.Cod_Cue  AS Cod_Cue, ")
                loComandoSeleccionar.AppendLine("		Cuentas_Contables.Nom_Cue   AS Nom_Cue, ")
                loComandoSeleccionar.AppendLine("		Auxiliares.Cod_Aux          AS Cod_Aux,")
                loComandoSeleccionar.AppendLine("		Auxiliares.Nom_Aux          AS Nom_Aux")
                loComandoSeleccionar.AppendLine("FROM Cuentas_Contables,")
                loComandoSeleccionar.AppendLine("	Auxiliares")
                loComandoSeleccionar.AppendLine("WHERE Cuentas_Contables.Cod_Cue = '2.1.1.02.001'")
                loComandoSeleccionar.AppendLine("	AND Auxiliares.Tipo = 'Proveedor'")
                loComandoSeleccionar.AppendLine("	AND SUBSTRING(Auxiliares.Cod_Aux,1,3) = ('PE0')")
            End If
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Cue")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                        "No se Encontraron Registros para los Parámetros Especificados. ", _
                         vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                         "350px", _
                         "200px")
            End If

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rCuentas_Contables", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rCuentas_Contables.ReportSource = loObjetoReporte


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
' MJP: 15/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MVP: 05/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------' 
' MAT: 16/05/11: Mejora de la vista de diseño, Ajuste del Select
'-------------------------------------------------------------------------------------------'
