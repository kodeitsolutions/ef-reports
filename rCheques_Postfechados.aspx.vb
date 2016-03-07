'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCheques_Postfechados"
'-------------------------------------------------------------------------------------------'
Partial Class rCheques_Postfechados
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("   SELECT  ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Documento, ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Cod_Caj, ")
            loComandoSeleccionar.AppendLine("               Cajas.Nom_Caj, ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Tipo, ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Referencia, ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Comentario, ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Doc_ori, ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Mon_Hab,  ")
            loComandoSeleccionar.AppendLine("               Movimientos_Cajas.Status  ")
            loComandoSeleccionar.AppendLine("   FROM Movimientos_Cajas Movimientos_Cajas ")
            loComandoSeleccionar.AppendLine("   JOIN Cajas AS Cajas ON Cajas.Cod_Caj = Movimientos_Cajas.Cod_Caj ")
            loComandoSeleccionar.AppendLine("   WHERE 		Movimientos_Cajas.Tipo = 'Cheque' ")
            loComandoSeleccionar.AppendLine("               AND Movimientos_Cajas.Tip_Ori = 'Cobros'  ")
            loComandoSeleccionar.AppendLine("               AND Movimientos_Cajas.Referencia NOT IN (SELECT Referencia FROM Renglones_Depositos)  ")

            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Cod_Caj BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Cod_Con BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Fec_Ini BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cajas.Status In (" & lcParametro5Desde &")")
            loComandoSeleccionar.AppendLine("ORDER BY   CONVERT(nchar(30), Movimientos_Cajas.Fec_Ini,112) DESC, Movimientos_Cajas.Cod_Caj, " & lcOrdenamiento)
'me.mEscribirConsulta(loComandoSeleccionar.ToString)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCheques_Postfechados", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCheques_Postfechados.ReportSource = loObjetoReporte

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
' CMS: 20/05/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 17/07/09: Metodo de Ordenamiento, Verificacion de registros
'-------------------------------------------------------------------------------------------'
' CMS: 19/16/10: Se agrego el campo estatus del movimiento de caja y el filtro correspondiente
'-------------------------------------------------------------------------------------------'
' MAT:  28/04/11: Mejora de la vista de Diseño
'-------------------------------------------------------------------------------------------'
