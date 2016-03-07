'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCMatriz_SucursalesSaldosClientes"
'-------------------------------------------------------------------------------------------'
Partial Class rCMatriz_SucursalesSaldosClientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Cod_Cli AS Cod_Matriz,")
            loComandoSeleccionar.AppendLine(" 		Nom_Cli As Nom_Matriz,")
            loComandoSeleccionar.AppendLine(" 		Status AS Status_Matriz,")
            loComandoSeleccionar.AppendLine(" 		Mon_Sal AS Saldo_Matriz")
            loComandoSeleccionar.AppendLine(" INTO #tempMatriz")
            loComandoSeleccionar.AppendLine(" FROM Clientes")
            loComandoSeleccionar.AppendLine(" WHERE Tip_Cli = 'Matriz'")

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Cod_Cli AS Cod_Sucursal,")
            loComandoSeleccionar.AppendLine(" 		Nom_Cli As Nom_Sucursal,")
            loComandoSeleccionar.AppendLine(" 		Matriz,")
            loComandoSeleccionar.AppendLine(" 		Status AS Status_Sucursal,")
            loComandoSeleccionar.AppendLine(" 		Mon_Sal AS Saldo_Sucursal")
            loComandoSeleccionar.AppendLine(" INTO #tempSucursal")
            loComandoSeleccionar.AppendLine(" FROM Clientes")
            loComandoSeleccionar.AppendLine(" WHERE Tip_Cli = 'Sucursal'")

            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 			ROW_NUMBER() OVER(PARTITION BY Cod_Matriz ORDER BY Cod_Matriz, " & lcOrdenamiento & ") AS Fila,")
            loComandoSeleccionar.AppendLine(" Cod_Matriz,")
            loComandoSeleccionar.AppendLine(" Nom_Matriz,")
            loComandoSeleccionar.AppendLine(" Status_Matriz,")
            loComandoSeleccionar.AppendLine(" Saldo_Matriz,")
            loComandoSeleccionar.AppendLine(" Cod_Sucursal,")
            loComandoSeleccionar.AppendLine(" Nom_Sucursal,")
            loComandoSeleccionar.AppendLine(" Matriz,")
            loComandoSeleccionar.AppendLine(" Status_Sucursal,")
            loComandoSeleccionar.AppendLine(" Saldo_Sucursal")
            loComandoSeleccionar.AppendLine(" From #tempMatriz")
            loComandoSeleccionar.AppendLine(" JOIN #tempSucursal ON  #tempSucursal.Matriz = #tempMatriz.Cod_Matriz")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 			Cod_Matriz between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cod_Sucursal between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Status_Matriz IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine(" 			AND Status_Sucursal IN (" & lcParametro2Desde & ")")

            'loComandoSeleccionar.AppendLine(" ORDER BY Cod_Matriz, Cod_Sucursal")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Matriz," & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCMatriz_SucursalesSaldosClientes", laDatosReporte)

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

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrCMatriz_SucursalesSaldosClientes.ReportSource = loObjetoReporte

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
' CMS: 31/08/09: Programacion inicial
'-------------------------------------------------------------------------------------------'