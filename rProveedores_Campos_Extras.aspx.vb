'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rProveedores_Campos_Extras"
'-------------------------------------------------------------------------------------------'

Partial Class rProveedores_Campos_Extras
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
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT			 Proveedores.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("                Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico1, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico2, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico3, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico4, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico5, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico6, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico7, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico8, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico9, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Logico10, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble1, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble2, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble3, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble4, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble5, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble6, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble7, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble8, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble9, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Doble10, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter1, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter2, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter3, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter4, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter5, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter6, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter7, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter8, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter9, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Caracter10, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Fecha1, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Fecha2, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Fecha3, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Fecha4, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Fecha5, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo1, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo2, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo3, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo4, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo5, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo6, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo7, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo8, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo9, ")
            loComandoSeleccionar.AppendLine("                Campos_Extras.Memo10 ")
            loComandoSeleccionar.AppendLine("FROM			 Proveedores ")
            loComandoSeleccionar.AppendLine("LEFT JOIN Campos_Extras ON (Campos_Extras.Cod_Reg = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("                           AND Campos_Extras.Origen = 'Proveedores')")
            loComandoSeleccionar.AppendLine("WHERE			 Proveedores.Cod_Pro between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Tip between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Zon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Cla between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Proveedores.Cod_Ven between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro5Hasta)
      
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProveedores_Campos_Extras", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProveedores_Campos_Extras.ReportSource = loObjetoReporte


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
' MAT: 14/03/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RAC: 23/03/11: Se modificaron los separadores de los campos numericos y su tamaño, tambien
'                se modifico el formato de los campos de tipo fecha de la forma: dd/mm/aa.
'-------------------------------------------------------------------------------------------'
