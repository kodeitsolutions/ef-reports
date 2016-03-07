'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_Campos_Extras"
'-------------------------------------------------------------------------------------------'

Partial Class rArticulos_Campos_Extras
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
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden


            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT			 Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("                Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico1, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico2, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico3, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico4, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico5, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico6, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico7, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico8, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico9, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Logico10, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble1, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble2, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble3, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble4, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble5, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble6, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble7, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble8, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble9, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Doble10, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter1, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter2, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter3, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter4, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter5, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter6, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter7, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter8, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter9, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Caracter10, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Fecha1, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Fecha2, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Fecha3, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Fecha4, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Fecha5, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo1, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo2, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo3, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo4, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo5, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo6, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo7, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo8, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo9, ")
            loComandoSeleccionar.AppendLine("                Campos_Articulos.Memo10 ")
            loComandoSeleccionar.AppendLine("FROM			 Campos_Articulos,Articulos ")
            loComandoSeleccionar.AppendLine("WHERE			 Articulos.Cod_Art between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Articulos.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 				 AND Articulos.Cod_Tip between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Articulos.Cod_Cla between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("                AND Campos_Articulos.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            'Dim SCADENA = Convert.ToString(loComandoSeleccionar) 

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_Campos_Extras", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_Campos_Extras.ReportSource = loObjetoReporte


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
' RAC: 16/03/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RAC: 23/03/11: Se modifico el archivo rpt para: cambiar el tamaño de los campos numericos,
'                cambiar los separadores y cambiar el formato de las fechas en la forma: 
'                dd/mm/aa
'-------------------------------------------------------------------------------------------'
