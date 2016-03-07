'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rClientes_Campos_Extras"
'-------------------------------------------------------------------------------------------'

Partial Class rClientes_Campos_Extras
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


            loComandoSeleccionar.AppendLine("SELECT			 Clientes.Cod_Cli, ")
            loComandoSeleccionar.AppendLine("                Clientes.Nom_Cli, ")
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
            loComandoSeleccionar.AppendLine("FROM			 Campos_Extras,Clientes ")
            loComandoSeleccionar.AppendLine("WHERE			 Clientes.Cod_Cli between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Tip between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Zon between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Cla between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				 AND Clientes.Cod_Ven between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				 AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("                AND Campos_Extras.Cod_Reg = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("                AND Campos_Extras.Origen = 'Clientes'")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
           

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rClientes_Campos_Extras", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrClientes_Campos_Extras.ReportSource = loObjetoReporte


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
' RAC: 23/03/11: Se modificaron las etiquetas en el rpt, de campos numericos para separar 
'                en orden correcto los decimales de los miles, y se modificaron las etiquetas 
'                de fechas para mostarlas en el orden: dd/mm/aa.
'-------------------------------------------------------------------------------------------'
