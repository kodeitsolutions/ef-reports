'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOpciones"
'-------------------------------------------------------------------------------------------'

Partial Class rOpciones

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = ""

        If (lcParametro2Desde = "''" Or lcParametro2Desde = "' '" Or lcParametro2Desde = "'0'") Then

            lcParametro2Hasta = "'zzzzzzzz'"

        Else

            lcParametro2Hasta = lcParametro2Desde

        End If



        'Dim laModulos As ArrayList = cusAplicacion.goServicios.mObtenerModulosAplicativo()

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT		Cod_Opc, ")
            loComandoSeleccionar.AppendLine("				Nom_Opc, ")
            loComandoSeleccionar.AppendLine("				Status, ")
            loComandoSeleccionar.AppendLine("				Modulo, ")
            loComandoSeleccionar.AppendLine("				(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Opciones, ")
            loComandoSeleccionar.AppendLine("				(Case Tip_Opc When 'C' Then 'Carácter' When 'N' Then 'Numérico' When 'L' Then 'Lógico' When 'F' Then 'Fecha' Else 'Memo' End) as Tipo, ")
            loComandoSeleccionar.AppendLine("				Val_Num, ")
            loComandoSeleccionar.AppendLine("				Val_Car, ")
            loComandoSeleccionar.AppendLine("				Val_Fec, ")
            loComandoSeleccionar.AppendLine("				Val_Log, ")
            loComandoSeleccionar.AppendLine("				Val_Mem, ")
            loComandoSeleccionar.AppendLine("				Comentario ")
            loComandoSeleccionar.AppendLine(" FROM			Opciones ")
            loComandoSeleccionar.AppendLine(" WHERE			Cod_Opc     Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("				And Tip_Opc IN      (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("				And Modulo  Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("				And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY      " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOpciones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrOpciones.ReportSource = loObjetoReporte

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
' MJP: 09/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP: 11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP: 14/07/08 : Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' JJD: 10/01/09 : Ajuste al reporte
'-------------------------------------------------------------------------------------------'
' GCR: 26/03/09 : Adicion de campos y ajustes al diseño.
'-------------------------------------------------------------------------------------------'
' CMS: 31/08/09: Metodo de ordenamiento, verificacionde registros, boton imprimir
'-------------------------------------------------------------------------------------------'
' JJD: 05/11/10: Se incluyo el filtro del modulo.
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'