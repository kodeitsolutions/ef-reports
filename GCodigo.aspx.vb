'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "GCodigo"
'-------------------------------------------------------------------------------------------'
Partial Class GCodigo

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

Private Function mToIniCap(ByVal lcCadena1 As string) As String 
	
	Dim lcAux As String
	
	lcAux = lcCadena1
	
	If lcAux.Contains("_") Then
		 lcAux = lcAux.ToUpper.Remove(1) & lcAux.ToLower.Substring(1,3) & lcAux.ToUpper.Substring(4,1) & lcAux.ToLower.Substring(5)
	Else
		lcAux = lcAux.ToUpper.Remove(1) & lcAux.ToLower.Substring(1)
	End If

    Return lcAux

End Function



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = cusAplicacion.goReportes.paParametrosIniciales(0)
            Dim lcParametro0Hasta As String = cusAplicacion.goReportes.paParametrosFinales(0)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            Dim loComandoSeleccionar2 As New StringBuilder()
            Dim loComandoSeleccionar3 As New StringBuilder()
            
            DIM lcMatrizCampo AS NEW Arraylist() 
            DIM lcMatrizTipo AS NEW Arraylist() 

			lcMatrizCampo.AddRange(Split(lcParametro0Desde,","))
			lcMatrizTipo.AddRange(Split(lcParametro0Hasta,","))

			For i As Integer = 0 To (lcMatrizCampo.Count - 1)
				'Select Case lcMatrizTipo(i) 
				'	Case "char"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.txt" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")
				'	Case "bit"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.chk" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")
				'	Case "datetime"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.txt" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")
				'	Case "text"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.txt" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")
				'	Case "decimal"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.txt" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")
				'	Case "int"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.txt" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")
				'	Case "*char"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.cbo" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")
				'	Case "*text"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.cbo" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")
				'	Case "*decimal"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.cbo" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")
				'	Case "*int"							
				'			loComandoSeleccionar.AppendLine("Me.mRegistrarControl(Me.cbo" & mToIniCap(lcMatrizCampo(i)) & ",		""" & lcMatrizCampo(i) & """)")														
				'End	Select
				
				'loComandoSeleccionar2.AppendLine("laParametros(" & i & ")		=	(""" & lcMatrizCampo(i) & """)")
				
				loComandoSeleccionar3.AppendLine("Protected Sub txt" & mToIniCap(lcMatrizCampo(i)) & "_mResultadoBusquedaNoValido(ByVal sender As vis1Controles.txtNormal, ByVal lcNombreCampo As String, ByVal lnIndice As Integer) Handles txt" & mToIniCap(lcMatrizCampo(i)) & ".mResultadoBusquedaNoValido")
				loComandoSeleccionar3.AppendLine("")
				loComandoSeleccionar3.AppendLine("		Me.wbcAdministradorMensajeModal.mMostrarMensajeModal(""Advertencia"",  _")
				loComandoSeleccionar3.AppendLine("														""Mesaje."", _ ")
				loComandoSeleccionar3.AppendLine("														vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _ ")
				loComandoSeleccionar3.AppendLine("														""500px"", ""250px"")")
				loComandoSeleccionar3.AppendLine("")
				loComandoSeleccionar3.AppendLine("		Me.txt" & mToIniCap(lcMatrizCampo(i)) & ".mLimpiarCampos()")
				loComandoSeleccionar3.AppendLine("")
				loComandoSeleccionar3.AppendLine("End Sub")
				
			Next i
			
			'me.mEscribirConsulta(loComandoSeleccionar.ToString())
			'me.mEscribirConsulta(loComandoSeleccionar2.ToString())
			me.mEscribirConsulta(loComandoSeleccionar3.ToString())




            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("GCodigo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvGCodigo.ReportSource = loObjetoReporte

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
' CMS: 15/09/08: Codigo inicial.
'-------------------------------------------------------------------------------------------'