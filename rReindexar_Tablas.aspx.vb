'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports System.Data.SqlClient

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rReindexar_Tablas"
'-------------------------------------------------------------------------------------------'
Partial Class rReindexar_Tablas		   
    Inherits vis2formularios.frmReporte
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
   
   public property pcResultado As DataSet
   	   
   	   Get
			  Return Viewstate("pcResultado")
   	   End Get
   	   
   	   Set(ByVal value As DataSet)
			Viewstate("pcResultado") = Value
   	   End Set
   	   
   End Property
Public Function mObtenerTodosLocal(ByVal lcCadenaConexion  As String, ByVal lcComandoSelect As String, _
                                         ByVal lcNombreTabla As String) As DataSet

    Dim         loConexionExterna As New SqlConnection(lcCadenaConexion)
	Dim TimeOut As SqlCommand = loConexionExterna.CreateCommand() 
	
	TimeOut.CommandText = lcComandoSelect
	TimeOut.CommandTimeout = 0           
	 
    Dim         loDataAdapter     As New SqlDataAdapter( TimeOut)
    Dim         loDataSet         As New DataSet

    Try
		   
          loConexionExterna.Open()
          loDataAdapter.FillSchema(loDataSet,SchemaType.Source,lcNombreTabla)
          loDataAdapter.Fill(loDataSet, lcNombreTabla)
          loConexionExterna.Close()
    Catch loExcepcion As Exception

          Throw New Exception("Error en mObtenerTodosLocal: " + loExcepcion.Message, loExcepcion.InnerException)

    End Try
Return loDataSet

End Function


    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

			
				loComandoSeleccionar.AppendLine(" DECLARE Cur CURSOR FOR      ")
				loComandoSeleccionar.AppendLine("      SELECT Name From Sys.Tables ORDER BY Name")
				loComandoSeleccionar.AppendLine("   ")
				loComandoSeleccionar.AppendLine(" DECLARE @Consulta NVARCHAR(MAX),  ")
				loComandoSeleccionar.AppendLine(" 		@Numero INT,  ")
				loComandoSeleccionar.AppendLine(" 		@Temp NVARCHAR(MAX)  ")
				loComandoSeleccionar.AppendLine("   ")
				loComandoSeleccionar.AppendLine(" SET		@Consulta = ''  ")
				loComandoSeleccionar.AppendLine(" SET		@Numero = 0  ")
				loComandoSeleccionar.AppendLine("   ")
				loComandoSeleccionar.AppendLine(" OPEN Cur      ")
				loComandoSeleccionar.AppendLine("  -- Consiguiendo la cadena a ejecutar  ")
				loComandoSeleccionar.AppendLine("      FETCH NEXT FROM Cur INTO @Temp      ")
				loComandoSeleccionar.AppendLine("      WHILE @@FETCH_STATUS = 0      ")
				loComandoSeleccionar.AppendLine("          BEGIN      ")
				loComandoSeleccionar.AppendLine("  		      SET @Consulta = @Consulta  + 'DBCC DBREINDEX (''' + LTRIM(RTRIM(@Temp)) + ''') '+ nchar(13) + ''  ")
				loComandoSeleccionar.AppendLine("  		      SET @Numero = @Numero + 1 ")
				loComandoSeleccionar.AppendLine("  		      FETCH NEXT FROM Cur INTO @Temp      ")
				loComandoSeleccionar.AppendLine("          End      ")
				loComandoSeleccionar.AppendLine("   CLOSE Cur      ")
				loComandoSeleccionar.AppendLine("   DEALLOCATE Cur     ")
				loComandoSeleccionar.AppendLine("   ")							   
				loComandoSeleccionar.AppendLine("  EXEC sp_executesql @Consulta  ")
				loComandoSeleccionar.AppendLine("   ")
				loComandoSeleccionar.AppendLine("  SET   @Consulta  = @Consulta  + NCHAR(13) + CAST(@Numero AS VARCHAR) + ' Tablas afectadas'")
				loComandoSeleccionar.AppendLine("   ")							   				
				loComandoSeleccionar.AppendLine("  SELECT CAST (@Consulta AS TEXT) AS Consulta  ")

			Dim Conexion As String
			Dim laDatosReporte As DataSet
			If NOt Me.IsPostBack  Then

				Conexion = goDatos.pcCadenaConexion()   						
				laDatosReporte = mObtenerTodosLocal(Conexion.ToString, loComandoSeleccionar.ToString(),  "curReportes")
				pcResultado = laDatosReporte 
            
            Else
				
				laDatosReporte  = pcResultado
								
			End If
			

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rReindexar_Tablas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrReindexar_Tablas.ReportSource = loObjetoReporte

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
' CMS: 10/03/10: Codigo inicial.
'-------------------------------------------------------------------------------------------'