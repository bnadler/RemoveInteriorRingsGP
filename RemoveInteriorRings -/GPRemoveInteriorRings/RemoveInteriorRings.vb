' Copyright 2010 ESRI
' 
' All rights reserved under the copyright laws of the United States
' and applicable international laws, treaties, and conventions.
' 
' You may freely redistribute and use this sample code, with or
' without modification, provided you include the original copyright
' notice and use restrictions.
' 
' See the use restrictions at http://help.arcgis.com/en/sdk/10.0/usageRestrictions.htm
' 

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Geoprocessing
Imports ESRI.ArcGIS.DataManagementTools
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.ADF.CATIDs

Namespace GPRemoveInteriorRings
    Public Class RemoveInteriorRingsFunction : Implements IGPFunction2

        ' Local members
        Private m_ToolName As String = "RemoveInteriorRings" 'Function Name
        Private m_MetaDataFile As String = "RemoveInteriorRings_area.xml" 'Metadata file
        Private m_Parameters As IArray ' Array of Parameters
        Private m_GPUtilities As New GPUtilities ' GPUtilities object


#Region "IGPFunction2 Members"
       
        ' This is the location where the parameters to the Function Tool are defined. 
        ' This property returns an IArray of parameter objects (IGPParameter). 
        ' These objects define the characteristics of the input and output parameters. 
        Public ReadOnly Property ParameterInfo() As IArray Implements IGPFunction2.ParameterInfo
            Get
                'Array to the hold the parameters
                Dim pParameters As IArray = New ArrayClass()

                'Define domain to limit valid inputs
                Dim pGPDomain As IGPFeatureClassDomain = New GPFeatureClassDomain
                pGPDomain.AddType(esriGeometryType.esriGeometryPolygon)

                'Input DataType is GPFeatureLayerType
                Dim inputParameter As IGPParameterEdit3 = New GPParameterClass()
                ' Set Input Parameter properties
                With inputParameter
                    .DataType = New GPFeatureLayerTypeClass()
                    .Value = New GPFeatureLayerClass()
                    .Direction = esriGPParameterDirection.esriGPParameterDirectionInput
                    .ParameterType = esriGPParameterType.esriGPParameterTypeRequired
                    .DisplayName = "Input Features"
                    .Name = "input_features"
                    .Domain = CType(pGPDomain, IGPDomain)
                End With
                pParameters.Add(inputParameter)

                ' Output parameter (Derived) and data type is DEFeatureClass
                Dim outputParameter As IGPParameterEdit3 = New GPParameterClass()

                ' Set output parameter properties
                With outputParameter
                    .DataType = New GPFeatureLayerTypeClass()
                    .Value = New GPFeatureLayerClass()
                    .Direction = esriGPParameterDirection.esriGPParameterDirectionOutput
                    .ParameterType = esriGPParameterType.esriGPParameterTypeDerived
                    .DisplayName = "Output FeatureClass"
                    .Name = "out_featureclass"
                    .AddDependency("Input_Features")
                End With

                'Create a new schema object - 
                'schema means the structure or design of the feature class (field information, geometry information, extent)
                Dim outputSchema As IGPFeatureSchema = New GPFeatureSchemaClass()
                With outputSchema
                    .FieldsRule = esriGPSchemaFieldsType.esriGPSchemaFieldsFirstDependency
                    .FeatureTypeRule = esriGPSchemaFeatureType.esriGPSchemaFeatureFirstDependency
                End With
                Dim schema As IGPSchema = CType(outputSchema, IGPSchema)
                schema.GenerateOutputCatalogPath = False

                'Clone the schema from the dependency. 
                'This means update the output with the same schema as the input feature class (the dependency). 
                schema.CloneDependency = True

                'Apply schema to output
                outputParameter.Schema = CType(outputSchema, IGPSchema)
                pParameters.Add(outputParameter)

                Return pParameters
            End Get
        End Property

        ' This method will update the output parameter value with the additional area field.
        Public Sub UpdateParameters(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal pEnvMgr As ESRI.ArcGIS.Geoprocessing.IGPEnvironmentManager) Implements ESRI.ArcGIS.Geoprocessing.IGPFunction2.UpdateParameters
            m_Parameters = paramvalues

        End Sub
        ' Validate: This will validate each parameter and return messages.
        ' This method will check that a given set of parameter values are of the 
        ' appropriate number, DataType, and Value.
        Public Function Validate(ByVal paramvalues As IArray, ByVal updateValues As Boolean, ByVal envMgr As IGPEnvironmentManager) As IGPMessages Implements IGPFunction2.Validate

            If m_Parameters Is Nothing Then
                m_Parameters = ParameterInfo()
            End If

            ' Call InternalValidate (Basic Validation). Are all the required parameters supplied?
            ' Are the Values to the parameters the correct data type?
            Dim validateMsgs As IGPMessages
            validateMsgs = m_GPUtilities.InternalValidate(m_Parameters, paramvalues, updateValues, True, envMgr)

            ' Return the messages
            Return validateMsgs
        End Function
        ' Called after returning from the internal validation routine. You can examine the messages created from internal validation and change them if desired. 
        Public Sub UpdateMessages(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal pEnvMgr As ESRI.ArcGIS.Geoprocessing.IGPEnvironmentManager, ByVal Messages As ESRI.ArcGIS.Geodatabase.IGPMessages) Implements ESRI.ArcGIS.Geoprocessing.IGPFunction2.UpdateMessages

            ' Check for error messages
            Dim msg As IGPMessage
            msg = CType(Messages, IGPMessage)

            If msg.IsError() Then
                'MsgBox(msg.Description)
                Return
            End If

        End Sub
        ' Execute: Execute the function given the array of the parameters
        Public Sub Execute(ByVal paramvalues As IArray, ByVal trackcancel As ITrackCancel, ByVal envMgr As IGPEnvironmentManager, ByVal message As IGPMessages) Implements IGPFunction2.Execute

           
            ' Get the first Input Parameter
            Dim parameter As IGPParameter = CType(paramvalues.Element(0), IGPParameter)

            ' UnPackGPValue. This ensures you get the value either form the dataelement or GpVariable (ModelBuilder)
            Dim parameterValue As IGPValue = m_GPUtilities.UnpackGPValue(parameter)

            ' Open the Input Dataset - use DecodeFeatureLayer as the input might be
            ' a layer file or a feature layer from ArcMap
            Dim inputFeatureClass As IFeatureClass = Nothing
            Dim qf As IQueryFilter = Nothing
            m_GPUtilities.DecodeFeatureLayer(parameterValue, inputFeatureClass, qf)

            If inputFeatureClass Is Nothing Then
                message.AddError(2, "Could not open input dataset.")
                Return
            End If

            Dim featcount As Integer
            featcount = inputFeatureClass.FeatureCount(Nothing)

            ' Set the properties of the UI Step Progresson bar
            Dim pStepPro As IStepProgressor
            pStepPro = CType(trackcancel, IStepProgressor)
            pStepPro.MinRange = 0
            pStepPro.MaxRange = featcount
            pStepPro.StepValue = 1
            pStepPro.Message = "Remove Rings"
            pStepPro.Position = 0
            pStepPro.Show()

            Try
                'Open update cursor & Enumerate
                Dim updateCursor As IFeatureCursor = inputFeatureClass.Update(Nothing, False)
                Dim updateFeature As IFeature = updateCursor.NextFeature()

                Do While Not updateFeature Is Nothing

                    'Get exterior rings from input feature
                    Dim pPoly As IPolygon4 = CType(updateFeature.Shape, IPolygon4)
                    Dim pExtRingBag As IGeometryBag = pPoly.ExteriorRingBag
                    Dim pExtGeomCollection As IGeometryCollection = CType(pExtRingBag, IGeometryCollection)

                    'Create new empty shape
                    Dim pNewPoly As IPolygon4 = New PolygonClass
                    Dim NewGeomCollection As IGeometryCollection = CType(pNewPoly, IGeometryCollection)

                    'Copy exterior rings to empty Shape
                    NewGeomCollection.AddGeometryCollection(pExtGeomCollection)

                    'Apply New Shape to Feature
                    updateFeature.Shape = CType(pNewPoly, IGeometry)

                    'updated shape
                    updateCursor.UpdateFeature(updateFeature)
                    updateFeature.Store()
                    pStepPro.Step()

                    'Next feature
                    updateFeature = updateCursor.NextFeature()
                Loop

                ' Release the cursor to remove the lock on the input data.  
                System.Runtime.InteropServices.Marshal.ReleaseComObject(updateCursor)

            Catch ex As Exception
                message.AddError(2, ex.Message)
            End Try

            pStepPro.Hide()
        End Sub
        ' Set the name of the function tool. 
        ' This name appears when executing the tool at the command line or in scripting. 
        ' This name should be unique to each toolbox and must not contain spaces.
        Public ReadOnly Property Name() As String Implements IGPFunction2.Name
            Get
                Return m_ToolName
            End Get
        End Property

        ' Set the function tool Display Name as seen in ArcToolbox.
        Public ReadOnly Property DisplayName() As String Implements IGPFunction2.DisplayName
            Get
                Return "Remove Interior Rings"
            End Get
        End Property

        ' This is the function name object for the Geoprocessing Function Tool. 
        ' This name object is created and returned by the Function Factory.
        ' The Function Factory must first be created before implementing this property.
        Public ReadOnly Property FullName() As IName Implements IGPFunction2.FullName
            Get
                ' Add RemoveInteriorRings.FullName getter implementation
                Dim functionFactory As IGPFunctionFactory = New RemoveInteriorRingsFunctionFactory()
                'INSTANT VB NOTE: The local variable name was renamed since Visual Basic will not uniquely identify class members when local variables have the same name:
                Return CType(functionFactory.GetFunctionName(m_ToolName), IName)
            End Get
        End Property

        ' This is used to set a custom renderer for the output of the Function Tool.
        Public Function GetRenderer(ByVal pParam As IGPParameter) As Object Implements IGPFunction2.GetRenderer
            Return Nothing
        End Function

        ' This is the unique context identifier in a [MAP] file (.h). 
        ' ESRI Knowledge Base article #27680 provides more information about creating a [MAP] file. 
        Public ReadOnly Property HelpContext() As Integer Implements IGPFunction2.HelpContext
            Get
                Return 0
            End Get
        End Property

        ' This is the path to a .chm file which is used to describe and explain the function and its operation. 
        Public ReadOnly Property HelpFile() As String Implements IGPFunction2.HelpFile
            Get
                Return ""
            End Get
        End Property

        ' This is used to return whether the function tool is licensed to execute.
        Public Function IsLicensed() As Boolean Implements IGPFunction2.IsLicensed
            Return True
        End Function

        ' This is the name of the (.xml) file containing the default metadata for this function tool. 
        ' The metadata file is used to supply the parameter descriptions in the help panel in the dialog. 
        ' If no (.chm) file is provided, the help is based on the metadata file. 
        ' ESRI Knowledge Base article #27000 provides more information about creating a metadata file.
        Public ReadOnly Property MetadataFile() As String Implements IGPFunction2.MetadataFile
            Get
                Return m_MetaDataFile
            End Get
        End Property

        ' This is the class id used to override the default dialog for a tool. 
        ' By default, the Toolbox will create a dialog based upon the parameters returned 
        ' by the ParameterInfo property.
        Public ReadOnly Property DialogCLSID() As UID Implements IGPFunction2.DialogCLSID
            Get
                Return Nothing
            End Get
        End Property
#End Region
        'Whatever....
#Region "IGPFunction Members"

        Public Function GetRenderer1(ByVal pParam As ESRI.ArcGIS.Geoprocessing.IGPParameter) As Object Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.GetRenderer
            Return Nothing
        End Function

        Public ReadOnly Property ParameterInfo1() As ESRI.ArcGIS.esriSystem.IArray Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.ParameterInfo
            Get
                Return ParameterInfo()
            End Get
        End Property

        Public ReadOnly Property DialogCLSID1() As ESRI.ArcGIS.esriSystem.UID Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.DialogCLSID
            Get
                Return DialogCLSID
            End Get
        End Property

        Public ReadOnly Property DisplayName1() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.DisplayName
            Get
                Return DisplayName
            End Get
        End Property

        Public Sub Execute1(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal trackcancel As ESRI.ArcGIS.esriSystem.ITrackCancel, ByVal envMgr As ESRI.ArcGIS.Geoprocessing.IGPEnvironmentManager, ByVal message As ESRI.ArcGIS.Geodatabase.IGPMessages) Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.Execute
            Call Execute(paramvalues, trackcancel, envMgr, message)
        End Sub

        Public ReadOnly Property FullName1() As ESRI.ArcGIS.esriSystem.IName Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.FullName
            Get
                FullName1 = FullName
            End Get
        End Property

        Public ReadOnly Property HelpContext1() As Integer Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.HelpContext
            Get
                Return HelpContext
            End Get
        End Property

        Public ReadOnly Property HelpFile1() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.HelpFile
            Get
                Return HelpFile
            End Get
        End Property

        Public Function IsLicensed1() As Boolean Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.IsLicensed
            Return IsLicensed()
        End Function

        Public ReadOnly Property MetadataFile1() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.MetadataFile
            Get
                Return MetadataFile
            End Get
        End Property

        Public ReadOnly Property Name1() As String Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.Name
            Get
                Return Name
            End Get
        End Property

        Public Function Validate1(ByVal paramvalues As ESRI.ArcGIS.esriSystem.IArray, ByVal updateValues As Boolean, ByVal envMgr As ESRI.ArcGIS.Geoprocessing.IGPEnvironmentManager) As ESRI.ArcGIS.Geodatabase.IGPMessages Implements ESRI.ArcGIS.Geoprocessing.IGPFunction.Validate
            Return Validate(paramvalues, updateValues, envMgr)
        End Function

#End Region
    End Class

    '////////////////////////////
    ' Function Factory Class
    '//////////////////////////
    <Guid("2554BFC7-94F9-4d28-B3FE-14D17599B35A"), ComVisible(True)> _
    Public Class RemoveInteriorRingsFunctionFactory : Implements IGPFunctionFactory
        Private m_GPFunction As IGPFunction

        ' Register the Function Factory with the ESRI Geoprocessor Function Factory Component Category.
#Region "Function Factory"

        <ComRegisterFunction()> _
        Private Shared Sub Reg(ByVal regKey As String)
            GPFunctionFactories.Register(regKey)
        End Sub

        <ComUnregisterFunction()> _
        Private Shared Sub Unreg(ByVal regKey As String)
            GPFunctionFactories.Unregister(regKey)
        End Sub
#End Region

        ' Utility Function added to create the function names.
        Private Function CreateGPFunctionNames(ByVal index As Long) As IGPFunctionName

            Dim functionName As IGPFunctionName = New GPFunctionNameClass()
            'INSTANT VB NOTE: The local variable name was renamed since Visual Basic will not uniquely identify class members when local variables have the same name:
            Dim name_Renamed As IGPName

            Select Case index
                Case (0)
                    name_Renamed = CType(functionName, IGPName)
                    name_Renamed.Category = "RemoveInteriorRings"
                    name_Renamed.Description = "RemoveInterior Rings for Polygon Feature Class"
                    name_Renamed.DisplayName = "Remove Rings"
                    name_Renamed.Name = "RemoveInteriorRings"
                    name_Renamed.Factory = Me
            End Select

            Return functionName
        End Function

        ' Implementation of the Function Factory
#Region "IGPFunctionFactory Members"

        ' This is the name of the function factory. 
        ' This is used when generating the Toolbox containing the function tools of the factory.
        Public ReadOnly Property Name() As String Implements IGPFunctionFactory.Name
            Get
                Return "RemoveInteriorRings"
            End Get
        End Property

        ' This is the alias name of the factory.
        Public ReadOnly Property [Alias]() As String Implements IGPFunctionFactory.Alias
            Get
                Return "RemoveRings"
            End Get
        End Property

        ' This is the class id of the factory. 
        Public ReadOnly Property CLSID() As UID Implements IGPFunctionFactory.CLSID
            Get
                Dim id As UID = New UIDClass()
                id.Value = Me.GetType().GUID.ToString("B")
                Return id
            End Get
        End Property

        ' This method will create and return a function object based upon the input name.
        Public Function GetFunction(ByVal Name As String) As IGPFunction Implements IGPFunctionFactory.GetFunction
            Select Case Name
                Case ("RemoveInteriorRings")
                    m_GPFunction = New RemoveInteriorRingsFunction()
            End Select

            Return m_GPFunction
        End Function

        ' This method will create and return a function name object based upon the input name.
        Public Function GetFunctionName(ByVal Name As String) As IGPName Implements IGPFunctionFactory.GetFunctionName
            Dim gpName As IGPName = New GPFunctionNameClass()

            Select Case Name
                Case ("RemoveInteriorRings")
                    Return CType(CreateGPFunctionNames(0), IGPName)
            End Select
            Return gpName
        End Function

        ' This method will create and return an enumeration of function names that the factory supports.
        Public Function GetFunctionNames() As IEnumGPName Implements IGPFunctionFactory.GetFunctionNames
            ' Add CalculateFunctionFactory.GetFunctionNames implementation
            Dim nameArray As IArray = New EnumGPNameClass()
            nameArray.Add(CreateGPFunctionNames(0))
            Return CType(nameArray, IEnumGPName)
        End Function

        ' This method will create and return an enumeration of GPEnvironment objects. 
        ' If tools published by this function factory required new environment settings, 
        'then you would define the additional environment settings here. 
        ' This would be similar to how parameters are defined. 
        Public Function GetFunctionEnvironments() As IEnumGPEnvironment Implements IGPFunctionFactory.GetFunctionEnvironments
            Return Nothing
        End Function

#End Region
    End Class

End Namespace
