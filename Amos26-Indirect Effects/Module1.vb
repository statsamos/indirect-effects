Imports System
Imports Microsoft.VisualBasic
Imports Amos
Imports AmosEngineLib
Imports AmosEngineLib.AmosEngine.TMatrixID
Imports MiscAmosTypes
Imports MiscAmosTypes.cDatabaseFormat
Imports System.Xml

<System.ComponentModel.Composition.Export(GetType(Amos.IPlugin))>
Public Class CustomCode
    Implements IPlugin
    Public isDebug As Boolean = False

    'This plugin was written January 2018 by John Lim for James Gaskin.
    'This plugin was updated 2022 by Joseph Steed

    Public Function Name() As String Implements IPlugin.Name
        Return "Indirect Effects"
    End Function

    Public Function Description() As String Implements IPlugin.Description
        Return "Creates matrices of all possible standardized and unstandardized indirect effects in the model."
    End Function

    Public Function Mainsub() As Integer Implements IPlugin.MainSub

        'Settings to get bootstrap estimates.
        pd.GetCheckBox("AnalysisPropertiesForm", "StandardizedCheck").Checked = True
        pd.GetCheckBox("AnalysisPropertiesForm", "DoBootstrapCheck").Checked = True
        pd.GetTextBox("AnalysisPropertiesForm", "BootstrapText").Text = "2000"
        pd.GetCheckBox("AnalysisPropertiesForm", "ConfidenceBCCheck").Checked = True
        pd.GetTextBox("AnalysisPropertiesForm", "ConfidenceBCText").Text = "90"

        'Create the html output that combines standardized/unstandardized indirect effects.
        CreateOutput()

    End Function

    ' This takes in the path list for an indirect effect, grabs all of the standard regression weights, and multiplies them together for each path and returns the final value calculcated effect.
    Function getStandardizedIndirectEffectForPath(paths As String(), tableStandardizedRegression As XmlElement, numRegression As Integer) As Double

        'I dim a value: standardizedIndirectEffect for the path (as double)
        Dim standardizedIndirectEffect As Double
        Dim theSize As Integer
        theSize = paths.Length - 1
        'grab each path standardized regression weight in the list (should be the number of paths minus one)
        'for i in pathsArray length -1 (0 to 3 for a thing that has 4)
        standardizedIndirectEffect = 1.0
        Dim firstStrings(theSize) As String
        Dim secondStrings(theSize) As String

        For i = 0 To UBound(paths) - 1
            firstStrings(i) = paths(i)
            secondStrings(i) = paths(i + 1)
        Next i
        'MsgBox(ArrayToString(firstStrings) & " " & ArrayToString(secondStrings))

        ' Iterate through the pairs array and multiply the standardized regression weight from each pair
        For i = LBound(firstStrings, 1) To UBound(firstStrings, 1)
            If isDebug Then
                MsgBox(i)
            End If
            For y = 1 To numRegression 'Iterate through Standardized Regression Weights to match each path
                If firstStrings(i) = MatrixName(tableStandardizedRegression, y, 2) And secondStrings(i) = MatrixName(tableStandardizedRegression, y, 0) Then
                    Dim currentPath As Double = MatrixElement(tableStandardizedRegression, y, 3)
                    standardizedIndirectEffect = standardizedIndirectEffect * currentPath 'Multiply standardized estimates with current path standardized estimates
                End If
            Next
        Next

        'multiply all of them together
        Return standardizedIndirectEffect
    End Function


    'Create the html output that combines standardized/unstandardized indirect effects.
    Sub CreateOutput()

        pd.AnalyzeCalculateEstimates()

        'Get regression weights, standardized regression weights, and user-defined estimands w/bootstrap confidence intervals xml tables from the output.
        Dim tableRegression As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Regression Weights:']/table/tbody")
        ' Standardized Regression Weights 
        Dim tableStandardizedRegression As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='estimates']/div[@ntype='scalars']/div[@nodecaption='Standardized Regression Weights:']/table/tbody")
        Dim tableBootstrap As XmlElement = GetXML("body/div/div[@ntype='models']/div[@ntype='model'][position() = 1]/div[@ntype='group'][position() = 1]/div[@ntype='bootstrap']/div[@ntype='bootstrapconfidence']/div[@ntype='biascorrected']/div[@ntype='scalars']/div[@nodecaption='User-defined estimands:']/table/tbody")
        Dim numRegression As Integer = GetNodeCount(tableRegression)
        Dim numBootstrap As Integer = GetNodeCount(tableBootstrap)
        Dim standardizedIndirectEffects(numBootstrap - 1) As Double 'Array to hold standardized indirect effects

        For x = 1 To numBootstrap 'Iterate through the table of bootstrap estimates.
            Dim paths As String() = Strings.Split(MatrixName(tableBootstrap, x, 0), " --> ") 'Split variable names into array.
            standardizedIndirectEffects(x - 1) = getStandardizedIndirectEffectForPath(paths, tableStandardizedRegression, numRegression)
            'For y = 1 To numRegression 'Iterate through Standardized Regression Weights to match first path.
            ' If paths(0) = MatrixName(tableStandardizedRegression, y, 2) And paths(1) = MatrixName(tableStandardizedRegression, y, 0) Then
            ' Dim firstPath As Double = MatrixElement(tableStandardizedRegression, y, 3)
            ' For z = 1 To numRegression 'Iterate through Standardized Regression Weights again to match second path.
            ' If paths(1) = MatrixName(tableStandardizedRegression, z, 2) And paths(2) = MatrixName(tableStandardizedRegression, z, 0) Then
            ' Dim secondPath As Double = MatrixElement(tableStandardizedRegression, z, 3)
            ' MsgBox("firstPath: " & firstPath & "; secondPath: " & secondPath & "; MatrixName(tableBootstrap, x, 0): " & MatrixName(tableBootstrap, x, 0) & "; paths: " & ArrayToString(paths) & "; MatrixName(tableStandardizedRegression, z, 2): " & MatrixName(tableStandardizedRegression, z, 2) & "; MatrixName(tableStandardized, y, 2): " & MatrixName(tableStandardizedRegression, y, 2) & "; MatrixName(tableStandardized, y, 0): " & MatrixName(tableStandardizedRegression, y, 0))
            ' standardizedIndirectEffects(x - 1) = firstPath * secondPath 'Multiply standardized estimates together to get indirect effect
            ' End If
            ' Next
            '     End If
            'Next

        Next


        'Delete the output file if it exists
        If (System.IO.File.Exists("IndirectEffects.html")) Then
            System.IO.File.Delete("IndirectEffects.html")
        End If

        'Start the debugger for the html output
        Dim debug As New AmosDebug.AmosDebug
        Dim resultWriter As New TextWriterTraceListener("IndirectEffects.html")
        Trace.Listeners.Clear()
        Trace.Listeners.Add(resultWriter)

        debug.PrintX("<html><body><h1>Indirect Effects</h1><hr/>")
        If isDebug Then

            debug.PrintX("<br>")
            debug.PrintX("<br>")
            debug.PrintX(tableBootstrap.ToString())
            debug.PrintX("<br>")
            debug.PrintX(ArrayToString(standardizedIndirectEffects))
            debug.PrintX("<br>")
        End If
        debug.PrintX("<br>")
        debug.PrintX("<br>")

        'Populate model fit measures in data table
        debug.PrintX("<table><tr><th>Indirect Path</th><th>Unstandardized Estimate</th><th>Lower</th><th>Upper</th><th>P-Value</th><th>Standardized Estimate</th></tr><tr>")

        For i = 1 To numBootstrap
            debug.PrintX("<td>" + MatrixName(tableBootstrap, i, 0) + "</td>") 'Name of indirect path
            debug.PrintX("<td>" + MatrixElement(tableBootstrap, i, 3).ToString("#0.000")) 'Estimate
            debug.PrintX("<td>" + MatrixElement(tableBootstrap, i, 4).ToString("#0.000")) 'Lower
            debug.PrintX("<td>" + MatrixElement(tableBootstrap, i, 5).ToString("#0.000")) 'Upper
            debug.PrintX("<td>" + MatrixElement(tableBootstrap, i, 6).ToString("#0.000")) 'P-Value

            'Output the significance significance with the standardized estimate
            If MatrixName(tableBootstrap, i, 6) = "***" Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "***</td>")
            ElseIf MatrixName(tableBootstrap, i, 6) = "" Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) = 0 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) < 0.001 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "***</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) < 0.01 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "**</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) < 0.05 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "*</td>")
            ElseIf MatrixElement(tableBootstrap, i, 6) < 0.1 Then
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "&#x271D;</td>")
            Else
                debug.PrintX("<td>" + standardizedIndirectEffects(i - 1).ToString("#0.000") + "</td>")
            End If

            debug.PrintX("</tr>")
        Next

        'References
        debug.PrintX("</table><h3>References</h3>Significance of Estimates:<br>*** p < 0.001<br>** p < 0.010<br>* p < 0.050<br>&#x271D; p < 0.100<br>")
        debug.PrintX("<p>--If you would like to cite this tool directly, please use the following:")
        debug.PrintX("Gaskin, J., James, M., Lim, J, & Steed, J. (2022), ""Indirect Effects"", AMOS Plugin. <a href=""http://statwiki.gaskination.com"">Gaskination's StatWiki</a>.</p>")

        'Write style And close
        debug.PrintX("<style>table{border:1px solid black;border-collapse:collapse;}td{border:1px solid black;text-align:center;padding:5px;}th{text-weight:bold;padding:10px;border: 1px solid black;}</style>")
        debug.PrintX("</body></html>")

        'Take down our debugging, release file, open html
        Trace.Flush()
        Trace.Listeners.Remove(resultWriter)
        resultWriter.Close()
        resultWriter.Dispose()
        Process.Start("IndirectEffects.html")

    End Sub

    'Get a string element from an xml table.
    Function MatrixName(eTableBody As XmlElement, row As Long, column As Long) As String

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixName = e.InnerText
        Catch ex As Exception
            MatrixName = ""
        End Try

    End Function

    'Get a number from an xml table.
    Function MatrixElement(eTableBody As XmlElement, row As Long, column As Long) As Double

        Dim e As XmlElement

        Try
            e = eTableBody.ChildNodes(row - 1).ChildNodes(column) 'This means that the rows are not 0 based.
            MatrixElement = CDbl(e.GetAttribute("x"))
        Catch ex As Exception
            MatrixElement = 0
        End Try

    End Function

    'Use an output table path to get the xml version of the table.
    Function GetXML(path As String) As XmlElement

        Dim doc As Xml.XmlDocument = New Xml.XmlDocument()
        doc.Load(Amos.pd.ProjectName & ".AmosOutput")
        Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(doc.NameTable)
        Dim eRoot As Xml.XmlElement = doc.DocumentElement

        GetXML = eRoot.SelectSingleNode(path, nsmgr)

    End Function

    'Get the number of rows in an xml table.
    Function GetNodeCount(table As XmlElement) As Integer

        Dim nodeCount As Integer = 0

        'Handles a model with zero correlations
        Try
            nodeCount = table.ChildNodes.Count
        Catch ex As NullReferenceException
            nodeCount = 0
        End Try

        GetNodeCount = nodeCount

    End Function

    'Set all the path values to null.
    Sub ClearPaths()
        'Set paths back to null.
        For Each variable As PDElement In pd.PDElements 'Iterate through the paths in the model
            If variable.IsPath Then
                If (variable.Variable1.IsLatentVariable And variable.Variable2.IsLatentVariable) Or (variable.Variable1.IsObservedVariable And variable.Variable2.IsObservedVariable) Then
                    variable.Value1 = variable.Variable1.NameOrCaption + " | " + variable.Variable2.NameOrCaption 'Change path to names of the connected variables.
                End If
            End If
        Next
    End Sub

    'Set path values to concatenated names of connected variables.
    Sub NamePaths()
        For Each variable As PDElement In pd.PDElements 'Iterate through the paths in the model
            If variable.IsPath Then
                If (variable.Variable1.IsLatentVariable And variable.Variable2.IsLatentVariable) Or (variable.Variable1.IsObservedVariable And variable.Variable2.IsObservedVariable) Then
                    variable.Value1 = variable.Variable1.NameOrCaption + " | " + variable.Variable2.NameOrCaption 'Change path to names of the connected variables.
                End If
            End If
        Next
    End Sub

    Function ArrayToString(arr() As Double) As String
        Dim str As String
        Dim i As Integer

        ' Loop through the array and add each element to the string with a comma separator
        For i = LBound(arr) To UBound(arr)
            str = str & arr(i) & "<br>"
        Next i

        ' Remove the last comma from the string
        If Len(str) > 0 Then
            str = Left(str, Len(str) - 1)
        End If

        ArrayToString = str
    End Function

    Function ArrayToString(arr() As String) As String
        Dim str As String
        Dim i As Integer

        ' Loop through the array and add each element to the string with a comma separator
        For i = LBound(arr) To UBound(arr)
            str = str & arr(i) & "<br>"
        Next i

        ' Remove the last comma from the string
        If Len(str) > 0 Then
            str = Left(str, Len(str) - 1)
        End If

        ArrayToString = str
    End Function


End Class


