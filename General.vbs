Public Sub main_1()
    'References set to Excel object library and Word object library
    Dim ExcelApp As Excel.Application
    Set ExcelApp = Excel.Application
    Dim currentDirectory As String
    Dim fileName As String
    
    fileName = "General.xlsm"
    
    Application.Workbooks("General").Activate
    
    currentDirectory = CurDir
    
    Debug.Print "=============================================="
    Debug.Print Format(Now(), "mm/dd/yy - hh:mm am/pm")
    Debug.Print "CurrentDirectory = " & currentDirectory
End Sub

Public Sub Demo_Object_FileSystemObject()
    
    Debug.Print "================Demo_Object_FileSystemObject()================="
    Debug.Print Format(Now(), "mm/dd/yy - hh:mm am/pm")
    
    Dim Object_FileSystemObject As New FileSystemObject
    Dim DesktopFolder As Folder
    Dim eachFile As File
    Dim DesktopPath As String, UserProfilePath As String
    
    UserProfilePath = Environ$("USERPROFILE")
    DesktopPath = Object_FileSystemObject.BuildPath(UserProfilePath, "Downloads")
    
    Debug.Print "UserProfilePath    = " & UserProfilePath   'UserProfilePath = C:\Users\fred
    Debug.Print "DesktopPath        = " & DesktopPath       'DesktopPath     = C:\Users\fred\Desktop
    
    Set DesktopFolder = Object_FileSystemObject.GetFolder(DesktopPath)
    
    For Each eachFile In DesktopFolder.Files
        Debug.Print "eachFile.Path = " & eachFile.Path
    Next eachFile
    
End Sub

Public Sub Demo_Write_Text_File()
    
    On Error GoTo Label_Error_Handler
    
    Debug.Print "================Demo_Write_Text_File()================="
    Debug.Print Format(Now(), "mm/dd/yy - hh:mm:ss am/pm")
    
    Dim FSO As New FileSystemObject, output_file As Object
    Dim path_userProfile As String, path_outputFolder As String
    Dim output_fileName As String
    Dim path_output_fileName As String
    
    path_userProfiles = Environ$("USERPROFILE")
    path_outputFolder = FSO.BuildPath(path_userProfiles, "Downloads")
    output_fileName = "temp_" & Format(Now(), "yyyy-mm-dd-hh-mm") & ".txt"
    path_output_fileName = FSO.BuildPath(path_outputFolder, output_fileName)
    
    Debug.Print "path_outputFolder = " & path_outputFolder
    Debug.Print "output_fileName = " & output_fileName
    Debug.Print "path_output_fileName = " & path_output_fileName
    
    If FSO.FolderExists(path_outputFolder) Then
        Debug.Print """" & path_outputFolder & """" & " exist"
    Else
        Debug.Print """" & path_outputFolder & """" & " is invalid"
    End If
    
    If FSO.FileExists(path_output_fileName) Then
        
        Debug.Print path_output_fileName & " already exist"
        Set output_file = FSO.OpenTextFile(path_output_fileName, ForAppending)
        output_file.WriteLine Format(Now(), "yyyy-mm-dd-hh-mm-ss = ") & "7,8,9"
    Else
        Set output_file = FSO.CreateTextFile(path_output_fileName)
        output_file.WriteLine Format(Now(), "yyyy-mm-dd-hh-mm-ss = ") & "Hello, world"
        output_file.WriteLine "1,2,3"
        output_file.WriteLine "5,5,6"
    End If
    
    Debug.Print "FSO.GetFileVersion(path_output_fileName) = " & FSO.GetFileVersion(path_output_fileName)
    Debug.Print "FSO.GetParentFolderName (path_output_fileName) = " & FSO.GetParentFolderName(path_output_fileName)
    Debug.Print "FSO.GetFileName(path_output_fileName) = " & FSO.GetFileName(path_output_fileName)
    Debug.Print "FSO.GetBaseName(path_output_fileName) = " & FSO.GetBaseName(path_output_fileName)
    Debug.Print "FSO.GetExtensionName(path_output_fileName) = " & FSO.GetExtensionName(path_output_fileName)
    
    Debug.Print "FSO.GetParentFolderName(path_output_fileName) = " & FSO.GetParentFolderName(path_output_fileName)
    
    GoTo Label_Exit_Sub

Label_Error_Handler:
    
    Debug.Print Err.Number & ", " & Err.Description & ", " & Err.Number
    Err.Clear

Label_Exit_Sub:
    output_file.Close

End Sub

Public Sub CreatePivotTableAndCharts()

    Debug.Print "================Demo CreatePivotTableAndCharts()================="
    Debug.Print Format(Now(), "mm/dd/yy - hh:mm:ss am/pm")

    Dim PCaches      As PivotCaches
    Dim PCache       As PivotCache
    Dim NewWS        As Worksheet
    Dim PTables      As PivotTables
    Dim PTable       As PivotTable
    Dim PChartShape  As Shape
    Dim PChart       As Chart
    Dim PChartSheet  As Chart

    Application.Workbooks("General").Activate

    'Create PivotCache
    Set PCaches = ThisWorkbook.PivotCaches
    Set PCache = _
    PCaches.Create(xlDatabase, ThisWorkbook.Worksheets(1).Range("A1").CurrentRegion)

    'Create PivotTable
    Set NewWS = ThisWorkbook.Worksheets.Add
    Set PTables = NewWS.PivotTables
    Set PTable = PTables.Add(PCache, NewWS.Range("A1"))
    PTable.AddFields RowFields:=Array("Gardener", "Crop Description")
    PTable.AddDataField PTable.PivotFields("Harvest Units"), "Sum of Harvest Units", xlSum

    'Create Embdedded Chart
    Set PChartShape = NewWS.Shapes.AddChart2(XlChartType:=xlColumnStacked)
    Set PChart = PChartShape.Chart
    PChart.SetSourceData PTable.TableRange1
    
    'Create Sheet Chart
    Set PChartSheet = ThisWorkbook.Charts.Add
    PChartSheet.SetSourceData PTable.TableRange1

End Sub

Sub Demo_Create_Folder()

    Debug.Print "================Demo_Create_Folder================="
    Debug.Print Format(Now(), "mm/dd/yy - hh:mm:ss am/pm")
    
    Dim MyFSO As New FileSystemObject
    Dim path_userProfiles As String
    
    path_userProfiles = Environ$("USERPROFILE")
    path_downloads_folder = MyFSO.BuildPath(path_userProfiles, "Downloads")
    path_new_folder = MyFSO.BuildPath(path_downloads_folder, "Test_1")
    
    If MyFSO.FolderExists(path_new_folder) Then
        Debug.Print path_new_folder & " is already already exist"
    Else
        MyFSO.CreateFolder (path_new_folder)
        Debug.Print path_new_folder & " was created at " & Format(Now(), "mm/dd/yy hh:mm:ss")
    End If
    
End Sub

Sub Demo_Get_File_Names()

    Debug.Print "================Demo_Get_File_Names================="
    Debug.Print Format(Now(), "mm/dd/yy - hh:mm:ss am/pm")
    
    Dim MyFSO As New FileSystemObject
    Dim MyFile As File
    Dim MyFolder As Folder
    Dim path_userProfiles As String: path_userProfiles = Environ$("USERPROFILE")
    Dim path_downloads_folder As String: path_downloads_folder = MyFSO.BuildPath(path_userProfiles, "Downloads")
    
    Set MyFolder = MyFSO.GetFolder(path_downloads_folder)
    
    For Each MyFile In MyFolder.Files
        Debug.Print MyFile.Name
    Next MyFile

End Sub

Sub Demo_Get_Sub_Folder_Names()
    
    Debug.Print "================Demo_Get_Sub_Folder_Names================="
    Debug.Print Format(Now(), "mm/dd/yy - hh:mm:ss am/pm")
    
    Dim MyFSO As New FileSystemObject
    Dim path_userProfiles As String: path_userProfiles = Environ$("USERPROFILE")
    Dim path_downloads_folder As String: path_downloads_folder = MyFSO.BuildPath(path_userProfiles, "Downloads")
    Dim Current_Folder As Folder: Set Current_Folder = MyFSO.GetFolder(path_downloads_folder)
    Dim Sub_Folder As Folder
    
    For Each Sub_Folder In Current_Folder.SubFolders
        Debug.Print Sub_Folder.Name
    Next Sub_Folder

End Sub
