VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnalyticsTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**
' * This is intended to gather data in how VBA code is being used. This can capture
' * code use, new users, code completion, errors, timing, trends and other metrics.
' *
' * @author Robert Todar <robert@roberttodar.com>
' * @ref {Library} Microsoft Scripting Runtime
' * @ref {Class Module} JSON
' */
Option Explicit

'/**
' * State is a {Type} containing all the private variables for this class.
' */
Private Type state
    stats As Scripting.Dictionary
    startTime As Double
    wasStarted As Boolean
    wasReported As Boolean
    isDisabled As Boolean
End Type

'/**
' * GUILD is unique ids. These are used to access specifc record to append end times to it.
' */
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

'/**
' * property {Long} CoCreateGuid - API Needed to generate UUID.
' * property {state} this - The private properties stored in one place.
' * property {String} filePath - This is where the JSON is stored.
' */
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Private this As state

' Change this to where you want to save the analytics to.
Private Const filePath As String = "C:\test\Analytics.json"

'/**
' * Create new objects on init. Need to make sure xml is valid, will disable if not.
' */
Private Sub Class_Initialize()
    ' All standard stats are collected in this
    ' dictionary. This keeps it dynamic for any
    ' custom stats that can be passed through.
    Set this.stats = New Scripting.Dictionary
    
    ' Make sure there is a file to write to.
    ' If there isn't, or user doesn't have access
    ' then this class is "Disabled" so that errors
    ' won't occur.
    createAnalyticsFileIfNeeded
End Sub

'**/ Check that the file exists */
Private Sub createAnalyticsFileIfNeeded()
    ' Need to make sure the file exists and can be written to otherwise
    ' errors could occur.
    Dim fso As New FileSystemObject
    If Not fso.FileExists(filePath) Then
        ' If the file doesn't exists then it will try to create it.
        ' If it can't be created then it will set the state `isDisabled` to true.
        ' Each function will check that state and will only continue if it is false.
        ' This ensures errors won't occur since the file can't be appended to.
        If Not buildOutFilePath(filePath) = True Then
            this.isDisabled = True
            Debug.Print "Analytics are not being tracked. Unable to create filepath: " & filePath
            Exit Sub
        End If
    End If
    
    ' Need to double check that the user has write access to this folder
    ' otherwise it would cause errors. Disable if user doesn't.
    Dim folderPath As String
    folderPath = fso.GetParentFolderName(filePath)
    If Not hasWriteAccessToFolder(folderPath) Then
        this.isDisabled = True
        Debug.Print "Analytics are not being tracked. User does not have write access to " & filePath
        Exit Sub
    End If
End Sub

'/**
' * If analytics are being tracked and this object is destroyed then it will complete on it's own.
' */
Private Sub Class_Terminate()
    ' This class will finalalize on terminate. But if
    ' the code was already finalized by calling that function then
    ' this doesn't need to run again.
    If this.wasStarted And Not this.wasReported Then
        FinalizeStats
    End If
End Sub

'/**
' * This is the main init by the user. It logs the codeName and starts everything off.
' * @param {String} codeName - The name of the code (function) being run.
' * @param {Double} [minutesToSave] - Allows user to mark ahead of time the time saved.
' */
Public Sub TrackStats(ByVal codeName As String, Optional ByVal minutesToSave As Variant)
    If this.isDisabled Then
        Exit Sub
    End If
    
    ' Display an alert letting the developer know analytics are being tracked.
    Debug.Print "Tracking Analitics For " & codeName
    
    ' Note that is was started to make sure the stats are logged.
    ' These are all the standard stats that are to be collected.
    ' Add or remove from this list as desired.
    this.wasStarted = True
    AddStat "event", "onUse"
    AddStat "id", UUID
    AddStat "codeName", codeName
    AddStat "username", Environ("UserName")
    AddStat "date", CStr(Date)
    AddStat "time", CStr(Time)
    
    ' This field can be ommited and included in the finalize method.
    ' Only add if it was passed in.
    If Not IsMissing(minutesToSave) Then
        AddStat "minutesSaved", minutesToSave
    End If
    
    ' Start the tracking of the runtime. Do this after all values have been
    ' added so that that part of the process isn't included in the total time.
    this.startTime = timer
End Sub

' Adds a stat to the main `onUse` event stat. This can be
' used within this class or outside of it as well.
Public Sub AddStat(ByVal key As String, ByVal value As Variant)
    this.stats(key) = value
End Sub

' Log custom events. These can be whatever data is needed to be loged.
' This logs immediatly on the call.
Public Sub LogEvent(ByVal eventName As String, ParamArray keyValuePairs() As Variant)
    ' Check to see that key/value pairs passed in (an even number).
    If (UBound(keyValuePairs) - LBound(keyValuePairs) + 1) Mod 2 <> 0 Then
        Debug.Print "[Analitics.LogEvent] Unable to log. Invalid parameters: expecting key/value pairs, but received an odd number of arguments."
        Exit Sub
    End If
    
    ' Add standard stats are added by default.
    ' Add or remove as desired.
    Dim stats As New Scripting.Dictionary
    Set stats = New Scripting.Dictionary
    stats.Add "event", eventName
    stats.Add "id", UUID
    stats.Add "username", Environ("UserName")
    stats.Add "date", CStr(Date)
    stats.Add "time", CStr(Time)
    
    ' This step Adds all the key value pairs that were passed in.
    Dim index As Long
    For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        stats.Add keyValuePairs(index), keyValuePairs(index + 1)
    Next index
    
    ' Log a formatted version to the immediate window.
    Debug.Print JSON.Stringify(stats, 2)
    
    ' Append the record to the end of the JSON file.
    appendToTextFile filePath, JSON.Stringify(stats) & ","
End Sub

'/**
' * This finilized the analytics and adds all the finished runtimes.
' * This is also called from Class_Terminate if user doesn't manully call this method.
' *
' * @param {Double} [minutesToSave] - Allows user to add time saved.
' */
Public Sub FinalizeStats(Optional ByVal minutesToSave As Variant)
    If this.isDisabled Then
        Exit Sub
    End If
    
    ' Get the total runtime, rounding this to the third decimal place.
    AddStat "runtimeSeconds", Math.Round(timer - this.startTime, 3)
    
    ' Mark as reported. This is to ensure it does not run twice
    this.wasReported = True
    
    ' This argument doesn't have to come through.
    ' It can be also be added to the `TrackStats` function.
    ' If it is not passed in then do not add it.
    If Not IsMissing(minutesToSave) Then
        AddStat "minutesSaved", minutesToSave
    End If
    
    ' Log a formatted version to the immediate window.
    Debug.Print JSON.Stringify(this.stats, 2)
    
    ' Append record to log file. Note, this file is JSON but will not have needed
    ' Array [] syntax due to appending is much faster then rewriting the entire file.
    ' Will need to run function to fix this.
    appendToTextFile filePath, JSON.Stringify(this.stats) & ","
End Sub

'/**
' * Creates an unique id. This is added to each log record.
' * This is used to try to mimic data stored in a database.
' * Not actually used by anything but might be needed by the user
' * for reporting reasons.
' *
' * @author Gus Molina (open sourced)
' * @copywrite 2000
' * @see <https://stackoverflow.com/questions/7031347/how-can-i-generate-guids-in-excel>
' */
Private Function UUID() As String
    Dim udtGUID As GUID
    If (CoCreateGuid(udtGUID) = 0) Then
        UUID = _
            String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
            String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
            String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
            IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
            IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
            IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
            IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
            IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
            IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
            IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
            IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function

'/**
' * Build entire folder path. The Standard CreateFolder only works for one level.
' * @ref {Library} Microsoft Scripting Runtime
' * @param {String} fullPath - The path that needs to get created.
' * @returns {Boolean} True if no errors occured and path was created.
' */
Private Function buildOutFilePath(ByVal fullPath As String) As Boolean
    On Error GoTo Catch
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    ' Clean path make sure to only have '\' in the path.
    Dim absolutePath As String
    absolutePath = fso.GetAbsolutePathName(fullPath)

    ' Split the folder into each folder name.
    Dim folderNames() As String
    folderNames = split(absolutePath, "\")
    
    ' Loop each folder and make folder path if it doesn't already exist.
    Dim index As Integer
    For index = LBound(folderNames, 1) To UBound(folderNames, 1) - 1
        ' This builds the path in steps.
        Dim currentPath As String
        currentPath = currentPath & folderNames(index) & "\"
        If Not fso.FolderExists(currentPath) Then
            fso.CreateFolder currentPath
        End If
    Next index
    
    ' Lastly, if a file was included, create it if it doesn't exist.
    If Len(fso.GetExtensionName(absolutePath)) > 0 Then
        If Not fso.FileExists(absolutePath) Then
            fso.CreateTextFile fullPath
        End If
    End If
    
    buildOutFilePath = True
    Exit Function
Catch:
    ' Any errors will return false.
End Function

'/**
' * Appends to any type of text file.
' * @ref {Library} Microsoft Scripting Runtime
' * @param {String} filePath - Path to the file to write to.
' */
Private Function appendToTextFile(ByVal filePath As String, ByVal message As String) As Boolean
    On Error GoTo Catch
    Dim fso As New FileSystemObject
    If Not fso.FileExists(filePath) Then
        buildOutFilePath filePath
    End If
    
    Dim ts As TextStream
    Set ts = fso.OpenTextFile(filePath, ForAppending, True)
    ts.WriteLine message
    
    appendToTextFile = True
    Exit Function
Catch:
    ' Errors will return false.
End Function

'/**
' * Attempts to Create a text file and write to it to see if user has write access.
' * @ref {Library} Microsoft Scripting Runtime
' * @example: HasWriteAccessToFolder("C:\Program Files") ~> True || False
' */
Private Function hasWriteAccessToFolder(ByVal folderPath As String) As Boolean
    On Error GoTo Catch
    ' Make sure folder exists, this function returns false if it does not
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Exit Function
    End If

    ' Get unique temp filepath, don't want to overwrite something that already exists
    Do
        Dim count As Integer
        Dim filePath As String
        
        filePath = fso.BuildPath(folderPath, "TestWriteAccess" & count & ".tmp")
        count = count + 1
    Loop Until Not fso.FileExists(filePath)
    
    ' Attempt to create the tmp file, error returns false
    fso.CreateTextFile(filePath).Write ("Test Folder Access")
    Kill filePath
    
    ' No error, able to write to file; return true!
    hasWriteAccessToFolder = True
Catch:
End Function
