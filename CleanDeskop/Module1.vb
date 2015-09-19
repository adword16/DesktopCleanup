Imports System.ComponentModel
Imports System.Globalization
Imports System.IO

Module Module1
    Private _AnyFileErrorHapend As Boolean
    Private Const OLD_DIR_PRE_FIX As String = "old_"

    Sub Main(args() As String)
        Dim toKeepDirectorys = {"keep"}
        Dim oldPath As String = Nothing
        Dim maxKeepTimeinDays As Integer = 30
        Try
            Dim dir = args.ToList.FirstOrDefault
            If String.IsNullOrEmpty(dir) Then
                Throw New InvalidEnumArgumentException("please call with path to desktop as argument")
            End If

            If Not Directory.Exists(dir) Then
                Throw New DirectoryNotFoundException(dir)
            End If


            If args.Count > 1 Then
                For i = 1 To args.Count - 1
                    If args(i).StartsWith("-oldPath", True, CultureInfo.InvariantCulture) Then
                        oldPath = GetNextArgument(args, i)
                        If Not Directory.Exists(oldPath) Then
                            Throw New DirectoryNotFoundException(oldPath)
                        End If
                    End If
                Next
            End If

            ConsoleLog(String.Format("Cleaning Desktop '{0}'...", dir))

            If String.IsNullOrWhiteSpace(oldPath) Then
                ForcCleanDirectory(dir, toKeepDirectorys)
            Else
                MoveTolOldDirectory(dir, oldPath, toKeepDirectorys)
                CleanOldDirectorys(oldPath, maxKeepTimeinDays)
            End If

            If _AnyFileErrorHapend Then
                Console.ReadLine()
            End If

        Catch ex As Exception
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine(ex.ToString)
            Console.ReadLine()

        End Try
    End Sub

    Private Sub CleanOldDirectorys(ByVal oldPath As String, ByVal maxKeepTimeinDays As Integer)
        Dim dirs = Directory.GetDirectories(oldPath)
        For Each dirPath In dirs
            Dim dirName As String = Path.GetFileName(dirPath)
            If dirName.StartsWith(OLD_DIR_PRE_FIX, True, CultureInfo.InvariantCulture) Then
                Dim dirNameWithOutPrefix = dirName.Substring(OLD_DIR_PRE_FIX.Length)
                Dim dirDate = Date.Parse(dirNameWithOutPrefix)
                If CInt(Now.Subtract(dirDate).TotalDays) >= maxKeepTimeinDays Then
                    Try
                        Directory.Delete(dirPath, True)
                    Catch ex As Exception
                        ConsoleLogFileActionException(ex)
                    End Try
                End If
            End If
        Next
    End Sub

    Private Sub ConsoleLog(str As String)
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine(str)
    End Sub

    Private Sub MoveTolOldDirectory(directoryPath As String, oldPath As String, toKeepDirectorys As String())
        Dim oldPathDir = Path.Combine(oldPath, OLD_DIR_PRE_FIX & GetCurrentTimeDirectoryName())
        If Not Directory.Exists(oldPathDir) Then
            Directory.CreateDirectory(oldPathDir)
        End If

        For Each path In Directory.GetDirectories(directoryPath)
            Dim currentDirectoryName = IO.Path.GetFileName(path)
            Dim ignoreDir = toKeepDirectorys.Any(Function(t) String.Compare(t, currentDirectoryName, True) = 0)
            If Not ignoreDir Then
                Dim toMoveToDirectory = IO.Path.Combine(oldPathDir, currentDirectoryName)

                'if file ealready exist create a directory like "dir(1)"
                If Directory.Exists(toMoveToDirectory) Then
                    For index As Integer = 1 To Integer.MaxValue - 1
                        Dim newDirectoryName = String.Format("{0}({1})", currentDirectoryName, index)
                        toMoveToDirectory = IO.Path.Combine(oldPathDir, newDirectoryName)
                        If Not Directory.Exists(toMoveToDirectory) Then
                            Exit For
                        End If
                    Next
                End If

                Try
                    Directory.Move(path, toMoveToDirectory)
                Catch ex As IOException
                    ConsoleLogFileActionException(ex)
                End Try

            End If
        Next

        For Each filePath In Directory.GetFiles(directoryPath)
            Dim newFilePath As String = Path.Combine(oldPathDir, Path.GetFileName(filePath))

            'if file ealready exist create a file like "filename(1).ext"
            If File.Exists(newFilePath) Then
                For index As Integer = 1 To Integer.MaxValue - 1
                    Dim newFileName As String = String.Format("{0}({1}){2}", Path.GetFileNameWithoutExtension(filePath), index, Path.GetExtension(filePath))
                    newFilePath = Path.Combine(oldPathDir, newFileName)
                    If Not File.Exists(newFilePath) Then
                        Exit For
                    End If
                Next
            End If

            Try
                File.Move(filePath, newFilePath)
            Catch ex As IOException
                ConsoleLogFileActionException(ex)
            End Try

        Next
    End Sub

    Private Sub ConsoleLogFileActionException(ioException As IOException)
        _AnyFileErrorHapend = true
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine(ioException.Message)
    End Sub

    Private Function GetCurrentTimeDirectoryName() As String
        Return Format(Now.Date, "yyyy-MM-dd")
    End Function

    Private Sub ForcCleanDirectory(directoryPath As String, toKeepDirectorys As String())
        For Each path In Directory.GetDirectories(directoryPath)
            Dim dirName = IO.Path.GetDirectoryName(path)
            Dim ignoreDir = toKeepDirectorys.Any(Function(t) String.Compare(t, dirName, True) = 0)
            If Not ignoreDir Then
                Directory.Delete(path, True)
            End If
        Next
        For Each file In Directory.GetFiles(directoryPath)
            IO.File.Delete(file)
        Next
    End Sub

    Private Function GetNextArgument(strArray As String(), currentIndex As Integer) As String
        If strArray.Count <= currentIndex + 1 Then
            Return String.Empty
        End If
        Return strArray(currentIndex + 1)
    End Function
End Module
