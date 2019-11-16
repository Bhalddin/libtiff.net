Imports System.Diagnostics
Imports System.IO
Imports System.Drawing
Imports BitMiracle.LibTiff.Classic

Namespace BitMiracle.LibTiff.Samples
    Public NotInheritable Class MultiplePagesToSides
        Public Shared watchFolder As FileSystemWatcher
        Private Sub New()
        End Sub
        Private Class Size
            Sub New(height As Integer, width As Integer)
                _height = height
                _width = width
            End Sub

            Private _width As Integer
            Private _height As Integer

            Property Height As Integer
                Get
                    Return _height
                End Get
                Set(ByVal Value As Integer)
                    _height = Value
                End Set
            End Property

            Property Width As Integer
                Get
                    Return _width
                End Get
                Set(ByVal Value As Integer)
                    _width = Value
                End Set
            End Property

        End Class
        Private Class Offset
            Sub New(offsetX As Integer, offsetY As Integer)
                _offsetX = offsetX
                _offsetY = offsetY
            End Sub

            Private _offsetY As Integer
            Private _offsetX As Integer

            Property OffsetX As Integer
                Get
                    Return _offsetX
                End Get
                Set(ByVal Value As Integer)
                    _offsetX = Value
                End Set
            End Property

            Property OffsetY As Integer
                Get
                    Return _offsetY
                End Get
                Set(ByVal Value As Integer)
                    _offsetY = Value
                End Set
            End Property

        End Class
        Public Shared Sub Main()
            InitializeFolderWatcher()
            Console.WriteLine("Press X key to exit application.")

            Dim exitConsole As Boolean = False

            Do While exitConsole = False

                Dim consoleKey As ConsoleKeyInfo = Console.ReadKey()

                'if the key pressed is x then

                If consoleKey.Key = System.ConsoleKey.X Then

                    exitConsole = True 'exit the console

                End If

            Loop
            watchFolder.EnableRaisingEvents = False
        End Sub
        Private Shared Sub InitializeFolderWatcher()
            watchFolder = New System.IO.FileSystemWatcher()

            'this is the path we want to monitor
            watchFolder.Path = My.Settings.FolderToWatch

            'Add a list of Filter we want to specify
            'make sure you use OR for each Filter as we need to
            'all of those 

            watchFolder.NotifyFilter = IO.NotifyFilters.DirectoryName
            watchFolder.NotifyFilter = watchFolder.NotifyFilter Or
                                       IO.NotifyFilters.FileName
            watchFolder.NotifyFilter = watchFolder.NotifyFilter Or
                                       IO.NotifyFilters.Attributes

            ' add the handler to each event
            AddHandler watchFolder.Changed, AddressOf LogChange
            AddHandler watchFolder.Created, AddressOf LogChange
            AddHandler watchFolder.Deleted, AddressOf LogChange

            ' add the rename handler as the signature is different
            AddHandler watchFolder.Renamed, AddressOf LogRename

            'Set this property to true to start watching
            watchFolder.EnableRaisingEvents = True
            Console.WriteLine("Tiff Image Process Agent started....")
            Console.WriteLine("Watching folder {0}", Path.GetFullPath(My.Settings.FolderToWatch))
        End Sub
        Private Shared Sub LogChange(ByVal source As Object, ByVal e As _
                        System.IO.FileSystemEventArgs)
            If Not Path.GetExtension(e.FullPath).ToLower = ".tif" Then Return
            If e.ChangeType = IO.WatcherChangeTypes.Changed Then
                Console.WriteLine("File {0} has been changed", e.FullPath)
                Dim bmpFile As String = String.Concat(Path.GetDirectoryName(e.FullPath), "\",
                                                      Path.GetFileNameWithoutExtension(e.Name),
                                                      ".bmp")
                Console.WriteLine("Changing {0} file...", bmpFile)
                ConvertImage(e.FullPath, bmpFile)
                Console.WriteLine("File {0} has been Changed successfully!", bmpFile)
            End If
            If e.ChangeType = IO.WatcherChangeTypes.Created Then
                Console.WriteLine("File {0} has been created", e.FullPath)
                Dim bmpFile As String = String.Concat(Path.GetDirectoryName(e.FullPath), "\",
                                                      Path.GetFileNameWithoutExtension(e.Name),
                                                      ".bmp")
                Console.WriteLine("creating {0} file...", bmpFile)
                ConvertImage(e.FullPath, bmpFile)
                Console.WriteLine("File {0} has been created successfully!", bmpFile)
            End If
            If e.ChangeType = IO.WatcherChangeTypes.Deleted Then
                Console.WriteLine("File {0} has been created", e.FullPath)
                Dim bmpFile As String = String.Concat(Path.GetDirectoryName(e.FullPath), "\",
                                                      Path.GetFileNameWithoutExtension(e.Name),
                                                      ".bmp")
                If File.Exists(bmpFile) Then
                    Console.WriteLine("Deleting {0} file...", bmpFile)
                    File.Delete(bmpFile)
                    Console.WriteLine("File {0} deleted successfully!", bmpFile)
                End If
            End If
        End Sub
        Public Shared Sub LogRename(ByVal source As Object, ByVal e As _
                            System.IO.RenamedEventArgs)
            Console.WriteLine("File {0} has been renamed", e.OldFullPath)
            Dim oldBmpFile As String = String.Concat(Path.GetDirectoryName(e.FullPath), "\",
                                                  Path.GetFileNameWithoutExtension(e.Name),
                                                  ".bmp")
            Dim NewBmpFile As String = String.Concat(Path.GetDirectoryName(e.FullPath), "\",
                                                  Path.GetFileNameWithoutExtension(e.Name),
                                                  ".bmp")
            If File.Exists(oldBmpFile) Then
                Console.WriteLine("renaming {0} file...", oldBmpFile)
                Console.WriteLine("New file name is : {0}", NewBmpFile)
                My.Computer.FileSystem.RenameFile(oldBmpFile, NewBmpFile)
                Console.WriteLine("File {0} renamed successfully!")
            End If
        End Sub
        Private Shared Sub ConvertImage(TiffFile As String, BmpFile As String)
            Using input As Tiff = Tiff.Open(TiffFile, "r")
                If input Is Nothing Then
                    Console.WriteLine("Could not open incoming image")
                    Return
                End If

                If input.IsTiled() Then
                    Console.WriteLine("Could not process tiled image")
                    Return
                End If

                Dim targetHeight As Integer = 0
                Dim targetWidth As Integer = 0

                Dim numberOfDirectories As Integer = input.NumberOfDirectories()
                Dim sizes As ArrayList = New ArrayList
                Dim rasters As ArrayList = New ArrayList
                Dim offsetPoints As ArrayList = New ArrayList

                For i As Short = 0 To numberOfDirectories - 1
                    input.SetDirectory(i)


                    ' Find the width and height of the image
                    Dim value As FieldValue() = input.GetField(TiffTag.IMAGEWIDTH)
                    Dim width As Integer = value(0).ToInt()

                    value = input.GetField(TiffTag.IMAGELENGTH)
                    Dim height As Integer = value(0).ToInt()

                    Dim size As Size = New Size(height, width)
                    sizes.Add(size)


                    Dim imageSize As Integer = height * width
                    Dim raster As Integer() = New Integer(imageSize - 1) {}


                    ' Read the image into the memory buffer
                    If Not input.ReadRGBAImage(width, height, raster) Then
                        Console.WriteLine("Could not read image")
                        Return
                    End If

                    rasters.Add(raster)
                    Console.WriteLine("Width:{0} Height:{1}", width, height)

                    Dim offset As Offset = New Offset(targetWidth, 0)
                    offsetPoints.Add(offset)

                    targetHeight = Math.Max(targetHeight, height)
                    targetWidth += width


                Next
                Console.WriteLine("New Image:- Width:{0} Height:{1}", targetWidth, targetHeight)

                'get offsetY
                For i As Short = 0 To numberOfDirectories - 1
                    DirectCast(offsetPoints(i), Offset).OffsetY = targetHeight / 2 - DirectCast(sizes(i), Size).Height / 2
                Next


                Using bmp As New Bitmap(targetWidth, targetHeight)
                    For i As Short = 0 To numberOfDirectories - 1
                        input.SetDirectory(i)
                        For x As Integer = 0 To DirectCast(sizes(i), Size).Width - 1
                            For y As Integer = 0 To DirectCast(sizes(i), Size).Height - 1
                                bmp.SetPixel(DirectCast(offsetPoints(i), Offset).OffsetX + x, DirectCast(offsetPoints(i), Offset).OffsetY + y,
                                             GetColor(x, y, rasters(i), sizes(i).width, sizes(i).height))
                            Next
                        Next
                        bmp.Save(BmpFile)
                    Next
                End Using
            End Using

        End Sub
        Private Shared Function GetColor(ByVal x As Integer, ByVal y As Integer, ByVal raster As Integer(), ByVal width As Integer, ByVal height As Integer) As Color
            Dim offset As Integer = (height - y - 1) * width + x
            Dim red As Integer = Tiff.GetR(raster(offset))
            Dim green As Integer = Tiff.GetG(raster(offset))
            Dim blue As Integer = Tiff.GetB(raster(offset))
            Return Color.FromArgb(red, green, blue)
        End Function
    End Class
End Namespace