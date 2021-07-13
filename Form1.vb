Option Explicit On
'Option Strict On

Imports System.Configuration
Imports System.ComponentModel
Imports System.IO

Public Class Form1

    Private Class Cue

        'Public Variables
        Public title As String
        Public uRL As String
        Public length As DateTime

        Public Shared Function GetWAVDuration(ByVal strPathAndFilename As String) As String

            REM *** SEE ALSO
            REM *** http://soundfile.sapp.org/doc/WaveFormat/
            REM *** CAUTION
            REM *** On the WaveFormat web page, the positions of the values
            REM *** SampleRate and BitsPerSample have been reversed.
            REM *** https://stackoverflow.com/questions/65588931/determine-wav-duration-time-in-visual-studio-2019/65616287#65616287

            REM *** DEFINE LOCAL VARIABLES

            REM *** Define Visual Basic BYTE Array Types
            Dim byNumberOfChannels() As Byte
            Dim bySamplesPerSec() As Byte
            Dim byBitsPerSample() As Byte
            Dim bySubChunkToSizeData() As Byte

            Dim nNumberOfChannels As Int16
            Dim nSamplesPerSec As Int16
            Dim nBitsPerSample As Int32
            Dim nSubChunkToSizeData As Int32

            Dim dNumberOfSamples As Double
            Dim nDurationInMillis As Int32

            Dim strDuration As String

            REM *** INITIALIZE LOCAL VARIABLES
            byNumberOfChannels = New Byte(2) {}
            bySamplesPerSec = New Byte(2) {}
            byBitsPerSample = New Byte(4) {}
            bySubChunkToSizeData = New Byte(4) {}

            nNumberOfChannels = 0
            nSamplesPerSec = 0
            nBitsPerSample = 0L
            nSubChunkToSizeData = 0L

            dNumberOfSamples = 0.0
            nDurationInMillis = 0

            strDuration = ""

            REM *** Initialize the return string value
            GetWAVDuration = ""

            REM ***************************************************************************

            REM *** Open the Input File for READ Operations
            Using fsFileStream = File.OpenRead(strPathAndFilename)

                REM *** Get the Number of Audio Channels
                fsFileStream.Seek(22, SeekOrigin.Begin)
                fsFileStream.Read(byNumberOfChannels, 0, 2)

                REM *** Get the Number of Bits Per Audio Sample
                fsFileStream.Seek(24, SeekOrigin.Begin)
                fsFileStream.Read(byBitsPerSample, 0, 4)

                REM *** Get the number of samples taken per second
                fsFileStream.Seek(34, SeekOrigin.Begin)
                fsFileStream.Read(bySamplesPerSec, 0, 2)

                REM *** Retrieve the size of the WAV data
                REM *** payload in the file
                fsFileStream.Seek(40, SeekOrigin.Begin)
                fsFileStream.Read(bySubChunkToSizeData, 0, 4)

            End Using

            REM *** Convert Values from their BYTE representation

            nNumberOfChannels = BitConverter.ToInt16(byNumberOfChannels, 0)
            nBitsPerSample = BitConverter.ToInt32(byBitsPerSample, 0)
            nSamplesPerSec = BitConverter.ToInt16(bySamplesPerSec, 0)
            nSubChunkToSizeData = BitConverter.ToInt32(bySubChunkToSizeData, 0)

            REM *** Compute the Duration of the WAV File
            REM *** Derives the duration in milliseconds

            REM *** Determine the number of Sound Samples 
            dNumberOfSamples = (nSubChunkToSizeData * 8) / (nNumberOfChannels * nBitsPerSample)
            nDurationInMillis = Convert.ToInt32(1000 * Convert.ToSingle(dNumberOfSamples) / Convert.ToSingle(nSamplesPerSec))

            REM *** Convert the time in Milliseconds to a string format
            REM *** represented by "hh:mm:ss"
            strDuration = ConvertMillisToTimeString(nDurationInMillis)

            REM *** POST METHOD RETURNS
            GetWAVDuration = strDuration

        End Function
    End Class
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    'Indicates current media panel to add controls to
    Private _CurrentMediaPanelName As String = Nothing

    'Used to give unique control names such as label1, label2, etc.
    Private _MediaPanelsAddedCount As Integer = 0

    Private _CueArray As Array

    '*********************Form Setup***********************************
    'Add menu to the form
    Private Sub CreateAudioCueToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateAudioCueToolStripMenuItem.Click

        'Dim x As Integer = 0
        'Do
        '    CreateMediaPanel()
        '    x += 1
        'Loop Until x > 50

        CreateMediaPanel()
        CreateMediaPlayer(_CurrentMediaPanelName, "D:\CFP\bumpers\categorical\opening 2010.wav")
        CreatePlayButton(_CurrentMediaPanelName)
        CreateDeleteButton(_CurrentMediaPanelName)
    End Sub


    '*********************Load Settings***********************************
    'Todo

    '*********************Create cues from settings***********************************

    'Add media panel to flow layout panel
    Public Sub CreateMediaPanel()

        Dim mediaPanel As Panel
        mediaPanel = New Panel()

        'Set panel properties
        With mediaPanel
            .BackColor = Color.White
            .Size = New Size(256, 144)
            .Name = "pnlMedia" + (_MediaPanelsAddedCount + 1).ToString
        End With

        'Add panel to flow layout panel
        flpMain.Controls.Add(mediaPanel)

        'Update panel variables
        _CurrentMediaPanelName = mediaPanel.Name
        _MediaPanelsAddedCount += 1


    End Sub

    'I don't know what goes here
    Private Sub flpMain_Paint(sender As Object, e As PaintEventArgs) Handles flpMain.Paint

    End Sub

    'Add player to the flow panel
    Public Sub CreateMediaPlayer(ByVal panelName As String, ByVal mediaPath As String)

        Dim mediaPlayer As AxWMPLib.AxWindowsMediaPlayer
        mediaPlayer = New AxWMPLib.AxWindowsMediaPlayer
        mediaPlayer.CreateControl()

        'Set properties
        With mediaPlayer
            .URL = mediaPath
            .settings.autoStart = False
            .Name = "mediaPlayer" + _MediaPanelsAddedCount.ToString
            '.uiMode = "mini"
            '.Size = New Size(130, 40)
            .Visible = False
            '.Dock = DockStyle.Bottom
            .Ctlcontrols.stop()
        End With

        'Loop through controls and add new label to passed panel
        For Each ControlObject As Control In flpMain.Controls
            If ControlObject.Name = panelName Then
                ControlObject.Controls.Add(mediaPlayer)
            End If
        Next
    End Sub

    'Add play button to the flow panel
    Public Sub CreatePlayButton(ByVal panelName As String)

        Dim playButton As Button
        playButton = New Button

        'set properties
        With playButton
            .AutoSize = True
            .Location = New Point(60, 60)
            .Name = "btnPlay" + _MediaPanelsAddedCount.ToString
            .Text = "Play"
        End With

        For Each ControlObject As Control In flpMain.Controls
            If ControlObject.Name = panelName Then
                ControlObject.Controls.Add(playButton)
            End If
        Next

        'Add handler for click events
        'AddHandler playButton.Click AddressOf DynamicPlayButton_Click

    End Sub

    'Add delete button to the flow panel
    Public Sub CreateDeleteButton(ByVal panelName As String)

        Dim deleteButton As Button
        deleteButton = New Button

        'set properties
        With deleteButton
            .AutoSize = True
            .Location = New Point(120, 60)
            .Name = "btnDelete" + _MediaPanelsAddedCount.ToString
            .Text = "Delete"
        End With

        For Each ControlObject As Control In flpMain.Controls
            If ControlObject.Name = panelName Then
                ControlObject.Controls.Add(deleteButton)
            End If
        Next

        'Add handler for click events
        AddHandler deleteButton.Click, AddressOf DynamicDeleteButton_Click

    End Sub

    '*********************Dynamic Event Handlers***********************************

    'Handle play button click requests
    Public Sub DynamicPlayButton_Click(ByVal sender As Object, ByVal e As EventArgs)

        Dim parentPanelName As String
        parentPanelName = Nothing


    End Sub

    'Handle play button click requests
    Public Sub DynamicDeleteButton_Click(ByVal sender As Object, ByVal e As EventArgs)

        Dim parentPanelName As String
        parentPanelName = Nothing

        'Remove handler from sender
        For Each controlObj As Control In flpMain.Controls
            For Each childControlObj As Control In controlObj.Controls
                If childControlObj.Name = sender.name Then
                    RemoveHandler childControlObj.Click, AddressOf DynamicDeleteButton_Click
                    parentPanelName = childControlObj.Parent.Name
                End If
            Next
        Next

        'Remove contact panel
        For Each controlObj As Control In flpMain.Controls
            If controlObj.Name = parentPanelName Then
                flpMain.Controls.Remove(controlObj)
                controlObj.Dispose()
            End If
        Next

    End Sub

End Class

'Application settings wrapper class. This class defines the settings we intend to use in our application.
NotInheritable Class FormSettings
    Inherits ApplicationSettingsBase

    <UserScopedSettingAttribute()>
    Public Property FormText() As String
        Get
            Return CStr(Me("FormText"))
        End Get
        Set(ByVal value As String)
            Me("FormText") = value
        End Set
    End Property

    <UserScopedSettingAttribute(), DefaultSettingValueAttribute("0, 0")>
    Public Property FormLocation() As Point
        Get
            Return CType(Me("FormLocation"), Point)
        End Get
        Set(ByVal value As Point)
            Me("FormLocation") = value
        End Set
    End Property

    <UserScopedSettingAttribute(), DefaultSettingValueAttribute("225, 200")>
    Public Property FormSize() As Size
        Get
            Return CType(Me("FormSize"), Size)
        End Get
        Set(ByVal value As Size)
            Me("FormSize") = value
        End Set
    End Property

    <UserScopedSettingAttribute(), DefaultSettingValueAttribute("LightGray")>
    Public Property FormBackColor() As Color
        Get
            Return CType(Me("FormBackColor"), Color)
        End Get
        Set(ByVal value As Color)
            Me("FormBackColor") = value
        End Set
    End Property
End Class
