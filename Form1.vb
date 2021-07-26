Option Explicit On
'Option Strict On

Imports System.Configuration
Imports System.ComponentModel
Imports System.IO
Imports PlaylistsNET.Content
Imports PlaylistsNET.Models

Public Class Form1

    Private Sub Create_Cue()

        'CreateMediaPanel()
        ''CreateMediaPlayer(_CurrentMediaPanelName, "D:\CFP\bumpers\categorical\opening 2010.wav")

        'CreatePlayButton(_currentMediaPanelName)
        'CreateDeleteButton(_currentMediaPanelName)

    End Sub


    'content and playlist are objects that hold all of the playlist data
    Private _content As New WplContent
    Private _playlist As New WplPlaylist

    'paths stores all of the file paths to play
    Private _paths As New List(Of String)

    'Indicates current media panel to add controls to
    Private _currentMediaPanelName As String = Nothing

    'Indicates current media player to add controls to
    Private _currentMediaPlayerName As String = Nothing

    'Used to give unique control names such as label1, label2, etc.
    Private _mediaPanelsAddedCount As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'stream is used to access the file
        Dim stream As New FileStream("PlayList.wpl", FileMode.Open, FileAccess.Read)
        'Dim myStreamWriter As New StreamWriter(stream)

        'loads the wpl content into the playlist
        _playlist = _content.GetFromStream(stream)

        Console.WriteLine(_playlist.ItemCount())

        'loads the paths from playlist
        _paths = _playlist.GetTracksPaths()

        Dim x As Integer = 0
        Do
            CreateMediaPanel()
            CreateMediaPlayer(_currentMediaPanelName, _paths.Item(x))
            CreatePlayButton(_currentMediaPanelName, _currentMediaPlayerName)
            CreateDeleteButton(_currentMediaPanelName)
            x += 1
        Loop Until x > _paths.Count - 1

        stream.Close()
        Me.Show()



    End Sub

    'read the wpl file



    '*********************Form Setup***********************************
    'Add menu to the form
    Private Sub CreateAudioCueToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateAudioCueToolStripMenuItem.Click

        'Dim x As Integer = 0
        'Do
        '    CreateMediaPanel()
        '    x += 1
        'Loop Until x > 50

        CreateMediaPanel()
        CreateMediaPlayer(_currentMediaPanelName, "D:\CFP\bumpers\categorical\opening 2010.wav")
        CreatePlayButton(_currentMediaPanelName, _currentMediaPlayerName)
        CreateDeleteButton(_currentMediaPanelName)
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
            .Name = "pnlMedia" + (_mediaPanelsAddedCount + 1).ToString
        End With

        'Add panel to flow layout panel
        flpMain.Controls.Add(mediaPanel)

        'Update panel variables
        _currentMediaPanelName = mediaPanel.Name
        _mediaPanelsAddedCount += 1


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
            .Name = "mediaPlayer" + _mediaPanelsAddedCount.ToString
            .uiMode = "full"
            .Size = New Size(230, 240)
            .Visible = True
            '.Dock = DockStyle.Bottom
            .Ctlcontrols.stop()
        End With

        'Loop through controls And add New label to passed panel
        For Each ControlObject As Control In flpMain.Controls
            If ControlObject.Name = panelName Then
                ControlObject.Controls.Add(mediaPlayer)
            End If
        Next

        'Update mediaplayer variable
        _currentMediaPlayerName = mediaPlayer.Name

    End Sub

    'Add play button to the flow panel
    Public Sub CreatePlayButton(ByVal panelName As String, ByVal playerName As String)

        Dim playButton As Button
        playButton = New Button

        'set properties
        With playButton
            .AutoSize = True
            .Location = New Point(60, 60)
            .Name = "btnPlay" + _mediaPanelsAddedCount.ToString
            .Text = "Play"
        End With

        For Each ControlObject As Control In flpMain.Controls
            If ControlObject.Name = panelName Then
                ControlObject.Controls.Add(playButton)
            End If
        Next

        'Add handler for click events
        AddHandler playButton.Click, AddressOf DynamicPlayButton_Click

    End Sub

    'Add delete button to the flow panel
    Public Sub CreateDeleteButton(ByVal panelName As String)

        Dim deleteButton As Button
        deleteButton = New Button

        'set properties
        With deleteButton
            .AutoSize = True
            .Location = New Point(120, 60)
            .Name = "btnDelete" + _mediaPanelsAddedCount.ToString
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

        'Identify the parentPanel where the button was clicked
        For Each controlObj As Control In flpMain.Controls
            For Each childControlObj As Control In controlObj.Controls
                If childControlObj.Name = sender.name Then
                    parentPanelName = childControlObj.Parent.Name
                End If
            Next
        Next

        'Play the associated player file

    End Sub

    'Handle delete button click requests
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

    Private Sub ReadPlaylistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReadPlaylistToolStripMenuItem.Click

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
