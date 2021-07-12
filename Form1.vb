Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    'Indicates current media panel to add controls to
    Private _CurrentMediaPanelName As String = Nothing

    'Used to give unique control names such as label1, label2, etc.
    Private _MediaPanelsAddedCount As Integer = 0

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

    Private Sub flpMain_Paint(sender As Object, e As PaintEventArgs) Handles flpMain.Paint

    End Sub
    'Add New media player to media panel
    Public Sub CreateMediaPlayer(ByVal panelName As String, ByVal mediaPath As String)

        Dim mediaPlayer As AxWMPLib.AxWindowsMediaPlayer
        mediaPlayer = New AxWMPLib.AxWindowsMediaPlayer
        mediaPlayer.CreateControl()

        'Set properties
        With mediaPlayer
            .URL = mediaPath
            .settings.autoStart = False
            .Name = "mediaPlayer" + (_MediaPanelsAddedCount).ToString
            '.uiMode = mini
            '.Size = (130, 40)
            .Dock = DockStyle.Bottom
        End With

        'Loop through controls and add new label to passed panel
        For Each ControlObject As Control In flpMain.Controls
            If ControlObject.Name = panelName Then
                ControlObject.Controls.Add(mediaPlayer)
            End If
        Next
    End Sub


    'Add new play button to panel
    Public Sub CreateMediaPlayButton(ByVal panelName As String)

    End Sub

    Private Sub CreateAudioCueToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateAudioCueToolStripMenuItem.Click

        'Dim x As Integer = 0
        'Do
        '    CreateMediaPanel()
        '    x += 1
        'Loop Until x > 50

        CreateMediaPanel()

        CreateMediaPlayer(_CurrentMediaPanelName, "D:\CFP\bumpers\categorical\opening 2010.wav")


    End Sub
End Class
