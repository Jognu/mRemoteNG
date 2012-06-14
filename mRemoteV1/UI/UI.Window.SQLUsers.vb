Imports WeifenLuo.WinFormsUI.Docking
Imports mRemoteNG.App.Runtime
Imports System.Data.SqlClient

Namespace UI
    Namespace Window
        Public Class SQLUsers
            Inherits UI.Window.Base

#Region "Form Init"
            Friend WithEvents clmSesUser As System.Windows.Forms.ColumnHeader
            Friend WithEvents cMenSession As System.Windows.Forms.ContextMenuStrip
            Private components As System.ComponentModel.IContainer
            Friend WithEvents cMenUsersRefresh As System.Windows.Forms.ToolStripMenuItem
            Friend WithEvents TimerSQLUsers As System.Windows.Forms.Timer
            Friend WithEvents lvUsers As System.Windows.Forms.ListView

            Private Sub InitializeComponent()
                Me.components = New System.ComponentModel.Container()
                Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SQLUsers))
                Me.lvUsers = New System.Windows.Forms.ListView()
                Me.clmSesUser = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
                Me.cMenSession = New System.Windows.Forms.ContextMenuStrip(Me.components)
                Me.cMenUsersRefresh = New System.Windows.Forms.ToolStripMenuItem()
                Me.TimerSQLUsers = New System.Windows.Forms.Timer(Me.components)
                Me.cMenSession.SuspendLayout()
                Me.SuspendLayout()
                '
                'lvUsers
                '
                Me.lvUsers.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                Me.lvUsers.BorderStyle = System.Windows.Forms.BorderStyle.None
                Me.lvUsers.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.clmSesUser})
                Me.lvUsers.ContextMenuStrip = Me.cMenSession
                Me.lvUsers.FullRowSelect = True
                Me.lvUsers.GridLines = True
                Me.lvUsers.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
                Me.lvUsers.Location = New System.Drawing.Point(0, -1)
                Me.lvUsers.MultiSelect = False
                Me.lvUsers.Name = "lvUsers"
                Me.lvUsers.ShowGroups = False
                Me.lvUsers.Size = New System.Drawing.Size(242, 174)
                Me.lvUsers.TabIndex = 0
                Me.lvUsers.UseCompatibleStateImageBehavior = False
                Me.lvUsers.View = System.Windows.Forms.View.Details
                '
                'clmSesUser
                '
                Me.clmSesUser.Text = "User"
                Me.clmSesUser.Width = 80
                '
                'cMenSession
                '
                Me.cMenSession.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cMenUsersRefresh})
                Me.cMenSession.Name = "cMenSession"
                Me.cMenSession.Size = New System.Drawing.Size(153, 48)
                '
                'cMenUsersRefresh
                '
                Me.cMenUsersRefresh.Image = Global.mRemoteNG.My.Resources.Resources.Refresh
                Me.cMenUsersRefresh.Name = "cMenUsersRefresh"
                Me.cMenUsersRefresh.Size = New System.Drawing.Size(152, 22)
                Me.cMenUsersRefresh.Text = Global.mRemoteNG.My.Language.strRefresh
                '
                'TimerSQLUsers
                '
                Me.TimerSQLUsers.Interval = 60000
                '
                'SQLUsers
                '
                Me.ClientSize = New System.Drawing.Size(242, 173)
                Me.Controls.Add(Me.lvUsers)
                Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
                Me.HideOnClose = True
                Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
                Me.Name = "SQLUsers"
                Me.TabText = "SQL Users"
                Me.Text = "SQL Users"
                Me.cMenSession.ResumeLayout(False)
                Me.ResumeLayout(False)

            End Sub
#End Region

#Region "Form Stuff"
            Private Sub SQLUsers_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
                ApplyLanguage()

                If My.Settings.UseSQLServer Then
                    TimerSQLUsers.Enabled = True
                    GetUsers()
                End If
            End Sub

            Private Sub ApplyLanguage()
                clmSesUser.Text = My.Language.strUser
                cMenUsersRefresh.Text = My.Language.strRefresh
                TabText = My.Language.strMenuSQLUsers
                Text = My.Language.strMenuSQLUsers
            End Sub
#End Region

#Region "Private Methods"
            Private Sub GetUsers()
                Try
                    If My.Settings.UseSQLServer Then
                        ClearList()

                        Dim sqlCon As SqlConnection
                        Dim sqlQuery As SqlCommand
                        Dim sqlRd As SqlDataReader

                        If My.Settings.SQLUser <> "" Then
                            sqlCon = New SqlConnection("Data Source=" & My.Settings.SQLHost & ";Initial Catalog=" & My.Settings.SQLDatabaseName & ";User Id=" & My.Settings.SQLUser & ";Password=" & Security.Crypt.Decrypt(My.Settings.SQLPass, App.Info.General.EncryptionKey))
                        Else
                            sqlCon = New SqlConnection("Data Source=" & My.Settings.SQLHost & ";Initial Catalog=" & My.Settings.SQLDatabaseName & ";Integrated Security=True")
                        End If

                        sqlCon.Open()

                        sqlQuery = New SqlCommand("SELECT * FROM tblUsers", sqlCon)
                        sqlRd = sqlQuery.ExecuteReader()

                        While sqlRd.Read
                            Dim tsTimeSpan As TimeSpan
                            tsTimeSpan = Now.Subtract(sqlRd.Item("LastLoad"))

                            If (tsTimeSpan.Minutes <= 1) Then
                                Dim lItem As New ListViewItem
                                lItem.Text = sqlRd.Item("Name")

                                AddToList(lItem)
                            End If
                        End While
                    End If
                Catch ex As Exception
                    MessageCollector.AddMessage(Messages.MessageClass.WarningMsg, "Import SQL Users failed" & vbNewLine & ex.Message, True)
                End Try
            End Sub


            Delegate Sub AddToListCB(ByVal [ListItem] As ListViewItem)
            Private Sub AddToList(ByVal [ListItem] As ListViewItem)
                If Me.lvUsers.InvokeRequired Then
                    Dim d As New AddToListCB(AddressOf AddToList)
                    Me.lvUsers.Invoke(d, New Object() {[ListItem]})
                Else
                    Me.lvUsers.Items.Add(ListItem)
                End If
            End Sub

            Delegate Sub ClearListCB()
            Private Sub ClearList()
                If Me.lvUsers.InvokeRequired Then
                    Dim d As New ClearListCB(AddressOf ClearList)
                    Me.lvUsers.Invoke(d)
                Else
                    Me.lvUsers.Items.Clear()
                End If
            End Sub
#End Region

#Region "Public Methods"
            Public Sub New(ByVal Panel As DockContent)
                Me.WindowType = Type.SQLUsers
                Me.DockPnl = Panel
                Me.InitializeComponent()
            End Sub
#End Region

            Private Sub cMenUsersRefresh_Click(sender As System.Object, e As System.EventArgs) Handles cMenUsersRefresh.Click, TimerSQLUsers.Tick
                Me.GetUsers()
            End Sub
        End Class
    End Namespace
End Namespace