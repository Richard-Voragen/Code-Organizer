Imports Newtonsoft.Json.Linq

Public Class AddFiles
    Public Property json As JObject


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Function GiveIndex(str As String, n As Integer) As Integer
        Try
            Return Integer.Parse(json(My.Forms.Form1.TabControl1.SelectedTab.Text)("Categories").ToString.Substring(My.Forms.Form1.GetNthIndex(json(My.Forms.Form1.TabControl1.SelectedTab.Text)("Categories").ToString, str + ",", n) - 1, 1))
        Catch ex As Exception
            Return GiveIndex(str, n + 1)
        End Try
    End Function

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Label1.Text = OpenFileDialog1.FileName.Substring(OpenFileDialog1.FileName.LastIndexOf("\"c) + 1)
        ListBox2.Items.Clear()
        Dim jsonChildren As List(Of JToken) = json.SelectToken(My.Forms.Form1.TabControl1.SelectedTab.Text).Children().ToList
        For Each j As JProperty In jsonChildren
            j.CreateReader()
            If j.Name.Equals(OpenFileDialog1.FileName.Substring(OpenFileDialog1.FileName.LastIndexOf("\"c) + 1)) = True Then
                Dim tags As String() = json(My.Forms.Form1.TabControl1.SelectedTab.Text)(j.Name)("Tags").ToString.Split(",")
                For Each i As String In tags
                    ListBox2.Items.Add(i)
                Next
            End If
        Next
        Process.Start(OpenFileDialog1.FileName)
    End Sub

    Private Sub AddFiles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PrintCategories()
    End Sub

    Dim selected As Integer = -1

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If (selected = -1) Then
            selected = GiveIndex(ListBox1.SelectedItem, 1)
            BackButton.Visible = True
            PrintSubCategories()
        ElseIf (selected <> -1) Then
            If ListBox2.Items.Contains(selected.ToString + ListBox1.SelectedItem) = False Then
                ListBox2.Items.Add(selected.ToString + ListBox1.SelectedItem)
            End If
        End If
    End Sub

    Private Sub BackButton_Click(sender As Object, e As EventArgs) Handles BackButton.Click
        BackButton.Visible = False
        PrintCategories()
    End Sub

    Private Sub DoneButton_Click(sender As Object, e As EventArgs) Handles DoneButton.Click
        If (ListBox2.Items.Count <> 0 And OpenFileDialog1.FileName <> "OpenFileDialog1") Then
            Dim tags As String = ""
            While ListBox2.Items.Count > 0
                tags += ListBox2.Items.Item(0) + ","
                ListBox2.Items.RemoveAt(0)
            End While
            Dim jMod As JObject = TryCast(json(My.Forms.Form1.TabControl1.SelectedTab.Text), JObject)
            Dim jAdd As JObject = New JObject(
                New JProperty(OpenFileDialog1.FileName.Substring(OpenFileDialog1.FileName.LastIndexOf("\"c) + 1), New JObject(
                              New JProperty("Tags", tags.TrimEnd(CChar(","))),
                              New JProperty("Location", OpenFileDialog1.FileName.Replace(My.Forms.Form1.location, "")))))
            Dim jsonString As String = jMod.ToString.TrimEnd(CChar("}")) + "," + jAdd.ToString.TrimStart(CChar("{"))
            json(My.Forms.Form1.TabControl1.SelectedTab.Text) = JObject.Parse(jsonString)
            My.Forms.Form1.json = json
            My.Forms.Form1.PrintCategories()

            System.IO.File.WriteAllText(My.Forms.Form1.location + "CodingOrganizer.json", json.ToString)
            BackButton.PerformClick()
            Label1.Text = ""
            ListBox2.Items.Clear()
            OpenFileDialog1.FileName = "OpenFileDialog1"
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            If (selected = -1 And TextBox1.Text <> "") Then
                Dim num() As String = json(My.Forms.Form1.TabControl1.SelectedTab.Text)("Categories").ToString.Split(",")
                json(My.Forms.Form1.TabControl1.SelectedTab.Text)("Categories") = json(My.Forms.Form1.TabControl1.SelectedTab.Text)("Categories").ToString + num.Length.ToString + TextBox1.Text + ","
                ListBox1.Items.Add(TextBox1.Text)
                TextBox1.Clear()
            ElseIf (TextBox1.Text <> "" And selected <> -1) Then
                json(My.Forms.Form1.TabControl1.SelectedTab.Text)("SubCategories") = json(My.Forms.Form1.TabControl1.SelectedTab.Text)("SubCategories").ToString + selected.ToString + TextBox1.Text + ","
                ListBox1.Items.Add(TextBox1.Text)
                TextBox1.Clear()
            End If
        End If
    End Sub

    Private Sub ListBox2_DoubleClick(sender As Object, e As EventArgs) Handles ListBox2.DoubleClick
        ListBox2.Items.Remove(ListBox2.SelectedItem)
    End Sub

    Public Sub PrintCategories()
        ListBox1.Items.Clear()
        For Each cat As String In json(My.Forms.Form1.TabControl1.SelectedTab.Text)("Categories").ToString.TrimEnd(","c).Split(",")
            ListBox1.Items.Add(cat.Substring(1))
        Next
        selected = -1
    End Sub

    Private Sub PrintSubCategories()
        ListBox1.Items.Clear()
        Dim tags As List(Of String) = My.Forms.Form1.LoadTags(selected, json(My.Forms.Form1.TabControl1.SelectedTab.Text)("SubCategories"))
        For Each i As String In tags
            If ListBox1.Items.Contains(i) = False Then
                ListBox1.Items.Add(i)
            End If
        Next
    End Sub

    Private Sub MoveUpDown(selected As Integer, move As Integer)
        Dim cursor = ListBox1.SelectedIndex
        If (((move = 1 And cursor > 0) Or (move = -1 And cursor < ListBox1.Items.Count - 1))) Then
            Dim output, input As String
            Dim first, second As Integer
            If (selected = -1) Then
                input = "Categories"
            Else
                input = "SubCategories"
            End If

            Dim catList As String() = json(My.Forms.Form1.TabControl1.SelectedTab.Text)(input).ToString.TrimEnd(","c).Split(",")
            If (selected = -1) Then
                first = My.Forms.Form1.FindIndexString(catList, GiveIndex(ListBox1.SelectedItem.ToString, 1).ToString + ListBox1.SelectedItem.ToString)
                second = first - move
            Else
                first = My.Forms.Form1.FindIndexString(catList, selected.ToString + ListBox1.SelectedItem.ToString)
                second = first - move
                While catList(second).Substring(0, 1) <> selected.ToString
                    second -= move
                End While
            End If
            Dim temp As String = catList(first)
            catList(first) = catList(second)
            catList(second) = temp

            output = ""
            For Each str As String In catList
                output += str + ","
            Next
            json(My.Forms.Form1.TabControl1.SelectedTab.Text)(input) = output
            IO.File.WriteAllText(My.Forms.Form1.location + "CodingOrganizer.json", json.ToString)
            If selected = -1 Then
                PrintCategories()
            Else
                PrintSubCategories()
            End If

            ListBox1.SelectedIndex = cursor
        End If
    End Sub

    Dim shift As Boolean = False

    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        If e.KeyCode = Keys.ShiftKey Then
            shift = True
        ElseIf shift And e.KeyCode = Keys.Up Then
            MoveUpDown(selected, 1)
        ElseIf shift And e.KeyCode = Keys.Down Then
            MoveUpDown(selected, -1)
        End If
    End Sub

    Private Sub ListBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyUp
        If e.KeyCode = Keys.ShiftKey Then
            shift = False
        End If
    End Sub
End Class