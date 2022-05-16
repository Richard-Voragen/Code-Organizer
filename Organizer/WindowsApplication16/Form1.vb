Imports Newtonsoft.Json.Linq

Public Class Form1
    'Public Property location As String = String.Format("{0}\Data\", Environment.CurrentDirectory).Substring(0, String.Format("{0}\Data\", Environment.CurrentDirectory).LastIndexOf("Organizer"))
    Public Property location As String = ""
    Dim jsonString As String = My.Computer.FileSystem.ReadAllText(location + "CodingOrganizer.json") 'use location + "CodingOrganizer.json" for 2 directories earlier
    Public Property json As JObject = JObject.Parse(jsonString)
    Dim tabbedIn As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PrintCategories()
    End Sub

    Public Function GetNthIndex(searchString As String, subStringToFind As String, n As Integer) As Integer
        Dim idx As Integer = searchString.IndexOf(subStringToFind)
        Dim count As Integer = 1
        While idx > -1 AndAlso count < n
            idx = searchString.IndexOf(subStringToFind, idx + 1)
            count += 1
        End While
        Return idx
    End Function

    Private Function GiveIndex(str As String, n As Integer) As Integer
        Try
            Return Integer.Parse(json(TabControl1.SelectedTab.Text)("Categories").ToString.Substring(GetNthIndex(json(TabControl1.SelectedTab.Text)("Categories").ToString, str + ",", n) - 1, 1))
        Catch ex As Exception
            Return GiveIndex(str, n + 1)
        End Try
    End Function

    Public Function FindIndexString(arr As String(), target As String) As Integer
        For i As Integer = 0 To arr.Length - 1
            If arr(i).Equals(target) Then
                Return i
            End If
        Next
        Return -1
    End Function

    Private Sub MoveUpDown(selected As Integer, tabbed As Boolean, move As Integer)
        Dim cursor = ListBox1.SelectedIndex
        If (tabbed = False And ((move = 1 And cursor > 0) Or (move = -1 And cursor < ListBox1.Items.Count - 1))) Then
            Dim output, input As String
            Dim first, second As Integer
            If (selected = -1) Then
                input = "Categories"
            Else
                input = "SubCategories"
            End If

            Dim catList As String() = json(My.Forms.Form1.TabControl1.SelectedTab.Text)(input).ToString.TrimEnd(","c).Split(",")
            If (selected = -1) Then
                first = FindIndexString(catList, GiveIndex(ListBox1.SelectedItem.ToString, 1).ToString + ListBox1.SelectedItem.ToString)
                second = first - move
            Else
                first = FindIndexString(catList, selected.ToString + ListBox1.SelectedItem.ToString)
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
            json(TabControl1.SelectedTab.Text)(input) = output
            IO.File.WriteAllText(My.Forms.Form1.location + "CodingOrganizer.json", json.ToString)
            If selected = -1 Then
                PrintCategories()
            Else
                PrintSubCategories()
            End If

            ListBox1.SelectedIndex = cursor
        End If
    End Sub

    Public Function LoadTags(cat As Integer, input As String) As List(Of String)
        Dim output As String() = Nothing
        output = input.TrimEnd(","c).Split(",")
        Dim ListOut As List(Of String) = New List(Of String)
        For Each str As String In output
            If (str.Substring(0, 1).CompareTo(cat.ToString) = 0) Then
                ListOut.Add(str.Substring(1))
            End If
        Next
        Return ListOut
    End Function

    Dim selected As Integer = -1

    Public Sub PrintCategories()
        ListBox1.Items.Clear()
        For Each cat As String In json(TabControl1.SelectedTab.Text)("Categories").ToString.TrimEnd(","c).Split(",")
            ListBox1.Items.Add(cat.Substring(1))
        Next
        selected = -1
        tabbedIn = False
    End Sub

    Public Sub PrintSubCategories()
        ListBox1.Items.Clear()
        Dim tags As List(Of String) = LoadTags(selected, json(TabControl1.SelectedTab.Text)("SubCategories"))
        For Each i As String In tags
            If ListBox1.Items.Contains(i) = False Then
                ListBox1.Items.Add(i)
            End If
        Next
    End Sub

    Dim output As List(Of String) = New List(Of String)

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        If selected = -1 Then
            selected = GiveIndex(ListBox1.SelectedItem, 1)
            PrintSubCategories()
        ElseIf selected <> -1 And tabbedIn = False Then
            Dim jsonChildren As List(Of JToken) = json.SelectToken(TabControl1.SelectedTab.Text).Children().ToList
            output.Clear()
            For Each j As JProperty In jsonChildren
                j.CreateReader()
                If j.Name.Contains("Categories") = False Then
                    Dim tags As List(Of String) = LoadTags(selected, json(TabControl1.SelectedTab.Text)(j.Name)("Tags"))
                    For Each i As String In tags
                        If i.Equals(ListBox1.SelectedItem) = True Then
                            If json(TabControl1.SelectedTab.Text)(j.Name)("Location").ToString.Contains(":"c) = True Then
                                output.Add(json(TabControl1.SelectedTab.Text)(j.Name)("Location").ToString)
                            Else
                                output.Add(My.Forms.Form1.location + json(TabControl1.SelectedTab.Text)(j.Name)("Location").ToString)
                            End If
                        End If
                    Next
                End If
            Next
            ListBox1.Items.Clear()
            For Each s As String In output
                ListBox1.Items.Add(s.Substring(s.LastIndexOf("\"c) + 1))
            Next
            tabbedIn = True
        ElseIf tabbedIn = True Then
            Process.Start(output.Item(ListBox1.SelectedIndex))
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BackButton.Click
        If tabbedIn = True Then
            tabbedIn = False
            PrintSubCategories()
        ElseIf selected <> -1 Then
            PrintCategories()
        End If
    End Sub

    Private Sub AddButton_Click(sender As Object, e As EventArgs) Handles AddButton.Click
        My.Forms.AddFiles.json = json
        My.Forms.AddFiles.Show()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        PrintCategories()
    End Sub

    Dim shift As Boolean = False

    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        If e.KeyCode = Keys.ShiftKey Then
            shift = True
        ElseIf shift And e.KeyCode = Keys.Up Then
            MoveUpDown(selected, tabbedIn, 1)
        ElseIf shift And e.KeyCode = Keys.Down Then
            MoveUpDown(selected, tabbedIn, -1)
        End If
    End Sub

    Private Sub ListBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyUp
        If e.KeyCode = Keys.ShiftKey Then
            shift = False
        End If
    End Sub
End Class
