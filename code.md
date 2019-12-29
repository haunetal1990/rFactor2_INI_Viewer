Was brauchen wir?
- 1 Listview mit 3 "Columns" (unter "View" auf "details" stellen.
- 1 Textbox (dort steht der Pfad drin; Muss aber auch nicht)

Was macht der Code?
Mit dem Code kann man die .mft-Dateien von den installierten Mods anzeigen. Es handelt sich hierbei einfach um .INI-Dateien. Zudem kann man die Spalten nach Namen sortieren lassen.

Imports System
Imports System.Xml
Public Class Form1

   Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32</br>
    
    Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32</br>
  
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ListView1.Items.Clear()
        If (TextBox1.Text.Length > 0) Then
            Dim sfile As String = Nothing
            Dim sPath As String = TextBox1.Text
            For Each sfile In My.Computer.FileSystem.GetFiles(sPath, FileIO.SearchOption.SearchAllSubDirectories, "*.mft")
                With ListView1.Items.Add(INI_ReadValueFromFile("COMPONENT", "Name", ListView1.Text, sfile))
                    .SubItems.Add(INI_ReadValueFromFile("COMPONENT", "Type", ListView1.Text, sfile))
                    .SubItems.Add(INI_ReadValueFromFile("COMPONENT", "Version", ListView1.Text, sfile))
                End With
            Next
        Else
            TextBox1.Text = Application.StartupPath()
            Dim sfile As String = Nothing
            Dim sPath As String = TextBox1.Text
            For Each sfile In My.Computer.FileSystem.GetFiles(sPath, FileIO.SearchOption.SearchAllSubDirectories, "*.mft")
                With ListView1.Items.Add(INI_ReadValueFromFile("COMPONENT", "Name", ListView1.Text, sfile))
                    .SubItems.Add(INI_ReadValueFromFile("COMPONENT", "Type", ListView1.Text, sfile))
                    .SubItems.Add(INI_ReadValueFromFile("COMPONENT", "Version", ListView1.Text, sfile))
                End With
            Next
        End If
    End Sub

    Public Class ListViewItemComparer
        Implements IComparer
        Private col As Integer
        Private order As SortOrder

        Public Sub New()
            col = 0
            order = SortOrder.Ascending
        End Sub

        Public Sub New(column As Integer, order As SortOrder)
            col = column
            Me.order = order
        End Sub

        Public Function Compare(x As Object, y As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim returnVal As Integer
            Try

                ' Attempt to parse the two objects as DateTime
                Dim firstDate As System.DateTime = DateTime.Parse(CType(x, ListViewItem).SubItems(col).Text)
                Dim secondDate As System.DateTime = DateTime.Parse(CType(y, ListViewItem).SubItems(col).Text)

                ' Compare as date
                returnVal = DateTime.Compare(firstDate, secondDate)

            Catch ex As Exception

                ' If date parse failed then fall here to determine if objects are numeric
                If IsNumeric(CType(x, ListViewItem).SubItems(col).Text) And
                IsNumeric(CType(y, ListViewItem).SubItems(col).Text) Then

                    ' Compare as numeric
                    returnVal = Val(CType(x, ListViewItem).SubItems(col).Text).CompareTo(
                  Val(CType(y, ListViewItem).SubItems(col).Text))

                Else
                    ' If not numeric then compare as string
                    returnVal = [String].Compare(CType(x, ListViewItem).SubItems(col).Text, CType(y, ListViewItem).SubItems(col).Text)
                End If
            End Try

            ' If order is descending then invert value
            If order = SortOrder.Descending Then
                returnVal *= -1
            End If

            Return returnVal

        End Function

    End Class

    ' Value to track which column was previously sorted
    Dim sortColumn As Integer = -1

    Private Sub Listview1_ColumnClick(sender As Object, e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick

        ' If current column is not the previously clicked column
        ' Add
        If Not e.Column = sortColumn Then
            ' Set the sort column to the new column
            sortColumn = e.Column
            'Default to ascending sort order
            ListView1.Sorting = SortOrder.Ascending
        Else
            'Flip the sort order
            If ListView1.Sorting = SortOrder.Ascending Then
                ListView1.Sorting = SortOrder.Descending
            Else
                ListView1.Sorting = SortOrder.Ascending
            End If
        End If
        'Set the ListviewItemSorter property to a new ListviewItemComparer object
        Me.ListView1.ListViewItemSorter = New ListViewItemComparer(e.Column, ListView1.Sorting)
        ' Call the sort method to manually sort
        ListView1.Sort()

    End Sub
    Private Function INI_ReadValueFromFile(ByVal strSection As String, ByVal strKey As String, ByVal strDefault As String, ByVal strFile As String) As String
        'Funktion zum Lesen
        'strSection = Sektion in der INI-Datei
        'strKey = Name des Schlüssels
        'strDefault = Standardwert, wird zurückgegeben, wenn der Wert in der INI-Datei nicht gefunden wurde
        'strFile = Vollständiger Pfad zur INI-Datei
        Dim strTemp As String = Space(1024), lLength As Integer
        lLength = GetPrivateProfileString(strSection, strKey, strDefault, strTemp, strTemp.Length, strFile)
        Return (strTemp.Substring(0, lLength))
    End Function

    Public Function INI_WriteValueToFile(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String, ByVal strFile As String) As Boolean
        'Funktion zum Schreiben
        'strSection = Sektion in der INI-Datei
        'strKey = Name des Schlüssels
        'strValue = Wert, der geschrieben werden soll
        'strFile = Vollständiger Pfad zur INI-Datei
        Return (Not (WritePrivateProfileString(strSection, strKey, strValue, strFile) = 0))
    End Function

End Class
