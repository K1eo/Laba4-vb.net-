Imports System.Data.OleDb
Public Class Access

    ' структура для зчитування всіх стовпчиків Access бази даних
    Public Structure straight
        Dim index As Integer
        Dim coefficients, Name, color_straight As String
    End Structure

    Dim names As New List(Of straight)


    'зчитування з Access таблиці
    Public Sub Input()
        Dim ins As Integer
        ins = 1
        names.RemoveRange(0, names.Count)
        Dim n As New List(Of straight)
        Dim r As New straight

        outFille()
        Dim coefficients As New OleDbCommand("select i,coefficients_ABC,Name,color_straight from Пряма_у_просторі", conn)
        Dim dr_coefficient As OleDbDataReader = coefficients.ExecuteReader
        While (dr_coefficient.Read)
            Try
                r.index = ins
                r.coefficients = dr_coefficient.Item("coefficients_ABC")
                r.Name = dr_coefficient.Item("Name")
                r.color_straight = dr_coefficient.Item("color_straight")
                ins += 1
                names.Add(r)
            Catch ex As Exception
            End Try
        End While

    End Sub

    ' файл для відкривання Access і занесення бази даних в conn
    Public Sub outFille()
        conn = New OleDbConnection
        Dim c As New OleDbCommand
        Try
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Валентин\Documents\BazaStraight.accdb;Persist Security Info=False;" 'задаємо параметри і адрес до бази даних
            conn.Open()                 'зчитуємо з бази даних
        Catch ex As Exception
            Console.WriteLine("База даних не знайдена")
        End Try
    End Sub
    ' вивід Access бази даних 
    Public Sub printAcessTable()

        Dim index As Integer
        index = 1
        Console.WriteLine("TABLE\n")
        ' outFille()
        Dim coefficients As New OleDbCommand("select i,coefficients_ABC,Name,color_straight from Пряма_у_просторі", conn)
        Dim dr_coefficient As OleDbDataReader = coefficients.ExecuteReader
        Dim s As String = "|_________|_________|_________|_________|"
        Console.WriteLine("{0}{1}{0}", s, vbCrLf + "|index    | coef_ABC|     Name|    color|" + Chr(10) + Chr(13))

        While (dr_coefficient.Read)
            Console.WriteLine("|{0,-9}|{1,9}|{2,9}|{3,9}|", index, dr_coefficient.Item("coefficients_ABC"), dr_coefficient.Item("Name"), dr_coefficient.Item("color_straight"))
            index += 1
        End While
        Console.WriteLine(s)
    End Sub

    ' вивід списка в табличній формі 
    Public Sub printTableList()

        Dim index As Integer
        index = 1

        Console.WriteLine("List\n")
        Dim s As String = "|_________|_________|_________|_________|"
        Console.WriteLine("{0}{1}{0}", s, vbCrLf + "|index    | coef_ABC|     Name|    color|" + Chr(10) + Chr(13))
        Dim i As Integer
        i = 0
        While i < names.Count
            Console.WriteLine("|{0, -9}|{1, 9}|{2, 9}|{3, 9}|", index, names(i).coefficients, names(i).Name, names(i).color_straight)
            i += 1
            index += 1
        End While
        Console.WriteLine(s)
    End Sub

    ' вивід списка 
    Public Sub printList()

        Dim index, i As Integer
        index = 1
        i = 0
        Console.WriteLine(names.Count)
        While i < names.Count
            Console.WriteLine("|{0, -9}|{1, 9}|{2, 9}|{3, 9}|", index, names(i).coefficients, names(i).Name, names(i).color_straight)
            i += 1
            index += 1
        End While

    End Sub

    'додавання елемента в список 
    Public Sub Add1()

        Dim r As New straight
        Dim string1, string2, string3 As String
        Console.Write("coefficient:    ")
        string1 = Console.ReadLine()
        Console.Write("Name: ")
        string2 = Console.ReadLine()
        Console.Write("color: ")
        string3 = Console.ReadLine()
        If string1 = "" Or string2 = "" Or string3 = "" Then
            Throw New System.Exception("дані введено некоректно")
        Else
            Dim count As Integer
            count = 1
            outFille()
            Dim coefficients As New OleDbCommand("select i from Пряма_у_просторі", conn)
            Dim dr_coefficient As OleDbDataReader = coefficients.ExecuteReader
            While dr_coefficient.Read
                count += 1
            End While
            r.index = count
            r.coefficients = string1
            r.Name = string2
            r.color_straight = string3
            names.Add(r)
        End If
    End Sub

    'додавання в Access базу даних
    Public Sub AddTable()

        Dim i, index As Integer
        i = 0
        index = 1
        Dim coefficients As New OleDbCommand("select i from Пряма_у_просторі", conn)
        Dim dr_coefficient As OleDbDataReader = coefficients.ExecuteReader
        RemovTable()
        While i < names.Count
            Dim c As New OleDbCommand
            c.Connection = conn
            c.CommandText = "insert into Пряма_у_просторі (i,coefficients_ABC,Name,color_straight) values('" & index & "','" & names(i).coefficients & "','" & names(i).Name & "', '" & names(i).color_straight & "')"
            c.ExecuteNonQuery()
            i += 1
            index += 1
        End While

    End Sub

    ' видалення усієї бази даних Access
    Public Sub RemovTable()
        Dim c As New OleDbCommand
        c.Connection = conn
        Dim coefficients As New OleDbCommand("select i from Пряма_у_просторі", conn)
        Dim dr_coefficient As OleDbDataReader = coefficients.ExecuteReader
        While dr_coefficient.Read
            c.CommandText = "delete from Пряма_у_просторі where i = " & dr_coefficient.Item("i")
            c.ExecuteNonQuery()
        End While
    End Sub

    ' сорутвання бульбашкою =)
    Public Sub bubbleSort()

        Dim n As New straight
        Dim i, j, flag As Integer
        flag = 1
        Console.WriteLine(names.Count)
        While flag > 0
            flag = 0

            For j = 1 To names.Count - 1
                If names(j - 1).Name > names(j).Name Then
                    n.index = names(j - 1).index
                    n.coefficients = names(j - 1).coefficients
                    n.Name = names(j - 1).Name
                    n.color_straight = names(j - 1).color_straight
                    names(j - 1) = names(j)
                    names(j) = n
                    flag = 1
                End If

            Next

        End While

    End Sub

End Class
