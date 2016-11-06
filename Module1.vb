Module Module1

    Sub Main()
        Dim index As Integer
        Dim A As New Access
        Do
            Console.WriteLine()
            Console.WriteLine("1.Input Access")
            Console.WriteLine("2.Add ArryList")
            Console.WriteLine("3.print ArryList")
            Console.WriteLine("4.print ArryListTable")
            Console.WriteLine("5.Sort")
            Console.WriteLine("6.output Baza")
            Console.WriteLine("7.print AccessTable")
            Console.WriteLine("8.Exit")
            Console.WriteLine()
            Console.Write("namber: ")
            index = Console.ReadLine()
            Select Case index
                Case 1
                    A.Input()
                Case 2
                    A.Add1()
                Case 3
                    A.printList()
                Case 4
                    A.printTableList()
                Case 5
                    A.bubbleSort()
                Case 6
                    A.AddTable()
                Case 7
                    A.printAcessTable()
                Case 8
                    Exit Sub
            End Select
        Loop While (True)

    End Sub

End Module
