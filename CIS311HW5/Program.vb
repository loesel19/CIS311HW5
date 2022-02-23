
Imports System.IO

Module Program

    'GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS
    'GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS
    Const BOOK_PATH = "C:\Users\aaloesel\Source\Repos\loesel19\CIS311HW5\CIS311HW5\books.txt"
    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES
    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES
    Dim myBooks As New List(Of clsBook)
    Sub Main(args As String())
        readFromFile()
        printInventoryReport()
        printTotalInventoryValue()
        printUnitPriceRange()
        printOverallStatistics()
    End Sub

    Sub readFromFile()
        Dim strCategory As String = ""
        Dim intQuantity As Integer = 0
        Dim sngPrice As Single = 0
        Dim strTitle As String
        'Dim sngInventoryTotal As Single
        Using reader As New StreamReader(BOOK_PATH, True)
            'loop through each line
            While (Not reader.EndOfStream)
                Dim strLine As String() = reader.ReadLine.Split(" ")
                strCategory = strLine(0)
                intQuantity = CInt(strLine(1))
                sngPrice = CSng(strLine(2))
                'reset strTitle to empty string so that the title does not contain the past title as well
                strTitle = ""
                'loop through all the remainder of the string to get the title
                For i As Integer = 3 To strLine.Count - 1
                    'avoid adding an extra space at the end of the title
                    If (i = strLine.Count - 1) Then
                        'here we just dont put a space at the end
                        strTitle = strTitle & strLine(i)
                    Else
                        strTitle = strTitle & strLine(i) & " "
                    End If
                Next
                'add to the list
                myBooks.Add(New clsBook(strCategory, intQuantity, sngPrice, strTitle, CSng(intQuantity * sngPrice)))
            End While
        End Using

    End Sub

    Sub printInventoryReport()
        Console.WriteLine(Space(33) & "Books 'R' Us")
        Console.WriteLine(Space(28) & "*** Inventory Report ***")
        Console.WriteLine(Space(25) & "-----------------------------")
        Console.WriteLine(Space(5) & "Title" & Space(22) & "Category" & Space(5) & "Quantity" & Space(5) & "Unit Cost" & Space(3) & "Extended Cost")
        Console.WriteLine(StrDup(29, "-") & Space(3) & StrDup(8, "-") & Space(5) & StrDup(8, "-") & Space(5) & StrDup(9, "-") & Space(3) & StrDup(13, "-"))
        'now we have to get a sorted generics list of all the books sorted by title
        Dim sortedBooks As Object
        sortedBooks = From books In myBooks Order By books.strTitle Select books
        For Each book In sortedBooks
            'we have to make a formatted string to print out the book information
            Dim strTemp As String
            'add title
            strTemp = String.Format("   {0,-29}   {1,2}   {2,9}   {3, 13}   {4,15:f2}", book.strTitle, book.strCategory, book.intQuantity, book.sngPrice, book.sngInventoryTotal)
            Console.WriteLine(strTemp)
        Next
    End Sub
    Sub printTotalInventoryValue()
        Console.WriteLine(StrDup(80, "-"))
        Console.WriteLine(Space(10) & "Total Inventory Value (Quantity * Unit Price) Statistics")
        Console.WriteLine(StrDup(80, "-"))

        For i As Integer = 1 To 4
            Console.WriteLine(String.Format("Those books in the range of {0,5:f2} to {1,5:f2} are:", 50 * (i - 1), 50 * i))
            'make a generic container sorted by price, and first assign it to the books with an inventory total under 50
            Dim priceSortedBooks As Object = Nothing
            priceSortedBooks = From books In myBooks Order By books.sngInventoryTotal Ascending Where books.sngInventoryTotal < (50 * i) And books.sngInventoryTotal > (50 * (i - 1))
                               Select books
            If Not priceSortedBooks.Equals(Nothing) Then
                Dim count As Integer = 0
                For Each book In priceSortedBooks
                    Dim strTemp As String
                    strTemp = String.Format("     {0,-29}    Price: {1,5:c}", book.strTitle, book.sngInventoryTotal)
                    Console.WriteLine(strTemp)
                    count += 1
                Next
                'this works
                If count = 0 Then
                    Console.WriteLine("(There are no books in this range)")
                End If
            Else
                'this will never print
                Console.WriteLine("(There are no books in this range)")
            End If
        Next
    End Sub

    Sub printUnitPriceRange()
        Console.WriteLine(StrDup(80, "-"))
        Console.WriteLine(Space(15) & "Unit Price Range by Category Statistics")
        Console.WriteLine(StrDup(80, "-"))
        Console.WriteLine("Category" & Space(6) & "# of Titles" & Space(9) & "Low" & Space(15) & "Ave" & Space(15) & "High")
        'now we need handles on the statistics we need to retreive from LINQ for F first
        Dim intTitles As Integer = Aggregate books In myBooks Where books.strCategory = "F" Into Count
        Dim sngLow As Single = Aggregate books In myBooks Where books.strCategory = "F" Select books.sngPrice Into Min()
        Dim sngAvg As Single = Aggregate books In myBooks Where books.strCategory = "F" Select books.sngPrice Into Average()
        Dim sngHigh As Single = Aggregate books In myBooks Where books.strCategory = "F" Select books.sngPrice Into Max()

        Console.WriteLine(String.Format(Space(4) & "{0} " & StrDup(12, ".") & "{1,2} " & StrDup(12, ".") & "{2,5:f2} " & StrDup(12, ".") & "{3, 5:f2} " & StrDup(12, ".") & "{4, 5:f2}",
"F", intTitles, sngLow, sngAvg, sngHigh))
        'now do it for N
        intTitles = Aggregate books In myBooks Where books.strCategory = "N" Into Count
        sngLow = Aggregate books In myBooks Where books.strCategory = "N" Select books.sngPrice Into Min()
        sngAvg = Aggregate books In myBooks Where books.strCategory = "N" Select books.sngPrice Into Average()
        sngHigh = Aggregate books In myBooks Where books.strCategory = "N" Select books.sngPrice Into Max()
        'print again
        Console.WriteLine(String.Format(Space(4) & "{0} " & StrDup(12, ".") & "{1,2} " & StrDup(12, ".") & "{2,5:f2} " & StrDup(12, ".") & "{3, 5:f2} " & StrDup(12, ".") & "{4, 5:f2}",
"N", intTitles, sngLow, sngAvg, sngHigh))

        'now do it for S
        intTitles = Aggregate books In myBooks Where books.strCategory = "S" Into Count
        sngLow = Aggregate books In myBooks Where books.strCategory = "S" Select books.sngPrice Into Min()
        sngAvg = Aggregate books In myBooks Where books.strCategory = "S" Select books.sngPrice Into Average()
        sngHigh = Aggregate books In myBooks Where books.strCategory = "S" Select books.sngPrice Into Max()
        'print again
        Console.WriteLine(String.Format(Space(4) & "{0} " & StrDup(12, ".") & "{1,2} " & StrDup(12, ".") & "{2,5:f2} " & StrDup(12, ".") & "{3, 5:f2} " & StrDup(12, ".") & "{4, 5:f2}",
"S", intTitles, sngLow, sngAvg, sngHigh))
    End Sub

    Sub printOverallStatistics()
        Console.WriteLine(StrDup(80, "-"))
        Console.WriteLine(Space(20) & "Overall Book Statistics")
        Console.WriteLine(StrDup(80, "-"))

        Dim sngTemp = Aggregate books In myBooks Select books.sngPrice Into Min()
        'get the cheapest books
        Dim objTemp As Object = From books In myBooks Where books.sngPrice = sngTemp
        Console.WriteLine(String.Format("The cheapest book title(s) at a unit price of {0:c} are:", sngTemp))
        For Each book In objTemp
            Console.WriteLine(Space(6) & book.strTitle)
        Next
        'now we do the priciest book
        sngTemp = Aggregate books In myBooks Select books.sngPrice Into Max()
        objTemp = From books In myBooks Where books.sngPrice = sngTemp
        Console.WriteLine(String.Format("The priciest book title(s) at a unit price of {0:c} are:", sngTemp))
        For Each book In objTemp
            Console.WriteLine(Space(6) & book.strTitle)
        Next

        'now least quantity
        sngTemp = Aggregate books In myBooks Select books.intQuantity Into Min()
        objTemp = From books In myBooks Where books.intQuantity = sngTemp
        Console.WriteLine(String.Format("The title(s) with the least quantity on hand at {0} are:", sngTemp))
        For Each book In objTemp
            Console.WriteLine(Space(6) & book.strTitle)
        Next

        'now highest quantity
        sngTemp = Aggregate books In myBooks Select books.intQuantity Into Max()
        objTemp = From books In myBooks Where books.intQuantity = sngTemp
        Console.WriteLine(String.Format("The title(s) with the most quantity on hand at {0} are:", sngTemp))
        For Each book In objTemp
            Console.WriteLine(Space(6) & book.strTitle)
        Next
    End Sub
End Module
