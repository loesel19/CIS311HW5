
Imports System.IO
Imports System.Text.RegularExpressions
Module Program
    '------------------------------------------------------------------------
    '-                  File Name: Program                                  -
    '-                  Part of Project: CIS311 HW5(bookstore inventory app)-
    '------------------------------------------------------------------------
    '-                      Written By: Andrew A. Loesel                    -
    '-                      Written On: February 23, 2022                   -
    '------------------------------------------------------------------------
    '- File Purpose:                                                        -
    '-                                                                      -
    '- The purpose of this file is to act as the command line interface for -
    '- the application. All input and output is handled within this file.   -
    '-                                                                      -
    '------------------------------------------------------------------------
    '- Program Purpose:                                                     -
    '-                                                                      -
    '- The purpose of this program is to create clsBook objects from a user -
    '- specified file and then print various reports about the books to the -
    '- command line interface. All statistics about the books will be from  -
    '- LINQ queries, and the only data collection used will be the list of  -
    '- clsBooks, myList. We make sure that no invalid entries will be shown -
    '- to the user, and once the reports are printed the application is     -
    '- still visible so that the user may view the reports we generated.    -
    '------------------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically):                         -
    '- myBooks - a List of clsBook objects, which will be populated with all-
    '-           the valid books from the user specified file.              -
    '- strBookPath - a string containing the user specified file path and   -
    '-               name of the file to create a report from.              -
    '------------------------------------------------------------------------
    'GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS
    'GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS GLOBAL CONSTANTS
    Const FICTION_CHARACTER = "F" 'denotes a book of the fiction category
    Const NONFICTION_CHARACTER = "N" 'denotes a book of the non fiction category
    Const SCIENCE_FICTION_CHARACTER = "S" 'denotes a book of the science fiction category
    Const PRICE_RANGE_MULTIPLIER = 50 'multiplier used while printing unit price ranges, means that the range will be intervals of this constant
    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES
    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES
    Dim myBooks As New List(Of clsBook)
    Dim strBookPath As String

    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    Sub Main(args As String())
        '------------------------------------------------------------------------
        '-                      Subprogram Name: Main                           -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: February 22, 2022                   -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subroutine is to control the flow of this applica-
        '- tion. First we try to get a file to read from the user in fileToReadI-
        '- nteraction, if we can not read the file we end the application execut-
        '- ion. If we can read the file we call readFromFile to try and populate-
        '- myBooks with clsBook objects. If there are no books in this list we  -
        '- do not print any reports and the application terminates. If we have  -
        '- some books in the list we can print the various reports using the sub-
        '- routines in the inner If block of this subroutine.                   -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- args - a string array of program arguments, not used in this applicat-
        '-        ion.                                                          -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- None                                                                 -
        '------------------------------------------------------------------------
        'change the color of the console
        Console.BackgroundColor = ConsoleColor.White
        Console.ForegroundColor = ConsoleColor.Black
        Console.Clear()
        'This program flow should fix the issue I had with assignment 3 the input procedure would
        'iterate until a good path was given by the user. If the file is not correct then the program
        'will halt execution after writing the end of execution method
        If FileToReadInteraction() Then
            ReadFromFile()
            'now check that some books have been added to the myBooks list
            If myBooks.Count > 0 Then
                PrintInventoryReport()
                PrintTotalInventoryValue()
                PrintUnitPriceRange()
                PrintOverallStatistics()
            Else
                Console.WriteLine("Unable to display report, no books were found in the file.")
            End If
        Else
            Console.WriteLine("Program execution has ended, please try again.")
        End If
    End Sub
    Function FileToReadInteraction() As Boolean
        '------------------------------------------------------------------------
        '-                      Subprogram Name: fileToReadInteraction          -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: February 23, 2022                   -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subprogram is to have the user input the name and-
        '- path of the file that contains book information. If the file does not-
        '- exist or does not contain the appropriate .txt extension then this   -
        '- function returns false. If the file exists the function will return  -
        '- true.                                                                -
        '- This function is a slightly altered version of the subroutine used in-
        '- my assignment 3.                                                     -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Retruns:                                                             -
        '- boolean - true if file can be opened and is a .txt, false otherwise. -
        '------------------------------------------------------------------------

        Console.WriteLine("Please enter the path and name of the file to process:")
        strBookPath = Console.ReadLine
        'see if file exists, if it does return true
        If System.IO.File.Exists(strBookPath) Then
            Return True
        Else
            Console.WriteLine("No such file found, check file path/name.")
            Return False
        End If
    End Function
    Sub ReadFromFile()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: readFromFile                   -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: February 22, 2022                   -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subprogram is to read every line of the specified-
        '- file, and try to create a clsBook object from the arguments in the   -
        '- line to add to the List of clsBook, myBooks. We by looping through   -
        '- each line of the file, checking that the first 3 arguments, category,-
        '- quantity and price are of the right type, and printing the user a mes-
        '- sage if not. We then create the title by concatonating elements (3)  -
        '- through (x) of split string array we create together, we then check t-
        '- this string, strTitle contains at least 1 alphanumeric character. If -
        '- all arguments were accepted we create a new clsBook entity and add it-
        '- to our global list of clsBook. If not we display a message telling   -
        '- the user the book was not added, and what line of the file it was on.-
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- blnContinue - a boolean whose value determines if the current file   -
        '-               line will continue to be processed or not.             -
        '- intLoopCount - an integer where the value represents the current iter-
        '-                ation in the while loop.                              -
        '- intQuantity - an integer that represents the quantity of books for   -
        '-               the current book.                                      -
        '- reader - The stream reader object we are using to read the file.     -
        '- sngPrice - a single representing the price per book of the current   -
        '-            book.                                                     -
        '- strCategory - a string representing the category of the current book.-
        '- strLine - an array of strings containing the split contents of the   -
        '-           current line in the file, split by space.                  -
        '- strTitle - a string representing the title of the current book.      -
        '------------------------------------------------------------------------
        Dim strCategory As String = ""
        Dim intQuantity As Integer = 0
        Dim sngPrice As Single = 0
        Dim strTitle As String
        Dim intLoopCount As Integer = 0
        'TODO add in input checking
        Using reader As StreamReader = New StreamReader(strBookPath, True)
            'loop through each line
            While Not reader.EndOfStream
                Dim blnContinue As Boolean = True
                'increment loop count preemptively so we do not have to worry about saying + 1 later when we want
                'to show what line we are on
                intLoopCount += 1
                Dim strLine As String() = reader.ReadLine.Split(" ")
                'we have to check that the file is in the right format, first lets make sure that the split string
                'array contains at least 4 elements
                If strLine.Count > 3 Then
                    'for the purpose of this program I will assume that the category will only be F N or S
                    '***************************
                    If strLine(0).Equals("S") Or strLine(0).Equals("N") Or strLine(0).Equals("F") Then
                        strCategory = strLine(0)
                    Else
                        Console.WriteLine(String.Format("First argument of line {0} must be either N, S or F.", intLoopCount))
                        blnContinue = False
                    End If
                    'use regex to make sure that the second argument of the line is only an integer
                    If Regex.IsMatch(strLine(1), "[0-9]+") Then
                            intQuantity = CInt(strLine(1))
                        Else
                            blnContinue = False
                            Console.WriteLine(String.Format("Second argument of line {0} must be an integer.", intLoopCount))
                        End If
                    'use regex to see if third argument is only a decimal or, an integer (if int it will be cast to a single properly)
                    If Regex.IsMatch(strLine(2), "[0-9]+\.?[0-9]") Then
                        sngPrice = CSng(strLine(2))
                    Else
                        blnContinue = False
                            Console.WriteLine(String.Format("Thrid argument of line {0} must be a decimal number or integer.", intLoopCount))
                        End If
                        'reset strTitle to empty string so that the title does not contain the past title as well
                        strTitle = ""
                        If blnContinue Then
                            'loop through all the remainder of the string to get the title
                            For i As Integer = 3 To strLine.Count - 1
                                'avoid adding an extra space at the end of the title
                                If (i = strLine.Count - 1) Then
                                    'here we just dont put a space at the end
                                    strTitle = strTitle & strLine(i).Trim
                                Else
                                    strTitle = strTitle & strLine(i).Trim & " "
                                End If
                            Next
                        End If
                        If blnContinue Then
                            'check to see that at least 1 character was in the book title
                            If Regex.IsMatch(strTitle, "[ a-zA-Z0-9]+") Then

                                'add to the list
                                myBooks.Add(New clsBook(strCategory, intQuantity, sngPrice, strTitle, CSng(intQuantity * sngPrice)))

                            Else
                                Console.WriteLine(String.Format("Fourth argument of line {0} must contain the title of the book, but is currently nothing.", intLoopCount))
                            End If
                        Else
                            Console.WriteLine(String.Format("Book from line {0} was not added to the report.", intLoopCount))
                        End If
                    Else
                        'here we print a message noting which line did not have enough entries in it
                        Console.WriteLine(String.Format("Line {0} of file did not contain enough arguments to make a complete book entry, book not added.", intLoopCount))
                End If
            End While
        End Using

    End Sub

    Sub PrintInventoryReport()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: printInventoryReport           -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: February 22, 2022                   -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subroutine is to print all the books in myBooks  -
        '- on the console. We first print a couple of formatted headers, then we-
        '- perform a LINQ query to get all books sorted by the title. We then   -
        '- do a for each loop on the results of the query and print all the info-
        '- rmation of each book in a format matching the headers. If the LINQ   -
        '- query returns nothing we tell the user that there are no books to dis-
        '- play.                                                                -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- sortedBooks - will hold the results of the LINQ query for all books  -
        '-               sorted by title.                                       -
        '- strTemp - a formatedd string that contains the information about the -
        '-           current book we want to display to the user.               -
        '------------------------------------------------------------------------
        Console.WriteLine(Space(33) & "Books 'R' Us")
        Console.WriteLine(Space(28) & "*** Inventory Report ***")
        Console.WriteLine(Space(25) & "-----------------------------")
        Console.WriteLine(Space(5) & "Title" & Space(22) & "Category" & Space(5) & "Quantity" & Space(5) & "Unit Cost" & Space(3) & "Extended Cost")
        Console.WriteLine(StrDup(29, "-") & Space(3) & StrDup(8, "-") & Space(5) & StrDup(8, "-") & Space(5) & StrDup(9, "-") & Space(3) & StrDup(13, "-"))
        'now we have to get a sorted generics list of all the books sorted by title
        Dim sortedBooks = From books In myBooks Order By books.strTitle Select books
        'make sure that there are books to be displayed, if none display a message
        'we should be handling this case in the readFromFile method, but it does not hurt anything to
        'double check it here
        If Not sortedBooks.Count = 0 Then
            For Each book In sortedBooks
                'we have to make a formatted string to print out the book information
                Dim strTemp As String
                'add title
                strTemp = String.Format("   {0,-29}   {1,2}   {2,9}   {3, 13}   {4,15:f2}", book.strTitle, book.strCategory, book.intQuantity, book.sngPrice, book.sngInventoryTotal)
                Console.WriteLine(strTemp)
            Next
        Else
            Console.WriteLine("(There are no books to display)")
        End If
    End Sub
    Sub PrintTotalInventoryValue()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: printTotalInventoryValue       -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: February 22, 2022                   -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subroutine is to display what books have an inven-
        '- tory total value in specific price ranges. We first print a few heade-
        '- rs. Next we loop 4 times, since we want to print out 4 price ranges. -
        '- The first 3 iterations Print the same header format with the lower an-
        '- d upper bounds first, and then we perform a LINQ query to get all boo-
        '- ks in the current price range. We check to see that the query contain-
        '- s at least 1 book, if it does not we print out a no books message. If-
        '- it has at least 1 book we iterate over every result in a for each    -
        '- loop and print out the title and inventory total. On the last outer  -
        '- for loop iteration we print just the lower bound in the header and   -
        '- lower bound and above, and our LINQ query is modified to get all     -
        '- books with a total price greater than the value, but printing results-
        '- is the same as the other outer loop iterations.                      -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- priceSortedBooks - a LINQ query result that contains the books within-
        '-                    the current price rang sorted by their inventory  -
        '-                    totals.                                           -
        '- strTemp - a string that holds the formatted information we will print-
        '-           on the console.                                            -
        '------------------------------------------------------------------------
        Console.WriteLine(StrDup(83, "-"))
        Console.WriteLine(Space(10) & "Total Inventory Value (Quantity * Unit Price) Statistics")
        Console.WriteLine(StrDup(83, "-"))

        For i As Integer = 1 To 4
            'We can only use the same case for i = 1 to 3, since for 4 we are just doing all the books above i * 50
            If i < 4 Then
                Console.WriteLine(String.Format("Those books in the range of {0,5:f2} to {1,5:f2} are:", PRICE_RANGE_MULTIPLIER * (i - 1), PRICE_RANGE_MULTIPLIER * i))
                'make a generic container sorted by price, and first assign it to the books with an inventory total under 50
                Dim priceSortedBooks = From books In myBooks Order By books.sngInventoryTotal Ascending Where books.sngInventoryTotal < (PRICE_RANGE_MULTIPLIER * i) And books.sngInventoryTotal > (PRICE_RANGE_MULTIPLIER * (i - 1))
                                       Select books
                'check that there were some books in this price range, if not display a message
                If Not priceSortedBooks.Count = 0 Then

                    For Each book In priceSortedBooks
                        Dim strTemp As String
                        strTemp = String.Format("     {0,-29}    Price: {1,5:c}", book.strTitle, book.sngInventoryTotal)
                        Console.WriteLine(strTemp)

                    Next

                Else
                    Console.WriteLine("(There are no books in this range)")
                End If
            Else
                Console.WriteLine(String.Format("Those books in the range of {0,5:f2} and above are:", PRICE_RANGE_MULTIPLIER * (i - 1)))
                Dim priceSortedBooks = From books In myBooks Order By books.sngInventoryTotal Ascending Where books.sngInventoryTotal > (PRICE_RANGE_MULTIPLIER * (i - 1))
                                       Select books
                If Not priceSortedBooks.Count = 0 Then

                    For Each book In priceSortedBooks
                        Dim strTemp As String
                        strTemp = String.Format("     {0,-29}    Price: {1,5:c}", book.strTitle, book.sngInventoryTotal)
                        Console.WriteLine(strTemp)

                    Next

                Else
                    Console.WriteLine("(There are no books in this range)")
                End If
            End If
        Next
    End Sub

    Sub PrintUnitPriceRange()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: printUnitPriceRange            -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: February 22, 2022                   -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subroutine is to print statistics for each catego-
        '- ry of books in the myBooks list. We first print some header lines.   -
        '- Next we create a variables for the statistics we want to display for -
        '- the current category : min, max, avg price and #of title. We use the -
        '- Aggregate LINQ query to get these statistics starting with fiction bo-
        '- oks. We then check to see that there are at least 1 title(s) in the  -
        '- category, if not we print a message. If there are then we print a for-
        '- matted line with all the statistics in it. We then repeat this proces-
        '- s for Non fiction and science fiction books.                         -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- intTitles - an integer representing the number of books in the curren-
        '-             t category.                                              -
        '- sngAvg - a single representing the average price of books in the curr-
        '-          ent category.                                               -
        '- sngHigh - a single representing the max price of a book in the curren-
        '-           t category.                                                -
        '- sngLow - a single representing the min price of a book in the current-
        '-          category.                                                   -
        '------------------------------------------------------------------------
        Console.WriteLine(StrDup(83, "-"))
        Console.WriteLine(Space(15) & "Unit Price Range by Category Statistics")
        Console.WriteLine(StrDup(83, "-"))
        Console.WriteLine("Category" & Space(6) & "# of Titles" & Space(9) & "Low" & Space(15) & "Ave" & Space(15) & "High")
        'now we need handles on the statistics we need to retreive from LINQ for F first
        Dim intTitles As Integer = Aggregate books In myBooks Where books.strCategory = FICTION_CHARACTER Into Count
        Dim sngLow As Single = Aggregate books In myBooks Where books.strCategory = FICTION_CHARACTER Select books.sngPrice Into Min()
        Dim sngAvg As Single = Aggregate books In myBooks Where books.strCategory = FICTION_CHARACTER Select books.sngPrice Into Average()
        Dim sngHigh As Single = Aggregate books In myBooks Where books.strCategory = FICTION_CHARACTER Select books.sngPrice Into Max()
        'check that count > 0 so that we have satistics to print out, if count = 0 then we can display some message
        'note as well that here we need to declare our low, avg and high variables before the if, but when we check the following
        'categories we do not need to do this, and can assign their values in the if, since they will not be used outside of the if block
        If intTitles > 0 Then
            Console.WriteLine(String.Format(Space(4) & "{0} " & StrDup(12, ".") & "{1,2} " & StrDup(12, ".") & "{2,5:f2} " & StrDup(12, ".") & "{3, 5:f2} " & StrDup(12, ".") & "{4, 5:f2}",
"F", intTitles, sngLow, sngAvg, sngHigh))
        Else
            'print a message
            Console.WriteLine(String.Format("(There are no books with category '{0}'.)", FICTION_CHARACTER))
        End If
        'now do it for N
        intTitles = Aggregate books In myBooks Where books.strCategory = NONFICTION_CHARACTER Into Count
        'do the check here now
        If intTitles > 0 Then
            sngLow = Aggregate books In myBooks Where books.strCategory = NONFICTION_CHARACTER Select books.sngPrice Into Min()
            sngAvg = Aggregate books In myBooks Where books.strCategory = NONFICTION_CHARACTER Select books.sngPrice Into Average()
            sngHigh = Aggregate books In myBooks Where books.strCategory = NONFICTION_CHARACTER Select books.sngPrice Into Max()
            'print again
            Console.WriteLine(String.Format(Space(4) & "{0} " & StrDup(12, ".") & "{1,2} " & StrDup(12, ".") & "{2,5:f2} " & StrDup(12, ".") & "{3, 5:f2} " & StrDup(12, ".") & "{4, 5:f2}",
"N", intTitles, sngLow, sngAvg, sngHigh))
        Else
            Console.WriteLine(String.Format("(There are no books with category '{0}'.", NONFICTION_CHARACTER))
        End If
        'now do it for S
        intTitles = Aggregate books In myBooks Where books.strCategory = SCIENCE_FICTION_CHARACTER Into Count
        'do the check again
        If intTitles > 0 Then
            sngLow = Aggregate books In myBooks Where books.strCategory = SCIENCE_FICTION_CHARACTER Select books.sngPrice Into Min()
            sngAvg = Aggregate books In myBooks Where books.strCategory = SCIENCE_FICTION_CHARACTER Select books.sngPrice Into Average()
            sngHigh = Aggregate books In myBooks Where books.strCategory = SCIENCE_FICTION_CHARACTER Select books.sngPrice Into Max()
            'print again
            Console.WriteLine(String.Format(Space(4) & "{0} " & StrDup(12, ".") & "{1,2} " & StrDup(12, ".") & "{2,5:f2} " & StrDup(12, ".") & "{3, 5:f2} " & StrDup(12, ".") & "{4, 5:f2}",
    "S", intTitles, sngLow, sngAvg, sngHigh))
        Else
            Console.WriteLine(String.Format("(There are no books with category '{0}'.", SCIENCE_FICTION_CHARACTER))
        End If
    End Sub

    Sub PrintOverallStatistics()
        '------------------------------------------------------------------------
        '-                      Subprogram Name: printOverallStatistics         -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: February 22, 2022                   -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subroutine is to print which books are the cheape-
        '- st, priciest, show up the least or most. We use the same approach for-
        '- all 4 statistics. We first set sngTemp to an Aggregate LINQ query to -
        '- get a value for cheapest, priciest book price, or fewest, most book  -
        '- quantities. We then set objTemp to a LINQ query result for books that-
        '- have either the price or quantity held in sngTemp. We then use a for -
        '- each to go through each book in objTemp and print out the title. Note-
        '- we never check that objTemp.count is > 0 since we know that if this  -
        '- sub is being called there are books in the global myBooks list, so   -
        '- each LINQ query must return at least 1 book.                         -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- None                                                                 -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- objTemp - holds the results of a LINQ query to find cheapest, pricies-
        '-           t, fewest and most books.                                  -
        '- sngTemp - a single which will hold the value for cheapest, priciest, -
        '-           fewest and most books.                                     -
        '------------------------------------------------------------------------
        Console.WriteLine(StrDup(83, "-"))
        Console.WriteLine(Space(20) & "Overall Book Statistics")
        Console.WriteLine(StrDup(83, "-"))

        Dim sngTemp = Aggregate books In myBooks Select books.sngPrice Into Min()
        'get the cheapest books
        Dim objTemp = From books In myBooks Where books.sngPrice = sngTemp
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
