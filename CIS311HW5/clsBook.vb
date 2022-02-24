Public Class clsBook
    '------------------------------------------------------------------------
    '-                  File Name: clsBook                                  -
    '-                  Part of Project: CIS311 HW5(bookstore inventory app)-
    '------------------------------------------------------------------------
    '-                      Written By: Andrew A. Loesel                    -
    '-                      Written On: February 22, 2022                   -
    '------------------------------------------------------------------------
    '- File Purpose:                                                        -
    '-                                                                      -
    '- The purpose of this file is to define a data container for the books.-
    '- There is no IO in this file, and only 1 subprogram which is a paramet-
    '- erized constructor for the clsBook type.
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
    '- intQuantity - the number of this specific book in the inventory.     -
    '- sngInventoryTotal - the total value of this specific book in the inve-
    '-                     ntory, intQuantity * sngPrice.                   -
    '- sngPrice - the price of this specific book, per book.                -
    '- strCategory - the category of this book, N for Non fiction, S for    -
    '-               Science fiction or F for Fiction.                      -
    '- strTitle - the title of this book.                                   -
    '------------------------------------------------------------------------

    'PROPERTIES PROPERTIES PROPERTIES PROPERTIES PROPERTIES PROPERTIES PROPERTIES PROPERTIES
    'PROPERTIES PROPERTIES PROPERTIES PROPERTIES PROPERTIES PROPERTIES PROPERTIES PROPERTIES
    Public Property strCategory As String
    Public Property intQuantity As Integer
    Public Property sngPrice As Single
    Public Property strTitle As String
    Public Property sngInventoryTotal As Single

    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    Public Sub New(ByVal strCategory As String, ByVal intQuantity As Integer, ByVal sngPrice As Single, ByVal strTitle As String, ByVal sngTotal As Single)
        '------------------------------------------------------------------------
        '-                      Subprogram Name: New                            -
        '------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                    -
        '-                      Written On: February 22, 2022                   -
        '------------------------------------------------------------------------
        '- Subprogram Purpose:                                                  -
        '-                                                                      -
        '- The purpose of this subroutine is to create a new instance of a clsBo-
        '- ok object, and set the clsBooks object's properties to the arguments -
        '- passed in to the subroutine.                                         -
        '------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                           -
        '- strCategory - will be set to the new object's strCategory property.  -
        '- intQuantity - will be set to the new object's intQuantity property.  -
        '- sngPrice - will be set to the new object's sngPrice property.        -
        '- strTitle - will be set to the new object's strTitle property.        -
        '- sngTotal - will be set to the new object's sngTotal property.        -
        '------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                          -
        '- None                                                                 -
        '------------------------------------------------------------------------
        Me.strCategory = strCategory
        Me.intQuantity = intQuantity
        Me.sngPrice = sngPrice
        Me.strTitle = strTitle
        Me.sngInventoryTotal = sngTotal
    End Sub

End Class
