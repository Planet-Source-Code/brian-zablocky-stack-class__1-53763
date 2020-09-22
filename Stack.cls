VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ======================================================================
'
'    Stack Class!  NOTICE: If you scroll way down you will find a
'    detailed discussion of the class.
'
' ----------------------------------------------------------------------
'
' This is my first successful stack.  You can PUSH items to the stack,
' then POP them off when you need them.  If you want to peek at an item
' but not POP it, then simply READ the item.
'
' A true stack looks like a stack of papers, as in you add new pages to
' the top.  When you remove a page, you take the first page, or the last
' one you added.  This is LAST-IN-FIRST-OUT, or simply LIFO.
'
' Another way of handling it is like adding new pages to the bottom of
' the stack, and still removing pages from the top.  This is a FIRST-IN-
' FIRST-OUT methodology, or simply FIFO.
'
'
' ======================================================================
'


Private Const DEFAULT_REDIM = 10    ' Used for initializing the stack limit
Private Const INCREMENT_REDIM = 3   ' Used for adding to the stack limit
' WARNING: These two constants must be at least 1 each.  However, the
' default constant should be set a little higher.  This speeds up access to
' the stacks by not causing so much redims



Private mvarStack() As Variant      ' The actual stack array.  Simple, eh?







Private Sub Class_Initialize()
    
    ReDim mvarStack(DEFAULT_REDIM)      ' Prepare the stack array.
                                        
            ' Changing the above line to utilize the options function is
            ' dumb because you cannot programmatically set the options
            ' until after this procedure has finished anyway.
    
    ' -----------------------------------------------------
    ' PUT OTHER STUFF BELOW THIS LINE
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
    
    
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - -
    ' PUT OTHER STUFF ABOVE THIS LINE
    ' -----------------------------------------------------
End Sub




' ===========================================================================
' ===========================================================================



Public Sub sPush(StackItem As Variant)
    '
    ' Adds a new value to the end of the stack array.  The value can then
    ' be popped of with sPop.
    '
    c = GetStackCount()                     ' How many stack elements?
    
    If c >= UBound(mvarStack()) Then                    ' If we are out of
        ReDim Preserve mvarStack(c + INCREMENT_REDIM)     ' elements in the
                                                        ' stack array, then
    End If                                              ' add some more.
    
    mvarStack(c) = StackItem                ' Put the item on the stack.
        
    'Discussion present below
End Sub



' ===========================================================================
' ===========================================================================



Public Function sPop() As Variant
    '
    ' Reads a value from the end of the stack (newest item) and then deletes
    ' the item from the stack.  You can use VB's VarType() to check the type.
    '
   
    c = GetStackCount()         ' How many stack elements?
    
    If c = 0 Then               '
        sPop = Empty            ' No elements? Then the request is Empty
        Exit Function           '
    End If                      '
    
    sPop = mvarStack(c - 1)     ' Read the item from the stack
    
    If c > DEFAULT_REDIM Then
        ReDim Preserve mvarStack(c)     'Free up some memory
    End If
    
    mvarStack(c - 1) = Empty    ' Delete the item from the stack
    
    ' Discussion present below
End Function


' ===========================================================================
' ===========================================================================


Public Function sRead(Optional StackItemNumber As Integer = 0) As Variant
    '
    ' Reads a specified value from the stack.  The item number is NOT an
    ' index, so the first stack value is 1, not 0.
    '
    c = GetStackCount()                     ' How many stack elements?
    
    If c = 0 Or StackItemNumber > c Then    ' If there are no elements in
        sRead = Empty                       ' the stack, or they ask for an
        Exit Function                       ' out of range index, then we
    End If                                  ' return EMPTY (duh)
    
    If StackItemNumber <= 0 Then                ' If the requested index <=
        sRead = mvarStack(c - 1)                ' 0, then return the last
    Else                                        ' item from the stack.  sRead
        sRead = mvarStack(StackItemNumber - 1)  ' with no args will return the
    End If                                      ' last item.
       
    ' Discussion present below
End Function



' ===========================================================================
' ===========================================================================



Public Sub ClearStack(Optional RedimValue As Integer = DEFAULT_REDIM)
    '
    ' Deletes all elements of a stack, and resets it to 10 values,
    ' unless the programmer specifies otherwise
    '
    
    ReDim mvarStack(RedimValue)
                            ' ReDim the stack.  This is supposed to
                            ' empty the array automatically, but I
                            ' don't trust it.
                            
    For i = 0 To UBound(mvarStack)  ' Make sure the array is empty. At
        mvarStack(i) = Empty        ' the expense of a little speed, we
    Next i                          ' gain a little confidence.
    
End Sub




' ===========================================================================
' ===========================================================================




Public Property Get StackCount() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StackCount
    StackCount = GetStackCount
End Property


Private Function GetStackCount() As Integer

    For i = LBound(mvarStack) To UBound(mvarStack)
        If mvarStack(i) <> Empty Then
            c = c + 1
        End If
    Next i
    
    GetStackCount = c
End Function










' ===========================================================================
' ===========================================================================
'
' DISCUSSIONS
'
' -----------------------------------------------------------------
'       General
' ---------------------
'
'       The stack array holds Variant values.  This means that
'       even an empty array will occupy some memory. The LEN
'       function doesn't show, but extra memory is used by the
'       compiler to show the VarType and other assorted stuff.
'
'       If you expect to store many items on the stack, consider
'       setting the constants to higher values.  An initially high
'       stack limit means more time before the ReDim's start.  A
'       higher ReDim value means you will have to ReDim less often.
'
'       When you PUSH an item to the stack, and the stack is full,
'       a ReDim is issued to add 1 or more spaces to the stack.  I
'       have chosen 3 as the number of spaces to add, but you may
'       choose to set a higher value.
'
'       The stack indexes have been accessed using integer variables,
'       so if you need more than 32767 stack items, then you will have
'       to go through this class and change the Integers to Long.  I
'       don't see where you would need to do this, unless you are doing
'       incredibly large binary file processing or something.
'
'















'
'
'   UNUSED PROCEDURES.  Here for your adjustment.  They are not ready
'   to go.  You must edit them yourself.  They may not work as is.
'
'   They were included as a possible version 2 item.
'
'
'










' UNUSED PROCEDURE ---- Here for future consideration
Private Function stk_GetItem(varID As Integer, Optional varRemove As Boolean = False) As Variant
        
        'r = -1
        'For i = 0 To UBound(mvarID)
        '       If mvarID(i) = varID Then
        '           r = varID
        '           Exit For
        '       End If
        'Next i
                        
        If varID < 0 Or varID > UBound(mvarStack) Then
            stk_GetItem = Nothing
        Else
            stk_GetItem = mvarStack(varID)
            
            If varRemove Then
                mvarStack(varID) = Empty
                'mvarID(r) = Empty
                stk_PackStack
            End If
        End If
        
End Function






' UNUSED PROCEDURE ---- Here for future consideration
Private Sub stk_PutItem(varID As Integer, varValue As Variant, Optional varInsert As Boolean = False)
        
                    
        If varID < 0 Then Exit Sub
        
        
        u = UBound(mvarStack)
        If varID > u Then
                ReDim Preserve mvarStack(varID + 3)
                mvarStack(varID) = varValue
        Else
            If varInsert Then
                If mvarStack(u) <> Empty Then ReDim Preserve mvarStack(u + 3)
                               
                For i = (UBound(mvarStack) - 1) To varID Step -1
                        mvarStack(i + 1) = mvarStack(i)
                Next i
            End If
            
            mvarStack(varID) = varValue
        End If
        
End Sub









' UNUSED PROCEDURE ---- Here for future consideration
Private Sub stk_PackStack()
        
        Do
            
            GoSub Pack_Message
        
            j = -1
            For i = 0 To UBound(mvarStack)
                If j < 0 Then
                        If mvarStack(i) = Empty Then
                            j = i
                        End If
                Else
                        If Not mvarStack(i) = Empty Then
                            mvarStack(j) = mvarStack(i)
                            mvarStack(i) = Empty
                            'mvarID(j) = mvarID(i)
                            'mvarID(i) = Empty
                        End If
                        j = i
                End If
            Next i
        Loop While j >= 0
        
Exit Sub




Pack_Message:
        '
        ' FOR DEBUG PURPOSES ONLY
        '
        m$ = ""
        For k = 0 To UBound(mvarStack)
            m$ = m$ & Trim(Str(k)) & "    " & mvarStack(k) & vbCrLf
        Next k
        MsgBox m$
        Return

End Sub

