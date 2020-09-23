Attribute VB_Name = "basErrorFilter"
Option Explicit

Public PreviousExceptionHandler As Long


' This function receives an exception code value and returns the
'  text description of the exception.
'
Public Function GetExceptionText(ByVal ExceptionCode As Long) As String
   Dim strExceptionString As String
   
   Select Case ExceptionCode
      Case EXCEPTION_ACCESS_VIOLATION
         strExceptionString = "Access Violation"
      
      Case EXCEPTION_DATATYPE_MISALIGNMENT
         strExceptionString = "Data Type Misalignment"
      
      Case EXCEPTION_BREAKPOINT
         strExceptionString = "Breakpoint"
      
      Case EXCEPTION_SINGLE_STEP
         strExceptionString = "Single Step"
      
      Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
         strExceptionString = "Array Bounds Exceeded"
      
      Case EXCEPTION_FLT_DENORMAL_OPERAND
         strExceptionString = "Float Denormal Operand"
      
      Case EXCEPTION_FLT_DIVIDE_BY_ZERO
         strExceptionString = "Divide By Zero"
      
      Case EXCEPTION_FLT_INEXACT_RESULT
         strExceptionString = "Floating Point Inexact Result"
      
      Case EXCEPTION_FLT_INVALID_OPERATION
         strExceptionString = "Invalid Operation"
      
      Case EXCEPTION_FLT_OVERFLOW
         strExceptionString = "Float Overflow"
      
      Case EXCEPTION_FLT_STACK_CHECK
         strExceptionString = "Float Stack Check"
      
      Case EXCEPTION_FLT_UNDERFLOW
         strExceptionString = "Float Underflow"
      
      Case EXCEPTION_INT_DIVIDE_BY_ZERO
         strExceptionString = "Integer Divide By Zero"
      
      Case EXCEPTION_INT_OVERFLOW
         strExceptionString = "Integer Overflow"
      
      Case EXCEPTION_PRIVILEGED_INSTRUCTION
         strExceptionString = "Privileged Instruction"
      
      Case EXCEPTION_IN_PAGE_ERROR
         strExceptionString = "In Page Error"
      
      Case EXCEPTION_ILLEGAL_INSTRUCTION
         strExceptionString = "Illegal Instruction"
      
      Case EXCEPTION_NONCONTINUABLE_EXCEPTION
         strExceptionString = "Non Continuable Exception"
      
      Case EXCEPTION_STACK_OVERFLOW
         strExceptionString = "Stack Overflow"
      
      Case EXCEPTION_INVALID_DISPOSITION
         strExceptionString = "Invalid Disposition"
      
      Case EXCEPTION_GUARD_PAGE_VIOLATION
         strExceptionString = "Guard Page Violation"
      
      Case EXCEPTION_INVALID_HANDLE
         strExceptionString = "Invalid Handle"
      
      Case EXCEPTION_CONTROL_C_EXIT
         strExceptionString = "Control-C Exit"
      
      Case Else
         strExceptionString = "Unknown (&H" & Right("00000000" & Hex(ExceptionCode), 8) & ")"
   
   End Select
   
   GetExceptionText = strExceptionString
End Function


' This function will be called when an unhandled exception occurs.
'  It raises an error so that it can be trapped with an ON ERROR statement
'  in the procedure that caused the exception.
'
Public Function ExceptionFilter( _
      ByRef ExceptionPtrs As EXCEPTION_POINTERS _
   ) As Long

   Dim Rec As EXCEPTION_RECORD
   Dim strException As String
   
   ' Get the current exception record.
   '
   Rec = ExceptionPtrs.pExceptionRecord
   
   ' If Rec.pExceptionRecord is not zero, then it is a nested exception and
   '  Rec.pExceptionRecord points to another EXCEPTION_RECORD structure.  Follow
   '  the pointers back to the original exception.
   '
   Do Until Rec.pExceptionRecord = 0
     ' A friendly declaration of CopyMemory.
     '
     CopyExceptionRecord Rec, Rec.pExceptionRecord, Len(Rec)
   Loop
   
   ' Translate the exception code into a user-friendly string.
   '
   strException = GetExceptionText(Rec.ExceptionCode)
   
   ' Raise an error to return control to the calling procedure.
   '
   Err.Raise 10000, "Exception", strException
    
End Function

