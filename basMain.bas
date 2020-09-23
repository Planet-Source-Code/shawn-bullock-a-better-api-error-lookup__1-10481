Attribute VB_Name = "basMain"
Option Explicit

Private PreviousExceptionHandler As Long


Public Sub Main()
   ' We are going to redirect all exception errors travel through our filter so that
   '  we can handle them within the confines of our own error trapping, as opposed to
   '  Windows crashing on us.  Some errors we have to exit gracefully on, but at
   '  least we aren't bringing down the whole system with us.  Because there might be
   '  a previous exceptionfilter, we record the address of it to PreviousExceptionFilter
   '  and restore it when we exit.  If there was none, it returns a 0 instead.  When we
   '  restore it, if it is 0, then it returns exception handling to Windows.
   '
   Let PreviousExceptionHandler = SetUnhandledExceptionFilter(AddressOf ExceptionFilter)
   
End Sub

Public Sub ExitApp()
    Call SetUnhandledExceptionFilter(PreviousExceptionHandler)
    End
End Sub
