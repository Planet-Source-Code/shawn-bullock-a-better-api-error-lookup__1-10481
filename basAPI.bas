Attribute VB_Name = "basAPI"
Option Explicit


Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
    
' This API function installs the custom exception handler
'
Declare Function SetUnhandledExceptionFilter Lib "kernel32" ( _
      ByVal lpTopLevelExceptionFilter As Long _
   ) As Long

' This API function is used to raise exceptions
'
Declare Sub RaiseException Lib "kernel32" ( _
      ByVal dwExceptionCode As Long, _
      ByVal dwExceptionFlags As Long, _
      ByVal nNumberOfArguments As Long, _
      lpArguments As Long _
   )

' Possible return values for the Unhandled Exception Filter
'
Public Const EXCEPTION_CONTINUE_EXECUTION = -1
Public Const EXCEPTION_CONTINUE_SEARCH = 0
Public Const EXCEPTION_EXECUTE_HANDLER = 1

' Maximum number of parameters an Exception_Record can have
'
Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15

' Structure that contains processor-specific register data
'
Type CONTEXT
   FltF0 As Double
   FltF1 As Double
   FltF2 As Double
   FltF3 As Double
   FltF4 As Double
   FltF5 As Double
   FltF6 As Double
   FltF7 As Double
   FltF8 As Double
   FltF9 As Double
   FltF10 As Double
   FltF11 As Double
   FltF12 As Double
   FltF13 As Double
   FltF14 As Double
   FltF15 As Double
   FltF16 As Double
   FltF17 As Double
   FltF18 As Double
   FltF19 As Double
   FltF20 As Double
   FltF21 As Double
   FltF22 As Double
   FltF23 As Double
   FltF24 As Double
   FltF25 As Double
   FltF26 As Double
   FltF27 As Double
   FltF28 As Double
   FltF29 As Double
   FltF30 As Double
   FltF31 As Double
   
   IntV0 As Double
   IntT0 As Double
   IntT1 As Double
   IntT2 As Double
   IntT3 As Double
   IntT4 As Double
   IntT5 As Double
   IntT6 As Double
   IntT7 As Double
   IntS0 As Double
   IntS1 As Double
   IntS2 As Double
   IntS3 As Double
   IntS4 As Double
   IntS5 As Double
   IntFp As Double
   IntA0 As Double
   IntA1 As Double
   IntA2 As Double
   IntA3 As Double
   IntA4 As Double
   IntA5 As Double
   IntT8 As Double
   IntT9 As Double
   IntT10 As Double
   IntT11 As Double
   IntRa As Double
   IntT12 As Double
   IntAt As Double
   IntGp As Double
   IntSp As Double
   IntZero As Double
   
   Fpcr As Double
   SoftFpcr As Double
   
   Fir As Double
   Psr As Long
   
   ContextFlags As Long
   Fill(4) As Long
End Type

' Structure that describes an exception
'
Type EXCEPTION_RECORD
   ExceptionCode As Long
   ExceptionFlags As Long
   pExceptionRecord As Long  ' Pointer to an EXCEPTION_RECORD structure
   ExceptionAddress As Long
   NumberParameters As Long
   ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type

' Structure that contains exception information that can be used by a debugger
'
Type EXCEPTION_DEBUG_INFO
   pExceptionRecord As EXCEPTION_RECORD
   dwFirstChance As Long
End Type

' The EXCEPTION_POINTERS structure contains an exception record with a
'  machine-independent description of an exception and a context record
'  with a machine-dependent description of the processor context at the
'  time of the exception
'
Type EXCEPTION_POINTERS
   pExceptionRecord As EXCEPTION_RECORD
   ContextRecord As CONTEXT
End Type

' Standard Exception Codes
'
Public Const EXCEPTION_ACCESS_VIOLATION             As Long = &HC0000005
Public Const EXCEPTION_DATATYPE_MISALIGNMENT        As Long = &H80000002
Public Const EXCEPTION_BREAKPOINT                   As Long = &H80000003
Public Const EXCEPTION_SINGLE_STEP                  As Long = &H80000004
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED        As Long = &HC000008C
Public Const EXCEPTION_FLT_DENORMAL_OPERAND         As Long = &HC000008D
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO           As Long = &HC000008E
Public Const EXCEPTION_FLT_INEXACT_RESULT           As Long = &HC000008F
Public Const EXCEPTION_FLT_INVALID_OPERATION        As Long = &HC0000090
Public Const EXCEPTION_FLT_OVERFLOW                 As Long = &HC0000091
Public Const EXCEPTION_FLT_STACK_CHECK              As Long = &HC0000092
Public Const EXCEPTION_FLT_UNDERFLOW                As Long = &HC0000093
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO           As Long = &HC0000094
Public Const EXCEPTION_INT_OVERFLOW                 As Long = &HC0000095
Public Const EXCEPTION_PRIVILEGED_INSTRUCTION       As Long = &HC0000096
Public Const EXCEPTION_IN_PAGE_ERROR                As Long = &HC0000006
Public Const EXCEPTION_ILLEGAL_INSTRUCTION          As Long = &HC000001D
Public Const EXCEPTION_NONCONTINUABLE_EXCEPTION     As Long = &HC0000025
Public Const EXCEPTION_STACK_OVERFLOW               As Long = &HC00000FD
Public Const EXCEPTION_INVALID_DISPOSITION          As Long = &HC0000026
Public Const EXCEPTION_GUARD_PAGE_VIOLATION         As Long = &H80000001
Public Const EXCEPTION_INVALID_HANDLE               As Long = &HC0000008
Public Const EXCEPTION_CONTROL_C_EXIT               As Long = &HC000013A

' This is a friendly declaration of the CopyMemory function.  It is used to copy
'  data into an EXTENSION_RECORD structure from a pointer to another structure.
'
Declare Sub CopyExceptionRecord Lib "kernel32" Alias "RtlMoveMemory" ( _
      pDest As EXCEPTION_RECORD, _
      ByVal LPEXCEPTION_RECORD As Long, _
      ByVal lngBytes As Long _
   )

