Imports System.Runtime.InteropServices
Module ModuleDuplex
    'Dim rtis As inter
    Public Structure PRINTER_INFO_9
        Dim pDevmode As Long '''' POINTER TO DEVMODE
    End Structure

    Public Structure DEVMODE
        Dim dmDeviceName As String
        Dim dmSpecVersion As Integer
        Dim dmDriverVersion As Integer
        Dim dmSize As Integer
        Dim dmDriverExtra As Integer
        Dim dmFields As Long
        Dim dmOrientation As Integer
        Dim dmPaperSize As Integer
        Dim dmPaperLength As Integer
        Dim dmPaperWidth As Integer
        Dim dmScale As Integer
        Dim dmCopies As Integer
        Dim dmDefaultSource As Integer
        Dim dmPrintQuality As Integer
        Dim dmColor As Integer
        Dim dmDuplex As Integer
        Dim dmYResolution As Integer
        Dim dmTTOption As Integer
        Dim dmCollate As Integer
        Dim dmFormName As String
        Dim dmUnusedPadding As Integer
        Dim dmBitsPerPel As Integer
        Dim dmPelsWidth As Long
        Dim dmPelsHeight As Long
        Dim dmDisplayFlags As Long
        Dim dmDisplayFrequency As Long
        Dim dmICMMethod As Long
        Dim dmICMIntent As Long
        Dim dmMediaType As Long
        Dim dmDitherType As Long
        Dim dmReserved1 As Long
        Dim dmReserved2 As Long
    End Structure

    Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Long) As Long
    Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, buffer As Long, ByVal pbSize As Long, pbSizeNeeded As Long) As Long
    Public Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Object, ByVal Command As Long) As Long
    Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String,
                                                                                            ByVal pDevModeOutput As Long, ByVal pDevModeInput As DEVMODE,
                                                                                            ByVal fMode As Long) As Long
    Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Object, pSource As Object, ByVal cbLength As Long)
    Public Const DM_IN_BUFFER = 8
    Public Const DM_OUT_BUFFER = 2

    'Public Function VarPtr(ByVal e As Object) As Integer
    '    Dim GC As rtis
    'End Function

    Public Sub SetPrinterProperty(ByVal sPrinterName As String, ByVal iPropertyType As Long)
        Dim PrinterName, sPrinter, sDefaultPrinter As String
        Dim Pinfo9 As PRINTER_INFO_9
        Dim hPrinter, nRet As Long
        Dim yDevModeData() As Byte
        Dim dm As DEVMODE

        '''' STROE THE CURRENT DEFAULT PRINTER
        sDefaultPrinter = sPrinterName

        '''' USE THE FULL PRINTER ADDRESS TO GET THE ADDRESS AND NAME MINUS THE PORT NAME
        PrinterName = "Brother HL-L2300D series"

        '''' OPEN THE PRINTER
        nRet = OpenPrinter(PrinterName, hPrinter, 0)

        '''' GET THE SIZE OF THE CURRENT DEVMODE STRUCTURE
        'nRet = DocumentProperties(0, hPrinter, PrinterName, 0, 0, 0)
        If (nRet < 0) Then MsgBox("Cannot get the size of the DEVMODE structure.") : Exit Sub

        '''' GET THE CURRENT DEVMODE STRUCTURE
        ReDim yDevModeData(nRet + 100)

        'nRet = DocumentProperties(0, hPrinter, PrinterName, Array.IndexOf(yDevModeData, 0), 0, DM_OUT_BUFFER)
        If (nRet < 0) Then MsgBox("Cannot get the DEVMODE structure.") : Exit Sub

        '''' COPY THE CURRENT DEVMODE STRUCTURE
        Call CopyMemory(dm, yDevModeData(0), Len(dm))

        '''' CHANGE THE DEVMODE STRUCTURE TO REQUIRED
        dm.dmDuplex = iPropertyType ' 1 = simplex, 2 = duplex

        '''' REPLACE THE CURRENT DEVMODE STRUCTURE WITH THE NEWLEY EDITED
        Call CopyMemory(yDevModeData(0), dm, Len(dm))

        '''' VERIFY THE NEW DEVMODE STRUCTURE
        'nRet = DocumentProperties(0, hPrinter, PrinterName, Array.IndexOf(yDevModeData, (0)), Array.IndexOf(yDevModeData, (0)), DM_IN_BUFFER Or DM_OUT_BUFFER)

        Pinfo9.pDevmode = Array.IndexOf(yDevModeData, (0))

        '''' SET THE DEMODE STRUCTURE WITH ANY CHANGES MADE
        nRet = SetPrinter(hPrinter, 9, Pinfo9, 0)
        If (nRet <= 0) Then MsgBox("Cannot set the DEVMODE structure.") : Exit Sub

        '''' CLOSE THE PRINTER
        nRet = ClosePrinter(hPrinter)

        '''' GET THE FULL PRINTER NAME FOR THE CUTE PDF WRITER
        sPrinter = "Brother HL-L2300D series"

        '''' CHECK TO MAKE SURE THE CUTEPDF WAS FOUND
        If sPrinter <> vbNullString Then
            '''' THIS FORCES EXCEL TO ACCEPT THE NEW CHANGES THAT HAVE BEEN MADE TO THE PRINTER SETTINGS
            '''' SET THE ACTIVE PRINTER TEMPERARILLY TO THE CUTE PDF WRITER
            'Application.ActivePrinter = sPrinter
            ''''' SET THE PRINTER BACK TO THE DEFAULY FOLLOW ME.
            'Application.ActivePrinter = sDefaultPrinter
        End If
    End Sub

    Public Sub SetDuplex(ByVal sPrinterName As String, iDuplex As Long)
        SetPrinterProperty(sPrinterName, iDuplex)
    End Sub
    Public Sub SetSimplex(ByVal sPrinterName As String, iDuplex As Long)
        SetPrinterProperty(sPrinterName, iDuplex)
    End Sub

End Module
