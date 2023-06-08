Imports System.Runtime.InteropServices
Module Module1
    Private hPrinter As IntPtr = New System.IntPtr()
    Private PrinterValues As PRINTER_DEFAULTS = New PRINTER_DEFAULTS()
    Private pinfo As PRINTER_INFO_2 = New PRINTER_INFO_2()
    Private dm As DEVMODE
    Private ptrDM As IntPtr
    Private ptrPrinterInfo As IntPtr
    Private sizeOfDevMode As Integer = 0
    Private lastError As Integer
    Private nBytesNeeded As Integer
    Private nRet As Long
    Private intError As Integer
    Private nJunk As System.Int32
    Private yDevModeData As IntPtr
    <DllImport("kernel32.dll", EntryPoint:="GetLastError", SetLastError:=False, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Function GetLastError() As Int32

    End Function

    <DllImport("winspool.Drv", EntryPoint:="ClosePrinter", SetLastError:=True, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Function ClosePrinter(ByVal hPrinter As IntPtr) As Boolean

    End Function
    <DllImport("winspool.Drv", EntryPoint:="DocumentPropertiesA", SetLastError:=True, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Function DocumentProperties(ByVal hwnd As IntPtr, ByVal hPrinter As IntPtr,
    <MarshalAs(UnmanagedType.LPStr)> ByVal pDeviceNameg As String, ByVal pDevModeOutput As IntPtr, ByRef pDevModeInput As IntPtr, ByVal fMode As Integer) As Integer

    End Function
    <DllImport("winspool.Drv", EntryPoint:="GetPrinterA", SetLastError:=True, CharSet:=CharSet.Ansi, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Function GetPrinter(ByVal hPrinter As IntPtr, ByVal dwLevel As Int32, ByVal pPrinter As IntPtr, ByVal dwBuf As Int32, <Out> ByRef dwNeeded As Int32) As Boolean

    End Function
    <DllImport("winspool.Drv", EntryPoint:="OpenPrinterA", SetLastError:=True, CharSet:=CharSet.Ansi, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Private Function OpenPrinter(
    <MarshalAs(UnmanagedType.LPStr)> ByVal szPrinter As String, <Out> ByRef hPrinter As IntPtr, ByRef pd As PRINTER_DEFAULTS) As Boolean

    End Function
    <DllImport("winspool.drv", CharSet:=CharSet.Ansi, SetLastError:=True)>
    Private Function SetPrinter(ByVal hPrinter As IntPtr, ByVal Level As Integer, ByVal pPrinter As IntPtr, ByVal Command As Integer) As Boolean

    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PRINTER_DEFAULTS
        Public pDatatype As Integer
        Public pDevMode As Integer
        Public DesiredAccess As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure PRINTER_INFO_2
        <MarshalAs(UnmanagedType.LPStr)>
        Public pServerName As String
        <MarshalAs(UnmanagedType.LPStr)>
        Public pPrinterName As String
        <MarshalAs(UnmanagedType.LPStr)>
        Public pShareName As String
        <MarshalAs(UnmanagedType.LPStr)>
        Public pPortName As String
        <MarshalAs(UnmanagedType.LPStr)>
        Public pDriverName As String
        <MarshalAs(UnmanagedType.LPStr)>
        Public pComment As String
        <MarshalAs(UnmanagedType.LPStr)>
        Public pLocation As String
        Public pDevMode As IntPtr
        <MarshalAs(UnmanagedType.LPStr)>
        Public pSepFile As String
        <MarshalAs(UnmanagedType.LPStr)>
        Public pPrintProcessor As String
        <MarshalAs(UnmanagedType.LPStr)>
        Public pDatatype As String
        <MarshalAs(UnmanagedType.LPStr)>
        Public pParameters As String
        Public pSecurityDescriptor As IntPtr
        Public Attributes As Int32
        Public Priority As Int32
        Public DefaultPriority As Int32
        Public StartTime As Int32
        Public UntilTime As Int32
        Public Status As Int32
        Public cJobs As Int32
        Public AveragePPM As Int32
    End Structure

    Private Const CCDEVICENAME As Short = 32
    Private Const CCFORMNAME As Short = 32

    <StructLayout(LayoutKind.Sequential)>
    Public Structure DEVMODE
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CCDEVICENAME)>
        Public dmDeviceName As String
        Public dmSpecVersion As Short
        Public dmDriverVersion As Short
        Public dmSize As Short
        Public dmDriverExtra As Short
        Public dmFields As Integer
        Public dmOrientation As Short
        Public dmPaperSize As Short
        Public dmPaperLength As Short
        Public dmPaperWidth As Short
        Public dmScale As Short
        Public dmCopies As Short
        Public dmDefaultSource As Short
        Public dmPrintQuality As Short
        Public dmColor As Short
        Public dmDuplex As Short
        Public dmYResolution As Short
        Public dmTTOption As Short
        Public dmCollate As Short
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=CCFORMNAME)>
        Public dmFormName As String
        Public dmUnusedPadding As Short
        Public dmBitsPerPel As Short
        Public dmPelsWidth As Integer
        Public dmPelsHeight As Integer
        Public dmDisplayFlags As Integer
        Public dmDisplayFrequency As Integer
    End Structure

    Private Const DM_DUPLEX As Integer = &H1000
    Private Const DM_IN_BUFFER As Integer = 8
    Private Const DM_OUT_BUFFER As Integer = 2
    Private Const PRINTER_ACCESS_ADMINISTER As Integer = &H4
    Private Const PRINTER_ACCESS_USE As Integer = &H8
    Private Const STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
    Private Const PRINTER_ALL_ACCESS As Integer = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

    'Dim PrinterData As Byte

    Public Function ChangePrinterSetting(ByVal PrinterName As String, ByVal PS As Printing.PrinterSettings) As Boolean
        If (CInt(PS.Duplex) < 1) OrElse (CInt(PS.Duplex) > 3) Then
            Throw New ArgumentOutOfRangeException("nDuplexSetting", "nDuplexSetting is incorrect.")
        Else
            dm = GetPrinterSettings(PrinterName)
            'dm.dmDefaultSource = CShort(PS.source)
            'dm.dmOrientation = CShort(PS.Orientation)
            'dm.dmPaperSize = CShort(PS.Size)
            dm.dmDuplex = CShort(PS.Duplex)
            Marshal.StructureToPtr(dm, yDevModeData, True)
            pinfo.pDevMode = yDevModeData
            pinfo.pSecurityDescriptor = IntPtr.Zero
            Marshal.StructureToPtr(pinfo, ptrPrinterInfo, True)
            lastError = Marshal.GetLastWin32Error()
            nRet = Convert.ToInt16(SetPrinter(hPrinter, 2, ptrPrinterInfo, 0))

            If nRet = 0 Then
                lastError = Marshal.GetLastWin32Error()
                Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
            End If

            If hPrinter <> IntPtr.Zero Then ClosePrinter(hPrinter)
            Return Convert.ToBoolean(nRet)
        End If
    End Function

    Private Function GetPrinterSettings(ByVal PrinterName As String) As DEVMODE
        'Dim PData As PrinterData = New PrinterData()
        Dim dm As DEVMODE
        Const PRINTER_ACCESS_ADMINISTER As Integer = &H4
        Const PRINTER_ACCESS_USE As Integer = &H8
        Const PRINTER_ALL_ACCESS As Integer = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
        PrinterValues.pDatatype = 0
        PrinterValues.pDevMode = 0
        PrinterValues.DesiredAccess = PRINTER_ALL_ACCESS
        nRet = Convert.ToInt32(OpenPrinter(PrinterName, hPrinter, PrinterValues))

        If nRet = 0 Then
            lastError = Marshal.GetLastWin32Error()
            Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
        End If

        GetPrinter(hPrinter, 2, IntPtr.Zero, 0, nBytesNeeded)

        If nBytesNeeded <= 0 Then
            Throw New System.Exception("Unable to allocate memory")
        Else

            If True Then
                'Dim fo As ptrPrinterIn = Marshal.AllocCoTaskMem(nBytesNeeded)
            End If

            ptrPrinterInfo = Marshal.AllocHGlobal(nBytesNeeded)
            nRet = Convert.ToInt32(GetPrinter(hPrinter, 2, ptrPrinterInfo, nBytesNeeded, nJunk))

            If nRet = 0 Then
                lastError = Marshal.GetLastWin32Error()
                Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
            End If

            pinfo = CType(Marshal.PtrToStructure(ptrPrinterInfo, GetType(PRINTER_INFO_2)), PRINTER_INFO_2)
            Dim Temp As IntPtr = New IntPtr()

            If pinfo.pDevMode = IntPtr.Zero Then
                Dim ptrZero As IntPtr = IntPtr.Zero
                sizeOfDevMode = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, ptrZero, ptrZero, 0)
                ptrDM = Marshal.AllocCoTaskMem(sizeOfDevMode)
                Dim i As Integer
                i = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, ptrDM, ptrZero, DM_OUT_BUFFER)

                If (i < 0) OrElse (ptrDM = IntPtr.Zero) Then
                    Throw New System.Exception("Cannot get DEVMODE data")
                End If

                pinfo.pDevMode = ptrDM
            End If

            intError = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, IntPtr.Zero, Temp, 0)
            yDevModeData = Marshal.AllocHGlobal(intError)
            intError = DocumentProperties(IntPtr.Zero, hPrinter, PrinterName, yDevModeData, Temp, 2)
            dm = CType(Marshal.PtrToStructure(yDevModeData, GetType(DEVMODE)), DEVMODE)

            If (nRet = 0) OrElse (hPrinter = IntPtr.Zero) Then
                lastError = Marshal.GetLastWin32Error()
                Throw New System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error())
            End If

            Return dm
        End If
    End Function
End Module
