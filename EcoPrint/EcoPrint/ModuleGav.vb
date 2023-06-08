Imports System.Drawing.Printing
Imports System.Runtime.InteropServices

Module ModuleGav



    <DllImport("Winspool.drv", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Function DeviceCapabilities(lpDeviceName As String, lpPort As String, iIndex As Integer, lpOutput As IntPtr, lpDevMode As IntPtr) As Integer
    End Function

    <DllImport("Winspool.drv", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Function OpenPrinter(pPrinterName As String, ByRef hPrinter As IntPtr, ByRef pDefault As PRINTER_DEFAULTS) As Boolean
    End Function

    <DllImport("Winspool.drv", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Function ClosePrinter(hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("Winspool.drv", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Function GetPrinter(hPrinter As IntPtr, Level As Integer, pPrinter As IntPtr, cbBuf As Integer, ByRef pcbNeeded As Integer) As Boolean
    End Function

    <DllImport("Winspool.drv", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Private Function SetPrinter(hPrinter As IntPtr, Level As Integer, pPrinter As IntPtr, Command As Integer) As Boolean
    End Function

    <DllImport("winspool.drv", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Function SetDefaultPrinter(Name As String) As Boolean
    End Function

    Public Const ERROR_INSUFFICIENT_BUFFER = 122

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PRINTER_INFO_2
        <MarshalAs(UnmanagedType.LPTStr)> Public pServerName As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pPrinterName As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pShareName As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pPortName As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pDriverName As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pComment As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pLocation As String

        Public pDevMode As IntPtr

        <MarshalAs(UnmanagedType.LPTStr)> Public pSepFile As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pPrintProcessor As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pDatatype As String
        <MarshalAs(UnmanagedType.LPTStr)> Public pParameters As String

        Public pSecurityDescriptor As IntPtr
        Public Attributes As Integer
        Public Priority As Integer
        Public DefaultPriority As Integer
        Public StartTime As Integer
        Public UntilTime As Integer
        Public Status As Integer
        Public cJobs As Integer
        Public AveragePPM As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure DEVMODE
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public pDeviceName As String
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
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)> Public dmFormName As String
        Public dmUnusedPadding As Short
        Public dmBitsPerPel As Integer
        Public dmPelsWidth As Integer
        Public dmPelsHeight As Integer
        Public dmNup As Integer
        Public dmDisplayFrequency As Integer
        Public dmICMMethod As Integer
        Public dmICMIntent As Integer
        Public dmMediaType As Integer
        Public dmDitherType As Integer
        Public dmReserved1 As Integer
        Public dmReserved2 As Integer
        Public dmPanningWidth As Integer
        Public dmPanningHeight As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure PRINTER_DEFAULTS
        Public pDatatype As IntPtr
        Public pDevMode As IntPtr
        Public DesiredAccess As Integer
    End Structure
    Dim STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
    Dim PRINTER_ACCESS_ADMINISTER As Integer = &H4
    Dim PRINTER_ACCESS_USE As Integer = &H8
    Dim PRINTER_ALL_ACCESS As Integer = (STANDARD_RIGHTS_REQUIRED Or (PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE))

    Public Const DC_PAPERS = 2
    Public Const DC_PAPERSIZE = 3
    Public Const DC_PAPERNAMES = 16
    Public Const DC_BINNAMES = 12
    Public Const DC_BINS = 6
    Public Const DEFAULT_VALUES = 0

    Public Const DM_ORIENTATION = &H1
    Public Const DM_PAPERSIZE = &H2
    Public Const DM_DUPLEX = &H1000

    <DllImport("Winspool.drv", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Function DocumentProperties(hwnd As IntPtr, hPrinter As IntPtr, pDeviceName As String, pDevModeOutput As IntPtr, pDevModeInput As IntPtr, fMode As Integer) As Integer
    End Function

    Public Const DM_UPDATE = 1
    Public Const DM_COPY = 2
    Public Const DM_PROMPT = 4
    Public Const DM_MODIFY = 8

    Public Const DM_IN_BUFFER = DM_MODIFY
    Public Const DM_IN_PROMPT = DM_PROMPT
    Public Const DM_OUT_BUFFER = DM_COPY
    Public Const DM_OUT_DEFAULT = DM_UPDATE

    Public Const HWND_BROADCAST As Integer = (&HFFFF)
    Public Const WM_DEVMODECHANGE As Integer = &H1B

    Public Const SMTO_NORMAL As Integer = &H0
    Public Const SMTO_BLOCK As Integer = &H1
    Public Const SMTO_ABORTIFHUNG As Integer = &H2
    Public Const SMTO_NOTIMEOUTIFNOTHUNG As Integer = &H8
    Public Const SMTO_ERRORONEXIT As Integer = &H20

    <DllImport("User32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Function SendMessageTimeout(hWnd As IntPtr, Msg As Integer, wParam As IntPtr, lParam As String, fuFlags As UInteger, uTimeout As UInteger, lpdwResult As IntPtr) As IntPtr
    End Function


    Private WithEvents PrintDocument1 As PrintDocument = New PrintDocument

    Public Function SetPrinterDef(ByVal PrinterName As Integer) As Boolean

        'SetDefaultPrinter(PrinterName)
        'Return PrinterName
    End Function


    Public Function GavPrinterSetting(ByVal PrinterName As String, ByVal MyDuplex As Integer) As Boolean
        Dim hPrinter As IntPtr
        Dim pDefaults As New PRINTER_DEFAULTS()
        Dim sPrinterName = PrinterSettings.InstalledPrinters(PrinterName)
        Dim bResult As Boolean = OpenPrinter(sPrinterName, hPrinter, pDefaults)
        If (bResult = True) Then
            Dim nBytesNeeded As Integer = -1
            Dim nSize As Integer = -1
            Dim pPrinterInfo As IntPtr = IntPtr.Zero
            GetPrinter(hPrinter, 2, IntPtr.Zero, 0, nBytesNeeded)
            pPrinterInfo = Marshal.AllocHGlobal(nBytesNeeded)
            nSize = nBytesNeeded
            If (Not GetPrinter(hPrinter, 2, pPrinterInfo, nSize, nBytesNeeded)) Then
                Marshal.ThrowExceptionForHR(Marshal.GetHRForLastWin32Error())
            End If
            Dim PrinterInfo As PRINTER_INFO_2 = CType(Marshal.PtrToStructure(pPrinterInfo, GetType(PRINTER_INFO_2)), PRINTER_INFO_2)
            Dim pDevMode As IntPtr = PrinterInfo.pDevMode
            Dim DevMode As DEVMODE = CType(Marshal.PtrToStructure(pDevMode, GetType(DEVMODE)), DEVMODE)
            DevMode.dmFields = DM_DUPLEX
            'DevMode.dmPaperSize = TryCast(ComboBox2.SelectedItem, ComboboxItem).PaperSize
            DevMode.dmDuplex = MyDuplex
            '&H1000

            Marshal.StructureToPtr(DevMode, pDevMode, False)
            Marshal.StructureToPtr(PrinterInfo, pPrinterInfo, False)
            Dim nReturn = DocumentProperties(IntPtr.Zero, hPrinter, sPrinterName, pDevMode, pDevMode, DM_IN_BUFFER Or DM_OUT_BUFFER)
            nReturn = Convert.ToInt32(SetPrinter(hPrinter, 2, pPrinterInfo, 2))
            SendMessageTimeout(New IntPtr(HWND_BROADCAST), WM_DEVMODECHANGE, IntPtr.Zero, sPrinterName, SMTO_NORMAL, 1000, IntPtr.Zero)
            Marshal.FreeHGlobal(pPrinterInfo)
            ClosePrinter(hPrinter)
        End If
    End Function

End Module
