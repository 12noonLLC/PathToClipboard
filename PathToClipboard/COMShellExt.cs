using System;
using System.Runtime.InteropServices;
using System.Text;


namespace MyCOMDefinitions
{
	public enum CMF : uint
	{
		CMF_NORMAL		= 0x00000000,
		CMF_DEFAULTONLY	= 0x00000001,
		CMF_VERBSONLY	= 0x00000002,
		CMF_EXPLORE		= 0x00000004,
		CMF_NOVERBS		= 0x00000008,
		CMF_CANRENAME	= 0x00000010,
		CMF_NODEFAULT	= 0x00000020,
		CMF_INCLUDESTATIC= 0x00000040,
		CMF_RESERVED	= 0xffff0000      // View specific
	}

	// GetCommandString uFlags
	public enum GCS: uint
	{
		VERBA =			0x00000000,     // canonical verb
		HELPTEXTA =		0x00000001,     // help text (for status bar)
		VALIDATEA =		0x00000002,     // validate command exists
		VERBW =			0x00000004,     // canonical verb (unicode)
		HELPTEXTW =		0x00000005,     // help text (unicode version)
		VALIDATEW =		0x00000006,     // validate command exists (unicode)
		UNICODE =		0x00000004,     // for bit testing - Unicode string
		VERB = VERBA,
		HELPTEXT = HELPTEXTA,
		VALIDATE = VALIDATEA
	}

	[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
	internal struct CMINVOKECOMMANDINFO
	{
		//NOTE: When SEE_MASK_HMONITOR is set, hIcon is treated as hMonitor
		public uint cbSize;						// sizeof(CMINVOKECOMMANDINFO)
		public CMIC fMask;						// any combination of CMIC_MASK_*
		public IntPtr hwnd;						// might be NULL (indicating no owner window)
		public IntPtr lpVerb;
		[MarshalAs(UnmanagedType.LPStr)]
		public string parameters;				// might be NULL (indicating no parameter)
		[MarshalAs(UnmanagedType.LPStr)]
		public string directory;				// might be NULL (indicating no specific directory)
		public int nShow;						// one of SW_ values for ShowWindow() API
		public uint dwHotKey;
		public IntPtr hIcon;
	}

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	internal struct CMINVOKECOMMANDINFOEX
	{
		public uint cbSize;
		public CMIC fMask;
		public IntPtr hwnd;
		[MarshalAs(UnmanagedType.LPStr)]
		public string verb;
		[MarshalAs(UnmanagedType.LPStr)]
		public string parameters;
		[MarshalAs(UnmanagedType.LPStr)]
		public string directory;
		public int nShow;
		public uint dwHotKey;
		public IntPtr hIcon;
		[MarshalAs(UnmanagedType.LPStr)]
		public string title;
		public IntPtr lpVerbW;
		public string parametersW;
		public string directoryW;
		public string titleW;
		POINT ptInvoke;
	}

	[Flags]
	internal enum CMIC : uint
	{
		CMIC_MASK_ICON = 0x00000010,
		CMIC_MASK_HOTKEY = 0x00000020,
		CMIC_MASK_NOASYNC = 0x00000100,
		CMIC_MASK_FLAG_NO_UI = 0x00000400,
		CMIC_MASK_UNICODE = 0x00004000,
		CMIC_MASK_NO_CONSOLE = 0x00008000,
		CMIC_MASK_ASYNCOK = 0x00100000,
		CMIC_MASK_NOZONECHECKS = 0x00800000,
		CMIC_MASK_FLAG_LOG_USAGE = 0x04000000,
		CMIC_MASK_SHIFT_DOWN = 0x10000000,
		CMIC_MASK_PTINVOKE = 0x20000000,
		CMIC_MASK_CONTROL_DOWN = 0x40000000
	}

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	public struct POINT
	{
		public int X;
		public int Y;
	}


	[ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("000214e8-0000-0000-c000-000000000046")]
	public interface IShellExtInit
	{
		void Initialize(
			 IntPtr /*LPCITEMIDLIST*/ pidlFolder,
			 IntPtr /*LPDATAOBJECT*/ pDataObj,
			 IntPtr /*HKEY*/ hKeyProgID);
	}


	[ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("000214e4-0000-0000-c000-000000000046")]
	public interface IContextMenu
	{
		/*
		 * The PreserveSig attribute indicates that the HRESULT or retval signature transformation that takes
		 * place during COM interop calls should be suppressed. When you do not apply PreserveSigAttribute
		 * (e.g. the GetCommandString method of IContextMenu), the failure HRESULT of the method needs to be
		 * thrown as a .NET exception. For example, Marshal.ThrowExceptionForHR(WinError.E_FAIL); When you
		 * apply the PreserveSigAttribute to a managed method signature, the managed and unmanaged signatures
		 * of the attributed method are identical (e.g. the QueryContextMenu method of IContextMenu).
		 * Preserving the original method signature is necessary if the member returns more than one success
		 * HRESULT value and you want to detect the different values.
		 * http://blogs.msdn.com/b/codefx/archive/2010/09/14/writing-windows-shell-extension-with-net-framework-4-c-vb-net-part-1.aspx
		 */
		[PreserveSig]
		int QueryContextMenu(
				IntPtr /*HMENU*/ hMenu,
				uint iMenu,
				uint idCmdFirst,
				uint idCmdLast,
				uint uFlags);

		void InvokeCommand(IntPtr pici);

		void GetCommandString(
				UIntPtr idCmd,
				uint uFlags,
				IntPtr pReserved,
				StringBuilder pszName,
				uint cchMax);

//		[MarshalAs(UnmanagedType.LPStr)]
	}
}
