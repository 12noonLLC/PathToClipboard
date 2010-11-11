using System;
using System.Collections.Generic;
using System.Text;

//using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;


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
		VERB =			GCS.VERBA,
		HELPTEXT =		GCS.HELPTEXTA,
		VALIDATE =		GCS.VALIDATEA
	}

	[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
	public struct INVOKECOMMANDINFO
	{
		//NOTE: When SEE_MASK_HMONITOR is set, hIcon is treated as hMonitor
		public UInt32 cbSize;						// sizeof(CMINVOKECOMMANDINFO)
		public UInt32 fMask;						// any combination of CMIC_MASK_*
		public IntPtr hwnd;						// might be NULL (indicating no owner window)
		public UInt32 verb;
		[MarshalAs(UnmanagedType.LPStr)]
		public string parameters;				// might be NULL (indicating no parameter)
		[MarshalAs(UnmanagedType.LPStr)]
		public string directory;				// might be NULL (indicating no specific directory)
		public Int32 nShow;						// one of SW_ values for ShowWindow() API
		public UInt32 dwHotKey;
		public IntPtr hIcon;
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
