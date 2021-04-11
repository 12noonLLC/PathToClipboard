using System;
using System.Text;

using System.Runtime.InteropServices;

namespace Win32Functions
{
	internal static class Constants
	{
		public const int MAX_FILE_LEN = 1024;

		public const int NOERROR = 0;
		public const int ERROR_MORE_DATA = 234;
		public const int ERROR_BAD_DEVICE = 1200;
		public const int ERROR_NO_NET_OR_BAD_PATH = 1203;
		public const int ERROR_NOT_CONNECTED = 2250;
		public const int STRSAFE_E_INSUFFICIENT_BUFFER = -2147024774;
	}

	#region Win32Enumerations

	public enum SEVERITY : uint
	{
		SUCCESS = 0,
		ERROR = 1
	}

	internal static class WinError
	{
		public const int S_OK = 0x0000;
		public const int S_FALSE = 0x0001;
		public const int E_FAIL = -2147467259;
		public const int E_INVALIDARG = -2147024809;
		public const int E_OUTOFMEMORY = -2147024882;
		public const int STRSAFE_E_INSUFFICIENT_BUFFER = -2147024774;

		public const uint SEVERITY_SUCCESS = 0;
		public const uint SEVERITY_ERROR = 1;

		/// <summary>
		/// Create an HRESULT value from component pieces.
		/// </summary>
		/// <param name="sev">The severity to be used</param>
		/// <param name="fac">The facility to be used</param>
		/// <param name="code">The error number</param>
		/// <returns>A HRESULT constructed from the above 3 values</returns>
//		public static int MAKE_HRESULT(uint sev, uint fac, uint code)
		public static UInt32 MAKE_HRESULT(SEVERITY sev, UInt32 fac, UInt32 code)
		{
			// (HRESULT) (((unsigned long)(sev)<<31) | ((unsigned long)(fac)<<16) | ((unsigned long)(code)))
			return ((UInt32)sev << 31) | (fac << 16) | code;
		}

	}

	public class Macros
	{
		public static int HighWord(int number)
		{
			return ((number & 0x80000000) == 0x80000000) ? (number >> 16) : ((number >> 16) & 0xffff);
		}

		public static int LowWord(int number)
		{
			return (number & 0xffff);
		}
	}


	// Make these constants 
	[Flags]
	internal enum MIIM : uint
	{
		STATE = 0x00000001,
		ID = 0x00000002,
		SUBMENU = 0x00000004,
		CHECKMARKS = 0x00000008,
		TYPE = 0x00000010,
		DATA = 0x00000020,
		STRING = 0x00000040,
		BITMAP = 0x00000080,
		FTYPE = 0x00000100
	}

	internal enum MFT : uint
	{
		STRING = 0x00000000,
		BITMAP = 0x00000004,
		MENUBARBREAK = 0x00000020,
		MENUBREAK = 0x00000040,
		OWNERDRAW = 0x00000100,
		RIGHTJUSTIFY = 0x00004000,
	}

	internal enum MFS : uint
	{
		ENABLED = 0x00000000,
		UNCHECKED = 0x00000000,
		UNHILITE = 0x00000000,
		GRAYED = 0x00000003,
		DISABLED = 0x00000003,
		CHECKED = 0x00000008,
		HILITE = 0x00000080,
		DEFAULT = 0x00001000,
	}

	// Use for AppendMenu and InsertMenu
	internal enum MF_WIN40 : uint
	{
		BYCOMMAND = 0x00000000,
		BYPOSITION = 0x00000400,

		INSERT = 0x00000000,
		CHANGE = 0x00000080,
		APPEND = 0x00000100,
		DELETE = 0x00000200,
		REMOVE = 0x00001000,

		SEPARATOR = 0x00000800,

		ENABLED = 0x00000000,
		GRAYED = 0x00000001,
		DISABLED = 0x00000002,

		UNCHECKED = 0x00000000,
		CHECKED = 0x00000008,
		USECHECKBITMAPS = 0x00000200,

		STRING = 0x00000000,
		BITMAP = 0x00000004,
		OWNERDRAW = 0x00000100,

		POPUP = 0x00000010,
		MENUBARBREAK = 0x00000020,
		MENUBREAK = 0x00000040,

		UNHILITE = 0x00000000,
		HILITE = 0x00000080,

		DEFAULT = 0x00001000,
		SYSMENU = 0x00002000,
		HELP = 0x00004000,
		RIGHTJUSTIFY = 0x00004000,
		MOUSESELECT = 0x00008000
	}

	internal enum CLIPFORMAT : uint
	{
		CF_TEXT = 1,
		CF_BITMAP = 2,
		CF_METAFILEPICT = 3,
		CF_SYLK = 4,
		CF_DIF = 5,
		CF_TIFF = 6,
		CF_OEMTEXT = 7,
		CF_DIB = 8,
		CF_PALETTE = 9,
		CF_PENDATA = 10,
		CF_RIFF = 11,
		CF_WAVE = 12,
		CF_UNICODETEXT = 13,
		CF_ENHMETAFILE = 14,
		CF_HDROP = 15,
		CF_LOCALE = 16,
		CF_MAX = 17,

		CF_OWNERDISPLAY = 0x0080,
		CF_DSPTEXT = 0x0081,
		CF_DSPBITMAP = 0x0082,
		CF_DSPMETAFILEPICT = 0x0083,
		CF_DSPENHMETAFILE = 0x008E,

		CF_PRIVATEFIRST = 0x0200,
		CF_PRIVATELAST = 0x02FF,

		CF_GDIOBJFIRST = 0x0300,
		CF_GDIOBJLAST = 0x03FF
	}

	internal enum NAME_INFO_LEVEL : uint
	{
		UNIVERSAL_NAME_INFO_LEVEL	= 0x000000001,
		REMOTE_NAME_INFO_LEVEL		= 0x000000002
	}

	#endregion

	#region Win32Structures

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
	internal struct MENUITEMINFO
	{
		public uint cbSize;
		public MIIM fMask;
		public MFT fType;
		public MFS fState;
		public uint wID;
		public IntPtr	/*HMENU*/	  hSubMenu;
		public IntPtr	/*HBITMAP*/   hbmpChecked;
		public IntPtr	/*HBITMAP*/	  hbmpUnchecked;
		public IntPtr	/*ULONG_PTR*/ dwItemData;
		public String dwTypeData;
		public uint cch;
		public IntPtr	/*HBITMAP*/ hbmpItem;

/*http://www.pinvoke.net/default.aspx/Structures/MENUITEMINFO.html
*/		// return the size of the structure
		internal static UInt32 SizeOf
		{
			// 48 bytes on Windows XP SP2
			// 80 bytes on Windows Server 2003 64-bit
			get { return (UInt32)Marshal.SizeOf(typeof(MENUITEMINFO)); }
		}
	}

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
	public struct UNIVERSAL_NAME_INFO
	{
		[MarshalAs(UnmanagedType.LPTStr)]
		public String lpUniversalName;
	}

	#endregion

	#region Win32 Imports

	public class Imports
	{
		//UNUSED
		//[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
		//internal static extern IntPtr lstrcpyn([Out] IntPtr pStringDest, IntPtr pStringSrc, UInt32 cchMax);


		[DllImport("shell32", CharSet = CharSet.Auto)]
		internal static extern uint DragQueryFile(IntPtr hDrop, uint iFile, StringBuilder buffer, uint cch);


		/// <summary>
		/// Free the specified storage medium.
		/// </summary>
		/// <param name="pmedium">
		/// Reference of the storage medium that is to be freed.
		/// </param>
		[DllImport("ole32.dll", CharSet = CharSet.Unicode)]
		public static extern void ReleaseStgMedium(ref System.Runtime.InteropServices.ComTypes.STGMEDIUM pmedium);


		[DllImport("user32")]
		internal static extern IntPtr CreatePopupMenu();

		[DllImport("user32")]
		internal static extern Int32 GetMenuItemCount(IntPtr hmenu);

		[DllImport("user32", CharSet = CharSet.Unicode)]
		internal static extern Boolean AppendMenu(IntPtr hmenu, MF_WIN40 uflags, IntPtr uIDNewItemOrSubmenu, string text);

		//superseded by InsertMenuItem() but can still use it
		[DllImport("user32", CharSet = CharSet.Unicode)]
		internal static extern Boolean InsertMenu(IntPtr hmenu, UInt32 position, MF_WIN40 uflags, IntPtr uIDNewItemOrSubmenu, string text);

// 12 Nov 2010: To get this to work, see shell context menu extension in http://1code.codeplex.com
//Can't get this to work (on 2003). Probably the string member of MENUITEMINFO. Maybe it should marshal as IntPtr.
//http://www.pinvoke.net/default.aspx/Interfaces/IContextMenu.html
//            /*
//             * explorer.exe Error: 0 : Exception: Cannot marshal field 'dwTypeData' of
//             * type 'Win32Functions.MENUITEMINFO': Struct or class fields cannot be of
//             * type StringBuilder. The same effect can usually be achieved by using a
//             * String field and preinitializing it to a string with length matching
//             * the length of the appropriate buffer.
//             */
//            Win32Functions.MENUITEMINFO menuitem = new Win32Functions.MENUITEMINFO();
//            menuitem.cbSize = Win32Functions.MENUITEMINFO.SizeOf;
//            menuitem.fMask = (uint)(Win32Functions.MIIM.ID | Win32Functions.MIIM.TYPE | Win32Functions.MIIM.STATE | Win32Functions.MIIM.STRING);
//            menuitem.wID = idCmdNext;
//            menuitem.fType = (uint)Win32Functions.MF.STRING;
//            menuitem.fState = (uint)Win32Functions.MF.ENABLED;
//            menuitem.dwTypeData = "TEST InsertMenuItem() " + menuitem.wID;
//            brc = Win32Functions.Imports.InsertMenuItem(hmenu, ixMenu, true, ref menuitem);
//            if (!brc)
//            {
//               int err = Marshal.GetLastWin32Error();
//               System.Diagnostics.Debug.WriteLine("InsertMenuItem failed = " + ((err == 87) ? "Incorrect param" : err.ToString()));
//            }
//            ++idCmdNext;
//            ++ixMenu;
		[DllImport("user32", CharSet = CharSet.Unicode, SetLastError = true)]
		internal static extern Boolean InsertMenuItem(IntPtr hmenu,
																	uint iItem,
																	[MarshalAs(UnmanagedType.Bool)]bool bByPosition,
																	ref MENUITEMINFO mii);


/*
.method private hidebysig static pinvokeimpl("mpr.dll" winapi) 
		  int32  marshal( unsigned int32)  WNetGetUniversalName(string lpLocalPath,
																				  int32  marshal( unsigned int32) dwInfoLevel,
																				  native int lpBuffer,
																				  int32&  marshal( unsigned int32) lpBufferSize) cil managed preservesig
*/
//		[DllImport("mpr.dll", CharSet = CharSet.Auto, SetLastError = true)]
//		internal static extern UInt32 WNetGetUniversalName(string lpLocalPath, NAME_INFO_LEVEL dwInfoLevel, IntPtr lpBuffer, ref UInt32 lpBufferSize);
//[DllImport("mpr.dll", CharSet = CharSet.Auto, SetLastError = true)]
//internal static extern UInt32 WNetGetUniversalName(string lpLocalPath, NAME_INFO_LEVEL dwInfoLevel, ref UNIVERSAL_NAME_INFO lpBuffer, ref UInt32 lpBufferSize);
		// REF: http://www.pinvoke.net/default.aspx/advapi32/WNetGetUniversalName.html
		[DllImport("mpr.dll")]
		[return: MarshalAs(UnmanagedType.U4)]
		internal static extern int WNetGetUniversalName(
																	 string lpLocalPath,
																	 [MarshalAs(UnmanagedType.U4)] int dwInfoLevel,
																	 IntPtr lpBuffer,
																	 [MarshalAs(UnmanagedType.U4)] ref int lpBufferSize);
	}

	#endregion

	#region Win32 Wrapper Classes

	namespace Wrappers
	{
		public class WNet
		{
/*
 *	This does not work, always returning NOT CONNECTED or BAD DEVICE.
 * http://msdn2.microsoft.com/en-us/library/aa385474.aspx
 * http://www.pinvoke.net/default.aspx/advapi32/WNetGetUniversalName.html
 * http://www.thescripts.com/forum/thread248531.html
 * http://forums.microsoft.com/MSDN/ShowPost.aspx?PostID=2157648&SiteID=1
 * http://www.utteraccess.com/forums/showflat.php?Cat=&Board=93&Number=1281885&Zf=&Zw=&Zg=2&Zl=b&Main=1281742&Search=true&where=&Zu=15840&Zd=l&Zn=&Zt=3e8&Zs=&Zy=
 */
			/*
			 * REF: http://www.pinvoke.net/default.aspx/advapi32/WNetGetUniversalName.html
			 */
			public static string WNetGetUniversalName(string strFilepath)
			{
				string retVal = null;

				/*
				 * Get the size of the string that will be returned.
				 */
				IntPtr pBuffer = IntPtr.Zero;
				try
				{
					// First, call WNetGetUniversalName to get the size.
					int size = 0;

					// Make the call.
					// Pass IntPtr.Size because the API doesn't like null, even though
					// size is zero.  We know that IntPtr.Size will be
					// aligned correctly.
					int apiRetVal = Imports.WNetGetUniversalName(strFilepath, (int)NAME_INFO_LEVEL.UNIVERSAL_NAME_INFO_LEVEL, (IntPtr)IntPtr.Size, ref size);
					// If the return value is not ERROR_MORE_DATA, then raise an exception.
					if (apiRetVal != Constants.ERROR_MORE_DATA)
					{
						// Throw an exception.
						Marshal.ThrowExceptionForHR(apiRetVal);
						return null;
					}

					// Allocate the memory.
					pBuffer = Marshal.AllocCoTaskMem(size);

					//UNIVERSAL_NAME_INFO uni = new UNIVERSAL_NAME_INFO();
					//cch = (UInt32)Marshal.SizeOf(uni);
					//UInt32 rc = Win32Functions.Imports.WNetGetUniversalName(strFilepath, Win32Functions.NAME_INFO_LEVEL.UNIVERSAL_NAME_INFO_LEVEL, ref uni, ref cch);
					apiRetVal = Imports.WNetGetUniversalName(strFilepath, (int)NAME_INFO_LEVEL.UNIVERSAL_NAME_INFO_LEVEL, pBuffer, ref size);
					if (apiRetVal != Constants.NOERROR)
					{
						// Throw an exception.
						Marshal.ThrowExceptionForHR(apiRetVal);
						return null;
					}

					// Now get the string.  It's all in the same buffer, but the pointer is first,
					// so offset the pointer by IntPtr.Size and pass to PtrToStringAnsi.
					retVal = Marshal.PtrToStringAnsi(new IntPtr(pBuffer.ToInt64() + IntPtr.Size), size);
					retVal = retVal.Substring(0, retVal.IndexOf('\0'));
				}
				finally
				{
					Marshal.FreeCoTaskMem(pBuffer);
				}

				return retVal;
			}
		}


		/*
		 * These classes are geared towards supporting shell context menu extensions.
		 * 
		 * T is the type of data the caller wants to associate with each menu command.
		 */
		public class Menu
		{
			private static uint _idFirstCommand = 0;
			public static uint FirstCommand
			{
				get { return _idFirstCommand; }
				set
				{
					/*
					 * This can be initialized again because the assembly stays loaded in the shell,
					 * and subsequent context menu displays have to reset the id of the first command.
					 */
//					if (_idFirstCommand != 0)
//						throw new ApplicationException("Cannot initialize FirstCommand again.");

					_idFirstCommand = value;
					NextCommandID = _idFirstCommand;
					_commandhandlers.Clear();
				}
			}

			public static uint NextCommandID { get; private set; }
			public static uint MenuPositionIx { get; set; }
	
			//----------------------------------------------------------------

			private IntPtr _hMenu = IntPtr.Zero;
			public static implicit operator IntPtr(Menu m)
			{
				return m._hMenu;
			}

			//----------------------------------------------------------------

			static Menu()
			{
				NextCommandID = 0;
			}

			public Menu(IntPtr hMenu)
			{
				System.Diagnostics.Debug.Assert(hMenu != IntPtr.Zero, "hMenu cannot be null.");
				_hMenu = hMenu;
			}


			/*
			 * This collection maps a command offset (not ID) to a delegate and the data to pass to the delegate
			 */
			public delegate void MenuDelegate(uint id, object data);
			protected static System.Collections.Generic.Dictionary<uint, System.Collections.Generic.KeyValuePair<MenuDelegate, object>> _commandhandlers = new System.Collections.Generic.Dictionary<uint, System.Collections.Generic.KeyValuePair<MenuDelegate, object>>();


			/// <summary>
			/// Add passed menu command handler to the dictionary of delegates with the menu id as its key.
			/// Also associate the passed data object with this handler.
			/// </summary>
			/// <typeparam name="T">Type of data associated with menu command</typeparam>
			/// <param name="handler">Delegate to call when this menu command is selected</param>
			/// <param name="data">Data associated with this menu command and passed to delegate</param>
			/// <returns>The 0-based offset of the command id.</returns>
			private static uint AddMenuCommandHandler<T>(MenuDelegate handler, T data)
			{
				System.Diagnostics.Debug.Assert(handler != null, "The passed handler cannot be null.");
				System.Diagnostics.Debug.Assert(_idFirstCommand != 0, "FirstCommand must be initialized.");
				System.Diagnostics.Debug.Assert(NextCommandID != 0, "NextCommand must be initialized.");

				// add this command id and its handler
				uint offsetCommand = NextCommandID - _idFirstCommand;
				_commandhandlers.Add(offsetCommand, new System.Collections.Generic.KeyValuePair<MenuDelegate, object>(handler, data));
				++NextCommandID;

				return offsetCommand;
			}

			//static private System.Collections.Generic.KeyValuePair<MenuDelegate, object> GetCommandHandler(uint idMenuCommandOffset)
			//{
			//   /*
			//    * It's either this or catch KeyNotFoundException thrown by dic[key].
			//    */
			//   System.Collections.Generic.KeyValuePair<MenuDelegate, object> handler;
			//   _commandhandlers.TryGetValue(idMenuCommandOffset, out handler);
			//   return handler;
			//}

			/// <summary>
			/// Calls the delegate associated with the passed menu command offset (not ID).
			/// </summary>
			/// <param name="offsetCommand">Offset of menu command</param>
			/// <returns>True if it finds the command offset and calls its handler.</returns>
			public static bool CallCommandHandler(uint offsetCommand)
			{
				/*
				 * It's either this or catch KeyNotFoundException thrown by dic[key].
				 */
				System.Collections.Generic.KeyValuePair<MenuDelegate, object> pairHandlerData;
				if (!_commandhandlers.TryGetValue(offsetCommand, out pairHandlerData))
					return false;

				MenuDelegate handler = pairHandlerData.Key;
				System.Diagnostics.Debug.Assert(handler != null);
				if (handler == null)
					return false;

				handler(offsetCommand, pairHandlerData.Value);
				return true;
			}



			public void AppendMenuSeparator()
			{
				bool brc = Imports.AppendMenu((IntPtr)_hMenu, MF_WIN40.SEPARATOR, IntPtr.Zero, string.Empty);
				System.Diagnostics.Debug.Assert(brc, "AppendMenu (separator) failed.");
				++NextCommandID;
			}

			public uint AppendMenuCommand<T>(string strText, MenuDelegate handler, T data)
			{
				System.Diagnostics.Debug.Assert(strText != string.Empty, "Text cannot be empty.");

				bool brc = Imports.AppendMenu((IntPtr)_hMenu, MF_WIN40.STRING, (IntPtr)NextCommandID, strText);
				System.Diagnostics.Debug.Assert(brc, "AppendMenu failed.");

				return AddMenuCommandHandler(handler, data);
			}

			public void AppendMenuPopup(MenuPopup menuPopup)
			{
				System.Diagnostics.Debug.Assert(menuPopup != null, "The menu cannot be null.");

				bool brc = Imports.AppendMenu(_hMenu, MF_WIN40.POPUP, menuPopup, menuPopup.Text);
				System.Diagnostics.Debug.Assert(brc, "AppendMenu failed.");
				++NextCommandID;		// this may not be necessary for popups, but I want to be safe
			}


			public void InsertMenuSeparator()
			{
				bool brc = Imports.InsertMenu((IntPtr)_hMenu, MenuPositionIx, MF_WIN40.SEPARATOR, IntPtr.Zero, string.Empty);
				System.Diagnostics.Debug.Assert(brc, "InsertMenu (separator) failed.");
				++NextCommandID;
				++MenuPositionIx;
			}

			public uint InsertMenuCommand<T>(string strText, MenuDelegate handler, T data)
			{
				System.Diagnostics.Debug.Assert(strText != string.Empty, "Text cannot be empty.");

				bool brc = Imports.InsertMenu((IntPtr)_hMenu, MenuPositionIx, MF_WIN40.BYPOSITION | MF_WIN40.STRING, (IntPtr)NextCommandID, strText);
				System.Diagnostics.Debug.Assert(brc, "InsertMenuCommand failed.");
				++MenuPositionIx;

				return AddMenuCommandHandler(handler, data);
			}

			public void InsertMenuPopup(MenuPopup menuPopup)
			{
				System.Diagnostics.Debug.Assert(menuPopup != null, "The menu cannot be null.");

				bool brc = Imports.InsertMenu(_hMenu, MenuPositionIx, MF_WIN40.BYPOSITION | MF_WIN40.POPUP, menuPopup, menuPopup.Text);
				System.Diagnostics.Debug.Assert(brc);
				++NextCommandID;		// this may not be necessary for popups, but I want to be safe
				++MenuPositionIx;
			}
		}


		public class MenuPopup : Menu
		{
			protected string _strText = string.Empty;
			public string Text
			{
				get { return _strText; }
			}


			public MenuPopup(string strText)
				: base(Imports.CreatePopupMenu())
			{
				System.Diagnostics.Debug.Assert(strText.Length > 0, "The menu text cannot be empty.");
				_strText = strText;
			}
		}
	}

	#endregion
}
