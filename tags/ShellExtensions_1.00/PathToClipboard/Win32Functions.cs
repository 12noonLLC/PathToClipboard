using System;
using System.Text;

using System.Runtime.InteropServices;

namespace Win32Functions
{
	abstract public class Constants
	{
		public const int MAX_FILE_LEN = 1024;

		public const int NOERROR = 0;
		public const int ERROR_MORE_DATA = 234;
		public const int ERROR_BAD_DEVICE = 1200;
		public const int ERROR_NOT_CONNECTED = 2250;
	}

	#region Win32Enumerations

	public enum SEVERITY : uint
	{
		SUCCESS = 0,
		ERROR = 1
	}

	public class Macros
	{
		static public UInt32 MAKE_HRESULT(SEVERITY sev, UInt32 fac, UInt32 code)
		{
			// (HRESULT) (((unsigned long)(sev)<<31) | ((unsigned long)(fac)<<16) | ((unsigned long)(code)))
			return ((UInt32)sev << 31) | (fac << 16) | code;
		}
	}


	// Make these constants 
	public enum MIIM : uint
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

	public enum MF : uint
	{
		INSERT = 0x00000000,
		CHANGE = 0x00000080,
		APPEND = 0x00000100,
		DELETE = 0x00000200,
		REMOVE = 0x00001000,
		BYCOMMAND = 0x00000000,
		BYPOSITION = 0x00000400,
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

	public enum CLIPFORMAT : uint
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

	public enum NAME_INFO_LEVEL : uint
	{
		UNIVERSAL_NAME_INFO_LEVEL	= 0x000000001,
		REMOTE_NAME_INFO_LEVEL		= 0x000000002
	}

	#endregion

	#region Win32Structures

	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
	public struct MENUITEMINFO
	{
		public UInt32 cbSize;
		public UInt32 fMask;
		public UInt32 fType;
		public UInt32 fState;
		public UInt32 wID;
		public IntPtr	/*HMENU*/	  hSubMenu;
		public IntPtr	/*HBITMAP*/   hbmpChecked;
		public IntPtr	/*HBITMAP*/	  hbmpUnchecked;
		public IntPtr	/*ULONG_PTR*/ dwItemData;
		[MarshalAs(UnmanagedType.LPTStr)]
		public String dwTypeData;
		public UInt32 cch;
		public IntPtr	/*HBITMAP*/ hbmpItem;

/*http://www.pinvoke.net/default.aspx/Structures/MENUITEMINFO.html
*/		// return the size of the structure
		static internal UInt32 SizeOf
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
		[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
		internal static extern IntPtr lstrcpyn([Out] IntPtr pStringDest, IntPtr pStringSrc, UInt32 cchMax);


		[DllImport("shell32", CharSet = CharSet.Auto)]
		internal static extern UInt32 DragQueryFile(UInt32 hDrop, UInt32 iFile, StringBuilder buffer, Int32 cch);


		[DllImport("user32")]
		internal static extern IntPtr CreatePopupMenu();

		[DllImport("user32")]
		internal static extern Int32 GetMenuItemCount(IntPtr hmenu);

		[DllImport("user32", CharSet = CharSet.Auto)]
		internal static extern Boolean AppendMenu(IntPtr hmenu, MF uflags, UInt32 uIDNewItemOrSubmenu, string text);

		//superseded by InsertMenuItem() but can still use it
		[DllImport("user32", CharSet = CharSet.Auto)]
		internal static extern Boolean InsertMenu(IntPtr hmenu, UInt32 position, MF uflags, IntPtr uIDNewItemOrSubmenu, string text);

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
		[DllImport("user32", CharSet = CharSet.Auto, SetLastError=true)]
		internal static extern Boolean InsertMenuItem(IntPtr hmenu, UInt32 iItem, Boolean bByPosition, ref MENUITEMINFO mii);


/*
.method private hidebysig static pinvokeimpl("mpr.dll" winapi) 
		  int32  marshal( unsigned int32)  WNetGetUniversalName(string lpLocalPath,
																				  int32  marshal( unsigned int32) dwInfoLevel,
																				  native int lpBuffer,
																				  int32&  marshal( unsigned int32) lpBufferSize) cil managed preservesig
*/
		[DllImport("mpr.dll", CharSet = CharSet.Auto, SetLastError = true)]
		internal static extern UInt32 WNetGetUniversalName(string lpLocalPath, NAME_INFO_LEVEL dwInfoLevel, IntPtr lpBuffer, ref UInt32 lpBufferSize);
//[DllImport("mpr.dll", CharSet = CharSet.Auto, SetLastError = true)]
//internal static extern UInt32 WNetGetUniversalName(string lpLocalPath, NAME_INFO_LEVEL dwInfoLevel, ref UNIVERSAL_NAME_INFO lpBuffer, ref UInt32 lpBufferSize);
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
			static public string WNetGetUniversalName(string strFilepath)
			{
				/*
				 * Get the size of the string that will be returned.
				 */
				UInt32 cb = 2;
				IntPtr pBuffer = Marshal.AllocCoTaskMem((int)cb);
				try
				{
					//UNIVERSAL_NAME_INFO uni = new UNIVERSAL_NAME_INFO();
					//cch = (UInt32)Marshal.SizeOf(uni);
					//UInt32 rc = Win32Functions.Imports.WNetGetUniversalName(strFilepath, Win32Functions.NAME_INFO_LEVEL.UNIVERSAL_NAME_INFO_LEVEL,
					//                                                         ref uni, ref cch);
					UInt32 rc = Win32Functions.Imports.WNetGetUniversalName(strFilepath, Win32Functions.NAME_INFO_LEVEL.UNIVERSAL_NAME_INFO_LEVEL,
					                                                         pBuffer, ref cb);
					if (rc == Constants.ERROR_NOT_CONNECTED)
						return Constants.ERROR_NOT_CONNECTED.ToString();
					if (rc != Constants.ERROR_MORE_DATA)
						return rc.ToString();

					//UNIVERSAL_NAME_INFO uni = new UNIVERSAL_NAME_INFO();
					//cch = (UInt32)Marshal.SizeOf(uni);
					//rc = Win32Functions.Imports.WNetGetUniversalName(strFilepath, Win32Functions.NAME_INFO_LEVEL.UNIVERSAL_NAME_INFO_LEVEL,
					//                                                   ref uni, ref cch);
				}
				finally
				{
					Marshal.FreeCoTaskMem(pBuffer);
				}
				System.Diagnostics.Trace.TraceInformation("Path len = {0}", cb);

				StringBuilder sb = new StringBuilder(Win32Functions.Constants.MAX_FILE_LEN);

//				Win32Functions.Imports.WNetGetUniversalName(strFilepath, Win32Functions.NAME_INFO_LEVEL.UNIVERSAL_NAME_INFO_LEVEL, out sb, ref cch);

				return sb.ToString();
			}
		}


		/*
		 * These classes are geared towards supporting shell context menu extensions.
		 * 
		 * T is the type of data the caller wants to associate with each menu command.
		 */
		public class Menu
		{
			static private uint _idFirstCommand = 0;
			static public uint FirstCommand
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
					_idNextCommand = _idFirstCommand;
					_commandhandlers.Clear();
				}
			}

			static private uint _idNextCommand = 0;
			static public uint NextCommand
			{
				get { return _idNextCommand; }
			}

			//----------------------------------------------------------------

			static private uint _ixMenuPosition = 0;
			static public uint MenuPosition
			{
				get { return _ixMenuPosition; }
				set { _ixMenuPosition = value; }
			}
	
			//----------------------------------------------------------------

			private IntPtr _hMenu = IntPtr.Zero;
			public IntPtr HMENU
			{
				get { return _hMenu; }
			}

			//----------------------------------------------------------------


			public Menu(IntPtr hMenu)
			{
				System.Diagnostics.Debug.Assert(hMenu != IntPtr.Zero, "hMenu cannot be null.");
				_hMenu = hMenu;
			}


			public delegate void MenuDelegate(uint id, object data);
			static protected System.Collections.Generic.Dictionary<uint, System.Collections.Generic.KeyValuePair<MenuDelegate, object>> _commandhandlers = new System.Collections.Generic.Dictionary<uint, System.Collections.Generic.KeyValuePair<MenuDelegate, object>>();


			/// <summary>
			/// Add passed menu command handler to the dictionary of delegates with the menu id as its key.
			/// Also associate the passed data object with this handler.
			/// </summary>
			/// <typeparam name="T">Type of data associated with menu command</typeparam>
			/// <param name="handler">Delegate to call when this menu command is selected</param>
			/// <param name="data">Data associated with this menu command and passed to delegate</param>
			/// <returns>The 0-based offset of the command id.</returns>
			static private uint AddMenuCommandHandler<T>(MenuDelegate handler, T data)
			{
				System.Diagnostics.Debug.Assert(handler != null, "The passed handler cannot be null.");
				System.Diagnostics.Debug.Assert(_idFirstCommand != 0, "FirstCommand must be initialized.");
				System.Diagnostics.Debug.Assert(_idNextCommand != 0, "NextCommand must be initialized.");

				// add this command id and its handler 
				_commandhandlers.Add(_idNextCommand - _idFirstCommand, new System.Collections.Generic.KeyValuePair<MenuDelegate, object>(handler, data));
				++_idNextCommand;

				return _idNextCommand - _idFirstCommand - 1;
			}

			static public System.Collections.Generic.KeyValuePair<MenuDelegate, object> GetCommandHandler(uint idMenuCommandOffset)
			{
				/*
				 * It's either this or catch KeyNotFoundException thrown by dic[key].
				 */
				System.Collections.Generic.KeyValuePair<MenuDelegate, object> handler;
				_commandhandlers.TryGetValue(idMenuCommandOffset, out handler);
				return handler;
			}

			static public void CallCommandHandler(uint ixCmd)
			{
				System.Collections.Generic.KeyValuePair<MenuDelegate, object> bag = Menu.GetCommandHandler(ixCmd);
				MenuDelegate handler = bag.Key;
				if (handler != null)
					handler(ixCmd, bag.Value);
			}



			public void AppendMenuSeparator()
			{
				bool brc = Win32Functions.Imports.AppendMenu((IntPtr)_hMenu, Win32Functions.MF.SEPARATOR, 0, string.Empty);
				System.Diagnostics.Debug.Assert(brc, "AppendMenu (separator) failed.");
				++_idNextCommand;
			}

			public uint AppendMenuCommand<T>(string strText, MenuDelegate handler, T data)
			{
				System.Diagnostics.Debug.Assert(strText != string.Empty, "Text cannot be empty.");

				bool brc = Win32Functions.Imports.AppendMenu((IntPtr)_hMenu, Win32Functions.MF.STRING, (UInt32)Menu.NextCommand, strText);
				System.Diagnostics.Debug.Assert(brc, "AppendMenu failed.");

				return AddMenuCommandHandler(handler, data);
			}

			public void AppendMenuPopup(MenuPopup menuPopup)
			{
				System.Diagnostics.Debug.Assert(menuPopup != null, "The menu cannot be null.");

				bool brc = Win32Functions.Imports.AppendMenu(_hMenu, Win32Functions.MF.POPUP, (UInt32)menuPopup.HMENU, menuPopup.Text);
				System.Diagnostics.Debug.Assert(brc, "AppendMenu failed.");
				++_idNextCommand;		// this may not be necessary for popups, but I want to be safe
			}


			public void InsertMenuSeparator()
			{
				bool brc = Win32Functions.Imports.InsertMenu((IntPtr)_hMenu, Menu.MenuPosition, Win32Functions.MF.SEPARATOR, IntPtr.Zero, string.Empty);
				System.Diagnostics.Debug.Assert(brc, "InsertMenu (separator) failed.");
				++_idNextCommand;
				++_ixMenuPosition;
			}

			public uint InsertMenuCommand<T>(string strText, MenuDelegate handler, T data)
			{
				System.Diagnostics.Debug.Assert(strText != string.Empty, "Text cannot be empty.");

				bool brc = Win32Functions.Imports.InsertMenu((IntPtr)_hMenu, Menu.MenuPosition, Win32Functions.MF.BYPOSITION | Win32Functions.MF.STRING, (IntPtr)Menu.NextCommand, strText);
				System.Diagnostics.Debug.Assert(brc, "InsertMenuCommand failed.");
				++_ixMenuPosition;

				return AddMenuCommandHandler(handler, data);
			}

			public void InsertMenuPopup(MenuPopup menuPopup)
			{
				System.Diagnostics.Debug.Assert(menuPopup != null, "The menu cannot be null.");

				bool brc = Win32Functions.Imports.InsertMenu(_hMenu, Menu.MenuPosition, Win32Functions.MF.BYPOSITION | Win32Functions.MF.POPUP, menuPopup.HMENU, menuPopup.Text);
				System.Diagnostics.Debug.Assert(brc);
				++_idNextCommand;		// this may not be necessary for popups, but I want to be safe
				++_ixMenuPosition;
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
				: base(Win32Functions.Imports.CreatePopupMenu())
			{
				System.Diagnostics.Debug.Assert(strText.Length > 0, "The menu text cannot be empty.");
				_strText = strText;
			}
		}
	}

	#endregion
}
