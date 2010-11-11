/*
 * Must register as a COM object:
 * C:\Windows\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe /v MyAssembly.dll
 *
 * Pass /regfile:test.txt to dump the registry entries to a file instead.
 *
 * Must add to GAC:
 * "C:\Program Files (x86)\Microsoft Visual Studio 8\SDK\v2.0\Bin\gacutil.exe" /if MyAssembly.dll
 *
 * Associated with HKCR/ * shellex/ContextMenuHandlers/
 * 
 * Writing Windows Shell Extension with .NET Framework 4
 * http://blogs.msdn.com/b/codefx/archive/2010/09/14/writing-windows-shell-extension-with-net-framework-4-c-vb-net-part-1.aspx
 * http://blogs.msdn.com/b/codefx/archive/2010/10/10/writing-windows-shell-extension-with-net-framework-4-c-vb-net-part-2.aspx
 * 
 * Registering Shell Extension Handlers
 * http://msdn.microsoft.com/en-us/library/cc144110.aspx
 *
 * Creating Shortcut Menu Handlers
 * http://msdn.microsoft.com/en-us/library/cc144171.aspx
 * 
 * Customizing a Shortcut Menu Using Dynamic Verbs
 * http://msdn.microsoft.com/en-us/library/ee453696.aspx
 */

using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace ShellExtension
{
	public abstract class ContextMenu : MyCOMDefinitions.IShellExtInit, MyCOMDefinitions.IContextMenu
	{
		private uint IDM_DISPLAY = 0;

		private System.Collections.Specialized.StringCollection _filepaths = new System.Collections.Specialized.StringCollection();


		#region IShellExtInit Members

		void MyCOMDefinitions.IShellExtInit.Initialize(IntPtr pidlFolder, IntPtr lpdobj, IntPtr hKeyProgID)
		{
			_filepaths.Clear();

			if (lpdobj == IntPtr.Zero)
				return;	//			return 0;

			IntPtr hDrop = IntPtr.Zero;	// HDROP representing collection of selected files
			System.Runtime.InteropServices.ComTypes.STGMEDIUM medium = new System.Runtime.InteropServices.ComTypes.STGMEDIUM();
			try
			{
				// Get info about the directory
				/*Doesn't work
								System.Windows.Forms.DataObject dataobject = new System.Windows.Forms.DataObject(System.Windows.Forms.DataFormats.FileDrop, lpdobj);
								if (dataobject.ContainsFileDropList())
								{
									System.Collections.Specialized.StringCollection files = dataobject.GetFileDropList();
								}
				*/
				System.Runtime.InteropServices.ComTypes.FORMATETC fmt = new System.Runtime.InteropServices.ComTypes.FORMATETC();
				fmt.cfFormat = (short)Win32Functions.CLIPFORMAT.CF_HDROP;
				fmt.ptd = IntPtr.Zero;
				fmt.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
				fmt.lindex = -1;
				fmt.tymed = System.Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL;

				System.Runtime.InteropServices.ComTypes.IDataObject dataObject = (System.Runtime.InteropServices.ComTypes.IDataObject)Marshal.GetObjectForIUnknown(lpdobj);
				dataObject.GetData(ref fmt, out medium);
				hDrop = medium.unionmember;
				if (hDrop == IntPtr.Zero)
					return;

				// how many files are selected? (Not necessarily any!)
				uint nSelected = Win32Functions.Imports.DragQueryFile(hDrop, uint.MaxValue, null, 0);
				if (nSelected == 0)
					return;

				// get each file path
				StringBuilder sb = new StringBuilder(Win32Functions.Constants.MAX_FILE_LEN);
				for (uint ix = 0; ix < nSelected; ++ix)
				{
					if (Win32Functions.Imports.DragQueryFile(hDrop, ix, sb, (uint)sb.Capacity) != 0)
						_filepaths.Add(sb.ToString());
				}
			}
			catch (Exception)
			{
			}
			finally
			{
				Win32Functions.Imports.ReleaseStgMedium(ref medium);
			}

//			return 0;
		}

		#endregion

		#region IContextMenu Members

		int MyCOMDefinitions.IContextMenu.QueryContextMenu(IntPtr hmenu, uint ixMenu, uint idCmdFirst, uint idCmdLast, uint uFlags)
		{
			// if the user is activating the default command, do nothing
			if ((uFlags & (uint)MyCOMDefinitions.CMF.CMF_DEFAULTONLY) != 0)
			{
				return (int)Win32Functions.Macros.MAKE_HRESULT(Win32Functions.SEVERITY.SUCCESS, 0, 0);
			}

			System.Diagnostics.Debug.Assert(idCmdFirst <= idCmdLast, "Windows bug: The first command is supposed to be less than or equal to the last command.");
			if (_filepaths.Count == 0)
			{
				return (int)Win32Functions.Macros.MAKE_HRESULT(Win32Functions.SEVERITY.SUCCESS, 0, 0);
			}

			// start menu command ids here
			uint idCmdNext = idCmdFirst;

			Win32Functions.Wrappers.Menu.FirstCommand = idCmdFirst;
			Win32Functions.Wrappers.Menu.MenuPosition = ixMenu;
			Win32Functions.Wrappers.Menu menuContext = new Win32Functions.Wrappers.Menu(hmenu);
			AddMenuCommands(menuContext, _filepaths);

			System.Diagnostics.Debug.Assert(Win32Functions.Wrappers.Menu.NextCommand - 1 <= idCmdLast, "Added more menu commands than permitted.");
			return (int)Win32Functions.Macros.MAKE_HRESULT(Win32Functions.SEVERITY.SUCCESS, 0, Win32Functions.Wrappers.Menu.NextCommand - Win32Functions.Wrappers.Menu.FirstCommand);
		}

		/// <summary>
		/// Function adds menu commands to the shell's context menu.
		/// </summary>
		/// <param name="menuContext">Menu object representing the shell's context menu.</param>
		/// <param name="filepaths">List of selected file paths.</param>
		/// <example>
		/// <code>
		/// menuContext.AppendMenuCommand("Test Menu class command 1", MyCommandHandler1);
		/// menuContext.AppendMenuCommand("Test Menu class command 2", MyCommandHandler2);
		/// menuContext.AppendMenuSeparator();
		///
		/// Win32Functions.Wrappers.MenuPopup menuPopup = new Win32Functions.Wrappers.MenuPopup("My nice test popup");
		/// menuPopup.AppendMenuCommand("Test SubMenu command 1", MyCommandHandler1);
		/// menuPopup.AppendMenuSeparator();
		/// menuPopup.AppendMenuCommand("Test SubMenu command 2", MyCommandHandler2);
		/// menuContext.AppendMenuPopup(menuPopup);
		///
		///
		/// menuContext.InsertMenuCommand("Test new menu class insert command 1", MyCommandHandler1);
		/// menuContext.InsertMenuCommand("Test new menu class insert command 2", MyCommandHandler2);
		/// menuContext.InsertMenuSeparator();
		///
		/// menuPopup = new Win32Functions.Wrappers.MenuPopup("My nice test insert popup");
		/// menuPopup.AppendMenuCommand("Test insert SubMenu command 1", MyCommandHandler1);
		/// menuPopup.InsertMenuSeparator();
		/// menuPopup.AppendMenuCommand("Test insert SubMenu command 2", MyCommandHandler2);
		/// menuContext.InsertMenuPopup(menuPopup);
		/// 
		/// public void MyCommandHandler1()
		/// {
		///	System.Windows.Forms.MessageBox.Show("Handler 1");
		/// }
		/// </code>
		/// </example>
		abstract protected void AddMenuCommands(Win32Functions.Wrappers.Menu menuContext, System.Collections.Specialized.StringCollection filepaths);


		void MyCOMDefinitions.IContextMenu.InvokeCommand(IntPtr pici)
		{
			try
			{
				/*
				 * This also works:
				 * Type typINVOKECOMMANDINFO = Type.GetType("MyCOMDefinitions.INVOKECOMMANDINFO");
				 * but it's more error-prone (if a type's name changes).
				 */
				Type typINVOKECOMMANDINFO = typeof(MyCOMDefinitions.INVOKECOMMANDINFO);
				MyCOMDefinitions.INVOKECOMMANDINFO ici = (MyCOMDefinitions.INVOKECOMMANDINFO)Marshal.PtrToStructure(pici, typINVOKECOMMANDINFO);
				/*
				 * First command: verb == 0
				 * Second command: verb == 1
				 * etc.
				 */
				Win32Functions.Wrappers.Menu.CallCommandHandler(ici.verb);
			}
			catch (Exception ex)
			{
				System.Diagnostics.Trace.TraceError("InvokeCommand exception: " + ex.Message);
			}
		}


		void MyCOMDefinitions.IContextMenu.GetCommandString(UIntPtr idcmd, uint uflags, IntPtr reserved, StringBuilder sbCommandString, UInt32 cch)
		{
			if (idcmd.ToUInt32() == IDM_DISPLAY)
			{
				string strText = string.Empty;

				if (((MyCOMDefinitions.GCS)uflags == MyCOMDefinitions.GCS.VERBA) || ((MyCOMDefinitions.GCS)uflags == MyCOMDefinitions.GCS.VERBW))
				{
					strText = GetCommandStringVerb();
				}
				else if (((MyCOMDefinitions.GCS)uflags == MyCOMDefinitions.GCS.HELPTEXTA) || ((MyCOMDefinitions.GCS)uflags == MyCOMDefinitions.GCS.HELPTEXTW))
				{
					strText = GetCommandStringHelp(Win32Functions.Wrappers.Menu.FirstCommand);
				}

				if (strText.Length > cch - 1)
				{
					Marshal.ThrowExceptionForHR(Win32Functions.Constants.STRSAFE_E_INSUFFICIENT_BUFFER);
				}

				sbCommandString.Clear();
				sbCommandString.Append(strText);
			}
		}

		/// <summary>
		/// Override this function to return specific text for the verb.
		/// </summary>
		/// <returns>The string to use as the verb</returns>
		virtual protected string GetCommandStringVerb()
		{
			return string.Empty;
		}

		/// <summary>
		/// Override this function to return specific text to be displayed in
		/// the status bar when the mouse hovers over the command with the
		/// passed offset.
		/// </summary>
		/// <param name="ixCmd">The offset of the menu command being hovered over.</param>
		/// <returns>The string to display in the status bar</returns>
		virtual protected string GetCommandStringHelp(uint ixCmd)
		{
			return string.Empty;
		}

		#endregion


		#region Registration

		/*
		 * These methods must be called by the derived class that is, ultimately, the extension.
		 * They're invoked by regasm.exe.
		 * 
		 * Creating Shell Extensions
		 * http://msdn.microsoft.com/en-us/library/cc144067(VS.85).aspx
		 * 
		 * Registering Shell Extension Handlers
		 * http://msdn.microsoft.com/en-us/library/cc144110(VS.85).aspx
		 * 
		 * Approved Shell Extensions
		 * http://msdn.microsoft.com/en-us/library/ms812054.aspx
		 */

		protected static void RegisterServerHelper(Type t, string[] filetypes)
		{
			string strGuid = t.GUID.ToString("B");

			try
			{
				/*
				 * Mark this extension as Approved.
				 */
				//Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true);
				//key.SetValue(strGuid, t.Name);
				//key.Close();

				/*
				 * Associate all specified filetypes with this handler.
				 * 
				 * Registry entry:
				 *		HKCR\<filetype>\ContextMenuHandlers\
				 */
				foreach (string filetype in filetypes)
				{
					string keynameFileType = GetFileType(filetype);

					Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(keynameFileType + @"\shellex\ContextMenuHandlers\" + strGuid);
					key.SetValue("", t.Name);
					key.Close();

					/*
					 * http://msdn.microsoft.com/en-us/library/bb762118.aspx
					 * SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL)
					 * 
					 * http://www.pinvoke.net/default.aspx/shell32.shchangenotify
					 * 
					[DllImport("shell32.dll")]
					static extern void SHChangeNotify(HChangeNotifyEventID wEventId, 
																  HChangeNotifyFlags uFlags, 
																  IntPtr dwItem1, 
																  IntPtr dwItem2);
					 */

					System.Console.WriteLine("Registered file type {0} (from {1})", keynameFileType, filetype);
				}

				System.Console.WriteLine("Registered GUID {0}", strGuid);
			}
			catch (Exception ex)
			{
				System.Console.WriteLine("Error registering extension: {0}", ex.Message);
			}
		}

		protected static void UnregisterServerHelper(Type t, string[] filetypes)
		{
			string strGuid = t.GUID.ToString("B");

			try
			{
				/*
				 * Unassociate all filetypes associated with this handler.
				 * 
				 * Registry entry:
				 *		HKCR\<filetype>\ContextMenuHandlers\
				 */
				foreach (string filetype in filetypes)
				{
					string keynameFileType = GetFileType(filetype);
					Microsoft.Win32.Registry.ClassesRoot.DeleteSubKeyTree(keynameFileType + @"\shellex\ContextMenuHandlers\" + strGuid);
					System.Console.WriteLine("Unregistered file type {0} (from {1})", keynameFileType, filetype);
				}

				/*
				 * Clear this extension as Approved.
				 */
				//Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true);
				//key.DeleteValue(strGuid);
				//key.Close();

				System.Console.WriteLine("Unregistered GUID {0}", strGuid);
			}
			catch (Exception ex)
			{
				System.Console.WriteLine("Error unregistering extension: {0}", ex.Message);
			}
		}

		private static string GetFileType(string filetype)
		{
			// If fileType starts with '.', try to read the default value of the 
			// HKCR\<File Type> key which contains the ProgID to which the file type 
			// is linked.
			if (filetype.StartsWith("."))
			{
				using (Microsoft.Win32.RegistryKey keyType = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(filetype))
				{
					if (keyType != null)
					{
						// If the key exists and its default value is not empty, use 
						// the ProgID as the file type.
						string defaultVal = keyType.GetValue(null) as string;
						if (!string.IsNullOrEmpty(defaultVal))
						{
							return defaultVal;
						}
					}
				}
			}
			return filetype;
		}

		#endregion
	}
}
