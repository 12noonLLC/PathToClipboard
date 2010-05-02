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
 */

using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace ShellExtension
{
	public abstract class ContextMenu : MyCOMDefinitions.IShellExtInit, MyCOMDefinitions.IContextMenu
	{
		private uint _hDrop = 0;						// HDROP representing collection of selected files


		#region IShellExtInit Members

		UInt32 MyCOMDefinitions.IShellExtInit.Initialize(IntPtr pidlFolder, IntPtr lpdobj, IntPtr hKeyProgID)
		{
			_hDrop = 0;

			if (lpdobj == IntPtr.Zero)
				return 0;

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
				System.Runtime.InteropServices.ComTypes.IDataObject dataObject = (System.Runtime.InteropServices.ComTypes.IDataObject)Marshal.GetObjectForIUnknown(lpdobj);
				System.Runtime.InteropServices.ComTypes.FORMATETC fmt = new System.Runtime.InteropServices.ComTypes.FORMATETC();
				fmt.cfFormat = (short)Win32Functions.CLIPFORMAT.CF_HDROP;
				fmt.ptd = IntPtr.Zero;
				fmt.dwAspect = System.Runtime.InteropServices.ComTypes.DVASPECT.DVASPECT_CONTENT;
				fmt.lindex = -1;
				fmt.tymed = System.Runtime.InteropServices.ComTypes.TYMED.TYMED_HGLOBAL;
				System.Runtime.InteropServices.ComTypes.STGMEDIUM medium = new System.Runtime.InteropServices.ComTypes.STGMEDIUM();
				dataObject.GetData(ref fmt, out medium);
				_hDrop = (uint)medium.unionmember.ToInt32();
			}
			catch (Exception)
			{
			}

			return 0;
		}

		#endregion

		#region IContextMenu Members

		UInt32 MyCOMDefinitions.IContextMenu.QueryContextMenu(IntPtr hmenu, UInt32 ixMenu, UInt32 idCmdFirst, UInt32 idCmdLast, UInt32 uFlags)
		{
			System.Diagnostics.Debug.Assert(idCmdFirst <= idCmdLast, "Windows bug: The first command is supposed to be less than or equal to the last command.");
			if (_hDrop == 0)
				return 0;

			// how many files are selected?
			uint nSelected = Win32Functions.Imports.DragQueryFile(_hDrop, 0xffffffff, null, 0);
			System.Diagnostics.Debug.Assert(nSelected > 0, "At least one file must be selected.");

			// get each file path
			System.Collections.Specialized.StringCollection filepaths = new System.Collections.Specialized.StringCollection();
			for (uint ix = 0; ix < nSelected; ++ix)
			{
				StringBuilder sb = new StringBuilder(Win32Functions.Constants.MAX_FILE_LEN);
				Win32Functions.Imports.DragQueryFile(_hDrop, ix, sb, sb.Capacity);
				filepaths.Add(sb.ToString());
			}

			// start menu command ids here
			uint idCmdNext = idCmdFirst;

			Win32Functions.Wrappers.Menu.FirstCommand = idCmdFirst;
			Win32Functions.Wrappers.Menu.MenuPosition = ixMenu;
			Win32Functions.Wrappers.Menu menuContext = new Win32Functions.Wrappers.Menu(hmenu);
			AddMenuCommands(menuContext, filepaths);

			System.Diagnostics.Debug.Assert(Win32Functions.Wrappers.Menu.NextCommand - 1 <= idCmdLast, "Added more menu commands than permitted.");
			return Win32Functions.Macros.MAKE_HRESULT(Win32Functions.SEVERITY.SUCCESS, 0, Win32Functions.Wrappers.Menu.NextCommand - Win32Functions.Wrappers.Menu.FirstCommand);
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


		UInt32 MyCOMDefinitions.IContextMenu.InvokeCommand(IntPtr pici)
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
			return 0;
		}


		UInt32 MyCOMDefinitions.IContextMenu.GetCommandString(UInt32 idcmd, UInt32 uflags, IntPtr reserved, IntPtr pszCommandString, UInt32 cch)
		{
			if (((uflags & (UInt32)MyCOMDefinitions.GCS.HELPTEXTA) != 0) ||
				 ((uflags & (UInt32)MyCOMDefinitions.GCS.HELPTEXTW) != 0))
			{
				string strText = GetCommandStringHelp(idcmd - Win32Functions.Wrappers.Menu.FirstCommand);

				IntPtr h = Marshal.StringToHGlobalAuto(strText);
				try
				{
					Win32Functions.Imports.lstrcpyn(pszCommandString, h, cch);
				}
				finally
				{
					Marshal.FreeHGlobal(h);
				}
			}
			return 0;
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
		 */

		static protected void RegisterServerHelper(Type t, string[] filetypes)
		{
			string strGuid = "{" + t.GUID.ToString() + "}";

			try
			{
				/*
				 * Mark this extension as Approved.
				 */
				Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true);
				key.SetValue(strGuid, t.Name);
				key.Close();

				/*
				 * Associate all specified filetypes with this handler.
				 * 
				 * Registry entry:
				 *		HKCR\<filetype>\ContextMenuHandlers\
				 */
				foreach (string filetype in filetypes)
				{
					key = Microsoft.Win32.Registry.ClassesRoot.CreateSubKey(filetype + @"\shellex\ContextMenuHandlers\" + t.Name + " " + strGuid);
					key.SetValue("", strGuid);
					key.Close();
					System.Console.WriteLine("Registered file type {0}", filetype);
				}

				System.Console.WriteLine("Registered GUID {0}", strGuid);
			}
			catch (Exception ex)
			{
				System.Console.WriteLine("Error registering extension: {0}", ex.Message);
			}
		}

		static protected void UnregisterServerHelper(Type t, string[] filetypes)
		{
			string strGuid = "{" + t.GUID.ToString() + "}";

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
					Microsoft.Win32.Registry.ClassesRoot.DeleteSubKey(filetype + @"\shellex\ContextMenuHandlers\" + t.Name + " " + strGuid);
					System.Console.WriteLine("Unregistered file type {0}", filetype);
				}

				/*
				 * Clear this extension as Approved.
				 */
				Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved", true);
				key.DeleteValue(strGuid);
				key.Close();

				System.Console.WriteLine("Unregistered GUID {0}", strGuid);
			}
			catch (Exception ex)
			{
				System.Console.WriteLine("Error unregistering extension: {0}", ex.Message);
			}
		}

		#endregion
	}
}
