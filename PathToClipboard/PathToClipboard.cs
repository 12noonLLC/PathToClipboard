using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace PathToClipboard	// WHY: PathToClipboardExtension doesn't work
{
	// Generate a new GUID and use it below in the class.
	[ComVisible(true), System.Runtime.InteropServices.Guid("2E820618-9430-421b-8FA5-12BD9E31EEF3")]
	public class PathToClipboard : ShellExtension.ContextMenu
	{
		override protected void AddMenuCommands(Win32Functions.Wrappers.Menu menuContext, System.Collections.Specialized.StringCollection filepaths)
		{
			if (filepaths.Count == 0)
				return;

			Win32Functions.Wrappers.MenuPopup menuPopup = new Win32Functions.Wrappers.MenuPopup("Path to clipboard");

			/*
			 * If there more than one file is selected, add an "ALL PATHS" option.
			 */
			if (filepaths.Count > 1)
			{
				// for each selected file, add all forms of its path to the menu, separated by--what else--separators
				string strPaths = string.Empty;
				string strNames = string.Empty;
				foreach (string strFilepath in filepaths)
				{
					strPaths += strFilepath + System.IO.Path.PathSeparator + System.Environment.NewLine;
					string strName = System.IO.Path.GetFileName(strFilepath);
					strNames += strName + System.Environment.NewLine;
				}

				// C:\bin\
				string strDir = System.IO.Path.GetDirectoryName(filepaths[0]) + System.IO.Path.PathSeparator;
				menuPopup.AppendMenuCommand(strDir + "\tDirectory", MyCommandHandler, strDir);
				menuPopup.AppendMenuCommand("All paths\tPath", MyCommandHandler, strPaths);
				menuPopup.AppendMenuCommand("All names\tFile", MyCommandHandler, strNames);
				menuPopup.AppendMenuSeparator();
			}

			foreach (string strFilepath in filepaths)
			{
				if (strFilepath != filepaths[0])
					menuPopup.AppendMenuSeparator();

				// C:\bin\test.txt
				menuPopup.AppendMenuCommand(strFilepath + "\tPath", MyCommandHandler, strFilepath);

				if (filepaths.Count == 1)
				{
					// C:\bin\
					string strDir = System.IO.Path.GetDirectoryName(strFilepath) + System.IO.Path.PathSeparator;
					menuPopup.AppendMenuCommand(strDir + "\tDirectory", MyCommandHandler, strDir);
				}

				// test.txt
				string strName = System.IO.Path.GetFileName(strFilepath);
				menuPopup.AppendMenuCommand(strName + "\tFile", MyCommandHandler, strName);

//UNC paths
//            // \\computer\c\bin\test.txt
////string strUNC = "UNC format";
//string strUNC = Win32Functions.Wrappers.WNet.WNetGetUniversalName(s);
//            _idsFilepaths.Add(menuPopup.AppendMenuCommand(strUNC, MyCommandHandler), strUNC);
			}

			menuContext.InsertMenuPopup(menuPopup);
		}

		override protected string GetCommandStringHelp(uint ixCmd)
		{
			return "Copy this text to the clipboard";
		}


		#region Menu Handlers

		/// <summary>
		/// Called by menu when the user selects a command.
		/// </summary>
		/// <typeparam name="T">Type of data stored with menu command</typeparam>
		/// <param name="ix">Offset of menu command</param>
		/// <param name="s">Data saved with command when it was added to the menu</param>
		public void MyCommandHandler(uint ixOffset, object s)
		{
			System.Windows.Forms.Clipboard.SetText(s.ToString());
		}

		#endregion

		#region Registration
		/*
			Use this instead of "*", "Folder".
			Could also be "Directory" instead of "Folder".
		 */
		static private string[] _filetypes = new string[] { "AllFilesystemObjects" };

		[System.Runtime.InteropServices.ComRegisterFunctionAttribute()]
		static public void RegisterServer(Type t)
		{
			ShellExtension.ContextMenu.RegisterServerHelper(t, _filetypes);
		}

		[System.Runtime.InteropServices.ComUnregisterFunctionAttribute()]
		static public void UnregisterServer(Type t)
		{
			ShellExtension.ContextMenu.UnregisterServerHelper(t, _filetypes);
		}

		#endregion
	}
}
