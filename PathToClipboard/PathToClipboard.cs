using System;
using System.Collections.Generic;
using System.Text;

using System.Runtime.InteropServices;

namespace PathToClipboard	// WHY: PathToClipboardExtension doesn't work
{
	// Generate a new GUID and use it below in the class.
	[ClassInterface(ClassInterfaceType.None)]
	[Guid("2E820618-9430-421b-8FA5-12BD9E31EEF3"), ComVisible(true)]
	public class PathToClipboard : ShellExtension.ContextMenu
	{
		private const string _textMenuCommand = "Path to clipboard";
		private readonly string _fmtMenuCommandDir = "{0}" + System.IO.Path.DirectorySeparatorChar + "\tDirectory";
		private const string _fmtMenuCommandPath = "{0}\tPath";
		private const string _fmtMenuCommandFile = "{0}\tFile";

		// We don't really use the verb. Presumably, it'd have to be set in the Registry so Windows could pass it to InvokeCommand().
		private const string _verb = "p2c";
		private const string _verbCanonicalName = "PathToClipboard";
		private const string _verbHelpText = "Copy this text to the clipboard";


		/// <summary>
		/// Add commands to the context menu.
		/// </summary>
		/// <param name="menuContext"></param>
		/// <param name="filepaths"></param>
		/// <returns>True if it adds any commands; false if not.</returns>
		override protected bool AddMenuCommands(Win32Functions.Wrappers.Menu menuContext, System.Collections.Specialized.StringCollection filepaths)
		{
			if (filepaths.Count == 0)
			{
				return false;
			}

			Win32Functions.Wrappers.MenuPopup menuPopup = new Win32Functions.Wrappers.MenuPopup(_textMenuCommand);

			/*
			 * If there more than one file is selected, add an "ALL PATHS" option.
			 */
			if (filepaths.Count > 1)
			{
				// for each selected file, add all forms of its path to the menu, separated by--what else--separators
				StringBuilder strPaths = new StringBuilder();
				StringBuilder strNames = new StringBuilder();
				foreach (string strFilepath in filepaths)
				{
					strPaths.AppendLine(strFilepath);
					string strName = System.IO.Path.GetFileName(strFilepath);
					strNames.AppendLine(strName);
				}

				// C:\bin\
				string strDir = System.IO.Path.GetDirectoryName(filepaths[0]);
				menuPopup.AppendMenuCommand(String.Format(_fmtMenuCommandDir, strDir), MyCommandHandler, strDir);
				menuPopup.AppendMenuCommand(String.Format(_fmtMenuCommandPath, "All paths"), MyCommandHandler, strPaths.ToString());
				//BMK TODO: "All names" command seems to crash Windows Explorer
				menuPopup.AppendMenuCommand(String.Format(_fmtMenuCommandFile, "All names"), MyCommandHandler, strNames.ToString());
				menuPopup.AppendMenuSeparator();
			}

			foreach (string strFilepath in filepaths)
			{
				if (strFilepath != filepaths[0])
					menuPopup.AppendMenuSeparator();

				// C:\bin\test.txt
				menuPopup.AppendMenuCommand(String.Format(_fmtMenuCommandPath, strFilepath), MyCommandHandler, strFilepath);

				if (filepaths.Count == 1)
				{
					// C:\bin\
					string strDir = System.IO.Path.GetDirectoryName(strFilepath);
					menuPopup.AppendMenuCommand(String.Format(_fmtMenuCommandDir, strDir), MyCommandHandler, strDir);
				}

				// test.txt
				string strName = System.IO.Path.GetFileName(strFilepath);
				menuPopup.AppendMenuCommand(String.Format(_fmtMenuCommandFile, strName), MyCommandHandler, strName);

				///*
				// * UNC path: \\COMPUTER\Documents\test.txt
				// */
				//string strUNC = Win32Functions.Wrappers.WNet.WNetGetUniversalName(strFilepath);
				//menuPopup.AppendMenuCommand(String.Format("UNC format", strUNC), MyCommandHandler, strUNC);
			}

			menuContext.InsertMenuPopup(menuPopup);

			/*
			 * Note: This is called when selecting a jump list's command.
			 */
			return true;
		}

		override protected string GetCommandStringVerb()
		{
			return _verb;
		}

		override protected string GetCommandStringCanonicalVerb()
		{
			return _verbCanonicalName;
		}

		override protected string GetCommandStringHelp(uint ixCmd)
		{
			return _verbHelpText;
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
		 * Use "Files" instead of "*", "Folder".
		 * Could also be "Directory" instead of "Folder".
		 * 
		 * AllFilesystemObjects
		 * *
		 * Folder
		 * Directory
		 * Folder
		 */
		static private string[] _filetypes = new string[] { "AllFilesystemObjects" };

		[ComRegisterFunction()]
		static public void RegisterServer(Type t)
		{
			ShellExtension.ContextMenu.RegisterServerHelper(t, _filetypes);
		}

		[ComUnregisterFunction()]
		static public void UnregisterServer(Type t)
		{
			ShellExtension.ContextMenu.UnregisterServerHelper(t, _filetypes);
		}

		#endregion
	}
}
