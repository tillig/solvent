/*
	Solvent - Power toys for the Solution Explorer
	Copyright (C) 2004  Travis Illig
	tillig@paraesthesia.com
	http://www.paraesthesia.com

	Permission is hereby granted, free of charge, to any person obtaining
	a copy of this software and associated documentation files (the "Software"),
	to deal in the Software without restriction, including without limitation
	the rights to use, copy, modify, merge, publish, distribute, sublicense,
	and/or sell copies of the Software, and to permit persons to whom the Software
	is furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included
	in all copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
	THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	SOFTWARE.
*/

using System;
using System.Diagnostics;
using Microsoft.Office.Core;
using Extensibility;
using System.Runtime.InteropServices;
using EnvDTE;

namespace Solvent
{
	/// <summary>
	/// Utility and helper methods for Solvent
	/// </summary>
	public class Utility
	{
		/// <summary>
		/// Opens a folder/file in Windows Explorer.
		/// </summary>
		/// <param name="folderPath">The path to the folder to open.</param>
		/// <returns>A <see cref="System.Diagnostics.Process"/> with the newly created Explorer item.</returns>
		public static System.Diagnostics.Process OpenItemInExplorer(string folderPath){
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.EnableRaisingEvents = false;
			proc.StartInfo.FileName = "explorer.exe";
			proc.StartInfo.Arguments = "\"" + folderPath + "\"";
			proc.Start();
			return proc;
		}


		/// <summary>
		/// Opens a command prompt at the provided folder location.
		/// </summary>
		/// <param name="folderPath">The path to the folder to open the command prompt at.</param>
		/// <returns>A <see cref="System.Diagnostics.Process"/> with the newly created command prompt.</returns>
		public static System.Diagnostics.Process CmdPromptHere(string folderPath){
			System.Diagnostics.Process proc = new System.Diagnostics.Process();
			proc.EnableRaisingEvents = false;
			proc.StartInfo.FileName = "cmd.exe";
			proc.StartInfo.Arguments = "/k cd /d \"" + folderPath + "\"";
			proc.Start();
			return proc;
		}

		
		/// <summary>
		/// Gets the selected <see cref="UIHierarchyItems"/> from a given window.
		/// </summary>
		/// <param name="applicationObject">The <see cref="_DTE"/> application object for the current add-in.</param>
		/// <param name="windowKind">The type of window to get the selected items from.  See <see cref="EnvDTE.Constants"/>.
		/// Example: EnvDTE.Constants.vsWindowKindSolutionExplorer</param>
		/// <returns>An <see cref="Array"/> with the selected <see cref="UIHierarchyItem"/> items.</returns>
		public static Array GetSelectedUIHierarchyItems(_DTE applicationObject, string windowKind){
			// Get the specified window
			Window explWin = applicationObject.Windows.Item(windowKind);

			// Get the UI Hierarchy
			UIHierarchy uiHierarchy = (UIHierarchy)explWin.Object;

			// Get the currently selected items
			Array selectedItems = uiHierarchy.SelectedItems as Array;

			return selectedItems;
		}

		
		/// <summary>
		/// Gets the path/filename for a selected <see cref="UIHierarchyItem"/> (from the Solution Explorer window).
		/// </summary>
		/// <param name="uihItem">The item to get the filename from.</param>
		/// <param name="getPathOnly">True to return only the path; false to return the full path and filename.</param>
		/// <returns>The path/filename for the selected item.</returns>
		/// <remarks>
		/// <para>
		/// Getting a path for a <see cref="UIHierarchyItem"/> is not as simple as one might envision.
		/// The method differs based on whether it's a <see cref="Project"/> or <see cref="ProjectItem"/>;
		/// what language it's in; whether it's a <see cref="ProjectItem"/> (if it's an item) in a project or directly in the
		/// solution; the kind of project (if it's a project); and so on.
		/// </para>
		/// <para>
		/// The path for items is based on the following criteria:
		/// </para>
		/// <list type="table">
		/// <listheader><term>Object Type</term><description>Relevant Path Properties</description></listheader>
		/// <item>
		///		<term>C#, VB.NET, J# Project</term>
		///		<description>
		///		<para>Complete Path:  Project.FullName</para>
		///		<para>Containing Folder:  Project.Properties.Item("FullPath").Value</para>
		///		</description>
		///	</item>
		/// <item>
		///		<term>C#, VB.NET, J# ProjectItem</term>
		///		<description>
		///		<para>Complete Path:  ProjectItem.Properties.Item("FullPath").Value</para>
		///		<para>Filename Only:  ProjectItem.Properties.Item("FileName").Value</para>
		///		</description>
		///	</item>
		/// <item>
		///		<term>C++ Project</term>
		///		<description>
		///		<para>Complete Path:  Project.FullName</para>
		///		<para>Containing Folder:  Project.Properties.Item("ProjectDirectory").Value</para>
		///		</description>
		///	</item>
		/// <item>
		///		<term>C++ ProjectItem</term>
		///		<description>
		///		<para>Complete Path:  ProjectItem.Properties.Item("FullPath").Value</para>
		///		<para>Filename Only:  ProjectItem.Properties.Item("ItemName").Value</para>
		///		<para>Note - folders in C++ projects are "virtual" and exist only in the
		///		project, so you can't get a path from them.</para>
		///		</description>
		///	</item>
		/// <item>
		///		<term>Setup Project</term>
		///		<description>
		///		<para>Complete Path:  Project.FullName</para>
		///		<para>Project Name:  Project.Properties.Item("Name").Value</para>
		///		<para>Note - when you get the complete path to a Setup project it includes the
		///		full path to the project with the filename... but without the file extension.  For
		///		example, <c>C:\projects\SetupProject\SetupProject</c> - where the project is in
		///		the <c>C:\projects\SetupProject</c> folder and the Setup project is <c>SetupProject.vdproj</c>.
		///		That means you either have to manually append the file extension to get the full
		///		path or just assume you can't get the full path and call it good.</para>
		///		</description>
		///	</item>
		/// <item>
		///		<term>Solution Items</term>
		///		<description>
		///		<para>Complete Path:  ProjectItem.get_FileNames(1)</para>
		///		<para>Note - the "Solution Items" folder is virtual so there's no path
		///		directly to it.  Solution Items themselves only allow you to get the
		///		filename.</para>
		///		</description>
		///	</item>
		/// </list>
		/// </remarks>
		public static string GetUIHierarchyItemFilename(UIHierarchyItem uihItem, bool getPathOnly){
			Debug.WriteLine("Determining filename for: " + uihItem.Name, "GetUIHierarchyItemFilename()");
			string retVal = "";
			object uihItemObject = uihItem.Object;
			try {
				if(uihItemObject is ProjectItem){
					// Cast as a ProjectItem
					ProjectItem projItem = (ProjectItem)uihItemObject;

					// Get the complete file path with filename
					if(projItem.Properties != null){
						// Regular project items have a FullPath property
						retVal = projItem.Properties.Item("FullPath").Value.ToString();
					}
					else{
						// Solution items only have FileNames attached
						retVal = projItem.get_FileNames(1);
					}
				}
				else if(uihItemObject is Project){
					// Cast as a Project
					Project proj = (Project)uihItemObject;

					// Get the complete file path with filename
					retVal = proj.FullName;
					
					// If it's a Setup project, we've got a special case - we can't get the filename
					// Setup projects have a "ProductName" property - others don't.
					try{
						if(proj.Properties.Item("ProductName").Value != null){
							Debug.WriteLine("This is a setup project.", "GetUIHierarchyItemFilename()");
							if(getPathOnly){
								// Remove the setup project name from the end of the path
								string setupName = proj.Properties.Item("Name").Value.ToString();
								retVal = retVal.Remove(retVal.LastIndexOf(setupName), setupName.Length);
							}
							else{
								// We can't get a filename for Setup projects
								retVal = "";
							}
						}
					}
					catch{
						Debug.WriteLine("This isn't a setup project.", "GetUIHierarchyItemFilename()");
					}
				}
			}
			catch(InvalidCastException){
				Debug.WriteLine("Unable to determine filename for: " + uihItem.Name, "GetUIHierarchyItemFilename()");
			}
			catch(Exception err){
				Debug.WriteLine("Error while getting full path for: " + uihItem.Name + "; " + err.Message, "GetUIHierarchyItemFilename()");
			}

			// Do some final cleanup
			if(retVal != ""){
				if(retVal.IndexOf(System.IO.Path.DirectorySeparatorChar) < 0){
					// If there's NO folder separator in the path, this is probably one of
					// those C++ "virtual" project folders and we can't get a path for it.
					retVal = "";
				}
				else if(getPathOnly && !retVal.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString())){
					// We're getting only the containing folder - lose the filename
					int lastPathSepIndex = retVal.LastIndexOf(System.IO.Path.DirectorySeparatorChar);
					retVal = retVal.Remove(lastPathSepIndex + 1, retVal.Length - lastPathSepIndex - 1);
				}
			}

			Debug.WriteLine((getPathOnly ? "Containing folder" : "Full path") + ": " + retVal, "GetUIHierarchyItemFilename()");

			return retVal;
		}

		
		/// <summary>
		/// Determines the visibility of a particular window.
		/// </summary>
		/// <param name="applicationObject">The <see cref="_DTE"/> application object for the current add-in.</param>
		/// <param name="windowKind">The type of window to determine visibility of.  See <see cref="EnvDTE.Constants"/>.
		/// Example: EnvDTE.Constants.vsWindowKindSolutionExplorer</param>
		/// <returns>True if the window is visible; false otherwise.</returns>
		public static bool WindowIsVisible(_DTE applicationObject, string windowKind){
			// Get the specified window
			Window explWin = applicationObject.Windows.Item(windowKind);

			return explWin.Visible;
		}
	}
}
