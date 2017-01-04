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

namespace Solvent {
	using System;
	using System.Diagnostics;
	using Microsoft.Office.Core;
	using Extensibility;
	using System.Runtime.InteropServices;
	using EnvDTE;
#if SENDTO
	// TODO: Change namespace when actual SendTo Menu is in effect
	using Paraesthesia.WinShell.SendToMenu_Stub;
#endif

	#region Read me for Add-in installation and setup information.

	// When run, the Add-in wizard prepared the registry for the Add-in.
	// At a later time, if the Add-in becomes unavailable for reasons such as:
	//   1) You moved this project to a computer other than which is was originally created on.
	//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
	//   3) Registry corruption.
	// you will need to re-register the Add-in by building the MyAddin21Setup project
	// by right clicking the project in the Solution Explorer, then choosing install.

	#endregion


	/// <summary>
	/// The primary class implementing the Solvent Add-in.
	/// </summary>
	/// <seealso class='IDTExtensibility2' />
	[GuidAttribute("E0C21DD5-61DD-4DFE-9715-AF99E2673EEB"), ProgId("Solvent.Connect")]
	public class Connect : Object, Extensibility.IDTExtensibility2, IDTCommandTarget {
		#region Constants

		/// <summary>
		/// The prefix for all commands.
		/// </summary>
		public const string CommandPrefix = "Solvent.Connect.";

#if SENDTO
		/// <summary>
		/// The prefix for SendTo commands.
		/// </summary>
		public const string CommandPrefixSendTo = "Solvent.Connect.SendTo";
#endif

		/// <summary>
		/// Gets the list of all available commands for this Add-In.
		/// </summary>
		/// <returns>An array of strings, each item being the name of a given command.</returns>
		public static readonly string[] CommandList = new string[]{
																	  "RecurseExpandContract",
																	  "ToggleRecursiveExpansion",
																	  "OpenSelectedItemContainingFolder",
																	  "OpenSelectedProjectContainingFolder",
																	  "OpenAllSubItems",
																	  "CmdInSelectedItemContainingFolder",
																	  "CmdInSelectedProjectContainingFolder"
																  };

		#endregion


		#region Variables

		private _DTE applicationObject = null;
		private AddIn addInInstance = null;
		private System.Resources.ResourceManager rm = null;
#if SENDTO
		private SendToMenuItem[] sendToMenuItems = null;
#endif
		#endregion


		#region Construction and Initialization

		/// <summary>
		///	Implements the constructor for the Add-in object.
		/// </summary>
		public Connect() {
			// Set up the resource manager
			try {
				rm = new System.Resources.ResourceManager("Solvent.Strings", this.GetType().Assembly);
			}
			catch(System.Exception err) {
				Debug.WriteLine("Error getting ResourceManager: " + err.ToString(), "Connect.Connect()");
				rm = null;
				throw err;
			}
		}

		#endregion


		#region IDTExtensibility2, IDTCommandTarget

		/// <summary>
		/// Implements the OnConnection method of the <see cref="Extensibility.IDTExtensibility2"/> interface.
		/// Receives notification that the Add-in is being loaded.
		/// </summary>
		/// <param name="application">Root object of the host application.</param>
		/// <param name="connectMode">Describes how the Add-in is being loaded.</param>
		/// <param name="addInInst">Object representing this Add-in.</param>
		/// <param name="custom">Custom parameters; not used in VS.NET 2003</param>
		/// <seealso class="IDTExtensibility2" />
		public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom) {
			applicationObject = (_DTE)application;
			addInInstance = (AddIn)addInInst;
			if(connectMode == Extensibility.ext_ConnectMode.ext_cm_UISetup) {
				DoUISetup();
			}
#if SENDTO
			if(connectMode == Extensibility.ext_ConnectMode.ext_cm_Startup || connectMode == Extensibility.ext_ConnectMode.ext_cm_AfterStartup){
				RefreshSendToMenu();
			}
#endif
		}


		/// <summary>
		/// Implements the OnDisconnection method of the IDTExtensibility2 interface.
		/// Receives notification that the Add-in is being unloaded.
		/// </summary>
		/// <param name="disconnectMode">Describes how the Add-in is being unloaded.</param>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		/// <seealso class="IDTExtensibility2" />
		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom) {
		}


		/// <summary>
		/// Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
		/// Receives notification that the collection of Add-ins has changed.
		/// </summary>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		/// <seealso class="IDTExtensibility2" />
		public void OnAddInsUpdate(ref System.Array custom) {
		}


		/// <summary>
		/// Implements the OnStartupComplete method of the IDTExtensibility2 interface.
		/// Receives notification that the host application has completed loading.
		/// </summary>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		/// <seealso class="IDTExtensibility2" />
		public void OnStartupComplete(ref System.Array custom) {
		}


		/// <summary>
		/// Implements the OnBeginShutdown method of the IDTExtensibility2 interface.
		/// Receives notification that the host application is being unloaded.
		/// </summary>
		/// <param name="custom">Array of parameters that are host application specific.</param>
		/// <seealso class="IDTExtensibility2" />
		public void OnBeginShutdown(ref System.Array custom) {
		}


		/// <summary>
		/// Implements the QueryStatus method of the IDTCommandTarget interface.
		/// This is called when the command's availability is updated
		/// </summary>
		/// <param name="commandName">The name of the command to determine state for.</param>
		/// <param name="neededText">Text that is needed for the command.</param>
		/// <param name="status">The state of the command in the user interface.</param>
		/// <param name="commandText">Text requested by the neededText parameter.</param>
		/// <seealso class="Exec" />
		public void QueryStatus(string commandName, EnvDTE.vsCommandStatusTextWanted neededText, ref EnvDTE.vsCommandStatus status, ref object commandText) {
			if(neededText == EnvDTE.vsCommandStatusTextWanted.vsCommandStatusTextWantedNone) {
				vsCommandStatus availableStatus = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported|vsCommandStatus.vsCommandStatusEnabled;
				vsCommandStatus disabledStatus = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported;
				bool solnExplIsVisible = Utility.WindowIsVisible(applicationObject, Constants.vsWindowKindSolutionExplorer);
				Array selectedItems = null;

				Debug.WriteLine("Querying command " + commandName, "Connect.QueryStatus()");
#if SENDTO
				if(commandName.StartsWith(CommandPrefixSendTo)){
					// For a SendTo command to be available, all selected items in
					// the Solution Explorer have to be able to be sent to the command
					// and the Solution Explorer window must be visible
					if(Utility.WindowIsVisible(applicationObject, EnvDTE.Constants.vsWindowKindSolutionExplorer)){
						// Get the selected items
						selectedItems = Utility.GetSelectedUIHierarchyItems(applicationObject, EnvDTE.Constants.vsWindowKindSolutionExplorer);

						if(selectedItems.Length < 1){
							// We have to have at least one item selected
							status = disabledStatus;
						}
						else{
							// Check all selected items - if any can't get the associated filename,
							// this command will be disabled.
							bool sendToItemEnabled = true;
							foreach(object o in selectedItems) {
								string filename = Utility.GetUIHierarchyItemFilename((UIHierarchyItem)o, false);
								if(filename == ""){
									sendToItemEnabled = false;
									break;
								}
							}

							if(sendToItemEnabled){
								// We found all the filenames; this item is enabled
								status = availableStatus;
							}
							else{
								// We couldn't find one or more filenames; this item is disabled
								status = disabledStatus;
							}
						}
					}
					else{
						// The Solution Explorer window is hidden; this command is disabled
						status = disabledStatus;
					}
				}
				else{
#endif
					// Handle other Solvent commands
					switch(commandName){
						case "Solvent.Connect.RecurseExpandContract":
							status = (solnExplIsVisible ? availableStatus : disabledStatus);
							break;
						case "Solvent.Connect.ToggleRecursiveExpansion":
							status = (solnExplIsVisible ? availableStatus : disabledStatus);
							break;
						case "Solvent.Connect.OpenSelectedItemContainingFolder":
						case "Solvent.Connect.CmdInSelectedItemContainingFolder":
							if(!solnExplIsVisible || GetSelectedItemContainingFolder() == ""){
								status = disabledStatus;
							}
							else{
								status = availableStatus;
							}
							break;
						case "Solvent.Connect.OpenSelectedProjectContainingFolder":
						case "Solvent.Connect.CmdInSelectedProjectContainingFolder":
							if(!solnExplIsVisible || GetSelectedItemContainingFolder() == ""){
								status = disabledStatus;
							}
							else{
								status = availableStatus;
							}
							break;
						case "Solvent.Connect.OpenAllSubItems":
							selectedItems = Utility.GetSelectedUIHierarchyItems(applicationObject, Constants.vsWindowKindSolutionExplorer);
							if(!solnExplIsVisible || selectedItems == null || selectedItems.GetLength(0) < 1){
								status = disabledStatus;
							}
							else{
								status = availableStatus;
							}
							break;
						default:
							Debug.WriteLine("Unhandled QueryStatus: " + commandName, "Connect.QueryStatus()");
							break;
					}
#if SENDTO
				}
#endif
				Debug.WriteLine("Status set to " + status.ToString(), "Connect.QueryStatus()");
			}
		}


		/// <summary>
		/// Implements the Exec method of the IDTCommandTarget interface.
		/// This is called when the command is invoked.
		/// </summary>
		/// <param name="commandName">The name of the command to execute.</param>
		/// <param name="executeOption">Describes how the command should be run.</param>
		/// <param name="varIn">Parameters passed from the caller to the command handler.</param>
		/// <param name="varOut">Parameters passed from the command handler to the caller.</param>
		/// <param name="handled">Informs the caller if the command was handled or not.</param>
		/// <seealso class="Exec" />
		public void Exec(string commandName, EnvDTE.vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled) {
			handled = false;
			if(executeOption == EnvDTE.vsCommandExecOption.vsCommandExecOptionDoDefault) {
				Debug.WriteLine("Handling command " + commandName, "Connect.Exec()");
#if SENDTO
				if(commandName.StartsWith(CommandPrefixSendTo)){
					// Handle SendTo Commands

					// Get the SendTo item based on the index of the command
					int sendToIndex = Int32.Parse(commandName.Replace(CommandPrefixSendTo, ""));
					SendToMenuItem sendToItem = sendToMenuItems[sendToIndex];
					Debug.WriteLine("SendTo command: " + sendToItem.DisplayName, "Connect.Exec()");

					// Get the selected items
					Array selectedItems = Utility.GetSelectedUIHierarchyItems(applicationObject, EnvDTE.Constants.vsWindowKindSolutionExplorer);
					foreach(object o in selectedItems) {
                    	string filename = Utility.GetUIHierarchyItemFilename((UIHierarchyItem)o, false);
						if(filename != ""){
							// Pass valid selected items to the SendTo command
							sendToItem.ExecuteSendTo(filename);
						}
					}
					handled = true;
				}
				else{
#endif
					// Handle non-SendTo commands
					switch(commandName){
						case "Solvent.Connect.RecurseExpandContract":
							RecurseExpandContract();
							handled = true;
							break;
						case "Solvent.Connect.ToggleRecursiveExpansion":
							ToggleRecursiveExpansion();
							handled = true;
							break;
						case "Solvent.Connect.OpenSelectedItemContainingFolder":
							OpenSelectedItemContainingFolder();
							handled = true;
							break;
						case "Solvent.Connect.OpenSelectedProjectContainingFolder":
							OpenSelectedProjectContainingFolder();
							handled = true;
							break;
						case "Solvent.Connect.CmdInSelectedItemContainingFolder":
							CmdInSelectedItemContainingFolder();
							handled = true;
							break;
						case "Solvent.Connect.CmdInSelectedProjectContainingFolder":
							CmdInSelectedProjectContainingFolder();
							handled = true;
							break;
						case "Solvent.Connect.OpenAllSubItems":
							OpenAllSubItems();
							handled = true;
							break;
						default:
							Debug.WriteLine("Unhandled command: " + commandName, "Connect.Exec()");
							break;
					}
#if SENDTO
				}
#endif
			}
		}

		#endregion


		#region UI Setup

		/// <summary>
		/// Sets up the UI elements of the Add-in (command bars, etc.)
		/// </summary>
		protected void DoUISetup(){
			Debug.WriteLine("Setting up UI elements", "Connect.DoUISetup()");

			// Create/get required objects
			object []contextGUIDS = new object[] { };
			Commands commands = applicationObject.Commands;
			_CommandBars commandBars = applicationObject.CommandBars;

			// Create the commands and UI elements
			try {
				// Create commands
				Command cmdToggleRecursiveExpansion = commands.AddNamedCommand(addInInstance, "ToggleRecursiveExpansion", rm.GetString("ToggleRecursiveExpansion"), rm.GetString("ToggleRecursiveExpansion.Description"), false, 102, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
				Command cmdRecurseExpandContract = commands.AddNamedCommand(addInInstance, "RecurseExpandContract", rm.GetString("RecurseExpandContract"), rm.GetString("RecurseExpandContract.Description"), false, 102, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
				Command cmdOpenSelectedItemContainingFolder = commands.AddNamedCommand(addInInstance, "OpenSelectedItemContainingFolder", rm.GetString("OpenSelectedItemContainingFolder"), rm.GetString("OpenSelectedItemContainingFolder.Description"), false, 104, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
				Command cmdOpenSelectedProjectContainingFolder = commands.AddNamedCommand(addInInstance, "OpenSelectedProjectContainingFolder", rm.GetString("OpenSelectedProjectContainingFolder"), rm.GetString("OpenSelectedProjectContainingFolder.Description"), false, 104, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
				Command cmdOpenAllSubItems = commands.AddNamedCommand(addInInstance, "OpenAllSubItems", rm.GetString("OpenAllSubItems"), rm.GetString("OpenAllSubItems.Description"), false, 103, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
				Command cmdCmdInSelectedProjectContainingFolder = commands.AddNamedCommand(addInInstance, "CmdInSelectedProjectContainingFolder", rm.GetString("CmdInSelectedProjectContainingFolder"), rm.GetString("CmdInSelectedProjectContainingFolder.Description"), false, 105, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
				Command cmdCmdInSelectedItemContainingFolder = commands.AddNamedCommand(addInInstance, "CmdInSelectedItemContainingFolder", rm.GetString("CmdInSelectedItemContainingFolder"), rm.GetString("CmdInSelectedItemContainingFolder.Description"), false, 105, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);

				// Add the Tools menu selections
				Debug.WriteLine("Adding items to Tools menu", "Connect.DoUISetup()");
				CommandBar toolsMenu = (CommandBar)commandBars["Tools"];
				CommandBar solventToolsMenu = (CommandBar)commands.AddCommandBar(rm.GetString("ToolBarMain"), EnvDTE.vsCommandBarType.vsCommandBarTypeMenu, toolsMenu, 1);
				CommandBarControl toolMenuButton = cmdToggleRecursiveExpansion.AddControl(solventToolsMenu, solventToolsMenu.Controls.Count + 1);
				toolMenuButton = cmdOpenAllSubItems.AddControl(solventToolsMenu, solventToolsMenu.Controls.Count + 1);


				// Add commands to the Solution Explorer context menu

				// Add commands to folder/container menus
				Debug.WriteLine("Adding commands to container menus", "Connect.OnConnection()");
				string[] containerMenuTitles = new string[]{
															   "Folder",
															   "Solution",
															   "Project",
															   "Reference Root",
															   "Web Reference Folder",
															   "Dependency node"
														   };
				Debug.Indent();
				foreach(string contextMenuTitle in containerMenuTitles){
					Debug.WriteLine("Setting up context menu: " + contextMenuTitle, "Connect.DoUISetup()");
					CommandBar contextMenu = (CommandBar)commandBars[contextMenuTitle];
					CommandBar solventContextMenu = (CommandBar)commands.AddCommandBar(rm.GetString("ContextMenuMain"), EnvDTE.vsCommandBarType.vsCommandBarTypeMenu, contextMenu, contextMenu.Controls.Count + 1);
					CommandBarControl contextMenuButton = cmdRecurseExpandContract.AddControl(solventContextMenu, solventContextMenu.Controls.Count + 1);
					switch(contextMenuTitle){
						case "Folder":
							contextMenuButton = cmdOpenSelectedItemContainingFolder.AddControl(solventContextMenu, solventContextMenu.Controls.Count + 1);
							contextMenuButton = cmdCmdInSelectedItemContainingFolder.AddControl(solventContextMenu, solventContextMenu.Controls.Count + 1);
							contextMenuButton = cmdOpenAllSubItems.AddControl(solventContextMenu, solventContextMenu.Controls.Count + 1);
							break;
						case "Project":
							contextMenuButton = cmdOpenSelectedProjectContainingFolder.AddControl(solventContextMenu, solventContextMenu.Controls.Count + 1);
							contextMenuButton = cmdCmdInSelectedProjectContainingFolder.AddControl(solventContextMenu, solventContextMenu.Controls.Count + 1);
							contextMenuButton = cmdOpenAllSubItems.AddControl(solventContextMenu, solventContextMenu.Controls.Count + 1);
							break;
					}
				}
				Debug.Unindent();

				// Add commands to item menu
				Debug.WriteLine("Adding commands to Item menu", "Connect.DoUISetup()");
				CommandBar itemContextMenu = (CommandBar)commandBars["Item"];
				CommandBar solventItemContextMenu = (CommandBar)commands.AddCommandBar(rm.GetString("ContextMenuMain"), EnvDTE.vsCommandBarType.vsCommandBarTypeMenu, itemContextMenu, itemContextMenu.Controls.Count + 1);
				CommandBarControl solventItemContextMenuButton = cmdOpenSelectedItemContainingFolder.AddControl(solventItemContextMenu, solventItemContextMenu.Controls.Count + 1);
				solventItemContextMenuButton = cmdCmdInSelectedItemContainingFolder.AddControl(solventItemContextMenu, solventItemContextMenu.Controls.Count + 1);
			}
			catch(System.Exception err) {
				Debug.WriteLine("Error adding commands to UI: " + err.ToString(), "Connect.DoUISetup()");
			}
#if SENDTO
			// Last but not least, refresh the SendTo menu
			RefreshSendToMenu();
#endif
		}


#if SENDTO
		/// <summary>
		/// Generates/regenerates the SendTo menu.
		/// </summary>
		protected void RefreshSendToMenu(){
			Debug.WriteLine("Starting RefreshSendToMenu", "Connect.RefreshSendToMenu()");

			// Set up required objects
			object []contextGUIDS = new object[] { };
			Commands commands = applicationObject.Commands;
			_CommandBars commandBars = applicationObject.CommandBars;

			// Remove any existing SendTo commands
			foreach(Command c in commands){
				try{
					if(c.Name != null && c.Name.StartsWith(CommandPrefixSendTo)){
						Debug.WriteLine("Deleting SendTo command " + c.Name, "Connect.RefreshSendToMenu()");
						try{
							c.Delete();
						}
						catch(Exception err){
							Debug.WriteLine("Error deleting SendTo command " + c.Name + ": " + err.Message, "Connect.RefreshSendToMenu()");
						}
					}
				}
				catch(Exception badCommandException){
						Debug.WriteLine("Error getting command name for " + c.Guid + ": " + badCommandException.Message, "Connect.RefreshSendToMenu()");
				}
			}

			// Get the updated SendTo menu items
			sendToMenuItems = SendToMenu.GetSendToMenu();

			// Build the new array of VS commands from SendTo items
			Command[] sendToMenuCommands = new Command[sendToMenuItems.Length];
			for (int i = 0; i < sendToMenuItems.Length; i++) {
            	sendToMenuCommands[i] = commands.AddNamedCommand(addInInstance, "SendTo" + i.ToString(), sendToMenuItems[i].DisplayName, sendToMenuItems[i].DisplayName, true, 3277, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
			}

#if DEBUG
			// List the new SendTo commands
			foreach(Command c in commands){
				try{
					if(c.Name != null && c.Name.StartsWith(CommandPrefixSendTo)){
						Debug.WriteLine("SendTo command: " + c.Name, "Connect.RefreshSendToMenu()");
					}
				}
				catch(Exception badCommandException){
					Debug.WriteLine("Error getting command name for " + c.Guid + ": " + badCommandException.Message, "Connect.RefreshSendToMenu()");
				}
			}
#endif

			// Add the SendTo items and/or menus
			Debug.WriteLine("Adding items to SendTo menus", "Connect.RefreshSendToMenu()");
			string[] baseContextMenuTitles = new string[]{
														   "Item",
														   "Folder",
														   "Project"
													   };

			Debug.Indent();
			foreach(string baseContextMenuTitle in baseContextMenuTitles){
				// Get the base context menu
				CommandBar contextMenu = (CommandBar)commandBars[baseContextMenuTitle];
				CommandBar solventContextMenu = null;
				try{
					foreach(CommandBarControl c in contextMenu.Controls) {
						if(c.Caption == rm.GetString("ContextMenuMain")){
							// This is the Solvent menu
							solventContextMenu = ((CommandBarPopup)c).CommandBar;
							break;
						}
 					}

				}
				catch(Exception noSolventMenuException){
					Debug.WriteLine("Unable to get Solvent menu under " + baseContextMenuTitle + ": " + noSolventMenuException.Message, "Connect.RefreshSendToMenu()");
					solventContextMenu = null;
				}
				if(solventContextMenu != null){
					// We have the Solvent menu
					Debug.WriteLine("Found Solvent menu for " + baseContextMenuTitle, "Connect.RefreshSendToMenu()");
					CommandBar sendToContextMenu = null;
					// Find the SendTo menu
					try{
						foreach(CommandBarControl c in solventContextMenu.Controls) {
							if(c.Caption == rm.GetString("SendTo")){
								// This is the SendTo menu
								sendToContextMenu = ((CommandBarPopup)c).CommandBar;
								break;
							}
						}

					}
					catch(Exception noSendToMenuException){
						Debug.WriteLine("Unable to get SendTo menu under " + baseContextMenuTitle + ": " + noSendToMenuException.Message, "Connect.RefreshSendToMenu()");
						sendToContextMenu = null;
					}
					if(sendToContextMenu == null){
						// Create the SendTo context menu
						Debug.WriteLine("Creating a new SendTo menu for " + baseContextMenuTitle, "Connect.RefreshSendToMenu()");
						sendToContextMenu = (CommandBar)commands.AddCommandBar(rm.GetString("SendTo"), EnvDTE.vsCommandBarType.vsCommandBarTypeMenu, solventContextMenu, solventContextMenu.Controls.Count + 1);
					}
					if(sendToContextMenu != null){
						// Add the commands to the SendTo menu
						Debug.WriteLine("Found SendTo menu for " + baseContextMenuTitle, "Connect.RefreshSendToMenu()");
						Debug.Indent();
						for (int i = 0; i < sendToMenuCommands.Length; i++) {
							CommandBarControl sendToMenuButton = sendToMenuCommands[i].AddControl(sendToContextMenu, sendToContextMenu.Controls.Count + 1);
							Debug.WriteLine("Added SendTo menu button for " + baseContextMenuTitle + ": " + sendToMenuButton.Caption, "Connect.RefreshSendToMenu()");
						}
						Debug.Unindent();
					}
				}
			}
			Debug.Unindent();

			Debug.WriteLine("Finishing RefreshSendToMenu", "Connect.RefreshSendToMenu()");
		}
#endif

		#endregion


		#region RecurseExpandContract/ToggleRecursiveExpansion



		/// <summary>
		/// Toggles recursive expand/contract for all selected items in the Solution Explorer.
		/// </summary>
		/// <remarks>
		/// <para>
		/// The "<see cref="RecurseExpandContract"/>" and "ToggleRecursiveExpansion" commands handle similar
		/// functions: they recursively expand or contract items in the Solution Explorer
		/// window.  "<see cref="RecurseExpandContract"/>" is called from the right-click menu of the Solution
		/// Explorer window, while "ToggleRecursiveExpansion" is called from the Tools menu.
		/// </para>
		///
		/// <para>
		/// Normally "<see cref="RecurseExpandContract"/>" might only act on the item in the Solution Explorer
		/// that was right-clicked, but the user can select multiple items and right-click, so
		/// we'll act on all selected items; "ToggleRecursiveExpansion" should act on all selected
		/// items in the Solution Explorer by default.
		/// </para>
		/// </remarks>
		public void ToggleRecursiveExpansion(){
			Debug.WriteLine("Starting ToggleRecursiveExpansion", "Connect.ToggleRecursiveExpansion()");

			// Get the currently selected items in Solution Explorer
			Array selectedItems = Utility.GetSelectedUIHierarchyItems(applicationObject, Constants.vsWindowKindSolutionExplorer);
			if(selectedItems == null){
				return;
			}

			// Recursively expand/contract the entire hierarchy based on
			// the current state of the tree at the selected item node
			foreach(UIHierarchyItem uiHierarchyItem in selectedItems){
				UIHierarchyItems uiHierarchyItems = uiHierarchyItem.UIHierarchyItems;
				if(uiHierarchyItems != null && uiHierarchyItems.Count > 0){
					// Determine whether we're expanding or contracting the selected node
					bool expand = !(uiHierarchyItem.UIHierarchyItems.Expanded);
					DoRecursiveExpandContract(uiHierarchyItem, expand);
				}
			}

			Debug.WriteLine("Ending ToggleRecursiveExpansion", "Connect.ToggleRecursiveExpansion()");
		}


		/// <summary>
		/// Toggles recursive expand/contract for all selected items in the Solution Explorer from
		/// the Solution Explorer context menu.
		/// </summary>
		/// <seealso cref="ToggleRecursiveExpansion"/>
		public void RecurseExpandContract(){
			ToggleRecursiveExpansion();
		}


		/// <summary>
		/// Performs the recursive expand/contract on a given node in the UI Hierarchy.
		/// </summary>
		/// <param name="uihItem">The <see cref="UIHierarchyItem"/> that serves as the root for the expansion.</param>
		/// <param name="expand">True to expand the item; false to contract it.</param>
		protected void DoRecursiveExpandContract(UIHierarchyItem uihItem, bool expand){
			// Get the sub items
			UIHierarchyItems uiHierarchyItems = uihItem.UIHierarchyItems;

			// If there are no sub items to deal with, exit
			if(uiHierarchyItems == null || uiHierarchyItems.Count < 1){
				return;
			}

			// Expand/contract the stuff below this node first
			foreach(UIHierarchyItem uiHierarchyItem in uiHierarchyItems){
				DoRecursiveExpandContract(uiHierarchyItem, expand);
			}

			// Expand/contract this node
			uihItem.UIHierarchyItems.Expanded = expand;
		}

		#endregion


		#region OpenSelectedItemContainingFolder/OpenSelectedProjectContainingFolder

		/// <summary>
		/// Gets the full path to the currently selected item in the Solution Explorer
		/// </summary>
		/// <returns>The path to the currently selected item in the Solution Explorer.  For folders,
		/// this is the folder path; for web references, this is the path to the folder containing the WSDL;
		/// for files, this is the path to the containing folder.  Empty string if there's an error or if the
		/// selection is invalid.</returns>
		protected string GetSelectedItemContainingFolder(){
			Debug.WriteLine("Starting GetSelectedItemContainingFolder", "Connect.GetSelectedItemContainingFolder()");

			// Get the currently selected items in Solution Explorer and ensure we have exactly one item
			Array selectedItems = Utility.GetSelectedUIHierarchyItems(applicationObject, Constants.vsWindowKindSolutionExplorer);
			if(selectedItems == null || selectedItems.GetLength(0) != 1){
				Debug.WriteLine("No selected items; exiting early", "Connect.GetSelectedItemContainingFolder()");
				return "";
			}

			// Get the selected item
			UIHierarchyItem uihItem = (UIHierarchyItem)(selectedItems.GetValue(selectedItems.GetLowerBound(0)));

			// Get the path from the selected item and return it
			Debug.WriteLine("Ending GetSelectedItemContainingFolder", "Connect.GetSelectedItemContainingFolder()");
			return Utility.GetUIHierarchyItemFilename(uihItem, true);
		}


		/// <summary>
		/// Opens the containing folder for the selected item in Solution Explorer.
		/// </summary>
		public void OpenSelectedItemContainingFolder(){
			Debug.WriteLine("Starting OpenSelectedItemContainingFolder", "Connect.OpenSelectedItemContainingFolder()");
			string containingFolderPath = GetSelectedItemContainingFolder();
			if(containingFolderPath == ""){
				Debug.WriteLine("No containing folder path; exiting early", "Connect.OpenSelectedItemContainingFolder()");
				return;
			}
			Utility.OpenItemInExplorer(containingFolderPath);
			Debug.WriteLine("Ending OpenSelectedItemContainingFolder", "Connect.OpenSelectedItemContainingFolder()");
		}


		/// <summary>
		/// Opens the containing folder for the selected project in Solution Explorer.
		/// </summary>
		public void OpenSelectedProjectContainingFolder(){
			Debug.WriteLine("Starting OpenSelectedProjectContainingFolder", "Connect.OpenSelectedProjectContainingFolder()");
			string containingFolderPath = GetSelectedItemContainingFolder();
			if(containingFolderPath == ""){
				Debug.WriteLine("No containing folder path; exiting early", "Connect.OpenSelectedProjectContainingFolder()");
				return;
			}
			Utility.OpenItemInExplorer(containingFolderPath);
			Debug.WriteLine("Ending OpenSelectedProjectContainingFolder", "Connect.OpenSelectedProjectContainingFolder()");
		}

		#endregion


		#region OpenAllSubItems

		/// <summary>
		/// Opens all subitems, recursively, located below the currently selected item
		/// in the Solution Explorer window.
		/// </summary>
		public void OpenAllSubItems(){
			Debug.WriteLine("Starting OpenAllSubItems", "Connect.OpenAllSubItems()");

			// Get the currently selected items in Solution Explorer and ensure we have at least one item
			Array selectedItems = Utility.GetSelectedUIHierarchyItems(applicationObject, Constants.vsWindowKindSolutionExplorer);
			if(selectedItems == null || selectedItems.GetLength(0) < 1){
				Debug.WriteLine("No selected items; exiting early", "Connect.OpenAllSubItems()");
				return;
			}

			foreach(UIHierarchyItem uihItem in selectedItems){
				try{
					ProjectItems itemsToOpen = null;

					// Get the project items
					// TODO: Change to "if <object> is <type>" style of checking
					ProjectItem projItem = uihItem.Object as ProjectItem;
					Project proj = uihItem.Object as Project;
					if(projItem != null){
						itemsToOpen = projItem.ProjectItems;
					}
					else if(proj != null){
						itemsToOpen = proj.ProjectItems;
					}
					else{
						Debug.WriteLine("Hierarchy item is neither a ProjectItem nor a Project; no items to open", "Connect.OpenAllSubItems()");
						itemsToOpen = null;
					}

					// Open all sub items
					RecursiveOpenDocs(itemsToOpen);
				}
				catch(Exception err){
					Debug.WriteLine("Error getting project items for " + uihItem.Name + "; " + err.Message, "Connect.OpenAllSubItems()");
				}
			}

			Debug.WriteLine("Ending OpenAllSubItems", "Connect.OpenAllSubItems()");


		}


		/// <summary>
		/// Opens all items (and subitems) in a <see cref="ProjectItems"/> collection.
		/// </summary>
		/// <param name="projItems">The collection of items to open (including all subitems therein).</param>
		protected void RecursiveOpenDocs(ProjectItems projItems){
			// If there are no items to open, return
			if(projItems == null || projItems.Count < 1){
				return;
			}

			// Iterate through the collection of project items and open each one
			foreach(EnvDTE.ProjectItem projItem in projItems){
				// Open subitems of the current item, then the current item
				try{
					RecursiveOpenDocs(projItem.ProjectItems);
					Window itemWin = projItem.Open(EnvDTE.Constants.vsViewKindPrimary);
					itemWin.Activate();
					Debug.WriteLine("Opened item " + projItem.Name, "Connect.RecursiveOpenDocs()");
				}
				catch(Exception err){
					Debug.WriteLine("Error opening item " + projItem.Name + ";" + err.Message, "Connect.RecursiveOpenDocs()");
				}
			}
		}

		#endregion


		#region CmdInSelectedItemContainingFolder/CmdInSelectedProjectContainingFolder

		/// <summary>
		/// Opens a command prompt in the containing folder for the selected item in Solution Explorer.
		/// </summary>
		public void CmdInSelectedItemContainingFolder(){
			Debug.WriteLine("Starting CmdInSelectedItemContainingFolder", "Connect.CmdInSelectedItemContainingFolder()");
			string containingFolderPath = GetSelectedItemContainingFolder();
			if(containingFolderPath == ""){
				Debug.WriteLine("No containing folder path; exiting early", "Connect.CmdInSelectedItemContainingFolder()");
				return;
			}
			Utility.CmdPromptHere(containingFolderPath);
			Debug.WriteLine("Ending CmdInSelectedItemContainingFolder", "Connect.CmdInSelectedItemContainingFolder()");
		}


		/// <summary>
		/// Opens a command prompt in the containing folder for the selected project in Solution Explorer.
		/// </summary>
		public void CmdInSelectedProjectContainingFolder(){
			Debug.WriteLine("Starting CmdInSelectedProjectContainingFolder", "Connect.CmdInSelectedProjectContainingFolder()");
			string containingFolderPath = GetSelectedItemContainingFolder();
			if(containingFolderPath == ""){
				Debug.WriteLine("No containing folder path; exiting early", "Connect.CmdInSelectedProjectContainingFolder()");
				return;
			}
			Utility.CmdPromptHere(containingFolderPath);
			Debug.WriteLine("Ending CmdInSelectedProjectContainingFolder", "Connect.CmdInSelectedProjectContainingFolder()");
		}

		#endregion
	}
}