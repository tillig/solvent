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
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Solvent {
	/// <summary>
	/// Custom install/uninstall actions for the Solvent setup routine.
	/// </summary>
	[RunInstaller(true)]
	public class CustomInstallActions : System.Configuration.Install.Installer {
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Resources.ResourceManager rm = null;

		/// <summary>
		/// Creates and initializes the CustomInstallActions object.
		/// </summary>
		public CustomInstallActions() {
			// This call is required by the Designer.
			InitializeComponent();

			// Get a string resource manager
			rm = new System.Resources.ResourceManager("Solvent.Strings", this.GetType().Assembly);
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing ) {
			if( disposing ) {
				if(components != null) {
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}


		#region Component Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			components = new System.ComponentModel.Container();
		}
		#endregion


		/// <summary>
		/// Determines if Visual Studio is running
		/// </summary>
		/// <returns>True if VS (devenv.exe) is running; false otherwise</returns>
		private bool vsIsRunning(){
			bool retVal = false;
			Process[] processes = Process.GetProcesses();
			foreach(Process p in processes){
				try{
					if((p != null) && (p.MainModule != null)){
						string fileName = System.IO.Path.GetFileName(p.MainModule.FileName);
						if(String.Compare(fileName, "devenv.exe") == 0){
							retVal = true;
							break;
						}
					}
				}
				catch(Exception err){
					Trace.WriteLine("Error checking if Visual Studio is running; " + err.Message, "CustomInstallActions.vsIsRunning()");
				}
			}
			return retVal;
		}


		/// <summary>
		/// Checks to make sure VS.NET is not currently running when install happens
		/// </summary>
		/// <param name="stateSaver"></param>
		public override void Install(IDictionary stateSaver) {

#if !DEBUG
			// Ensure VS.NET isn't running
			while(vsIsRunning()){
				DialogResult dr = MessageBox.Show("You must close Visual Studio .NET before installing this product.  Please shut down all instances of VS and retry, or cancel the installation.", "VS.NET Running", MessageBoxButtons.RetryCancel);
				if(dr == DialogResult.Cancel){
					throw new Exception(rm.GetString("Installation.Exceptions.MustCloseVS"));
				}
			}
#endif

			// No VS instance; OK to install
			base.Install (stateSaver);
		}



		/// <summary>
		/// Uninstalls custom UI that was placed by the add-in
		/// </summary>
		/// <param name="savedState"></param>
		public override void Uninstall(IDictionary savedState) {
			Trace.WriteLine("Starting uninstall procedure", "CustomInstallActions.Uninstall()");
			DialogResult dr = DialogResult.OK;
			// Ensure VS.NET isn't running
			while(vsIsRunning()){
				dr = MessageBox.Show("Visual Studio .NET is still running.  Please shut down all instances of Visual Studio .NET to proceed, then click Retry.  Click Abort to cancel the uninstallation.  Click Ignore to finish the uninstallation and leave custom UI in place.", "VS.NET Still Running", MessageBoxButtons.AbortRetryIgnore);
				if(dr == DialogResult.Abort || dr == DialogResult.Ignore){
					Trace.WriteLine("User selected " + dr.ToString() + "; exiting VS check", "CustomInstallActions.Uninstall()");
					break;
				}
			}

			// Based on dialog result, perform uninstall actions
			switch(dr){
				case DialogResult.Abort:
					break;
				case DialogResult.Ignore:
					Trace.WriteLine("Running base uninstallation", "CustomInstallActions.Uninstall()");
					base.Uninstall(savedState);
					break;
				case DialogResult.OK:
				case DialogResult.Retry:
					Trace.WriteLine("Running base uninstallation", "CustomInstallActions.Uninstall()");
					base.Uninstall(savedState);
					EnvDTE.DTE dte = null;
					try{
						// Remove custom UI items, one for each command
						Trace.WriteLine("Removing custom UI elements", "CustomInstallActions.Uninstall()");
						Type dteType = Type.GetTypeFromProgID("VisualStudio.DTE.7.1");
						dte = System.Activator.CreateInstance(dteType, true) as EnvDTE.DTE;
						if(dte != null){
							// TODO: Handle removal of SendTo commands (just remove everything with Solvent.Connect at the start?)
							// Iterate through the collection of commands and remove them all
							foreach(string commandName in Connect.CommandList){
								try{
									dte.Commands.Item(Connect.CommandPrefix + commandName, -1).Delete();
								}
								catch(Exception excRemoving){
									Trace.WriteLine("Error removing command " + commandName + "; " + excRemoving.Message, "CustomInstallActions.Uninstall()");
								}
							}
						}
						else{
							Trace.WriteLine("Unable to get DTE reference.", "CustomInstallActions.Uninstall()");
						}
					}
					catch(Exception err){
						Trace.WriteLine("Error removing add-in UI elements: " + err.Message, "CustomInstallActions.Uninstall()");
					}
					finally{
						if(dte != null){
							dte.Quit();
							Marshal.ReleaseComObject(dte);
						}
					}
					break;
				default:
					break;
			}
			Trace.WriteLine("Finished uninstall procedure", "CustomInstallActions.Uninstall()");
		}

	}
}
