using System;
using System.Drawing;

namespace Paraesthesia.WinShell.SendToMenu_Stub {
	/// <summary>
	/// Stub class for emulating the finished SendToMenuItem class functionality.
	/// </summary>
	public class SendToMenuItem {
		#region Variables

		private string _displayName;
		private Bitmap _icon;

		#endregion

		#region Constructors

		/// <summary>
		/// Creates a default SendToMenuItem
		/// </summary>
		/// <param name="displayName">The display name of the item.</param>
		/// <param name="icon">The icon representing the item.</param>
		public SendToMenuItem(string displayName, Bitmap icon){
			_displayName = displayName;
			_icon = icon;
		}

		#endregion

		#region Properties

		/// <summary>
		/// Gets the display name of the item.
		/// </summary>
		public string DisplayName{
			get{
				return _displayName;
			}
		}

		/// <summary>
		/// Gets the icon representing the item.
		/// </summary>
		public Bitmap Icon{
			get{
				return _icon;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// Executes the "SendTo" operation on this item for the given filename.
		/// </summary>
		/// <param name="filename">The file to "send to" this item.</param>
		public void ExecuteSendTo(string filename){
			System.Diagnostics.Debug.WriteLine("Executed SendTo: " + this.DisplayName + ", Filename: " + filename);
		}

		#endregion
	}
}
