using System;

namespace Paraesthesia.WinShell.SendToMenu_Stub {
	/// <summary>
	/// Stub class for emulating the finished SendToMenu class functionality.
	/// </summary>
	public class SendToMenu {
		
		/// <summary>
		/// Gets a fake array of items as placeholders for SendTo menu items.
		/// </summary>
		/// <returns>An array of <see cref="SendToMenuItem"/> objects.</returns>
		public static SendToMenuItem[] GetSendToMenu(){
			Random rand = new Random();
			int count = rand.Next(3, 8);

			SendToMenuItem[] retVal = new SendToMenuItem[count];
			
			for(int i = 0; i < count; i++){
				retVal[i] = new SendToMenuItem("SendTo Item " + i.ToString(), null);
			}

			return retVal;

		}
	}
}
