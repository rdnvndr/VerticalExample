using System;
using System.Linq;
using System.Windows.Forms;
using Ascon.Vertical.Technology;
using Ascon.Integration;

namespace markersearch
{
	internal sealed class Program
	{
		[STAThread]
		private static void Main(string[] args)
		{
			string markerName = "ttplock";
			
			if (args.Count() < 1)
				return;
				
			string modelName    = args[0];		
			try {
				AuthenticationManager.Authenticate();
			} catch(Exception) {
				MessageBox.Show("Ошибка авторизации!");		
				return;
			}
			
			TechModel model = null;
			try {
				model	 = TechModel.Load(modelName);
			} catch(Exception) {
				MessageBox.Show("Ошибка загрузки модели!");		
				AuthenticationManager.Deauthenticate();
				return;
			}

			foreach (var cls in model.Classes)
			{
				foreach (var marker in cls.Markers) {
					if (marker.Key.Name == markerName) {
						MessageBox.Show(cls.Name);
					}
				}
				
				foreach (var member in cls.Members) 
				{
					foreach (var marker in member.Markers) {
						if (marker.Key.Name == markerName) {
							MessageBox.Show(cls.Name + "." + member.Name);
						}
					}
						
				}
			}
			model.Close();
			
			AuthenticationManager.Deauthenticate();
			MessageBox.Show("Выполнено!");				
		}
	}
}