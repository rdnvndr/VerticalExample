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
			string markerName = "isObsolete";
			
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

                        System.IO.StreamWriter logFile;
			try {
				logFile = new System.IO.StreamWriter(@"D:\commentlist.txt");
			} catch (Exception) {
				model.Close();
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка создания файла!");
				return;
			}

			foreach (var cls in model.Classes)
			{
				foreach (var marker in cls.Markers) {
					if (marker.Key.Name == markerName) {
                                                logFile.WriteLine("    " + cls.Name + ": " + cls.Markers.Get(markerName).ToString());
					}
				}
				
				foreach (var member in cls.Members) 
				{
					foreach (var marker in member.Markers) {
						if (marker.Key.Name == markerName) {
                                                       logFile.WriteLine("  " + cls.Name + "." + member.Name + ": " + member.Markers.Get(markerName).ToString());
						}
					}
						
				}
			}
			model.Close();
                        logFile.Close();
			
			AuthenticationManager.Deauthenticate();
			MessageBox.Show("Выполнено!");				
		}
	}
}