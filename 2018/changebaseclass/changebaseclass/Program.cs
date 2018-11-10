using System;
using System.Windows.Forms;
using System.Linq;
using Ascon.Integration;
using Ascon.Vertical.Technology;

namespace changebaseclass
{
	internal sealed class Program
	{
		[STAThread]
		private static void Main(string[] args)
		{
			if (args.Count() < 2)
				return;
				
			string modelName    = args[0];		
			string currentClass = args[1];			
			string baseClass = (args.Count() < 3) ? "" : args[2];
			
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
						
						
			if (baseClass != String.Empty)
				model.Classes[currentClass].BaseClass = model.Classes[baseClass];
			else
				model.Classes[currentClass].BaseClass = null;
			
			model.Save(modelName);
			model.Close();
			
			AuthenticationManager.Deauthenticate();
			MessageBox.Show("Выполнено!");				
		}
	}
}