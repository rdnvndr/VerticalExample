using System;
using System.Linq;
using System.Windows.Forms;
using Ascon.Vertical.Technology;
using Ascon.Integration;

namespace resavetp
{
	internal sealed class Program
	{
		[STAThread]
		private static void Main(string[] args)
		{			
			if (args.Count() < 2)
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
			
			for (int i = args.Count() - 1; i>0; --i) {
				try {
					TechDocument document = TechDocument.Load(args[i], model);
					document.Save(args[i]);
					document.Close();
				} catch(Exception) {
					MessageBox.Show("Ошибка загрузки документа!");		
					continue;
				}
				
			}
			model.Close();
			
			AuthenticationManager.Deauthenticate();
			MessageBox.Show("Выполнено!");				
		}
	}
}
