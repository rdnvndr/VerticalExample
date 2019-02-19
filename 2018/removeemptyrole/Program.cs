using System;
using System.Linq;
using System.Windows.Forms;
using Ascon.Vertical.Technology;
using Ascon.Integration;

namespace removeemptyrole
{
	internal sealed class Program
	{
		[STAThread]
		private static void Main(string[] args)
		{
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
				var missingRoles = cls.Permissions.Where(
					p => p.Key != Guid.Empty 
					&& AuthenticationManager.Roles.All(r => r.Id != p.Key)
				).ToList();
				
				foreach (var missingRole in missingRoles)
					cls.Permissions.Clear(missingRole.Key);

				foreach (var member in cls.Members) { 
					missingRoles = member.Permissions.Where(
						p => p.Key != Guid.Empty 
						&& AuthenticationManager.Roles.All(r => r.Id != p.Key)
					).ToList();
					foreach (var missingRole in missingRoles) 
						member.Permissions.Clear(missingRole.Key);
				}
			}
			model.Save(modelName);
			model.Close();
			
			AuthenticationManager.Deauthenticate();
			MessageBox.Show("Выполнено!");				
		}
	}
}