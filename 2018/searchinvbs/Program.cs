using System;
using System.Windows.Forms;
using Ascon.Integration;
using Ascon.Vertical.Technology;
using System.Text.RegularExpressions;

namespace searchinvbs
{
	internal sealed class Program
	{
		[STAThread]
		private static void Main(string[] args)
		{
			try {
				AuthenticationManager.Authenticate();			
			} catch (Exception) {
				MessageBox.Show("Ошибка авторизации!");
				return;
			}
			
			TechModel model;
			try {
				model = TechModel.Load("d:\\structure.vtp");
			} catch (Exception) {
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка загрузки модели!");
				return;
			}
			
			System.IO.StreamWriter logFile;
			try {
				logFile = new System.IO.StreamWriter(@"D:\log.txt");
			} catch (Exception) {
				model.Close();
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка создания файла!");
				return;
			}
			
			int count = 0;
			foreach (TechClass cls in model.Classes) {			
				foreach (TechClassMember mbr in cls.Members) {
					var prop = mbr as TechClassProperty;
					if (prop != null) {
							var getter = model.VBSFunctions[model.VBSFunctions.GetPropertyGetterName(prop)];						
							var setter = model.VBSFunctions[model.VBSFunctions.GetPropertySetterName(prop)];
							if (getter != String.Empty || setter != String.Empty) {
								logFile.WriteLine(cls.Name + "." + mbr.Name);
								logFile.WriteLine("-----------------------------------------------------------");
								logFile.WriteLine(getter);
								logFile.WriteLine(setter);
							}
					}  
					
					var func = mbr as TechClassFunction;
					if (func != null) {
						var body = model.VBSFunctions[model.VBSFunctions.GetFunctionName(func)];
						if (body != String.Empty) {
							logFile.WriteLine(cls.Name + "." + mbr.Name);
							logFile.WriteLine("-----------------------------------------------------------");
							logFile.WriteLine(body);
						}
					}
				}
			}
			model.Close();
			AuthenticationManager.Deauthenticate();
			logFile.Close();
			MessageBox.Show("Выполнено!");				
		}
	}
}