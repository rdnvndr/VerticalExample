using System;
using System.Windows.Forms;
using System.Linq;
using Ascon.Integration;
using Ascon.Vertical.Technology;

namespace namelist
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
				logFile = new System.IO.StreamWriter(@"D:\namelist.txt");
			} catch (Exception) {
				model.Close();
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка создания файла!");
				return;
			}
			
			foreach (TechClass cls in model.Classes) {
				if (cls.DisplayName != String.Empty) {
					logFile.WriteLine("    " + cls.Name + "(Экранное имя): " + cls.DisplayName);
				}
				if (cls.Description != String.Empty) {
					logFile.WriteLine("    " + cls.Name + "(Комментарий): " + cls.Description);
				}
				foreach (var group in cls.Appearance.AttributesGroups) {
					if (group.Name != String.Empty) {
						logFile.WriteLine("    " + cls.Name + "(Группа): " +  group.Name);
					}
				}
					
				foreach (TechClassMember mbr in cls.Members) {
					if (mbr.Description != String.Empty) {
						logFile.WriteLine("    " + cls.Name + "." + mbr.Name + "(Экранное имя): " + mbr.DisplayName);
					}
					if (mbr.Description != String.Empty) {
						logFile.WriteLine("    " + cls.Name + "." + mbr.Name + "(Комментарий): " + mbr.Description);
					}
				}
			}
			
			foreach (var flt in model.Filters) {
				if (flt.DisplayName != String.Empty) {
					logFile.WriteLine("    " + flt.Name + "(Экранное имя): " + flt.DisplayName);
				}
				if (flt.Description != String.Empty) {
					logFile.WriteLine("    " + flt.Name + "(Комментарий): " + flt.Description);
				}
			}
			
			foreach (var func in model.Functions) {
				if (func.DisplayName != String.Empty) {
					logFile.WriteLine("    " + func.Name + "(Экранное имя): " + func.DisplayName);
				}
				if (func.Description != String.Empty) {
					logFile.WriteLine("    " + func.Name + "(Комментарий): " + func.Description);
				}
			}
			
			foreach (var num in model.Numerators) {
				if (num.DisplayName != String.Empty) {
					logFile.WriteLine("    " + num.Name + "(Экранное имя): " + num.DisplayName);
				}
				if (num.Description != String.Empty) {
					logFile.WriteLine("    " + num.Name + "(Комментарий): " + num.Description);
				}
			}
			
			foreach (var marker in model.Markers) {
				if (marker.DisplayName != String.Empty) {
					logFile.WriteLine("    " + marker.Name + "(Экранное имя): " + marker.DisplayName);
				}
				if (marker.Description != String.Empty) {
					logFile.WriteLine("    " + marker.Name + "(Комментарий): " + marker.Description);
				}
			}
			
			foreach (var tree in model.Presentations) {
				if (tree.DisplayName != String.Empty) {
					logFile.WriteLine("    " + tree.Name + "(Экранное имя): " + tree.DisplayName);
				}
				if (tree.Description != String.Empty) {
					logFile.WriteLine("    " + tree.Name + "(Комментарий): " + tree.Description);
				}
			}
			
			model.Close();
			AuthenticationManager.Deauthenticate();
			logFile.Close();
			MessageBox.Show("Выполнено!");				
		}
	}
}