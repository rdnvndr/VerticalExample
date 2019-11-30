using System;
using System.Windows.Forms;
using Ascon.Integration;
using Ascon.Vertical.Technology;
using System.Text.RegularExpressions;

namespace listoftostring
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
				logFile = new System.IO.StreamWriter(@"D:\tostringlist.txt");
			} catch (Exception) {
				model.Close();
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка создания файла!");
				return;
			}
			
			logFile.WriteLine("СПИСОК ФУНКЦИЙ ToString:");
			logFile.WriteLine("--------------------------------------");
			foreach (TechClass cls in model.Classes) {
				if (cls.ToStringFunction.Body != String.Empty) {
					logFile.WriteLine(cls.Name);
                			logFile.WriteLine("--------------------------------------");
                                        logFile.WriteLine(cls.ToStringFunction.Body);
                			logFile.WriteLine("--------------------------------------");
                                        logFile.WriteLine();

				}					
			}
			model.Close();
			AuthenticationManager.Deauthenticate();
			logFile.Close();
			MessageBox.Show("Выполнено!");				
		}
	}
}