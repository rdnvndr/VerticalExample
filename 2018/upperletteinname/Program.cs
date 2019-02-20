using System;
using System.Windows.Forms;
using Ascon.Integration;
using System.Text.RegularExpressions;
using Ascon.Vertical.Technology;

namespace upperletteinname
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
			
			Regex rx = new Regex(@"^[a-z0-9_]+$");
			
			
			foreach (TechClass cls in model.Classes) {
				if (!rx.IsMatch(cls.Name)) {
						logFile.WriteLine(cls.Name);
					}
				foreach (TechClassMember mbr in cls.Members) {
					if (!rx.IsMatch(mbr.Name)) {
						logFile.WriteLine(cls.Name + "." + mbr.Name );
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