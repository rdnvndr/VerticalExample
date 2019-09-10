using System;
using System.Windows.Forms;
using Ascon.Integration;
using Ascon.Vertical.Technology;
using System.Text.RegularExpressions;

namespace searchlistvalue
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
				logFile = new System.IO.StreamWriter(@"D:\commentlist.txt");
			} catch (Exception) {
				model.Close();
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка создания файла!");
				return;
			}
			
			logFile.WriteLine("СПИСОК КОММЕНТАРИЕВ:");
			logFile.WriteLine("--------------------------------------");
			foreach (TechClass cls in model.Classes) {
				if (cls.Description != String.Empty) {
					logFile.WriteLine("    " + cls.Name + ": " + cls.Description);
				}
					
				foreach (TechClassMember mbr in cls.Members) {
					if (mbr.Description != String.Empty) {
						logFile.WriteLine("    " + cls.Name + "." + mbr.Name + ": " + mbr.Description);
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