using System;
using System.Windows.Forms;
using Ascon.Integration;
using Ascon.Vertical.Technology;
using System.Text.RegularExpressions;

namespace searchattrbad
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
				logFile = new System.IO.StreamWriter(@"D:\attr.txt");
			} catch (Exception) {
				model.Close();
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка создания файла!");
				return;
			}
			
			foreach (TechClass cls in model.Classes) {
				foreach (TechClassMember mbr in cls.Members) {
					if (mbr.Type == TechClassMemberType.Property) {
						var prop = mbr as TechClassProperty;
						var mRestriction = prop.ValueRestrictions as TechMeasurandValueRestrictions;
						if (mRestriction != null && mRestriction.Precision >= 0) {
							logFile.WriteLine(cls.Name + "." + prop.Name);
							continue;
						}
						var dRestriction = prop.ValueRestrictions as TechMeasurandValueRestrictions;
						if (dRestriction != null && mRestriction.Precision >= 0) {
							logFile.WriteLine(cls.Name + "." + prop.Name);
                                                        continue;
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