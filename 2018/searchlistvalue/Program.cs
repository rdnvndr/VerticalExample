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
				System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\attr.txt");
			} catch (Exception) {
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка загрузки модели!");
				return;
			}
			
			System.IO.StreamWriter logFile;
			try {
				logFile = new System.IO.StreamWriter(@"D:\attrlist.txt");
			} catch (Exception) {
				model.Close();
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка создания файла!");
				return;
			}
			
			logFile.WriteLine("СПИСОК АТРИБУТОВ СО СПИСКАМИ ЗНАЧЕНИЙ:");
			logFile.WriteLine("--------------------------------------");
			foreach (TechClass cls in model.Classes) {
				foreach (TechClassMember mbr in cls.Members) {
					if (mbr.Type != TechClassMemberType.Function) {
						var attr = mbr as TechClassAttribute;
						
						var doubleRestr = attr.ValueRestrictions as TechDoubleValueRestrictions;
						if (doubleRestr != null && doubleRestr.Type == TechValueRestrictionType.List) {
							logFile.WriteLine("    " + cls.Name + "." + attr.Name);
						}
						
						var intRestr = attr.ValueRestrictions as TechIntegerValueRestrictions;
						if (intRestr != null && intRestr.Type == TechValueRestrictionType.List) {
							logFile.WriteLine("    " + cls.Name + "." + attr.Name);
						}
						
						var measRestr = attr.ValueRestrictions as TechMeasurandValueRestrictions;
						if (measRestr != null && measRestr.Type == TechValueRestrictionType.List) {
							logFile.WriteLine("    " + cls.Name + "." + attr.Name);
						}
						
						var strRestr = attr.ValueRestrictions as TechStringValueRestrictions;
						if (strRestr != null && strRestr.Type == TechValueRestrictionType.List) {
							logFile.WriteLine("    " + cls.Name + "." + attr.Name);
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