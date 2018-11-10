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
				System.IO.StreamWriter file = new System.IO.StreamWriter(@"D:\attr.txt");
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
			
			int count = 0;
			foreach (TechClass cls in model.Classes) {
				foreach (TechClassMember mbr in cls.Members) {
					if (mbr.Type == TechClassMemberType.Property) {
						var prop = mbr as TechClassProperty;
						
						var getter = model.VBSFunctions[model.VBSFunctions.GetPropertyGetterName(prop)];
						bool isBadGetter = !String.IsNullOrEmpty(getter) 
							&& !Regex.IsMatch(getter, "function\\s+" + prop.Name + "_get\\s*\\(",	RegexOptions.IgnoreCase);
						
						if (isBadGetter) 
						{
							logFile.WriteLine(cls.Name + "." + prop.Name + " (Геттер)");
							
							Regex rx = new Regex("function\\s+([a-zA-Z0-9_-]*)\\s*\\(",  
							                     RegexOptions.Compiled | RegexOptions.IgnoreCase);
            				Match matche = rx.Match(getter);
            				var func = matche.Groups[1].Captures[0].ToString();
            				
            				Regex rgx = new Regex("function\\s+" + func + "\\s*\\(");
      						getter = rgx.Replace(getter, "function " + prop.Name + "_get(");
      						
      						Regex rgy = new Regex(func + "\\s*=");
      						getter = rgy.Replace(getter, prop.Name + "_get =");
      						
      						logFile.WriteLine(func);
      						logFile.WriteLine(getter);
      						
						}
						
						var setter = model.VBSFunctions[model.VBSFunctions.GetPropertySetterName(prop)];
						bool isBadSetter = !String.IsNullOrEmpty(setter) 
							&& !Regex.IsMatch(setter, "function\\s+" + prop.Name + "_set\\s*\\(",	RegexOptions.IgnoreCase);
						
						if (isBadSetter) 
						{
							logFile.WriteLine(cls.Name + "." + prop.Name + " (Сеттер)");
							
							Regex rx = new Regex("function\\s+([a-zA-Z0-9_-]*)\\s*\\(",  
							                    RegexOptions.Compiled | RegexOptions.IgnoreCase);
							
            				Match matche = rx.Match(setter);
            				var func = matche.Groups[1].Captures[0].ToString();
							
							Regex rgx = new Regex("function\\s+" + func + "\\s*\\(");
      						getter = rgx.Replace(setter, "function " + prop.Name + "_set(");
      						
      						Regex rgy = new Regex(func + "\\s*=");
      						getter = rgy.Replace(setter, prop.Name + "_set =");
      						
      						logFile.WriteLine(func);
      						logFile.WriteLine(setter);
						}
						
						if (isBadGetter || isBadSetter) {
							++count;
							logFile.WriteLine();
						}
					}
				}
			}
			model.Close();
			AuthenticationManager.Deauthenticate();
			logFile.WriteLine("ИТОГО АТРИБУТОВ: " + count);
			logFile.Close();
			MessageBox.Show("Выполнено!");				
		}
	}
}