using System;
using System.Linq;
using System.IO;
using Ascon.Integration;
using Ascon.Vertical.Technology;

namespace changesystem
{
	internal sealed class Program
	{
		[STAThread]
		public static void Main(string[] args)
		{
			var fileName = (args.Count() > 0) ? args[0] : String.Empty;
			if (File.Exists(fileName)) {
				try {
					AuthenticationManager.Authenticate();
				} catch (Exception) {
					Console.Write("Ошибка авторизации!");
					Console.Write("Нажмите любую клавишу для продолжения . . . ");
					Console.ReadKey(true);
					return;
				}
				Console.WriteLine("Введите имя класса, атрибута или функции в формате\n<Имя класс>.<Имя атрибута или функции>:");
				var item = Console.ReadLine();
				Console.WriteLine();
				var items = item.Split('.');
				
				if (items.Count() >= 1) {
					var model = TechModel.Load(fileName);
					if (model != null) {
						String className = items[0];
						TechClass cls = model.Classes[className];
						if (cls != null) {
							if  (items.Count() >= 2) {
								String attrName = items[1];
								TechClassMember mbr = cls.Members[attrName];
								if (mbr != null) {
									mbr.IsSystem = !mbr.IsSystem;
									model.Save(fileName);	
								} else {
									Console.WriteLine("Атрибут или функция с указанным именем не существует");
								}
							} else {
								cls.IsSystem = !cls.IsSystem;
								model.Save(fileName);	
							}
						} else {
							Console.WriteLine("Класс с указанным именем не существует");					
						}
						model.Close();
					} else {
						Console.WriteLine("Невозможно открыть объектную модель Вертикаль");
					}
				}
				AuthenticationManager.Deauthenticate();
			} else {
				Console.WriteLine("Указанный файл не существует");
			}
			
			Console.Write("Нажмите любую клавишу для продолжения . . . ");
			Console.ReadKey(true);
		}		
	}
}