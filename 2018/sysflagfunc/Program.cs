using System;
using System.Linq;
using System.IO;
using Ascon.Integration;
using Ascon.Vertical.Technology;

namespace sysflagfunc
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
				Console.WriteLine("Введите имя глобальноый функции:");
				var funcName = Console.ReadLine();
								
				if (!String.IsNullOrEmpty(funcName)) {
					var model = TechModel.Load(fileName);
					if (model != null) {
						var globalFunc = model.Functions[funcName];
						if (globalFunc != null) {
							globalFunc.IsSystem = !globalFunc.IsSystem;
							model.Save(fileName);
						} else {
							Console.WriteLine("Глобальная функция с указанным именем не существует");					
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