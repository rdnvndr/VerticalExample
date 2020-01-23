using System;
using System.Linq;
using System.IO;
using System.Windows.Forms;
using Ascon.Vertical.Technology;
using Ascon.Integration;

namespace changesystemmarker
{
	internal sealed class Program
	{
		[STAThread]
		private static void Main(string[] args)
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
				Console.WriteLine("Введите имя маркера:");
				var markername = Console.ReadLine();
				Console.WriteLine();
				if (markername.Count() >= 1) {
					var model = TechModel.Load(fileName);
					if (model != null) {
						var marker = model.Markers[markername];
						if (marker != null) {
							marker.IsSystem = !marker.IsSystem;
							model.Save(fileName);	
						} else {
							Console.WriteLine("Маркер с указанным именем не существует");
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