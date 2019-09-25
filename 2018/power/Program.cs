using System;
using System.Linq;
using System.Windows.Forms;
using Ascon.Vertical.Technology;
using Ascon.Integration;

namespace power
{
	internal sealed class Program
	{
		[STAThread]
		private static void Main(string[] args)
		{			
			try {
				AuthenticationManager.Authenticate();
			} catch(Exception) {
				MessageBox.Show("Ошибка авторизации!");		
				return;
			}
			
			try {
				TechDocument document = TechDocument.Load("test.vtp");
                var dseFilter = document.Model.Filters["dse"];
				var dse = document.Objects.Root.Children.Find(dseFilter).FirstOrDefault();
				if (dse != null) {
				   	var operFilter = document.Model.Filters["operations"];
					var oper = dse.Children.Find(operFilter).FirstOrDefault();
					if (oper != null) {
						var equipmentFilter = document.Model.Filters["equipment"];
						var equipment = oper.Children.Find(equipmentFilter).FirstOrDefault();
						if (equipment != null) {
							MessageBox.Show("Мощность оборудования: " 
							                + equipment.Attributes["power"].Value.ToString()
							                + " Вт");
						}
					}
				}
				document.Close();
			} catch(Exception) {
				MessageBox.Show("Ошибка загрузки документа!");		
			}
				
			
			AuthenticationManager.Deauthenticate();
		}
	}
}
