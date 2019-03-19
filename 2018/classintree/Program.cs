using System;
using System.Collections;
using System.Linq;
using System.Windows.Forms;

using Ascon.Vertical;
using Ascon.Vertical.Technology;
using Ascon.Integration;

namespace copytree
{
	class Program
	{
		static bool childrenPrsn(TechPresentationModelItem parent, String str) 
		{			
			foreach (var item in parent.Items) {
				var filter = parent.Presentation.Model.Filters[item.Name];	
				if (filter != null) {
					for (int i=0; i<filter.Count; ++i) {
						var childCls = filter[i];
						if (str == childCls.Name) {
							return true;
						}
					}
				} else if (str == item.Name) {
					return true;
				}
				if (childrenPrsn(item, str))
					return true;
			}
						
			return false;
		}
		
		[STAThread]
		private static void Main(string[] args)
		{
			string modelName = @"d:\str.vtp";
			string str = "sub_material";
			try {
				AuthenticationManager.Authenticate();
			} catch(Exception) {
				MessageBox.Show("Ошибка авторизации!");		
				return;
			}
			
			TechModel model = null;
			try {
				model	 = TechModel.Load(modelName);
			} catch(Exception) {
				MessageBox.Show("Ошибка загрузки модели!");		
				AuthenticationManager.Deauthenticate();
				return;
			}
			
			var presenations = model.Presentations;
			foreach (var prs in presenations) {				
				foreach (var item in prs.Items) {
					bool ret = false;
					var filter = model.Filters[item.Name];
					if (filter != null) {
						for (int i=0; i<filter.Count; ++i) {
							var childCls = filter[i];
							if (str == childCls.Name) {
								ret = true;
								break;
							}
						}
					} else if (str == item.Name) {
						ret = true;
					}
					
					if (!ret)
						ret = childrenPrsn(item, str);
					
					if (ret) {
						MessageBox.Show("Найден класс в дереве:\n    " + prs.DisplayName);	
						break;
					}
						
				}
			}
			model.Close();
			
			AuthenticationManager.Deauthenticate();
			MessageBox.Show("Выполнено!");				
		}
	}
}