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
			static void childrenPrsn(TechPresentationModelItem oldParentItem, TechPresentationModelItem newParentItem) {
			var model = oldParentItem.Presentation.Model;
			foreach (var oldItem in oldParentItem.Items) {
				var filter = model.Filters[oldItem.Name];
				if (filter != null) {
					for (int i=0; i<filter.Count; ++i) {
						var childCls = filter[i];
						var parentCls = model.Classes[newParentItem.Name];
						if (parentCls != null && model.ClassesRelations.Contains(parentCls, childCls)) {
							var newItem = newParentItem.Items.Add(childCls.Name);
							newItem.Description = childCls.Name + ": " + childCls.Description;
							childrenPrsn(oldItem, newItem);
						}
					}
				} else {
					var childCls  = model.Classes[oldItem.Name];
					var parentCls = model.Classes[newParentItem.Name];
					if (parentCls != null && model.ClassesRelations.Contains(parentCls, childCls)) {
						var newItem = newParentItem.Items.Add(oldItem.Name);
						newItem.Description = oldItem.Description;
						childrenPrsn(oldItem, newItem);
					}
				}
			}
			
			return;
		}
		
		[STAThread]
		private static void Main(string[] args)
		{
			string modelName = @"e:\str.vtp";
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
			var prs = model.Presentations[0];
			
			if (prs != null) {
				var newPrsn = presenations.Add("test");		
				foreach (var oldItem in  prs.Items) {
					var filter = model.Filters[oldItem.Name];
					if (filter != null) {
						for (int i=0; i<filter.Count; ++i) {
							var childCls = filter[i];
							var newItem = newPrsn.Items.Add(childCls.Name);
							newItem.Description = childCls.Name + ": " + childCls.Description;
							childrenPrsn(oldItem, newItem);
						}
					} else {
						var newItem = newPrsn.Items.Add(oldItem.Name);
						newItem.Description = oldItem.Description;
						childrenPrsn(oldItem, newItem);
					}
				}
			}
						
					
			model.Save(modelName);
			model.Close();
			
			AuthenticationManager.Deauthenticate();
			MessageBox.Show("Выполнено!");				
		}
	}
}