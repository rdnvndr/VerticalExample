using System;
using System.Collections;
using System.Linq;
using System.Windows.Forms;

using Ascon.Vertical;
using Ascon.Vertical.Technology;
using Ascon.Integration;
namespace restorestate{
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

                        OpenFileDialog openFileDialog1 = new OpenFileDialog  
                        {  
                            InitialDirectory = @"c:\",  
                            Title = "Выбор файла техпроцесса",  
  
                            CheckFileExists = true,  
                            CheckPathExists = true,  
  
                            DefaultExt = "vtp",  
                            Filter = "Файлы ТП (*.vtp;*.ttp)|*.vtp;*.ttp",  
                            FilterIndex = 4,  
                            RestoreDirectory = true,  
  
                            ReadOnlyChecked = true,  
                            ShowReadOnly = true  
                        };  
                        
                        if (openFileDialog1.ShowDialog() != DialogResult.OK)  
                        {  
                            AuthenticationManager.Deauthenticate();
                            return;   
                        }  
                           
                        TechDocument doc = null;
                        try {
			    doc = TechDocument.Load(openFileDialog1.FileName);
			 } catch (Exception) {
                            doc.Close();
    			    AuthenticationManager.Deauthenticate();
		  	    MessageBox.Show("Ошибка загрузки техпроцесса!");
			    return;
			}

                        var root = doc.Objects.Root;
                        var prevData = root.Attributes["prev_package_part"].ComplexValue as TechFileValue;
                        var data = root.Attributes["package_part"].ComplexValue as TechFileValue;

                        if (prevData != null && data != null) {
                            // удалить текущую статистику.
                            var fileId = data.LinkedFileId;
                            data.Clear();
                            doc.LinkedFiles.Remove(fileId);
                            // заменить на предыдущую статистику.
                            var prevFileId = prevData.LinkedFileId;
                            data.LinkedFileId = prevFileId;
                            prevData.LinkedFileId = Guid.Empty;
                            doc.Save(openFileDialog1.FileName);
                            MessageBox.Show("Выполнено!");
                        } else {
                            MessageBox.Show("Статистика отсутсвует!");
                        }
                       
			doc.Close();
			AuthenticationManager.Deauthenticate();
							
		}
	}
}