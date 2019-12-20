using System;
using System.Collections;
using System.Linq;
using System.Windows.Forms;

using Ascon.Vertical;
using Ascon.Vertical.Technology;
using Ascon.Integration;
namespace undoconfirm{
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

                        if (doc.Extensions.EngineeringChangesManager.State 
                            == TechEngineeringChangesTechnologyState.Approved) 
                        {
                            doc.Extensions.EngineeringChangesManager.CancelLatestApprove();
                            doc.Save( openFileDialog1.FileName);
                        } else {
                            MessageBox.Show("Техпроцесс не утвержден!");
                        }  
                       
			doc.Close();
			AuthenticationManager.Deauthenticate();
			MessageBox.Show("Выполнено!");				
		}
	}
}