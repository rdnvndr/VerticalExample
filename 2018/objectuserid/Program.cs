﻿using System;
using System.Collections;
using System.Linq;
using System.Windows.Forms;

using Ascon.Vertical;
using Ascon.Vertical.Technology;
using Ascon.Integration;
namespace upperletteinname
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
			
			System.IO.StreamWriter logFile;
			try {
				logFile = new System.IO.StreamWriter(@"D:\log1.txt");
			} catch (Exception) {
				AuthenticationManager.Deauthenticate();
				MessageBox.Show("Ошибка создания файла!");
				return;
			}

			TechDocument doc = TechDocument.Load(@"D:\test.vtp");
                	    foreach (var child in doc.Objects) {
		                logFile.WriteLine(child.ToString() + " (" + child.Class.Name + "): " + child.CreatorId);
                            }

			doc.Close();

			AuthenticationManager.Deauthenticate();
			logFile.Close();
			MessageBox.Show("Выполнено!");				
		}
	}
}