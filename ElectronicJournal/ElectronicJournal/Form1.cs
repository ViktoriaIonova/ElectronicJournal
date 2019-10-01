using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlServerCe;

namespace ElectronicJournal
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        public Form1()
        {
            InitializeComponent();
        }
    
        connect ConnectDB = new connect();
        DataSet ds = new DataSet();
        private void Form1_Load(object sender, EventArgs e)
        {
          // metroTextBox1.Text = "KimVV";
          // metroTextBox2.Text = "654321";
            

        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            try
            {
                SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand.CommandText = "SELECT Prava FROM teacher WHERE Login='" + metroTextBox1.Text + "' AND Password = '" + metroTextBox2.Text + "'";
                thisCommand.Connection = ConnectDB.ConnectOnDb();
                int prava = (int)thisCommand.ExecuteScalar();

                if (prava == 0)
                {
                    Form jou = new Journal();
                    jou.Owner = this;
                    jou.Show();
                    this.Hide();
                }
                if (prava == 1)
                {
                    Form ad = new adminpanel();
                    ad.Owner = this;
                    ad.Show();
                    this.Hide();
                }
            }
            catch
            {
                //MessageBox.Show("Неправильные данные!");
                MetroFramework.MetroMessageBox.Show(this, "Неправильные данные!", " ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }   
           
        }
    }
}

/*
 * Copyright 2018, 2019 Виктория Ионова
 * This file is part of Electronic journal.

    Electronic journal  is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    Electronic journal is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with Foobar.  If not, see <https://www.gnu.org/licenses/>.
 */

