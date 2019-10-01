using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlServerCe;
using System.IO;



namespace ElectronicJournal
{
    class connect
    {
        SqlCeConnection SqlConnectionDB;
        SqlCeCommand ComandDB;
        SqlCeDataReader ReaderDB;

        static Encoding code = Encoding.UTF8;
        string path = File.ReadAllText("1.ini", code); 

       
        public SqlCeConnection ConnectOnDb()
        {
            SqlConnectionDB = new SqlCeConnection(path);
            SqlConnectionDB.Open();
            return SqlConnectionDB;
        }

        public void ConnectClose()
        {
            SqlConnectionDB.Close();
        }

        public void ConnectDispose()
        {
            SqlConnectionDB.Dispose();
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
