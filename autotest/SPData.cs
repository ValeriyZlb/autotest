using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace autotest
{
    class SPData
    {
        public string DataFiles;
        public int ColCount, RowCout, dataCol;
        public SPData(string dataf)
        {
            DataFiles = dataf;
            ColCount = default;
            dataCol = default;
            RowCout = default;
        }
        public string[] ReadTitle()
        {
            string Buffer = ReadFile();
            string[] ReadStrs = Buffer.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            
            if (ReadStrs.Length > 5)
            {
                RowCout = ReadStrs.Length - 6;
                Buffer = ReadStrs[5].Replace("\r", "");
                ReadStrs = Buffer.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                ColCount = ReadStrs.Length;
                return ReadStrs;
            }
            else
            {
                ReadStrs[0] = "notitle";
                return ReadStrs;
            }
        }
        public string[,] ReadData()
        {
            string Buffer = ReadFile().Replace("\r", " ");
            // TODO Нужна проверка на ошибку открытия файла
            string[] ReadStrs = Buffer.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (ReadStrs.Length > 5)
            {
                RowCout = ReadStrs.Length - 7;
                int FldLen = FieldLength(ReadStrs[6]);
                ColCount = ReadStrs[6].Length / FldLen;
                string[,] data = new string[RowCout + 1, ColCount];
                for (int row = 0; row < RowCout + 1; row++)
                    for (int col = 0; col < ColCount; col++)
                    {
                        Buffer = ReadStrs[row + 6].Substring(col * FldLen, FldLen);
                        Buffer = Buffer.Trim(' ');
                        if (col > 1) Buffer = Buffer.Replace(".", ",");
                        data[row, col] = Buffer;
                    }
                return data;
            }
            return null;
            /*
            if (ReadTitle()[0] != "notitle")
            {
                int colCount = ReadTitle().Length;
                string[,] rdata = new string[strCount, colCount];
                string[] rs = ReadFile().Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                int tabLen = rs[7].Length / colCount;
                for (int i = nstr; i < rs.Length - 7; i++)
                {
                    RowCout = i+1;
                    for (int j = 0; j < rs[7 + i].Length / tabLen; j++)
                    {
                        dataCol = rs[7 + i].Length / tabLen;
                        rdata[i, j] = rs[7 + i].Substring(j * tabLen, tabLen).Trim(' ', '\r');
                        // Заменяем точки на запятые, кроме полей даты и времени
                        if (j > 1) rdata[i, j] = rdata[i, j].Replace('.', ',');
                    }
                }
                return rdata;
            }
            return null;
            */
        }
        string ReadFile()
        {
            FileStream fin;
            try
            {
                fin = new FileStream(DataFiles, FileMode.Open);
            }
            catch (IOException exc)
            {
                Console.WriteLine("Ошибка открытия файла: \n" + exc.Message);
                return "FileOpenError";
            }

            byte[] bytes = new byte[fin.Length];
            for (int i = 0; i < fin.Length; i++)
                bytes[i] = (byte)fin.ReadByte();
            fin.Close();

            return Encoding.GetEncoding(866).GetString(bytes);
        }
        int FieldLength(string str)
        {
            for (int i = str.IndexOf(' '); i < str.Length; i++)
                if (str[i] != ' ') return i;
            return -1;
        }
        string[,] StringsToData(string[] strs)
        {
            return null;
        }

    }
}
