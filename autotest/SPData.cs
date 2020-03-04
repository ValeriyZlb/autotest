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
        public int titleCol, dataCol, dataRow;
        public SPData(string dataf)
        {
            DataFiles = dataf;
            titleCol = default;
            dataCol = default;
            dataRow = default;
        }
        public string[] ReadTitle()
        {
            string[] rs = ReadFile().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (rs.Length > 5)
            {
                string[] result = rs[5].Replace("\r", "").Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                titleCol = result.Length;
                return result;
            }
            else
            {
                rs[0] = "notitle";
                return rs;
            }
        }
        public string[,] ReadData(int nstr, int strCount)
        {
            if (ReadTitle()[0] != "notitle")
            {
                int colCount = ReadTitle().Length;
                string[,] rdata = new string[strCount, colCount];
                string[] rs = ReadFile().Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                int tabLen = rs[7].Length / colCount;
                for (int i = nstr; i < rs.Length - 7; i++)
                {
                    dataRow = i+1;
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

    }
}
