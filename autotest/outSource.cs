﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace autotest
{
    class outSource
    {
        public string DeviceName;
        public string DataFiles;
        public string StartProperties;
        public outSource(string devname, string devid, string dataf)
        {
            DeviceName = devname;
            StartProperties = devid;
            DataFiles = dataf;
        }
        public int RequestData(int RequestTime, TextBox consoleOut)
        {
            string RqTime = String.Format("/S={0:d2}000900 /F={1:d2}000800", RequestTime, RequestTime+1);
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = @"C:\Program Files (x86)\Logika\Spnet95\Sphone95.exe";
            consoleOut.AppendText(DateTime.Now + " Опрос: " + DeviceName + "...\r\n");
            startInfo.Arguments = "/P=" + StartProperties + RqTime;
            startInfo.UseShellExecute = true;
            var procStat = Process.Start(startInfo);
            procStat.WaitForExit();
            return 0;
        }
        public string[] ReadTitle()
        {
            string[] rs = ReadFile().Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (rs.Length > 5)
                return rs[5].Replace("\r", "").Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
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
                    //ws = rs[6 + i].Split(new char[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < rs[7 + i].Length / tabLen; j++)
                        rdata[i, j] = rs[7 + i].Substring(j * tabLen, tabLen).Replace(" ", "").Replace("\r", "");
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