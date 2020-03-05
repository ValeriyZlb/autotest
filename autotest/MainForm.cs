using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace autotest
{
    public partial class MainForm : Form
    {
        SPDevices[] Devices = new SPDevices[5];
        SPData[] DevData = new SPData[5];
        public MainForm()
        {
            InitializeComponent();
        }

        private void start_button_Click(object sender, EventArgs e)
        {
            string[] files = { @"C:\Users\Valera\Desktop\Test\Газ.txt",
                               @"C:\Users\Valera\Desktop\Test\Газ (техн).txt",
                               @"C:\Users\Valera\Desktop\Test\Вода.txt",
                               @"C:\Users\Valera\Desktop\Test\ПТВМ.txt",
                               @"C:\Users\Valera\Desktop\Test\ДЕ.txt"};

            string[] DevicesID = { "327", "328", "329", "330", "331" };
            string[] DeviceNames = { "СПГ761 (Газ)", "СПГ761.2(Газ (техн))", "СПТ961(Вода)", "СПТ961(ПТВМ)", "СПТ961(ДЕ)" };

            for (int i = 0; i < 5; i++)
            {
                Devices[i] = new SPDevices(DeviceNames[i], DevicesID[i]);
                DevData[i] = new SPData(files[i]);
            }

            string[][,] data = new string[5][,];
            for (int i = 0; i < 5; i++)
            {
                //int a = Devices[i].RequestData(-1, console_textBox);
                data[i] = DevData[i].ReadData();
            
                if (data[i] != null) console_textBox.AppendText(DateTime.Now + " " + Devices[i].DeviceName + " > Полученны данные (" + (DevData[i].RowCout).ToString() + " из 24)\r\n");
                else console_textBox.AppendText(DateTime.Now + " " + Devices[i].DeviceName + " > Error:  Данные отстутствуют\r\n");
            }

            //Объявляем приложение
            Excel.Application ex = new Microsoft.Office.Interop.Excel.Application();
            //Отобразить Excel
            ex.Visible = true;
            ex.Workbooks.Open(@"C:\Users\Valera\Desktop\Test\template.xls",
                              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing);

            //Отключить отображение окон с сообщениями
            ex.DisplayAlerts = true;
            //Получаем первый лист документа (счет начинается с 1)
            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

            // Выводим содержимое массива с данными на первой страничке, один за другим, через строку
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i * 27 + 1, 1] = DevData[i].DataFiles;
                Excel.Range rngData = sheet.Range[sheet.Cells[i * 27 + 2, 1], sheet.Cells[i * 27 + DevData[i].RowCout + 2, DevData[i].ColCount]];
                rngData.Value = data[i];
            }
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Settings frm = new Settings();
            frm.ShowDialog();
        }
    }
}
class WorkFiles
{
    public string[,] ReadData(string Filename, out int strCount)
    {
        strCount = 0;
        FileStream fin;
        string[,] data;
        try
        {
            fin = new FileStream(Filename, FileMode.Open);
        }
        catch (IOException exc)
        {
            Console.WriteLine("Ошибка открытия файла: \n" + exc.Message);
            data = new string[0,0];
            return data;
        }

        int StartString = 5;    // Номер строки с которого нужно считывать данные
        long flen = fin.Length;
        byte[] bytes = new byte[flen];
        for (int i = 0; i < flen; i++)
            bytes[i] = (byte)fin.ReadByte();
        string s = Encoding.GetEncoding(866).GetString(bytes);
        string[] ss = s.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        strCount = ss.Length - StartString;
        if (ss.Length > 6)
        {
            string[] title = ss[6].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            fin.Close();

            data = new string[ss.Length - StartString, title.Length];
            for (int i = StartString; i < ss.Length; i++)
            {
                string[] temp = ss[i].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                for (int j = 0; j < temp.Length; j++)
                    data[i - StartString, j] = temp[j];
            }
            return data;
        }
        data = new string [1, 1];
        data[0, 0] = "error";
        return data;
    }
}