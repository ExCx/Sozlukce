using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sozlukce
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        bool bitti = false;
        string path;

        private async void button1_Click(object sender, EventArgs e)
        {
            if (bitti)
            {
                System.Diagnostics.Process.Start(path);
                MessageBox.Show("Hadi kal sağlıcakla");
                Close();
            }
            else if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                button1.Text = $"Hallediyorum...";
                path = openFileDialog1.FileName;
                XLWorkbook workbook = null;
                try
                {
                    workbook = new XLWorkbook(path);
                }
                catch
                {
                    MessageBox.Show("Bi ibnelik oldu. (Excel kapalı dimi?)");
                    Close();
                }
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed();
                HttpClient client = new HttpClient();
                foreach (var row in rows)
                {
                    var phrase = row.Cell(1).GetString();

                    HttpResponseMessage response = await client.GetAsync($"http://localhost:8080/translate?phrase={phrase}");
                    if (response != null)
                    {
                        response.EnsureSuccessStatusCode();
                        var jsonString = await response.Content.ReadAsStringAsync();
                        var result = JsonConvert.DeserializeObject<TurengResponse>(jsonString);
                        var i = 2;
                        if (result.count == 0)
                        {
                            row.Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
                            row.Cell(1).Style.Font.FontColor = XLColor.White;
                            row.Cell(1).Style.Font.Bold = true;
                        }
                        foreach (var resultPhrase in result.phrases)
                        {
                            row.Cell(i).Value = resultPhrase.target;
                            i++;
                        }
                    }
                }
                workbook.Save();
                workbook.Dispose();
                button1.Text = $"Bitti. Bas da açayım.";
                bitti = true;
            }
        }

        class TurengResponse
        {
            public int count { get; set; }
            public Phrase[] phrases { get; set; }
        }
        class Phrase
        {
            public string source { get; set; }
            public string target { get; set; }
            public string category { get; set; }
            public string type { get; set; }
        }
    }
}