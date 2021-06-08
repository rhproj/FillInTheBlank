using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;

namespace DocFiller
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(comboTitle.Text))
            {
                MessageBox.Show("Не выбран шаблон документа");
            }
            else
            {
                //нужно собрать те данные, кот-е заменим в док-те:
                var DictToFill = new Dictionary<string, string>
                {
                    {"<ChPosition>", comboChPosittion.Text },
                    {"<ChRank>", comboChRank.Text },
                    {"<ChName>", comboChName.Text },
                    {"<Position>", comboPosition.Text },
                    {"<Signature>",tbSignature.Text },
                    {"<Title>",comboTitle.Text },
                    {"<Date>",dateTimePicker1.Value.ToString("dd.MM.yyyy") },
                    {"<Name>", tbName.Text }
                };
                var wordHelper = new WordHelper($"{comboTitle.Text}.docx");     //"Template.docx"

                wordHelper.Replace(DictToFill);

                MessageBox.Show("Документ создан!");
            }
        }

        private void tbSignature_TextChanged(object sender, EventArgs e)
        {
            tbName.Text = tbSignature.Text;
        }
    }
}
