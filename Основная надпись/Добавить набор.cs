using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Основная_надпись
{
    public partial class Добавить_набор : Form
    {
        Form1 form1;
        public Добавить_набор(Form1 f)
        {
            InitializeComponent();
            this.form1 = f;
            dataGridView1.Rows.Add("Разработал");
            dataGridView1.Rows.Add("Проверил");
            dataGridView1.Rows.Add("Тех. контроль");
            dataGridView1.Rows.Add("");
            dataGridView1.Rows.Add("Нормоконтроль");
            dataGridView1.Rows.Add("Утвердил");
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1[0, 3].ReadOnly = false;
            dataGridView1.Columns[2].Visible = false;
        }
       
        private void В_набор()
        {
             Шаблон_набора новыйНабор = new Шаблон_набора();
             новыйНабор.Имя_набора = textBox1.Text;

            if (textBox1.Text != "" && textBox1.Text != string.Empty)
            {
                новыйНабор.разработал.фамилия_разработчика = dataGridView1[1, 0].Value?.ToString();
                новыйНабор.разработал.подпись_разработчика = dataGridView1[2, 0].Value?.ToString();

                новыйНабор.проверил.фамилия_проверяющего = dataGridView1[1, 1].Value?.ToString();
                новыйНабор.проверил.подпись_проверяющего = dataGridView1[2, 1].Value?.ToString();

                новыйНабор.т_контр.фамилия_Тконтр = dataGridView1[1, 2].Value?.ToString();
                новыйНабор.т_контр.подпись_Тконтр = dataGridView1[2, 2].Value?.ToString();

                новыйНабор.своёПоле.должность_своёПоле = dataGridView1[0, 3].Value?.ToString();
                новыйНабор.своёПоле.фамилия_своёПоле = dataGridView1[1, 3].Value?.ToString();
                новыйНабор.своёПоле.подпись_своёПоле = dataGridView1[2, 3].Value?.ToString();

                новыйНабор.н_контр.фамилия_Нконтр = dataGridView1[1, 4].Value?.ToString();
                новыйНабор.н_контр.подпись_Нконтр = dataGridView1[2, 4].Value?.ToString();

                новыйНабор.утвердил.фамилия_утвердил = dataGridView1[1, 5].Value?.ToString();
                новыйНабор.утвердил.подпись_утвердил = dataGridView1[2, 5].Value?.ToString();


                form1.наборы.Add(новыйНабор);
                form1.listBox1.Items.Add(textBox1.Text);
            }
            else
            {
                MessageBox.Show("Имя набора не может быть пустым!", "Введите имя набора", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            В_набор();
        }

        private void Показать_подпись_в_PictureBox(int i)
        {
            openFileDialog1.Filter = "PNG(*.png)|*.png";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                dataGridView1[2, i].Value = openFileDialog1.FileName;
                switch (i)
                {
                    case 0:
                        pictureBox1.Image = Image.FromFile(dataGridView1[2, i].Value.ToString());
                        break;
                    case 1:
                        pictureBox2.Image = Image.FromFile(dataGridView1[2, i].Value.ToString());
                        break;
                    case 2:
                        pictureBox3.Image = Image.FromFile(dataGridView1[2, i].Value.ToString());
                        break;
                    case 3:
                        pictureBox4.Image = Image.FromFile(dataGridView1[2, i].Value.ToString());
                        break;
                    case 4:
                        pictureBox5.Image = Image.FromFile(dataGridView1[2, i].Value.ToString());
                        break;
                    case 5:
                        pictureBox6.Image = Image.FromFile(dataGridView1[2, i].Value.ToString());
                        break;
                }
            }
        }

        private void Подпись_Разраб_Click(object sender, EventArgs e)
        {
            Показать_подпись_в_PictureBox(0);
        }

        private void Подпись_Проверил_Click(object sender, EventArgs e)
        {
            Показать_подпись_в_PictureBox(1);
        }

        private void Подпись_Тконтр_Click(object sender, EventArgs e)
        {
            Показать_подпись_в_PictureBox(2);
        }

        private void ПустаяСтрока_Подпись_Click(object sender, EventArgs e)
        {
            Показать_подпись_в_PictureBox(3);
        }

        private void Подпись_Нконтр_Click(object sender, EventArgs e)
        {
            Показать_подпись_в_PictureBox(4);
        }

        private void Подпись_Утвердил_Click(object sender, EventArgs e)
        {
            Показать_подпись_в_PictureBox(5);
        }
    }
}
