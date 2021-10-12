using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;
using Kompas6API5;
using System.Runtime.InteropServices;

namespace Основная_надпись
{
    public partial class Form1 : Form
    {          
        public List<Шаблон_набора> наборы = JsonConvert.DeserializeObject<List<Шаблон_набора>>(Deserialized());
        static string fullPath = Path.GetFullPath("наборы.json");
        static string Deserialized()
        {
            string deserialized = File.ReadAllText(fullPath);
            return deserialized;
        }
          public Form1()
          {
              InitializeComponent();
              dgv_вставка.Rows.Add("Разработал");
              dgv_вставка.Rows.Add("Проверил");
              dgv_вставка.Rows.Add("Тех. контроль");
              dgv_вставка.Rows.Add("");
              dgv_вставка.Rows.Add("Нормоконтроль");
              dgv_вставка.Rows.Add("Утвердил");
              dgv_вставка.Columns[2].Visible = false;
              dgv_вставка.Columns[0].ReadOnly = true;
              dgv_вставка.Columns[1].ReadOnly = true;
              for (int i = 0; i < наборы.Count; i++)
              {
                 listBox1.Items.Add(наборы[i].Имя_набора);
              }
               dgv_вставка.Columns[3].Width = 70;
               dgv_вставка.Columns[3].ReadOnly = false;
          }

        private void Добавить_Click(object sender, EventArgs e)
        {
            Добавить_набор добавитьНабор = new Добавить_набор(this);
            добавитьНабор.ShowDialog();
        }

        public void Картинки_подписей_из_DVG_в_PictureBox(DataGridView dgv)
        {
            for (int i = 0; i < 6; i++)
            {
                pictureBox1.Image = Image.FromFile(dgv[2, i].Value?.ToString());
            }                
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                dgv_вставка[1, 0].Value = наборы[listBox1.SelectedIndex].разработал.фамилия_разработчика;
                dgv_вставка[2, 0].Value = наборы[listBox1.SelectedIndex].разработал.подпись_разработчика;

                dgv_вставка[1, 1].Value = наборы[listBox1.SelectedIndex].проверил.фамилия_проверяющего;
                dgv_вставка[2, 1].Value = наборы[listBox1.SelectedIndex].проверил.подпись_проверяющего;

                dgv_вставка[1, 2].Value = наборы[listBox1.SelectedIndex].т_контр.фамилия_Тконтр;
                dgv_вставка[2, 2].Value = наборы[listBox1.SelectedIndex].т_контр.подпись_Тконтр;

                dgv_вставка[0, 3].Value = наборы[listBox1.SelectedIndex].своёПоле.должность_своёПоле;
                dgv_вставка[1, 3].Value = наборы[listBox1.SelectedIndex].своёПоле.фамилия_своёПоле;
                dgv_вставка[2, 3].Value = наборы[listBox1.SelectedIndex].своёПоле.подпись_своёПоле;

                dgv_вставка[1, 4].Value = наборы[listBox1.SelectedIndex].н_контр.фамилия_Нконтр;
                dgv_вставка[2, 4].Value = наборы[listBox1.SelectedIndex].н_контр.подпись_Нконтр;

                dgv_вставка[1, 5].Value = наборы[listBox1.SelectedIndex].утвердил.фамилия_утвердил;
                dgv_вставка[2, 5].Value = наборы[listBox1.SelectedIndex].утвердил.подпись_утвердил;
                Картинки_подписей_из_DVG_в_PictureBox(dgv_вставка);
                Дата_вНаборе();
            }
        }

        private void Сохранить_Click(object sender, EventArgs e)
        {
            string json = JsonConvert.SerializeObject(наборы);
            File.WriteAllText(fullPath, json);
            MessageBox.Show("Сохранение прошло успешно!", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Очистить_DGV(int cellIndex)
        {
            for (int i = 0; i < 6; i++)
            {
                dgv_вставка.Rows[i].Cells[cellIndex].Value = null;
            }
        }

        private void Удалить_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                var result = MessageBox.Show("Вы действительно хотите удалить выбранный набор?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {

                    наборы.RemoveAt(listBox1.SelectedIndex);
                    listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                    Очистить_DGV(1);
                    Очистить_DGV(2);
                    Картинки_подписей_из_DVG_в_PictureBox(dgv_вставка);

                }
            }
            else if (listBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Для удаления необходимо выбрать набор!", "Набор не выбран", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Дата_вНаборе()
        {
            for (int i = 0; i < 6; i++)
            {
                if (dgv_вставка[1, i].Value != null)
                {
                    dgv_вставка[3, i].Value = dateTimePicker1.Value.ToShortDateString();
                }
                else if (dgv_вставка[1, i].Value == null)
                {
                    dgv_вставка[3, i].Value = null;
                }
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            Дата_вНаборе();
        }

        private void Изменить_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                Изменить_набор изменитьНабор = new Изменить_набор(this);
                изменитьНабор.ShowDialog();
            }
            else
            {
                MessageBox.Show("Для изменения необходимо выбрать набор!", "Набор не выбран", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Заполнить_Click(object sender, EventArgs e)
        {
            Работа_с_Компасом работа_С_Компасом = new Работа_с_Компасом();
            if (listBox1.SelectedIndex != -1)
            {
                try
                {
                    KompasObject kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                    ksDocument2D doc = (ksDocument2D)kompas.ActiveDocument2D();                 
                    if (doc != null)
                    {
                        int width = работа_С_Компасом.ReturnWidth(kompas, doc);

                        for (int i = 0; i < 6; i++)
                        {
                            работа_С_Компасом.Фамилия_В_Штамп(kompas, doc, i, dgv_вставка[1, i].Value?.ToString());
                        }

                        работа_С_Компасом.Фамилия_В_Штамп(kompas, doc, 6, dgv_вставка[0, 3].Value?.ToString()); //заполняется пустая строка

                        for (int i = 0; i < 6; i++)
                        {
                            работа_С_Компасом.Дата_В_Штамп(kompas, doc, i, dgv_вставка[3, i].Value?.ToString());
                        }

                        for (int i = 0; i < 6; i++)
                        {
                            if (dgv_вставка[2, i].Value != null)
                            {
                                работа_С_Компасом.ВставкаПодписи(kompas, doc, dgv_вставка[2, i].Value.ToString(), width, i);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Нет активного документа чертежа!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch
                {
                    MessageBox.Show("КОМПАС не открыт!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Для заполнения необходимо выбрать набор!", "Набор не выбран", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
        }

        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Добавить_набор добавитьНабор = new Добавить_набор(this);
            добавитьНабор.ShowDialog();
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                var result = MessageBox.Show("Вы действительно хотите удалить выбранный набор?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {

                    наборы.RemoveAt(listBox1.SelectedIndex);
                    listBox1.Items.RemoveAt(listBox1.SelectedIndex);
                    Очистить_DGV(1);
                    Очистить_DGV(2);
                    Картинки_подписей_из_DVG_в_PictureBox(dgv_вставка);

                }
            }
            else if (listBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Для удаления необходимо выбрать набор!", "Набор не выбран", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void изменитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                Изменить_набор изменитьНабор = new Изменить_набор(this);
                изменитьНабор.ShowDialog();
            }
            else
            {
                MessageBox.Show("Для изменения необходимо выбрать набор!", "Набор не выбран", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string json = JsonConvert.SerializeObject(наборы);
            File.WriteAllText(fullPath, json);
            MessageBox.Show("Сохранение прошло успешно!", "Сохранение", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            О_программе о_Программе = new О_программе();
            о_Программе.ShowDialog();
        }

        private void руководствоToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
