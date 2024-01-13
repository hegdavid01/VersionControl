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
using UserMaintence.Entities;

namespace UserMaintence
{
    public partial class Form1 : Form
    {
        BindingList<User> users = new BindingList<User>();

        public Form1()
        {
            InitializeComponent();

            button2.Text = Resource1.SaveToFileButtonText;

            label1.Text = Resource1.FullName;  
            button1.Text = Resource1.Add;

            listBox1.DataSource = users;
            listBox1.ValueMember = "ID";
            listBox1.DisplayMember = "FullName";

            var u = new User()
            {
                FullName = textBox1.Text,
            };
            users.Add(u);
        }



        private void button2_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Title = "Fájl mentése";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (StreamWriter writer = new StreamWriter(saveFileDialog.FileName))
                    {
                        foreach (var user in users) 
                        {
                            writer.WriteLine("ID: {user.ID}, FullName: {user.FullName}");
                        }
                    MessageBox.Show("Adatok sikeresen mentve a fájlba.");
                    }
                }


            }
            
            
        }
    }
}
