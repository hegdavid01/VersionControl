﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace week04
{
    public partial class Form1 : Form
    {
        List<Flat> Flats;
        
        public Form1()
        {
            InitializeComponent();
            LoadData();
        }
    
        private void LoadData()
        {
            Flats = Context.Flats.ToList();
        }
    }
}
