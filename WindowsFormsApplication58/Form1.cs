﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication58
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<List<string>> myList = new List<List<string>>();

            //string path = @"C:\План_закупки_товаров\ПЛАН_ЗАКУПКИ_ТОВАРОВ.docx";
            //string path2 = @"C:\План_закупки_товаров\2.docx";
            //string path3 = @"C:\План_закупки_товаров\3.docx";


            string path = Environment.CurrentDirectory + @"\" + "ПЛАН_ЗАКУПКИ_ТОВАРОВ.docx";
            string path2 = Environment.CurrentDirectory + @"\" + "2.docx";
            myList.Add(new List<string> { "ph_table1",
         
        
      
                
   
            
               "1",	"51.64.2",	"3020198",	"Коммутатор неуправляемый TP-LINKTL-SF1016D 16xUTP 10/100",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "1",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
            
             "1",	"51.64.2",	"3020198",	"Ноутбук Lenovo IdeaPad V580 15.6 (1366x768) Intel Core i5-3230M(2.6 GHz)/6GB/1TB/DVD±RW/nVidia GT 740M 1GB/WiFi/BT/FPR/Cam/Win 8",	"Указаны в извещении и документации к закупке",	"796", 	"кол", "1",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
            
                  "1",	"51.64.2",	"3020198",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 4Gb / 1000Gb / DVD-RW / 450W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "3",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
            
             "1",	"51.64.2",	"3020201",	"Монитор LED 24 BenQ GL2460",	"Указаны в извещении и документации к закупке",	"796", 	"кол", "1",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
           
             
                  "1",	"51.64.2",	"3020198",	"Коммутатор неуправляемый TP-LINKTL-SF1016D 16xUTP 10/100",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "1",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
            
             "1",	"51.64.2",	"3020201",	"Ноутбук Lenovo IdeaPad V580 15.6 (1366x768) Intel Core i5-3230M(2.6 GHz)/6GB/1TB/DVD±RW/nVidia GT 740M 1GB/WiFi/BT/FPR/Cam/Win 8",	"Указаны в извещении и документации к закупке",	"796", 	"кол", "1",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
           "2",	"51.64.2",	"3020354",	"Монитор LED 24 BenQ GL2460",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "1",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Май 2014 г.",	"котировка",	"нет",
            
            "2",	"51.64.2",	"3020321",	"ИБП Powercom BNT-1000AP",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "2",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"нет",
             "3",	"51.64.2",	"3020321",	"ИБП APC BX800CI-RS",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "1",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"нет",
            
              "3",	"51.64.2",	"3020201",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 4Gb / 1000Gb / DVD-RW / 450W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "122",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
            
              "4",	"51.64.2",	"3020202",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 450W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "200",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
               "4",	"51.64.2",	"3020203",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 460W",	"Указаны в извещении и документации к закупке",	"796", 	"кол", "250",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
               "4",	"51.64.3",	"3020204",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 470W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "207",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
               "4",	"51.64.2",	"3020205",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 480W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "208",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
                "4",	"51.64.2",	"3020206",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 490W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "208",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
            
                  "5",	"51.64.2",	"3020202",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 450W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "200",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
               "5",	"51.64.2",	"3020203",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 460W",	"Указаны в извещении и документации к закупке",	"796", 	"кол", "250",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
               "5",	"51.64.3",	"3020204",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 470W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "207",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
               "5",	"51.64.2",	"3020205",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 480W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "208",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
                "5",	"51.64.2",	"3020206",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 8Gb / 1000Gb / DVD-RW / 490W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "209",	"45277571000",	"г. Москва",	"135 000,00",	"Май 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
              //"6",	"51.64.2",	"3020201",	"ПЭВМ PC Office Intel Core i5-4430 (3.20GHz) / 4Gb / 1000Gb / DVD-RW / 450W",	"Указаны в извещении и документации к закупке",	"796", 	"штук", "122",	"45277571000",	"г. Москва",	"130 000,00",	"Март 2014 г",	"Март/Апрель 2014 г.",	"котировка",	"да",
            
                /*
              "5",	"23.20.11",	"2320212",	"Бензин АИ-95",	"Указаны в извещении и документации к закупке",	"112", 	"литры", "3125",	"45277571000",	"г. Москва",	"308225,00",	"Июнь 2014 г.",	"Октябрь 2014 г.",	"котировка",	"да",
              "6",	"23.20.11",	"2320212",	"Бензин АИ-92",	"Указаны в извещении и документации к закупке",	"112", 	"литры", "4125",	"45277571000",	"г. Москва",	"308225,00",	"Июнь 2014 г.",	"Октябрь 2014 г.",	"котировка",	"да",
              "6",	"23.20.15",	"2320230",	"Дизельное топливо",	"Указаны в извещении и документации к закупке",	"112", 	"литры", "1625",	"45277571000",	"г. Москва",	"308225,00",	"Июнь 2014 г.",	"Октябрь 2014 г.",	"котировка",	"да",
           
                  "7",	"23.20.15",	"4560000",	"Закупка услуг по подготовке и согласованию проектной документации на строительство погрузочно-разгрузочной площадки",	"Указаны в извещении и документации к закупке",	" ",	"услуга",	" ",	"45277571000",	"г. Москва",	"1 707 500,00",	"Июнь/Июль 2014 г.",	"Ноябрь/Декабрь 2015 г.",	"конкурс",	"нет",
                */
          
                });
           
            Dictionary<string, string> dict = new Dictionary<string, string>();
          dict.Add("PH_Year", "2015");

           dict.Add("PH_DD", "24");

           dict.Add("PH_DM", "апреля");
           dict.Add("PH_DY", "2015");
           

            FindCont.Find_table(myList, path, path2);

            GenText.GentText(dict,myList, path2, path2);

        }
    }
}
