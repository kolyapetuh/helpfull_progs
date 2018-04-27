/*
 * Created by SharpDevelop.
 * User: noone
 * Date: 13.04.2018
 * Time: 17:19
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;

namespace ExcelWriter
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		//запись в файл Excel
		void ExcelWrite()
		{
			
	
			// Создаём экземпляр нашего приложения
		    Excel.Application excelApp = new Excel.Application();
		    // Создаём экземпляр рабочий книги Excel
		    //Excel.Name="MyFile";
		    Excel.Workbook workBook;
		    // Создаём экземпляр листа Excel
		    Excel.Worksheet workSheet;  
		 	//создаём лист и рабочую книгу
		    workBook = excelApp.Workbooks.Add();
		    workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
		    
		    //Заполнение таблицы
		          // Заполняем первую строку числами от 1 до 10
		     
		     //вывод заголовков		     
		     workSheet.Cells[1, 1] = "Заголовок 1";		     
		     workSheet.Cells[1, 2] = "Заголовок 2";		     
		     workSheet.Cells[1, 3] = "Заголовок 3";
		     workSheet.Cells[1, 4] = "Заголовок 4";
		     workSheet.Cells[1, 5] = "Заголовок 5";
		     
		     /*
		      //вариант 2 вывода заголовков
		      int what_doyouwant=5;
		      string[] cnames=new string{ "Заголовок1","Заголовок2","Заголовок3","Заголовок4","Заголовок5"}
		      for (int k=0;k<what_doyouwant;k++)
		      	workSheet.Cells[1, k+1] = cnames[k]; 
		       
		      */
		     
		     
		     //вывод всей информации	
		     Random rnd=new Random();
		     for (int k=0;k<50;k++)
		     {	
				//запись данных
				for (int j=0;j<5;j++)
					//workSheet.Cells[k+2, j+1]=Convert.ToString(rnd.Next(1,1000000));
					//workSheet.Cells.form
					workSheet.Cells[k+2, j+1]="Перенеси меня 300 раз на новую строку 20 раз, попробуй, рас рас рас";
		     }

 		     
		     //стиль для заголовка
 		     Excel.Style style = excelApp.ActiveWorkbook.Styles.Add("NewStyle");
 		     style.Font.Size=12;//размер шрифта
 		     style.Font.Bold=true;
 		     style.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray); //цвет
			 style.Interior.Pattern = Excel.XlPattern.xlPatternSolid; //тип заливки
			 //выравнивание
			 style.HorizontalAlignment=Excel.XlHAlign.xlHAlignCenter;
			 style.VerticalAlignment=Excel.XlVAlign.xlVAlignCenter;
 		    // style.Borders.LineStyle=Excel.XlLineStyle.xlContinuous;

 		     //Excel.ta
 		     //границы ячееек и установка ширины по самой длинной ячейке
 		     Excel.Range rng =(Excel.Range) workSheet.Range[workSheet.Cells[1, 1],workSheet.Cells[51, 5]];
 		     rng.EntireColumn.AutoFit();//автоподбор длины по содержимому (не работает тк. после указан размер колонки)
 		     rng.EntireRow.WrapText=true; //автоперенос слов
 		     rng.Rows.ColumnWidth=25; //ширина
 		     rng.Columns.RowHeight=50;//высота
 		    // rng.Height=100;
 		     Excel.Borders border = rng.Borders; 	//границы	     
 		     border.LineStyle = Excel.XlLineStyle.xlContinuous;
 		     //вставка таблицы для rng
 		     workSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,rng,null,Excel.XlYesNoGuess.xlYes,null);
 		                               
 		     //стиль и границы заголовка
 		     rng=(Excel.Range) workSheet.Range[workSheet.Cells[1, 1],workSheet.Cells[1, 5]];
 		     rng.Style="NewStyle";
 		     border = rng.Borders; 
 		     border.LineStyle = Excel.XlLineStyle.xlContinuous;
 		   	 //border.LineStyle = Excel.XlLineStyle.xlContinuous;
		     
             // Открываем созданный excel-файл
		    excelApp.Visible = true; //делаем его видимым
		    excelApp.UserControl = true; //можно контролировать работу с файлом
		    //пробуем закрыть файл и если надо записываем его		    
		    try
		    {
		    	
		    	
		    	//excelApp.ActiveWorkbook.SaveCopyAs(@"flist.xlsx"); //сохранение с определённым именем
		    	//workBook.SaveCopyAs("flist.xlsx");	//сохранение с определённым именем
	    	
		    	excelApp.Workbooks.Close();
		    	excelApp.Quit();
		    
		    }
		    catch {}


		}
		void Button1Click(object sender, EventArgs e)
		{
			ExcelWrite();
		}
	}
}
