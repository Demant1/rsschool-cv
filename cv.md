# ***KULIASHOU DZMITRY***    
![](https://i.yapx.ru/NPAFgt.png) 
## @MyEmail: Demant01@gmail.com 

Brief information about me: 
>Hello everyone who is interested in this information. I decided to take these courses in order to acquire new skills and change my job in the future. Unfortunately, I have no experience in IT, but I have kept this idea for a very long time. I'm sure the time has come.
>
>I have such character traits as: 
>1. Perseverance 
>2. Hard work
>3. I am very motivated.

I know **C #**, had experience with **SQL DB**.\
This is a part of the code from a program I wrote for work (graph formation):
```
 if (radioButton1.Checked == true)
            {
                try
                {
                    SqlConnection sqlConn = new SqlConnection(@"Data Source=ADMINISTRATOR\KURSA4;Initial Catalog=Kursovoi;Integrated Security=True");
                    sqlConn.Open();

                    SqlCommand sqlCom = new SqlCommand("SELECT * FROM [dbo].[Приборы_КИПиА] WHERE (('" + StatClass.God + "'-ДатаПостЭкс)%10=0)", sqlConn);
                    SqlDataReader dr = sqlCom.ExecuteReader();
                    DataTable dt = new DataTable();
                    dt.Load(dr);
                    dataGridView1.DataSource = dt;
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Ошибка при составлении лога\n" + exc.Message);
                }
            }

            if (radioButton1.Checked == true)
            {
                try
                {
Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                    //Книга.
                    ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                    //Таблица.
                    ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
                    ExcelApp.Cells[2, 4] = "График диагностирования на " + StatClass.God + " год";
                    ExcelWorkSheet.Cells.get_Range("A2", "K2").Font.Bold = true;
                    ExcelWorkSheet.Cells.get_Range("A2", "K2").RowHeight = 50;
                    ExcelWorkSheet.Cells.get_Range("A2", "K2").Font.Size = 16;
                    ExcelWorkSheet.Cells.get_Range("A2", "K2").VerticalAlignment = 2;

                    string modelRange = "A2:K2";
                    var modelTable = ExcelWorkSheet.Cells[modelRange];
                    // Assign borders 
modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
             ExcelApp.Cells[3, 1] = "Код СИ";
                    ExcelApp.Cells[3, 2] = "Инвентарный номер";
                 ExcelApp.Cells[3, 3] = "Наименование СИ";
                 ExcelApp.Cells[3, 4] = "Тип";
                 ExcelApp.Cells[3, 5] = "Заводской номер";
                 ExcelApp.Cells[3, 6] = "Место установки";
                 ExcelApp.Cells[3, 7] = "Диапазон";
                 ExcelApp.Cells[3, 8] = "Класс точности";
                 ExcelApp.Cells[3, 9] = "Периодичность поверки";
                    ExcelApp.Cells[3, 10] = "Дата поверки";
                    ExcelApp.Cells[3, 11] = "Год ввода в эксплуатацию";
                    ExcelWorkSheet.Cells.get_Range("A3", "K3").Font.Bold = true;
                    ExcelWorkSheet.Cells.get_Range("A3", "K3").ColumnWidth = 20;
                    ExcelWorkSheet.Cells.get_Range("A3", "K3").HorizontalAlignment = 3;
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.ColumnCount; j++)
                        {
                            ExcelApp.Cells[i + 4, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                            var modelTable = ExcelWorkSheet.Cells[dataGridView1.Rows[i].Cells[j].Value];
                            // Assign borders 
                            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        }
                    }
                    ExcelApp.Cells[dataGridView1.Rows.Count + 5, 1] = "Ведущий инженер по КИПиСА Долженков Е.С. ______________";
                    ExcelApp.Cells[dataGridView1.Rows.Count + 5, 6] = DateTime.Today;
                    //Вызываем нашу созданную эксельку.
                    ExcelApp.Visible = true;
                    ExcelApp.UserControl = true; ;
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Ошибка при составлении лога\n" + exc.Message);
                }
            }
```
How i say early, i haven't practic in Web programming 
![](https://i.yapx.ru/NPDcds.png) \
 Higher education, graduated from **The Belarusian-Russian University** with a degree in _Information Technology Engineer_.
 #### *My English Level* = **A2**