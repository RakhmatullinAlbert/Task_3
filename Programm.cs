using System;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace Task3_OpenXML
{
    class Programm
    {

        static SpreadsheetDocument document = null;
        static WorkbookPart workbookPart = null;
        static Sheets thesheetcollection = null;
        static string fileName = "";
        static void Main(string[] args)
        {
            bool fileExsist = false;
            
            
            while (!fileExsist)
            {
                Console.WriteLine("Введите путь до Excel файла:");
                fileName = Console.ReadLine();
                //fileName = @"C:\Users\alber\Desktop\Prilozhenie_2.xlsx";
                int countErr = 0;
                try
                {
                    document = SpreadsheetDocument.Open(fileName, true);
                    workbookPart = document.WorkbookPart;
                    thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    fileExsist = true;
                    countErr++;
                    if (countErr > 20)
                    {
                        Exception exception = new Exception();
                    }
                }
                catch 
                {
                    Exception exception = new Exception();
                }
                finally { Console.Clear(); }
                
            }
           

            string[] menuItems = new string[] 
            { "1. Информация о клиентах по наименованию товара",
              "2. Изменение контактного лица",
              "3. \"Золотой\" клиент" ,
              "Выход" 
            };
            
            int row = Console.CursorTop;
            int col = Console.CursorLeft;
            int index = 0;
            string rez = "";
            
            
            
            while (true)
            {

                DrawMenu(menuItems, row, col, index, rez);
                switch (Console.ReadKey(true).Key)
                {
                    case ConsoleKey.DownArrow:
                        if (index < menuItems.Length - 1)
                            index++;
                        break;
                    case ConsoleKey.UpArrow:
                        if (index > 0)
                            index--;
                        break;
                    case ConsoleKey.Enter:
                        switch (index)
                        {
                            case 3:
                                rez = "Выбран выход из приложения";
                                Console.Clear();
                                document.Dispose();
                                return;
                            case 0:
                                try
                                {
                                    string[] product;
                                    int indexProduct;
                                    string[] client;
                                    int indexClient;
                                    string[] request; int indexRequest; string nameProduct;

                                    Console.WriteLine("Введите наименование товара:");
                                    nameProduct = Console.ReadLine();
                                    indexProduct = SearchTextInt(nameProduct, 0, 1);
                                    if (indexProduct == -1) { Exception exception = new Exception(); }
                                    product = GetCellsValue(0, indexProduct);
                                    indexRequest = SearchTextInt(product[0], 2, 1);
                                    if (indexRequest == -1) { Exception exception = new Exception(); }
                                    request = GetCellsValue(2, indexRequest);
                                    indexClient = SearchTextInt(request[2], 1, 0);
                                    if (indexClient == -1) { Exception exception = new Exception(); }
                                    client = GetCellsValue(1, indexClient);

                                    Console.WriteLine("");
                                    Console.WriteLine("Наименование организации: " + client[1]);
                                    Console.WriteLine("Адрес: " + client[2]);
                                    Console.WriteLine("Контактное лицо (ФИО): " + client[3]);
                                    Console.WriteLine("количество товара: " + request[4]);
                                    Console.WriteLine("Цена товара: " + product[3]);
                                    Console.WriteLine("Дата заказа: " + request[5]);
                                    Console.ReadKey();
                                    Console.Clear();
                                }
                                catch
                                {
                                    Console.Clear();
                                    Console.WriteLine($"Не найден товар, либо не существует заявки с наименованием товара");
                                    Console.ReadKey();
                                }
                                break;
                            case 1:

                                int countClient = GetCount(1);
                                int countClient2 = 0;
                                for (int i = 1; i <= countClient; i++)
                                {
                                    string[] Client = GetCellsValue(1, i);
                                    if (Client[0] == "")
                                    {
                                        break;
                                    }
                                    countClient2++;
                                }
                                string[] menuClients = new string[countClient2];
                                menuClients[0] = "Выход (Выберете организацию для изменения)";
                                for (int i = 2; i <= countClient2; i++)
                                {
                                    string[] Client = GetCellsValue(1, i);
                                    menuClients[i - 1] = $" Наименование организации: { Client[1]} | Адрес: { Client[2]} " +
                                        $"| Контактное лицо (ФИО): { Client[3]}  ";
                                }
                                int countClient3 = countClient2 + 1;
                                Console.Clear();
                                int result = 0;
                                int row1 = Console.CursorTop;
                                int col1 = Console.CursorLeft;
                                index = 0;

                                while (result == 0)
                                {
                                    DrawMenu(menuClients, row1, col1, index, rez);
                                    switch (Console.ReadKey(true).Key)
                                    {
                                        case ConsoleKey.DownArrow:
                                            if (index < menuClients.Length - 1)
                                                index++;
                                            break;
                                        case ConsoleKey.UpArrow:
                                            if (index > 0)
                                                index--;
                                            break;
                                        case
                                            ConsoleKey.Enter:
                                            switch (index)
                                            {
                                                case 0:
                                                    Console.Clear();
                                                    result = -1;
                                                    break;
                                                default:
                                                    result = index;
                                                    break;
                                            }
                                            break;
                                    }
                                }
                                
                                if (result > 0 )
                                {
                                    string[] rezClient = GetCellsValue(1, result + 1);
                                    string[] menuItems2 = new string[]
                                      { $" Выберете что изменить у организации: { rezClient[1]}",
                                    "Изменить Наименование организации",
                                    "Изменить ФИО контактного лица"
                                      };
                                    Console.Clear();
                                    int result2 = 0;
                                    index = 0;
                                    row1 = Console.CursorTop;
                                    col1 = Console.CursorLeft;
                                    while (result2 == 0)
                                    {
                                        DrawMenu(menuItems2, row1, col1, index, rez);
                                        switch (Console.ReadKey(true).Key)
                                        {
                                            case ConsoleKey.DownArrow:
                                                if (index < menuItems2.Length - 1)
                                                    index++;
                                                break;
                                            case ConsoleKey.UpArrow:
                                                if (index > 0)
                                                    index--;
                                                break;
                                            case
                                                ConsoleKey.Enter:
                                                switch (index)
                                                {
                                                    case 0:
                                                        Console.Clear();
                                                        break;
                                                    default:
                                                        result2 = index;
                                                        break;
                                                }
                                                break;
                                        }
                                    }
                                    if (result2 != 0)
                                    {

                                        if (result2 == 2)
                                            result2 = 3;
                                        string changeStr;
                                        Console.WriteLine("Введите новое значение:");
                                        changeStr = Console.ReadLine();
                                        SetCellsValue(1, result + 1, result2, changeStr);
                                        Console.WriteLine($"Изменено значение с {rezClient[result2]} на {changeStr}");
                                        Console.ReadKey();
                                    }
                                    Console.Clear();
                                }
                                break;
                            case 2:
                                string year;
                                string month;
                                DateTime dt;
                                DateTime dtReq;
                                while (true) 
                                { 
                                Console.WriteLine("Введите год (цифрами)");
                                year = Console.ReadLine();
                                Console.WriteLine("Введите месяц (цифрами)");
                                month = Console.ReadLine();
                                    
                                    if (month.Length == 1)
                                        month = "0" + month;

                                if( DateTime.TryParse("01." + month + "." + year    , out dt))
                                    {
                                        int countRequest = GetCount( 2);
                                        Dictionary<int, int> countRequestDictionary = new Dictionary<int, int>();
                                        
                                        int countClients = GetCount(1);
                                        for(int i = 2; i <= countClients; i++)
                                        {
                                            string[] Client = GetCellsValue(1, i);
                                            if (Client[0] != "")
                                                countRequestDictionary.Add(Int32.Parse(Client[0]), 0);
                                            else break;
                                        }
                                        
                                        for(int i = 2; i <= countRequest; i++)
                                        {
                                            string[] Requests = GetCellsValue(2, i);
                                            if (Requests[0] != "")
                                            {
                                                if (DateTime.TryParse(Requests[5], out dtReq))
                                                {
                                                    if (dtReq.Month == dt.Month)
                                                    {
                                                        
                                                        int r = Int32.Parse(Requests[2]);
                                                        countRequestDictionary[r]++;
                                                    }

                                                }
                                                else Console.WriteLine("Ошибка чтения даты в файле");
                                            }
                                            else break;
                                        }

                                        

                                        int[] idClient = new int[countClients];
                                        int max = 0;
                                        int countMax = 0;
                                        bool write = false;
                                        foreach(var person in countRequestDictionary)
                                        {
                                            if (person.Value > max && person.Value != 0)
                                            {
                                                max = person.Value;
                                                idClient[0] = person.Key;
                                                countMax = 0;
                                                write = true;
                                            }

                                            if (person.Value == max && person.Value != 0 )
                                            {
                                                idClient[countMax] = person.Key;
                                                countMax++;
                                            }
                                        }

                                        if (write)
                                        {
                                            for (int id = 1; id <= countMax; id++)
                                            {
                                                Console.WriteLine();
                                                string[] client;
                                                Console.WriteLine($"Клиент с наибольшим количеством заказов № {id}");
                                                int indexClient = SearchTextInt(idClient[id - 1].ToString(), 1, 0);
                                                if (indexClient != -1)
                                                {
                                                    client = GetCellsValue(1, indexClient);
                                                    Console.WriteLine($"Наименование организации: {client[1]}");
                                                    Console.WriteLine($"Адрес: { client[2]}");
                                                    Console.WriteLine($"Контактное лицо (ФИО): { client[3]}");
                                                }
                                            }
                                        }
                                        else { Console.WriteLine("Нет заявок в заданном месяце"); }
                                    }
                                else
                                    {
                                        Console.WriteLine("Неверно задан месяц либо год");
                                    }
                                break;
                                }
                                Console.ReadKey();
                                Console.Clear();
                                break;
                          
                        }
                        break;
                }
            }
        }






        private static void DrawMenu(string[] items, int row, int col, int index, string rez)
        {
            Console.SetCursorPosition(col, row);
            for (int i = 0; i < items.Length; i++)
            {
                if (i == index)
                {
                    Console.BackgroundColor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Black;
                }
                Console.WriteLine(items[i]);
                Console.ResetColor();
            }
            Console.WriteLine(rez);
            Console.WriteLine();
        }





        private static string SearchText( string searchText, int indexList, int indexColumn)
        {
            string result = null;
            Sheet sheet = (Sheet)thesheetcollection.ElementAt(indexList);
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

            foreach (Row row in sheetData.Elements<Row>())
            {
                Cell cell = row.Elements<Cell>().ElementAt(indexColumn);
                string cellvalue = cell.InnerText;
                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                {
                    cellvalue = workbookPart.SharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cellvalue)).InnerText;
                }

                if (cellvalue.ToUpper().Equals(searchText.ToUpper()))
                {
                    return row.RowIndex.ToString();
                }
            }
            return result;
        }

        private static int GetCount(int indexList)
        {
            int result = 1;
            Sheet sheet = (Sheet)thesheetcollection.ElementAt(indexList);
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            result = sheetData.Elements<Row>().Count();
            return result;
        }


        private static int SearchTextInt( string searchText, int indexList, int indexColumn)
        {
            int result = -1;
                try {
                    Sheet sheet = (Sheet)thesheetcollection.ElementAt(indexList);
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        Cell cell = row.Elements<Cell>().ElementAt(indexColumn);
                        string cellvalue = cell.InnerText;
                        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                        {
                            cellvalue = workbookPart.SharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cellvalue)).InnerText;
                        }

                        if (cellvalue.ToUpper().Equals(searchText.ToUpper()))
                        {
                            result = (int)Convert.ToInt32(row.RowIndex.Value);
                            return result;
                        }
                    }
                }
                catch{ }
                return result;
        }




        private static string[] GetCellsValue( int indexList, int indexRow)
        {
            string[] result = null;
            Sheet sheet = (Sheet)thesheetcollection.ElementAt(indexList);
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == indexRow);
            string valueCell = null;
            uint formatId;
            if (row != null)
            {
                result = new string[ (row.Elements<Cell>().Count()) ];
                int i = 0;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell != null) 
                    {
                        valueCell = cell.InnerText;
                        try 
                        {
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                            if (  sharedStringTablePart != null)
                            {
                                SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;
                                if (sharedStringTable != null) 
                                { 
                                    int index;
                                    if (int.TryParse(valueCell,out index))
                                    {
                                        valueCell = sharedStringTable.ElementAt(index).InnerText;
                                    }
                                }
                            }
                        }
                        CellFormats cellFormats = (CellFormats)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
                        CellFormat cellFormat = (CellFormat)cellFormats.ElementAt((int)cell.StyleIndex.Value);
                        formatId = cellFormat.NumberFormatId.Value;
                        if ((formatId >= 14 && formatId <= 22) || (formatId >= 45 && formatId <= 51) ||
                            (formatId >= 27 && formatId <= 31) || (formatId >= 34 && formatId <= 43) ||
                             formatId == 57 || formatId == 58)
                            valueCell = DateTime.FromOADate(double.Parse(valueCell)).ToString("dd.MM.yyyy");
                        }
                        catch
                        {
                            //
                        }
                        finally { result[i] = valueCell; }
                       
                    }
                    else
                        result[i] = " ";
                    i++;
                }
            }
            else
            {
                Console.WriteLine($"Строка с индексом {indexRow} не найдена");
            }
            return result;
        }


        private static bool SetCellsValue( int indexList, int indexRow, int indexColumn, string cellValue)
        {
            Sheet sheet = (Sheet)thesheetcollection.ElementAt(indexList);
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == indexRow);
            if (row != null)
            {
                    Cell cell = row.Elements<Cell>().ElementAt(indexColumn);
                    if (cell != null)
                    {
                        cell.DataType  = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.String );
                        cell.CellValue = new CellValue(cellValue);
                        workbookPart.Workbook.Save();
                        worksheetPart.Worksheet.Save();
                        document.Dispose();
                        document = SpreadsheetDocument.Open(fileName, true);
                        workbookPart = document.WorkbookPart;
                        thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    return true;
                    }
                    else
                    {
                        Console.WriteLine($"Столбец с индексом {indexColumn} не найден");
                    }
            }
            else
            {
                Console.WriteLine($"Строка с индексом {indexRow} не найдена");
            }
            return false;
        }
    }
}
