using ClosedXML.Excel;
using System;
using System.IO;

namespace Model_1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = "C:/Projects/ISRPO/Model_1/bin/Model.xlsx";

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                Console.WriteLine($"Лист: {worksheet.Name}");

                while (true)
                {
                    Console.WriteLine("=============================================");
                    Console.WriteLine("Предмет      Баллы");
                    var row = 1;
                    var column = 1;
                    var cellValue = worksheet.Cell(row, column).Value.ToString();
                    while (cellValue != "")
                    {
                        Console.Write($"{row}. {worksheet.Cell(row, column).Value.ToString()}   ---   ");
                        Console.Write(worksheet.Cell(row, column + 1).Value.ToString());
                        Console.WriteLine();
                        row++;
                        cellValue = worksheet.Cell(row, column).Value.ToString();
                    }

                    Console.WriteLine();
                    Console.WriteLine("Выберите доп. действие: ");
                    Console.WriteLine("1 - добавить предмет ");
                    Console.WriteLine("2 - удалить предмет ");
                    Console.Write("Действие: ");
                    try
                    {
                        int choice = int.Parse(Console.ReadLine());
                        Console.WriteLine();

                        if (choice == 1)
                        {
                            Console.Write("Введите название предмета: ");
                            string nameLesson = Console.ReadLine();
                            Console.Write("Введите баллы по предмету: ");
                            try
                            {
                                int markLesson = int.Parse(Console.ReadLine());
                                if (markLesson > 100 || markLesson < 0)
                                {
                                    Console.WriteLine("Количество баллов не может превышать 100 или быть меньше 0");
                                    continue;
                                }
                                while (cellValue != "")
                                {
                                    row++;
                                    cellValue = worksheet.Cell(row, column).Value.ToString();
                                }
                                worksheet.Cell(row, column).Value = nameLesson;
                                worksheet.Cell(row, column + 1).Value = markLesson;
                            }
                            catch (FormatException)
                            {
                                Console.WriteLine($"!!!Ошибка записи формата:");
                                Console.WriteLine($"Необходимо ввести номер предмета по таблице выше!!!");
                            }
                            
                        }
                        else if (choice == 2)
                        {
                            Console.Write("Введите номер предмета: ");
                            try
                            {
                                int numberLesson = int.Parse(Console.ReadLine());
                                worksheet.Row(numberLesson).Delete();
                            } catch (FormatException) {
                                Console.WriteLine($"!!!Ошибка записи формата:");
                                Console.WriteLine($"Необходимо ввести номер предмета по таблице выше!!!");
                            }

                        }
                        workbook.Save();
                    } catch (FormatException) {
                        Console.WriteLine($"Ошибка записи формата:");
                        Console.WriteLine($"Необходимо ввести номер действия");
                    }
                }
            }
        }
    }
}