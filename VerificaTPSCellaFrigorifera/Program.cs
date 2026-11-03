
using ClosedXML.Excel;



﻿using System;
using System.Drawing;
using ClosedXML.Excel; // Richiede il pacchetto NuGet ClosedXML

class Program
{
    static void Main()
    {

        string filePath = "C:/Documenti/VerificaTPSCellaFrigorifera.xlsx";



        // 2. CREAZIONE E SCRITTURA DEL FILE EXCEL
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("VerificaTPSCellaFrigorifero");




            // Formattazione intestazione (Riga 1)
            var headerRow = worksheet.Row(1);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Fill.BackgroundColor = XLColor.Yellow;
            headerRow.Style.Font.FontColor = XLColor.Blue;

            worksheet.Cell(1, 2).Value = "Rilevazione";
            worksheet.Cell(1, 3).Value = "Nibble";
            worksheet.Cell(1, 4).Value = "Hex";
            worksheet.Cell(1, 5).Value = "Note";

            worksheet.Cell(2, 2).Value = "Rilevazione n1";
            worksheet.Cell(3, 2).Value = "Rilevazione n2";
            worksheet.Cell(4, 2).Value = "Rilevazione n3";
            worksheet.Cell(5, 2).Value = "Rilevazione n4";
            worksheet.Cell(6, 2).Value = "Rilevazione n5";


            // RICHIESTA DEL NIBBLE

            int nibble;

            for (int i = 2; i <= 6; i++)
            {
                Console.WriteLine("Inserire il nibble");
                nibble = Convert.ToInt16(Console.ReadLine());

                worksheet.Cell(i, 3).Value = nibble;

                if (nibble == 0000)
                {
                    worksheet.Cell(i, 4).Value = "0";
                    
                }else if (nibble == 0001)
                {
                    worksheet.Cell(i, 4).Value = "1";


                }
                else if (nibble == 0010)
                {
                    worksheet.Cell(i, 4).Value = "2";
                    
                }
                else if (nibble == 0011)
                {
                    worksheet.Cell(i, 4).Value = "3";
                    
                }
                else if (nibble == 0100)
                {
                    worksheet.Cell(i, 4).Value = "4";
                  
                }
                else if (nibble == 0101)
                {
                    worksheet.Cell(i, 4).Value = "5";
                    
                }
                else if (nibble == 0110)
                {
                    worksheet.Cell(i, 4).Value = "6";
                   
                }
                else if (nibble == 0111)
                {
                    worksheet.Cell(i, 4).Value = "7";
                    
                }
                else if (nibble == 1000)
                {
                    worksheet.Cell(i, 4).Value = "8";
                  
                }
                else if (nibble == 1001)
                {
                    worksheet.Cell(i, 4).Value = "9";
                   
                }
                else if (nibble == 1010)
                {
                    worksheet.Cell(i, 4).Value = "A";
                 
                }
                else if (nibble == 1011)
                {
                    worksheet.Cell(i, 4).Value = "B";
                    
                }
                else if (nibble == 1100)
                {
                    worksheet.Cell(i, 4).Value = "C";
                  
                }
                else if (nibble == 1101)
                {
                    worksheet.Cell(i, 4).Value = "D";
                    
                }
                else if (nibble == 1110)
                {
                    worksheet.Cell(i, 4).Value = "E";
                    
                }
                else if (nibble == 1111)
                {
                    worksheet.Cell(i, 4).Value = "F"; 
                    
                }

                if (nibble <= 0010)
                {
                    worksheet.Cell(i,5).Value = "Attenzione, Rischio Congelamento";
                }
                else if (nibble >= 0011 || nibble <= 0110)
                {
                    worksheet.Cell(i, 5).Value = "Temperatura nella norma";
                }
                else if (nibble > 0110)
                {
                    worksheet.Cell(i, 5).Value = "Attenzione, Possibile Avaria";
                }



            }








            worksheet.Columns().AdjustToContents();

            // Salva il file
            workbook.SaveAs(filePath);
            Console.WriteLine($"\nIl File '{filePath}' è stato creato con successo!\n");
        }

    }
}











