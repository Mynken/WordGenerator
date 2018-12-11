using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Web;
using Wordgenerator.Models;
using Wordgenerator.Models.DAL.Additional;
using Wordgenerator.Models.DAL.Kontrahent;
using Xceed.Words.NET;

namespace Wordgenerator.Logic.Additional
{
    public class AdditionalIEGenerator
    {
        public string CreateIEDocument(AdditionalIE data, string path)
        {
            int pageSize = 8;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("uk-UA");
            string fileNamePath = string.Format(path + @"\Додаткова угода {0}.docx", data.FirstAgreementNumber);
            var nowDate = DateTime.Now.ToString("yyyyMMdd_HHmm");

            var document = DocX.Create(fileNamePath);

            document.MarginTop = 20;
            document.MarginRight = 40;

            document.InsertParagraph(string.Format("ДОДАТКОВА УГОДА № {0}", data.FirstAgreementNumber))
              .Font(new Font("Cambria"))
              .FontSize(13)
              .Bold()
              .Alignment = Alignment.center;


            document.InsertParagraph(string.Format("до Генерального договору № {0} від {1} року",
                data.GeneralAgreementType, data.GeneralAgreementDate.AddHours(data.TimeOffset).ToString("d MMMM yyyy")))
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Bold()
                .Alignment = Alignment.center;

            var headerInfo = document.AddTable(1, 2);
            headerInfo.Design = TableDesign.TableNormal;
            headerInfo.AutoFit = AutoFit.Window;
            headerInfo.Rows[0].Cells[0].Paragraphs[0].Append("м. Київ")
                 .Font(new Xceed.Words.NET.Font("Cambria"))
                 .FontSize(pageSize)
                 .Bold()
                .Alignment = Alignment.left;
            headerInfo.Rows[0].Cells[1].Paragraphs[0].Append(data.CityDate.AddHours(data.TimeOffset).ToString("d MMMM yyyy") + " року")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Bold()
               .Alignment = Alignment.right;

            document.InsertTable(headerInfo);

            document.InsertParagraph("Товариство з обмеженою відповідальністю \"КІНОМАНІЯ\",")
              .SpacingBefore(10d)
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append(" далі – Дистриб’ютор, в особі директора Буймістер Людмили Анатоліївни," +
                " яка діє на підставі Статуту, з однієї сторони, та")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            document.InsertParagraph(data.Kontrahent.FullName)
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append(", далі – Демонстратор, в особі " + data.Kontrahent.ActingUnder + ", з іншої сторони, а разом – Сторони,")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            document.InsertParagraph(string.Format("уклали цю Додаткову угоду, надалі – Угода, до Генерального договору № {0} від {1} року (далі - Договір)," +
                " домовились про таке:", data.GeneralAgreementType, data.GeneralAgreementDate.ToString("d MMMM yyyy")))
               .Font(new Xceed.Words.NET.Font("Cambria"))
               .FontSize(pageSize)
               .SpacingAfter(10)
               .Alignment = Alignment.both;


            document.InsertParagraph("РЕКВІЗИТИ ТА ПІДПИСИ СТОРІН")
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Bold()
            .SpacingAfter(10d)
            .Alignment = Alignment.center;

            var requsiteTable = document.AddTable(5, 2);

            requsiteTable.Design = TableDesign.TableNormal;
            requsiteTable.AutoFit = AutoFit.Window;
            requsiteTable.Rows[0].Cells[0].Paragraphs[0].Append("ДИСТРИБ’ЮТОР")
                 .Font(new Xceed.Words.NET.Font("Cambria"))
                 .FontSize(pageSize)
                 .Bold()
                 .SpacingAfter(10d)
                 .UnderlineStyle(UnderlineStyle.singleLine)
                .Alignment = Alignment.center;
            requsiteTable.Rows[0].Cells[1].Paragraphs[0].Append("ДЕМОНСТРАТОР")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .SpacingAfter(10d)
                .FontSize(pageSize)
                .UnderlineStyle(UnderlineStyle.singleLine)
                .Bold()
               .Alignment = Alignment.center;

            requsiteTable.Rows[1].Cells[0].Paragraphs[0].Append("Товариство з обмеженою відповідальністю \"КІНОМАНІЯ\"")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Bold()
               .Alignment = Alignment.center;
            requsiteTable.Rows[1].Cells[1].Paragraphs[0].Append(data.Kontrahent.FullName)
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Bold()
               .Alignment = Alignment.center;

            requsiteTable.Rows[2].Cells[0].Paragraphs[0].Append("Юридична адреса та адреса для листування: 01042," + "\n" +
                "м. Київ, вул. Іоанна Павла ІІ, б. 4/6, корп. \"А\", к. 821." + "\n" +
                "П/р № 26008364029900 в АТ \"УКРСИББАНК\", м. Харків," + "\n" +
                "МФО 351005." + "\n" +
                "П/р №26005455018547 в АТ „ОТП \"Банк\", МФО 300528. " + "\n" +
                "Ідентифікаційний код 32208748." + "\n" +
                "Свідоцтво про внесення суб’єкта кінематографії до Державного реєстру виробників, розповсюджувачів" +
                " і демонстраторів фільмів серії РУ № 000122 від 01.02.2012." + "\n" +
                "Свідоцтво ПДВ № 200024567 від 07.02.2012." + "\n" +
                "Дистриб’ютор є платником податку на прибуток на загальних підставах.")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .SpacingAfter(10d)
               .Alignment = Alignment.left;
            requsiteTable.Rows[2].Cells[1].Paragraphs[0].Append(string.Format("Місцезнаходження: {0}" + "\n" +
                "{1}" + "\n" +
                "МФО {2}" + "\n" +
                "Ідентифікаційний код {3}" + "\n" +
                "{4}" + "\n" +
                "Демонстратор є платником {5}", data.Kontrahent.Adress, data.Kontrahent.CurrentBankAccount, data.Kontrahent.Mfo, data.Kontrahent.IdentificationCode,
                data.Kontrahent.RegistrationLicense, data.Kontrahent.TaxInfo))
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .SpacingAfter(10d)
               .Alignment = Alignment.left;

            requsiteTable.Rows[3].Cells[0].Paragraphs[0].Append("Директор")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Bold()
                .SpacingAfter(20d)
               .Alignment = Alignment.left;
            requsiteTable.Rows[3].Cells[1].Paragraphs[0].Append(data.Kontrahent.Position)
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Bold()
                .SpacingAfter(20d)
               .Alignment = Alignment.left;

            requsiteTable.Rows[4].Cells[0].Paragraphs[0].Append("________________________________________________ Л.А. Буймістер")
               .Font(new Xceed.Words.NET.Font("Cambria"))
               .FontSize(pageSize)
               .Bold()
              .Alignment = Alignment.left;

            //if (kontrahent.Signature.Length >= 14)
            //{
            //    requsiteTable.Rows[4].Cells[1].Paragraphs[0].Append(string.Format("_________________________ {0}", kontrahent.Signature))
            //    .Font(new Xceed.Words.NET.Font("Cambria"))
            //    .FontSize(pageSize)
            //    .Bold()
            //   .Alignment = Alignment.left;
            //}
            //else
            //{
            //    requsiteTable.Rows[4].Cells[1].Paragraphs[0].Append(string.Format("_______________________________________________________ {0}", kontrahent.Signature))
            //    .Font(new Xceed.Words.NET.Font("Cambria"))
            //    .FontSize(pageSize)
            //    .Bold()
            //   .Alignment = Alignment.left;
            //}

            int allSpace = 70 - data.Kontrahent.Signature.Length;
            string signatureUnderline = "";

            for (int i = 1; i < allSpace; i++)
            {
                signatureUnderline += "_";
            }

            requsiteTable.Rows[4].Cells[1].Paragraphs[0].Append(string.Format("{0} {1}", signatureUnderline, data.Kontrahent.Signature))
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
             .Alignment = Alignment.left;



            document.InsertTable(requsiteTable);

            document.Save();

            return fileNamePath;
        }
    }
}
