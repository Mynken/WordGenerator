using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Wordgenerator.Models;
using Wordgenerator.Models.DAL;
using Wordgenerator.Models.DAL.Films;
using Wordgenerator.Models.DAL.Kontrahent;
using Xceed.Words.NET;
using System.Threading;
using System.Globalization;
using System.IO;

namespace Wordgenerator.Logic
{
    public class WordIEGenerator
    {
        public string CreateIEDocument(KontrahentIE kontrahent, Film film, ModelDoc dataForDoc, List<Trailer> trailers, string path)
        {
            int pageSize = dataForDoc.SessionModel.Count < 4 ? 8 : 7;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("uk-UA");
            string fileNamePath = "";

            var nowDate = DateTime.Now.ToString("yyyyMMdd_HHmm");

            if (dataForDoc.ThirdAgreementNumber != null)
            {
                fileNamePath = string.Format(path +
                    @"\{0}-{1}-{2}-{3}.docx", kontrahent.Number, film.Number, dataForDoc.ThirdAgreementNumber, nowDate);
            }
            else
            {
                fileNamePath = string.Format(path +
                    @"\{0}-{1}-{2}.docx", kontrahent.Number, film.Number, nowDate);
            }

            var document = DocX.Create(fileNamePath);

            document.MarginTop = 20;
            document.MarginRight = 40;

            if (dataForDoc.ThirdAgreementNumber != null)
            {
                document.InsertParagraph(string.Format("ДОДАТОК № {0}/{1}/{2}", kontrahent.Number, film.Number, dataForDoc.ThirdAgreementNumber))
                    .Font(new Xceed.Words.NET.Font("Cambria"))
                    .FontSize(13)
                    .Bold()
                    .Alignment = Alignment.center;
            }
            else
            {
                document.InsertParagraph(string.Format("ДОДАТОК № {0}/{1}", kontrahent.Number, film.Number))
                  .Font(new Xceed.Words.NET.Font("Cambria"))
                  .FontSize(13)
                  .Bold()
                  .Alignment = Alignment.center;
            }

            document.InsertParagraph(string.Format("до Генерального договору № {0} від {1} року",
                dataForDoc.GeneralAgreementType, dataForDoc.GeneralAgreementDate.AddHours(dataForDoc.TimeZoneOffset).ToString("d MMMM yyyy")))
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
            headerInfo.Rows[0].Cells[1].Paragraphs[0].Append(dataForDoc.FilmAgreeementDate.AddHours(dataForDoc.TimeZoneOffset).ToString("d MMMM yyyy") + " року")
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

            document.InsertParagraph(kontrahent.FullName)
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append(" далі – Демонстратор, в особі " + kontrahent.ActingUnder + ", з іншої сторони, а разом – Сторони,")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            document.InsertParagraph(string.Format("на виконання вимог Генерального Договору № {0} від {1} року (далі - Договір)," +
                " домовились про таке:", dataForDoc.GeneralAgreementType, dataForDoc.GeneralAgreementDate.ToString("d MMMM yyyy")))
               .Font(new Xceed.Words.NET.Font("Cambria"))
               .FontSize(pageSize)
               .SpacingAfter(10)
               .Alignment = Alignment.both;

            document.InsertParagraph("1.   ")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append("Відповідно до умов Договору Дистриб’ютор надає Демонстратору Право Демонстрування іноземного Фільму:")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            document.InsertParagraph(film.Name)
               .Font(new Xceed.Words.NET.Font("Cambria"))
               .FontSize(13)
               .Bold()
               .Alignment = Alignment.center;

            document.InsertParagraph(string.Format("який {0} на території України.", dataForDoc.DuplicatedLanguage))
               .Font(new Xceed.Words.NET.Font("Cambria"))
               .FontSize(pageSize)
               .Bold()
               .UnderlineStyle(UnderlineStyle.singleLine)
               .Alignment = Alignment.center;

            document.InsertParagraph("2.   ")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append(string.Format("Кінотеатр: {0}", dataForDoc.CinemaName))
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            document.InsertParagraph("3.   ")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append(string.Format("Місто: {0}", dataForDoc.City))
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            document.InsertParagraph("4.   ")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append(string.Format("Правовласник, рік: {0}", film.OwnerAndYear))
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            document.InsertParagraph("5.   ")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Bold()
             .Append(string.Format("Країна виробник: {0}", film.Country))
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Alignment = Alignment.both;

            document.InsertParagraph("6.   ")
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Bold()
            .Append(string.Format("Хронометраж: {0}", film.DurationTime))
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Alignment = Alignment.both;

            document.InsertParagraph("7.   ")
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Bold()
            .Append(string.Format("Мова демонстрування фільму:  {0}", film.Language))
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Alignment = Alignment.both;

            document.InsertParagraph("8.   ")
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Bold()
            .Append(string.Format("Період Демонстрування Фільму: з {0} року по {1} року",
            dataForDoc.DemonstrationPeriodFrom.AddHours(dataForDoc.TimeZoneOffset).ToString("d MMMM yyyy"),
            dataForDoc.DemonstrationPeriodTo.AddHours(dataForDoc.TimeZoneOffset).ToString("d MMMM yyyy")))
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Alignment = Alignment.both;

            document.InsertParagraph("9.   ")
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Bold()
            .Append(string.Format("Формат Фільмокопії: {0}", dataForDoc.FilmFormat))
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Alignment = Alignment.both;

            document.InsertParagraph("10. ")
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Bold()
            .Append("Сеанси щоденно:")
            .Font(new Xceed.Words.NET.Font("Cambria"))
            .FontSize(pageSize)
            .Alignment = Alignment.both;

            // Add a Table into the document and sets its values.
            var sessionTable = document.AddTable(dataForDoc.SessionModel.Count + 1, 5);
            sessionTable.AutoFit = AutoFit.Contents;
            sessionTable.Alignment = Alignment.center;
            sessionTable.Rows[0].Cells[0].Paragraphs[0].Append("№ тижня")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Alignment = Alignment.center;
            sessionTable.Rows[0].Cells[1].Paragraphs[0].Append("Дата початку")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Alignment = Alignment.center;
            sessionTable.Rows[0].Cells[2].Paragraphs[0].Append("Дата кінця")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Alignment = Alignment.center;
            sessionTable.Rows[0].Cells[3].Paragraphs[0].Append("Сеанси")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Alignment = Alignment.center;
            sessionTable.Rows[0].Cells[4].Paragraphs[0].Append("Дата оплати")
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Alignment = Alignment.center;
            if (dataForDoc.SessionModel.Count == 1)
            {
                sessionTable.Rows[1].Cells[0].Paragraphs[0].Append(string.Format("{0}", dataForDoc.SessionModel[0].NumberOfWeek))
                 .Font(new Xceed.Words.NET.Font("Cambria"))
                 .FontSize(pageSize)
                 .Alignment = Alignment.center;
                sessionTable.Rows[1].Cells[1].Paragraphs[0].Append(dataForDoc.SessionModel[0].StartDate.AddHours(dataForDoc.TimeZoneOffset).ToShortDateString())
                    .Font(new Xceed.Words.NET.Font("Cambria"))
                    .FontSize(pageSize)
                    .Alignment = Alignment.center;
                sessionTable.Rows[1].Cells[2].Paragraphs[0].Append(dataForDoc.SessionModel[0].EndDate.AddHours(dataForDoc.TimeZoneOffset).ToShortDateString())
                    .Font(new Xceed.Words.NET.Font("Cambria"))
                    .FontSize(pageSize)
                    .Alignment = Alignment.center;
                sessionTable.Rows[1].Cells[3].Paragraphs[0].Append(dataForDoc.SessionModel[0].SessionInfo)
                    .Font(new Xceed.Words.NET.Font("Cambria"))
                    .FontSize(pageSize)
                    .Alignment = Alignment.center;
                sessionTable.Rows[1].Cells[4].Paragraphs[0].Append(dataForDoc.SessionModel[0].PaymentDate.AddHours(dataForDoc.TimeZoneOffset).ToShortDateString())
                    .Font(new Xceed.Words.NET.Font("Cambria"))
                    .FontSize(pageSize)
                    .Alignment = Alignment.center;
            }
            else
            {
                for (int i = 1; i <= dataForDoc.SessionModel.Count; i++)
                {
                    sessionTable.Rows[i].Cells[0].Paragraphs[0].Append(string.Format("{0}", dataForDoc.SessionModel[i - 1].NumberOfWeek))
                        .Font(new Xceed.Words.NET.Font("Cambria"))
                        .FontSize(pageSize)
                        .Alignment = Alignment.center;
                    sessionTable.Rows[i].Cells[1].Paragraphs[0].Append(dataForDoc.SessionModel[i - 1].StartDate.AddHours(dataForDoc.TimeZoneOffset).ToShortDateString())
                        .Font(new Xceed.Words.NET.Font("Cambria"))
                        .FontSize(pageSize)
                        .Alignment = Alignment.center;
                    sessionTable.Rows[i].Cells[2].Paragraphs[0].Append(dataForDoc.SessionModel[i - 1].EndDate.AddHours(dataForDoc.TimeZoneOffset).ToShortDateString())
                        .Font(new Xceed.Words.NET.Font("Cambria"))
                        .FontSize(pageSize)
                        .Alignment = Alignment.center;
                    sessionTable.Rows[i].Cells[3].Paragraphs[0].Append(dataForDoc.SessionModel[i - 1].SessionInfo)
                        .Font(new Xceed.Words.NET.Font("Cambria"))
                        .FontSize(pageSize)
                        .Alignment = Alignment.center;
                    sessionTable.Rows[i].Cells[4].Paragraphs[0].Append(dataForDoc.SessionModel[i - 1].PaymentDate.AddHours(dataForDoc.TimeZoneOffset).ToShortDateString())
                        .Font(new Xceed.Words.NET.Font("Cambria"))
                        .FontSize(pageSize)
                        .Alignment = Alignment.center;
                }

            }

            document.InsertTable(sessionTable);

            document.InsertParagraph("11. ")
             .SpacingBefore(10d)
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Bold()
             .Append("Початок демонстрування відбувається у часових межах Сеансів, погоджених у п. 3.10 Договору: ")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Append(string.Format("для {0}", dataForDoc.TypeOfFilm))
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Bold()
             .Append(" – " + dataForDoc.CartoonFilmInfo)
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Alignment = Alignment.both;

            document.InsertParagraph("12. ")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append("Дистриб’ютор передає у тимчасове користування Демонстратору Фільмокопію – 1 (одна) шт.")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            document.InsertParagraph("13. ")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append("Демонстратор відповідно до умов Договору обов’язково демонструє такі Анонсні" +
               " ролики перед кожним сеансом Фільму та в такій послідовності:")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            for (int i = 0; i < trailers.Count; i++)
            {
                document.InsertParagraph(string.Format("\t{0}. {1}", i + 1, trailers[i].Name))
                    .Font(new Xceed.Words.NET.Font("Cambria"))
                    .FontSize(pageSize)
                    .Alignment = Alignment.left;
            }

            document.InsertParagraph("14. ")
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Bold()
              .Append(dataForDoc.RojaltiInfo)
              .Font(new Xceed.Words.NET.Font("Cambria"))
              .FontSize(pageSize)
              .Alignment = Alignment.both;

            document.InsertParagraph("15. ")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Bold()
             .Append("Демонстратор сплачує Роялті шляхом перерахування коштів на один із поточних рахунків Дистриб’ютора ")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Append("№26008364029900 в АТ „УкрСиббанк”, МФО 351005")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Bold()
             .Append(" або ")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Append("№26005455018547 в АТ „ОТП Банк”, МФО 300528")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Bold()
             .Append(" не пізніше ")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Append(dataForDoc.DaysInfo + " банківських днів після закінчення кожного тижня Демонстрування Фільму.")
             .Bold()
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Alignment = Alignment.both;

            document.InsertParagraph("16. ")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Bold()
             .Append("У випадках, не передбачених Додатком, Сторони керуються Договором та чинним законодавством України.")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Alignment = Alignment.both;

            document.InsertParagraph("17. ")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .Bold()
             .Append("Додаток складений у двох примірниках українською мовою, які мають однакову юридичну силу," +
             " по одному для кожної із Сторін. ")
             .Font(new Xceed.Words.NET.Font("Cambria"))
             .FontSize(pageSize)
             .SpacingAfter(10d)
             .Alignment = Alignment.both;

            document.InsertParagraph("18. РЕКВІЗИТИ ТА ПІДПИСИ СТОРІН")
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
            requsiteTable.Rows[1].Cells[1].Paragraphs[0].Append(kontrahent.FullName)
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
                "Демонстартор є платником {5}", kontrahent.Adress, kontrahent.CurrentBankAccount, kontrahent.Mfo, kontrahent.IdentificationCode,
                kontrahent.RegistrationLicense, kontrahent.TaxInfo))
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
            requsiteTable.Rows[3].Cells[1].Paragraphs[0].Append(kontrahent.Position)
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

            if(kontrahent.Signature.Length >= 14)
            {
                requsiteTable.Rows[4].Cells[1].Paragraphs[0].Append(string.Format("_________________________ {0}", kontrahent.Signature))
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Bold()
               .Alignment = Alignment.left;
            }
            else
            {
                requsiteTable.Rows[4].Cells[1].Paragraphs[0].Append(string.Format("_______________________________________________________ {0}", kontrahent.Signature))
                .Font(new Xceed.Words.NET.Font("Cambria"))
                .FontSize(pageSize)
                .Bold()
               .Alignment = Alignment.left;
            }


            document.InsertTable(requsiteTable);

            document.Save();

            return fileNamePath;
        }
    }
}