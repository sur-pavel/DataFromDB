using System;
using System.Collections.Generic;
using System.Linq;
using ManagedIrbis;
using ManagedIrbis.Search;
using ManagedIrbis.Batch;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DataFromDB
{
    class IrbisHandler
    {
        internal Logging logging;
        private string sequentialSpecification = "v910 : 'НП' and p(v686) and p(v606^a)";
        private string connectionString = "server=194.169.10.3;port=8888;user=1;password=1;";
        private BatchRecordReader reader;
        private WordHandler wordHandler;

        public IrbisHandler()
        {
            logging = new Logging();
            logging.CreateLogFile();
        }
        internal void Perform()
        {
            try
            {
                using (IrbisConnection connection = new IrbisConnection(connectionString))
                {
                    AnalizeDatabases(connection);
                    //SearchInMPDA(connection);
                }
            }
            catch (Exception ex)
            {
                logging.WriteLine(ex.Message);
                logging.WriteLine(ex.StackTrace);
            }
        }

        private void AnalizeDatabases(IrbisConnection connection)
        {
            DatabaseInfo[] dataBaseInfos = connection.ListDatabases("DBNAM2_2.MNU");
            Dictionary<DatabaseInfo, int> databasesList = new Dictionary<DatabaseInfo, int>();
            List<DatabaseInfo> errorDbList = new List<DatabaseInfo>();

            foreach (DatabaseInfo databaseInfo in dataBaseInfos)
            {
                connection.Database = databaseInfo.Name;
               try
                {                                   
                    databasesList.Add(databaseInfo, connection.GetMaxMfn());
                }
                catch (Exception ex)
                {                    
                    errorDbList.Add(databaseInfo);
                    continue;
                }
            }
            //var sortedDatabasesList =
            //    from entry in databasesList
            //    orderby entry.Value
            //    descending
            //    select entry;
            var sortedDatabasesList =
                from entry in databasesList
                orderby entry.Key.Name
                ascending
                select entry;
            foreach (var pair in sortedDatabasesList)
            {

                logging.WriteLine(pair.Key.Name.PadRight(15) + " has " + pair.Value.ToString().PadLeft(8) + "\t\t" + pair.Key.Description);
            }
            logging.WriteLine("\nDatabases with errors: ");
            List<DatabaseInfo> sortedList = !errorDbList.Equals(null)? errorDbList.OrderBy(db=>db.Name).ToList() : new List<DatabaseInfo>();
            foreach(DatabaseInfo errorDb in sortedList)
            {
                logging.WriteLine(errorDb.Name);
            }
            
        }

        private void SearchInMPDA(IrbisConnection connection)
        {
            connection.Database = "MPDA";
            SearchParameters parameters = new SearchParameters { SequentialSpecification = sequentialSpecification };
            int[] foundRecordsMfn = connection.SequentialSearch(parameters);
            logging.WriteLine("found Records: " + foundRecordsMfn.Length);
            reader = new BatchRecordReader(connection, "MPDA", 500, foundRecordsMfn);
            //SaveToExcel(reader);
            //SaveToWord(reader);
        }

        private void SaveToWord(BatchRecordReader reader)
        {
            wordHandler = new WordHandler();

            foreach (MarcRecord currentRecord in reader)
            {
                if (HasOneRubric(currentRecord))
                {
                    string invNums = FilterInventoryNumbers(currentRecord, "НП");
                    logging.WriteLine(invNums);
                    wordHandler.AddString(invNums);
                }

            }
            logging.WriteLine("adding invNums complete");
            wordHandler.SaveDoc();
            wordHandler.Quit();
            logging.WriteLine("Word File saved");
        }

        private bool HasOneRubric(MarcRecord currentRecord)
        {
            RecordField[] fields606 = currentRecord.Fields.GetField(606);
            logging.WriteLine("606 fields count: " + fields606.Length);
            if (fields606.Length < 2) return true;
            return false;
        }


        private void SaveToExcel(BatchRecordReader reader)
        {
            ExcelHandler excelHandler = new ExcelHandler();
            excelHandler.CreatExcelObject();
            foreach (MarcRecord currentRecord in reader)
            {
                excelHandler.AddRow(GetBriefDiscription(currentRecord));
            }

            logging.WriteLine("Add rows");
            excelHandler.SaveFile();
            logging.WriteLine("Save file");
        }

        private string GetSubField(MarcRecord currentRecord, int tag, char code)
        {

            string subField = currentRecord.FMA(tag, code)[0];
            return subField == null ? "" : subField;
        }

        private BriefDiscription GetBriefDiscription(MarcRecord currentRecord)
        {
            BriefDiscription brief = new BriefDiscription();
            brief.Mfn = currentRecord.Mfn;

            string[] workList = currentRecord.FMA(920);

            if (workList[0].Equals("SPEC"))
            {
                if (GetSubField(currentRecord, 961, 'z').Equals("ДА"))
                {
                    brief.Autor = WithComma(GetSubField(currentRecord, 961, 'a')) + WithComma(GetSubField(currentRecord, 961, 'b')) + GetSubField(currentRecord, 961, '1');
                }
                brief.Title = GetSubField(currentRecord, 461, 'c') + ". " + GetSubField(currentRecord, 200, 'v') + " " + GetSubField(currentRecord, 200, 'a');
                brief.Location = GetSubField(currentRecord, 461, 'd').Equals("") ? GetSubField(currentRecord, 210, 'a') : GetSubField(currentRecord, 461, 'd');
                brief.Year = GetSubField(currentRecord, 461, 'h').Equals("") ? GetSubField(currentRecord, 210, 'd') : GetSubField(currentRecord, 461, 'h');
            }
            else
            {
                brief.Autor = WithComma(GetSubField(currentRecord, 700, 'a')) + WithComma(GetSubField(currentRecord, 700, 'b')) + GetSubField(currentRecord, 700, '1');
                brief.Title = GetSubField(currentRecord, 200, 'a');
                brief.Location = GetSubField(currentRecord, 210, 'a');
                brief.Year = GetSubField(currentRecord, 210, 'd');
            };

            string[] invNumbers = currentRecord.FMA(910, 'b');
            brief.NumberOfCopies = invNumbers.Length;
            brief.FirstInvNum = GetSubField(currentRecord, 910, 'b');
            return brief;
        }

        private string WithComma(string v)
        {
            if (v.Equals("")) v = v + ", ";
            return v;
        }

        private string FilterInventoryNumbers(MarcRecord record, string fond)
        {
            return string.Join("\n", record.Fields.GetField(910)
                .Where(field =>
                {
                    var place = field.GetFirstSubFieldValue('d');
                    return place != null && place.Contains(fond);
                })
                .Select(field => field.GetFirstSubFieldValue('b'))
                .Where(number => !string.IsNullOrEmpty(number) && !number.Contains("-") && !number.Contains("/"))
                .Distinct()
                .ToArray());
        }
    }
}
