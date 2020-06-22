using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel.Composition;
using System.Net.Mail;
using System.Threading;
using System.Runtime.Remoting.Contexts;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Activities.Statements;
using System.Activities.Expressions;

namespace GetDataTableFromMail
{
    [Description("Return collection of Datatable if mailmessage contains any table, suggestion: use has HTML Table activity before this.")]
    public class GetMultipleTablesFromMailBody : NativeActivity
    {
        [RequiredArgument]
        [Category("Input")]
        public InArgument<MailMessage> Mailmessage { get; set; }

 //       [Category("Input")]
 //       [Description("Enter the table number if you want to get only single table in Datatable output section")]
 //       public InArgument<int> TableNumber { get; set; }

        [Category("Common")]
        public InArgument<int> Timeout { get; set; }

        [Category("Common")]
        public InArgument<bool> ContinueOnError { get; set; }

       [Category("Output")]
       public OutArgument<DataTable[]> DatatableArray { get; set; }

 //       [Category("Output")]

 //       public OutArgument<string> HtmlString { get; set; }

//        [Category("Output")]
//        [Description("Default table is first table for other field enter table number in TableNumber field")]
 //       public OutArgument<DataTable> Datatable { get; set; }
        protected override void CacheMetadata(NativeActivityMetadata metadata)
        {
            
            
                base.CacheMetadata(metadata);

           // metadata.AddValidationError("");

        //  if (Mailmessage == null) { throw new Exception("MailMessage cannot be Empty"); };

        //  if (OutDatatable != null && TableNumber==null) { throw new Exception("Please enter the TableNumber in input section of property, count should start from 1"); }
        }

        protected override async void Execute(NativeActivityContext context)
        {
            bool in_continueonerror = ContinueOnError.Get(context);

            try {
                CancellationToken cancellationToken = default;
                var timeout = Timeout.Get(context);
             
//                int tableNumber = TableNumber.Get(context);
                var task = GetHtmlMailBody(context);
                string htmlstring = task.Result;

                var out_NumberOfTables = NumberOfMailTables(htmlstring);

                int NumberOfTables = out_NumberOfTables.Result;

                if (NumberOfTables == 0) { throw new NullReferenceException("No table in mail Body"); };

 //               if (tableNumber - 1 > NumberOfTables) { throw new Exception("Total Number of tables in mail body is " + NumberOfTables.ToString() + " please enter the TableNumber equal or less then " + NumberOfTables.ToString()); }
 //               else
                {
                    var result_datatables = GetDatatable(htmlstring);

                    DataTable[] out_dataTables = result_datatables.Result;


                    if (await Task.WhenAny(task, Task.Delay(timeout, cancellationToken)) != task) { throw new TimeoutException(); };

//                    HtmlString.Set(context, htmlstring);

//                    if (tableNumber == 0) { Datatable.Set(context, out_dataTables[tableNumber]); }
 //                   else { Datatable.Set(context, out_dataTables[tableNumber - 1]); }
                    
                    DatatableArray.Set(context, out_dataTables);
                }
            }
 

            catch{ if (in_continueonerror == false) {  } }

        }

        private async Task<string> GetHtmlMailBody(NativeActivityContext context)
        {
          
            var mailMessage = Mailmessage.Get(context);
            var htmlstring = mailMessage.Headers.Get("HTMLBody");
            return htmlstring;
        }


        private async Task<int> NumberOfMailTables(string htmlstring) {

            string TableExpression = "<table[^>]*>(.*?)</table>";
            MatchCollection tables = Regex.Matches(htmlstring, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

            return tables.Count;
        }

        private async Task<DataTable[]> GetDatatable(string htmlinput)
        {
            const string HTML_TAG_PATTERN = "<.*?>";
           DataTable dt = null;
            DataRow dr = null;
            //    DataColumn dc = null;
            string TableExpression = "<table[^>]*>(.*?)</table>";
            string HeaderExpression = "<th[^>]*>(.*?)</th>";
            string RowExpression = "<tr[^>]*>(.*?)</tr>";
            string ColumnExpression = "<td[^>]*>(.*?)</td>";
            bool HeadersExist = false;
            int iCurrentColumn = 0;
            int iCurrentRow = 0;
            

            MatchCollection tables = Regex.Matches(htmlinput, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

            DataTable[] dataTables = new DataTable[tables.Count];

            

            int arraytableCount = 0;

            foreach (Match Table in tables)
            {
                // Reset the current row counter and the header flag    
                iCurrentRow = 0;
                HeadersExist = false;

                // Add a new table to the DataSet    
                dt = new DataTable();

                // Create the relevant amount of columns for this table (use the headers if they exist, otherwise use default names)    
                if (Table.Value.Contains("<th"))
                {
                    // Set the HeadersExist flag    
                    HeadersExist = true;

                    // Get a match for all the rows in the table    
                    MatchCollection Headers = Regex.Matches(Table.Value, HeaderExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    // Loop through each header element    
                    foreach (Match Header in Headers)
                    {
                        //dt.Columns.Add(Header.Groups(1).ToString); 

                        var header_data = Regex.Replace(Header.Groups[1].ToString(), HTML_TAG_PATTERN, string.Empty);

                        dt.Columns.Add(header_data);

                    }
                }
                else
                {
                    for (int iColumns = 1; iColumns <= Regex.Matches(Regex.Matches(Regex.Matches(Table.Value, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase)[0].ToString(), RowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase)[0].ToString(), ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase).Count; iColumns++)
                    {
                        dt.Columns.Add("Column " + iColumns);
                    }
                }

                // Get a match for all the rows in the table    
                MatchCollection Rows = Regex.Matches(Table.Value, RowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                // Loop through each row element    
                foreach (Match Row in Rows)
                {

                    // Only loop through the row if it isn't a header row    
                    if (!(iCurrentRow == 0 & HeadersExist == true))
                    {

                        // Create a new row and reset the current column counter    
                        dr = dt.NewRow();
                        iCurrentColumn = 0;

                        // Get a match for all the columns in the row    
                        MatchCollection Columns = Regex.Matches(Row.Value, ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                        // Loop through each column element    
                        foreach (Match Column in Columns)
                        {

                            DataColumnCollection columns = dt.Columns;

                            if (!columns.Contains("Column " + iCurrentColumn))
                            {
                                //Add Columns  
                                dt.Columns.Add("Column " + iCurrentColumn);
                            }
                            // Add the value to the DataRow    
                            var col_row = Regex.Replace(Column.Groups[1].ToString(), HTML_TAG_PATTERN, string.Empty);
                            dr[iCurrentColumn] = col_row;
                            // Increase the current column    
                            iCurrentColumn += 1;

                        }

                        // Add the DataRow to the DataTable    
                        dt.Rows.Add(dr);

                    }

                    // Increase the current row counter    
                    iCurrentRow += 1;
                }

                dataTables[arraytableCount] = dt;

                arraytableCount += 1;
            }
            return dataTables;
        }
    }


}
