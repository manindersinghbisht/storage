using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel.Composition;
using System.Data;
using System.ComponentModel;
using System.Activities.Statements;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Activities.Expressions;
using System.Drawing;
using System.Windows.Forms;

namespace Datatable_to_HTML
{
    [Description("convert a datatable to html text string")]
    public class Convert_DatTable_to_HTML : CodeActivity
    {
        
        [Category("Common")]
        [Description("enter True or False")]
        [DefaultValue(false)]
        public InArgument<bool> ContinueOnError { get; set; }

        [Category("Input")]
        [Description("insert a datatable")]
        public InArgument<DataTable> Input_datatable { get; set; }

        //Table body alignment...

         public enum txtAlign { 

            Center,
            Right,
            Left,
        } 

        [Category("Input table body alignment")]
        public txtAlign TableTextAlign { get; set; }

        //Table body color section...

        [Category("Input Table color")]
        [Description("Select the color from dropdown")]
        public Color  HeaderBackroundColor { get; set; }

        [Category("Input Table color")]
        [Description("Select the color from dropdown")]
 
        public Color HeaderTextColor { get; set; }

        [Category("Input Table color")]
        [Description("Select the color from dropdown")]
        public Color TableBodyTextColor { get; set; } 
        [Category("Input Table color")]
        [Description("Select the color from dropdown")]
        public Color TableBodyBackgroundColor { get; set; } 
        
        //output variable
        [Category("output")]
        [Description("output html body")]
        public OutArgument<string> HtmlBodyWithTable { get; set; }

        [Category("output")]
        [Description("output html table")]
        public OutArgument<string> HtmlTable { get; set; }



        protected override void Execute(CodeActivityContext context)
        {
            var in_continueOnError = ContinueOnError.Get(context);
            bool ActivityRun = false;

            if (in_continueOnError)
            {

                if (Input_datatable.Get(context) != null)
                {
                    ActivityRun = true;
                }
            }
            else
            {
                ActivityRun = true;

            }
            if (ActivityRun) { 
                var in_Datatable = Input_datatable.Get(context);

                string in_headTxtColor = HeaderTextColor.Name;

                var in_headerBackgroundTxtColor = HeaderBackroundColor.Name;

                var in_TableBodyTxtColor = TableBodyTextColor.Name;

                var in_TableBodyBgColor = TableBodyBackgroundColor.Name;

                var in_align = TableTextAlign;

                var out_result_body = "";
                var out_result_table = "";

                
                out_result_table= "<table style='border-collapse:collapse; text-align:" + in_align.ToString() + "' border='1'><tr style='color:" + in_headTxtColor.ToString() + ";background:" + in_headerBackgroundTxtColor.ToString() + "'>";
                
                out_result_body = "<html><head></head><body><table style='border-collapse:collapse; text-align:" + in_align.ToString() + "' border='1'><tr style='color:" + in_headTxtColor.ToString() + ";background:" + in_headerBackgroundTxtColor.ToString() + "'>";
                foreach (DataColumn col in in_Datatable.Columns)
                {
                    out_result_table =out_result_table+ "<th>" + col.ColumnName + "</th>";
                    out_result_body = out_result_body + "<th>" + col.ColumnName + "</th>";
                }
                out_result_table = out_result_table + "</tr>";
                out_result_body = out_result_body + "</tr>";

                foreach (DataRow row in in_Datatable.Rows)
                {
                    out_result_table =out_result_table+ "<tr style='color:" + in_TableBodyTxtColor.ToString() + ";background:" + in_TableBodyBgColor.ToString() + "'>";
                    out_result_body = out_result_body + "<tr style='color:" + in_TableBodyTxtColor.ToString() + ";background:" + in_TableBodyBgColor.ToString() + "'>";
                    foreach (DataColumn col in in_Datatable.Columns)
                    {
                        out_result_table=out_result_table+ "<td>" + row[col.ColumnName].ToString() + "</td>";
                        out_result_body = out_result_body + "<td>" + row[col.ColumnName].ToString() + "</td>";
                    }
                    out_result_table = out_result_table + "</tr>";
                    out_result_body = out_result_body + "</tr>";
                }


                out_result_table = out_result_table + "</table>";
                out_result_body = out_result_body + "</table></body></html>";

                HtmlTable.Set(context, out_result_table);
                HtmlBodyWithTable.Set(context, out_result_body);            
        };

        }
    }
}
