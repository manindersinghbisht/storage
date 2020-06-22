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

    namespace GetDataTableFromMail
{
        [Description("Return true if mailmessage has any table, if not returns false")]
        public class HasHTMLTable : CodeActivity
        {

        [RequiredArgument]
        [Category("Input")]
           public InArgument<MailMessage> Mailmessage { get; set; }

            [Category("Output")]
            public OutArgument<bool> HasHtmlTable { get; set; }

            protected override void CacheMetadata(CodeActivityMetadata metadata)
            {
                base.CacheMetadata(metadata);

            }


            protected override void Execute(CodeActivityContext context)
            {

            var mailMessage = Mailmessage.Get(context);

            var mailHtmlString = mailMessage.Headers.Get("HTMLBody");


                string TableExpression = "<table[^>]*>(.*?)</table>";
                MatchCollection tables = Regex.Matches(mailHtmlString, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                if (tables.Count > 0) { HasHtmlTable.Set(context, true); }
                else { HasHtmlTable.Set(context, false); }
            }
        }
    }
