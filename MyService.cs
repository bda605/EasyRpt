using Base.Enums;
using Base.Models;
using Base.Services;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;

namespace EasyRpt
{
    public class MyService
    {
        public void Run()
        {
            const string preLog = "EasyRpt: ";
            _Log.Info(preLog + "Start.");

            #region 1.read XpEasyRpt rows
            var info = "";
            var db = new Db();
            var rpts = db.GetJsons("select * from dbo.XpEasyRpt where Status=1");
            if (rpts == null)
            {
                info = "No XpEasyRpt Rows";
                goto lab_exit;
            }
            #endregion

            //send report loog
            var smtp = _Fun.Smtp;
            foreach (var rpt in rpts)
            {
                #region 2.set mailMessage
                var rptName = rpt["Name"].ToString();
                var email = new EmailDto()
                {
                    Subject = rptName,
                    ToUsers = _Str.ToList(rpt["ToEmails"].ToString()),
                    Body = "Hello, please check attached report.",
                };
                var msg = _Email.DtoToMsg(email, smtp);
                #endregion

                //3.sql to Memory Stream docx
                var ms = new MemoryStream();
                var docx = _Excel.GetMsDocxByFile(_Fun.DirRoot + "EasyRptData/" + rpt["TplFile"].ToString(), ms); //ms <-> docx
                _Excel.DocxBySql(rpt["Sql"].ToString(), docx, 1, db);
                docx.Dispose(); //must dispose, or get empty excel !!

                //4.set attachment
                ms.Position = 0;
                Attachment attach = new Attachment(ms, new ContentType(ContentTypeEstr.Excel));
                attach.Name = rptName + ".xlsx";
                msg.Attachments.Add(attach);

                //5.send email
                _Email.SendByMsgSync(msg, smtp);    //sync send for stream attachment !!
                ms.Close(); //close after send email, or get error: cannot access a closed stream !!

                _Log.Info(preLog + "Send " + rptName);
            }

            #region close db & log
        lab_exit:
            if (db != null)
                db.Dispose();
            if (info != "")
                _Log.Info(preLog + info);
            _Log.Info(preLog + "End.");
            #endregion
        }

    }//class
}
