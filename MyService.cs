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
            //read autoRpt rows
            var info = "";
            var db = new Db();
            var rpts = db.GetJsons("select * from dbo.EasyRpt where Status=1");
            if (rpts == null)
            {
                info = "No EasyRpt Rows";
                goto lab_exit;
            }

            //get mailMessage list
            //var msgs = new List<MailMessage>();
            var smtp = _Fun.Smtp;
            foreach (var rpt in rpts)
            {
                //check
                //set mailMessage
                var rptName = rpt["Name"].ToString();
                var email = new EmailDto()
                {
                    Subject = rptName,
                    ToUsers = _Str.ToList(rpt["ToEmails"].ToString()),
                    Body = "您好, 報表如附檔。",
                };
                var msg = _Email.DtoToMsg(email, smtp);

                //sql to docx(stream)
                var ms = new MemoryStream();
                var docx = _Excel.TplToDocx(ms, _Fun.DirRoot + "EasyRptData/" + rpt["TplFile"].ToString()); //ms <-> docx
                _Excel.ExportBySql(docx, rpt["Sql"].ToString(), db);
                docx.Dispose(); //must dispose, or get empty excel !!
                ms.Position = 0;

                //set attachment
                Attachment attach = new Attachment(ms, new ContentType(ContentTypeEstr.Excel));
                attach.Name = rpt["Name"].ToString() + ".xlsx";
                msg.Attachments.Add(attach);

                //send
                _Email.SendByMsgSync(msg, smtp);    //sync send for stream attachment !!
                ms.Close(); //close after send email, or get error: cannot access a closed stream !!

                _Log.Info($"Send EasyRpt: " + rptName);
            }


        lab_exit:
            if (db != null)
                db.Dispose();
            if (info != "")
                _Log.Info(info);
        }

    }//class
}
