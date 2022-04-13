using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Uniconta.API.Plugin;
using Uniconta.ClientTools.DataModel;
using Uniconta.API.Service;
using Uniconta.API.System;
using Uniconta.DataModel;
using Uniconta.Common;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using Uniconta.ClientTools.Page;
using System.Threading;

namespace CreateCalenderPoint
{
    public class AddToCalender : PageEventsBase
    {
        private CrmFollowUpClient CrmFollowUpClient;
        private CrmFollowUpClient CrmFollowUpClientLaest;
        // private CrmFollowUpClient[] CrmFollowUpClientList;
        public List<CrmFollowUpClient> CrmFollowUpClientList { get; set; }
        private CrudAPI crudAPI;
        public GridBasePage mypage;
       // public Uniconta.ClientTools.Page.FormBasePage test;
        public override bool OnMenuItemClicked(string ActionType, object sender, object arguments)
        {
            if (ActionType == "PostSaveGrid")
            {
                try
                {
                    CrmFollowUpClientList = mypage.gridControl.ItemsSource as List<CrmFollowUpClient>;

                  

                    foreach (var cm in CrmFollowUpClientList)
                    {
                        GenerateMailInvitation(cm);
                    }
                }
                catch (Exception ex)
                {
                   
                }
                return base.OnMenuItemClicked(ActionType, sender, arguments);
            }
            if (ActionType == "PreSaveGrid")
            {

                return base.OnMenuItemClicked(ActionType, sender, arguments);
            }
            if (ActionType == "PostSave")
            {
                try
                { 
                    crudAPI = api as CrudAPI;
                    CrmFollowUpClient = (CrmFollowUpClient)master;

                    if (CrmFollowUpClient.PrimaryKeyId == 0)
                    //der er tale om en oprettelse så data skal findes som sidste indsatte post
                    {
                        var CrmFollowUpClientsList = crudAPI.Query<CrmFollowUpClient>().Result.ToList();
                   ;
                        var Greateskey = -1;
                        foreach (var cm in CrmFollowUpClientsList)
                        {
                            if (cm.PrimaryKeyId > Greateskey)
                            {
                                Greateskey = cm.PrimaryKeyId;
                                CrmFollowUpClient = cm;
                            }
                        }                   

                    }
                    crudAPI.Read(CrmFollowUpClient);

                    GenerateMailInvitation(CrmFollowUpClient);

                    return base.OnMenuItemClicked(ActionType, sender, arguments);
                }
                //    return base.OnMenuItemClicked(ActionType, sender, arguments);
                catch (Exception ex)
                {
                    return base.OnMenuItemClicked(ActionType, sender, arguments); 
                }
            }
            else if (ActionType == "PostDelete")
            {
                CrmFollowUpClient = (CrmFollowUpClient)master;

                return true;
            }
            else
                return true;
        }

        private void GenerateMailInvitation(CrmFollowUpClient LocalcrmFollowUpClient)
        {
            //employe er ikke udfyldt korrekt, så derfor læses CrmFollowUpClient
            CrmFollowUpClientLaest = GetCrmFollowUpClient(LocalcrmFollowUpClient.PrimaryKeyId);

            if (!(CrmFollowUpClientLaest.Employee == null))
            {
                var body = ""; //der laves linieskift mellem de enkelte oplysninger
                if (!(CrmFollowUpClientLaest.FollowUpAction == ""))
                    body = body + CrmFollowUpClientLaest.FollowUpAction + " <br />";
                if (!(CrmFollowUpClientLaest.ContactEmail == ""))
                    body = body + CrmFollowUpClientLaest.ContactEmail + " <br />";
                if (!(CrmFollowUpClientLaest.ContactPerson == ""))
                    body = body + CrmFollowUpClientLaest.ContactPerson + " <br />";
                if (!(CrmFollowUpClientLaest.Phone == ""))
                    body = body + CrmFollowUpClientLaest.Phone;


                var body2 = "";  //her kan der ikke tilføjes liniebreak
                if (!(CrmFollowUpClientLaest.FollowUpAction == ""))
                    body2 = body2 + CrmFollowUpClientLaest.FollowUpAction + " ";
                if (!(CrmFollowUpClientLaest.ContactPerson == ""))
                    body2 = body2 + CrmFollowUpClientLaest.ContactPerson;


                var medarbejderinit = CrmFollowUpClientLaest.Employee;

                var MedarbejderListe = crudAPI.Query<EmployeeClient>().Result.ToList();

                var medarbejder = MedarbejderListe.Find(x => x.Number == medarbejderinit);

                if (!(medarbejder == null))
                {
                    if (!(medarbejder.Email == "") && (!(CrmFollowUpClientLaest.FollowUp == null) ))  //der skal være en Email på medarbejderen som skal modtage invitationen, og der skal være en opfølgningsdato

                        if (CrmFollowUpClientLaest.FollowUp >= System.DateTime.Now)     //Opfølgning tilbage i tiden bliver der ikke genereret mail på

                        SendHTMLEmailWithGoogleInvite("ht@mrc.dk", medarbejder.Email, "Uniconta opfølgning", body, body2, CrmFollowUpClientLaest.FollowUp);
                }
            };
        }


        public override void Init(object page, CrudAPI api, UnicontaBaseEntity master)
        {
            try
            {
                base.Init(page, api, master);

                mypage = (GridBasePage)page;
             
            }
            catch (Exception ex)
            {
                
            }
            crudAPI = api as CrudAPI;
           // this.api = api;
            this.master = master;
        }

      

        public CrmFollowUpClient GetCrmFollowUpClient(int PrimaryKeyId)
        {
            //NTT FIXET
            var crit = new List<PropValuePair>();
            var pair = PropValuePair.GenereteWhereElements("RowId", typeof(int), PrimaryKeyId.ToString());
            crit.Add(pair);
            var CrmFollowUpClients2 = crudAPI.Query<CrmFollowUpClient>(crit).Result;
            //return CrmFollowUpClients;


            //var CrmFollowUpClients = crudAPI.Query<CrmFollowUpClient>().Result.ToList();

            //var CrmFollowUpClienten = CrmFollowUpClients.Find(d => d.PrimaryKeyId == PrimaryKeyId);

            return CrmFollowUpClients2[0];
        }


        public static void SendHTMLEmailWithGoogleInvite(string from, string to, string subj, string body, string body2, DateTime date)
        {
            try
            {
                MailAddress t = new MailAddress(to);
                MailAddress f = new MailAddress(from);
                MailMessage message = new MailMessage(f, t);
                message.IsBodyHtml = true;
                message.Subject = subj;
                message.Body = body;
                // Use the application or machine configuration to get the
                // host, port, and credentials.
                SmtpClient client = new SmtpClient();
                client.Host = "192.168.100.11";
               

                client.Credentials = new NetworkCredential("ht@mrc.dk", "Rasmus96");
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = false;
                client.Port = 25252;
                client.UseDefaultCredentials = true;


                System.Net.Mime.ContentType typeHtml = new System.Net.Mime.ContentType("text/html");

                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(body, typeHtml);
                message.AlternateViews.Add(htmlView);

                StringBuilder str = new StringBuilder();
                str.AppendLine("BEGIN:VCALENDAR");
                str.AppendLine("PRODID:-//GeO");
                str.AppendLine("VERSION:2.0");
                str.AppendLine("METHOD:REQUEST");
                str.AppendLine("BEGIN:VEVENT");
                str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", date));
                str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", date));
                str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", date));
                str.AppendLine("LOCATION: " + body2);
                str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
                str.AppendLine(string.Format("DESCRIPTION;ENCODING=QUOTED-PRINTABLE:{0}", body));
                str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", body));
                str.AppendLine(string.Format("SUMMARY;ENCODING=QUOTED-PRINTABLE:{0}", message.Subject));
                str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", message.From.Address));
                str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE;CUTYPE=INDIVIDUAL;PARTSTAT=ACCEPTED:mailto:{1}", message.From.DisplayName, message.From.Address));
                str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE;CUTYPE=INDIVIDUAL;PARTSTAT=NEEDS-ACTION;ROLE=REQ-PARTICIPANT;SCHEDULE-STATUS=1.2:mailto:{1}", message.To[0].DisplayName, message.To[0].Address));
                str.AppendLine("BEGIN:VALARM");
                str.AppendLine("TRIGGER:-PT24H");
                str.AppendLine("ACTION:DISPLAY");
                str.AppendLine("DESCRIPTION;ENCODING=QUOTED-PRINTABLE:Reminder");
                str.AppendLine("END:VALARM");
                str.AppendLine("END:VEVENT");
                str.AppendLine("END:VCALENDAR");
                System.Net.Mime.ContentType type = new System.Net.Mime.ContentType("text/calendar");
                type.Parameters.Add("method", "REQUEST");
                type.Parameters.Add("name", "ginvite.ics");
                message.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(str.ToString(), type));



                client.Send(message);
            }

            catch (Exception ex)
            {
                if (Session.GlobalSession.LoggedIn)
                    Session.GlobalSession.LogOut();
            }






           


        }
    }
    }