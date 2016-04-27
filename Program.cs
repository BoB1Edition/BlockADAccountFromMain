using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.ObjectModel;
using System.Net;
using System.Security;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;
using System.DirectoryServices.AccountManagement;

namespace ADBlocker
{
    class Program
    {
        static string strBypassList = "Обходной лист (.+)";
        static string strCancelBypassList = "Уведомление об отмене обходного листа (.+)";
        static string strDate = ".+([0-9][0-9].[0-9][0-9].2016).+";

        static string strDomain = "domain.loc";
        static string strOU = "OU=company,DC=domain,DC=loc";
        static string strUser = "DomainAdm";
        static string strPassword = "Zzaq123Zzaq123";

        static string strEmailUser = "Bypass";
        static string strEmailPassword = "Qwerty_1234";
        static string strEmailDomain = "domain.loc";
        static string strEmailUrl = "Bypass@domain.ru";

        static string strFolderName = "Bypass";
        static string strEmailFilter = "from:Bypass@domain.ru";
        static string strEmailFromName = "Server";


        static void Main(string[] args)
        {
            Regex bypassList = new Regex(strBypassList, RegexOptions.IgnoreCase);
            Regex cancelBypassList = new Regex(strCancelBypassList, RegexOptions.IgnoreCase);
            Regex date = new Regex(strDate);
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, strDomain, strOU, strUser, strPassword);


            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
            service.Credentials = new NetworkCredential(strEmailUser, strEmailPassword, strEmailDomain);
            service.AutodiscoverUrl(strEmailUrl);

            Folder webtutor = getFolder(strFolderName, service);
            FindItemsResults<Item> result = service.FindItems(WellKnownFolderName.Inbox, strEmailFilter, new ItemView(int.MaxValue));
            //result.ToArray<Item>();
            foreach(EmailMessage res in result.Reverse())
            {
                if (res.From.Name == strEmailFromName)
                {
                    if (bypassList.IsMatch(res.Subject)) {
                        Match match = bypassList.Match(res.Subject);
                        if (match.Success)
                        {
                            Console.WriteLine("bypassList: " + match.Groups[1].Value);
                            UserPrincipal up = UserPrincipal.FindByIdentity(ctx, match.Groups[1].Value);
                            if (up == null) continue;
                            res.Load();
                            foreach (string str in res.Body.Text.Split('\n'))
                            {
                                match = date.Match(str);
                                if (match.Success)
                                {
                                    Console.WriteLine("Date: " + match.Groups[1].Value);
                                    if(up.AccountExpirationDate == null || Convert.ToDateTime(match.Groups[1].Value).AddDays(1)>up.AccountExpirationDate )
                                    up.AccountExpirationDate = Convert.ToDateTime(match.Groups[1].Value).AddDays(1);
                                    up.Save();
                                    up.Dispose();
                                    break;
                                }
                            }
                        }
                        res.Copy(webtutor.Id);
                    }
                    if(cancelBypassList.IsMatch(res.Subject)) {
                        Match match = cancelBypassList.Match(res.Subject);
                        if (match.Success)
                        {
                            Console.WriteLine("cancelBypassList: " + match.Groups[1].Value);
                            UserPrincipal up = UserPrincipal.FindByIdentity(ctx, match.Groups[1].Value);
                            if (up == null) continue;
                            up.AccountExpirationDate = null;
                            up.Save();
                            up.Dispose();
                        }
                        res.Copy(webtutor.Id);
                    }
                }
            }
            Environment.Exit(0);
        }
        static Folder getFolder(String name, ExchangeService service )
        {
            FindFoldersResults ffr = service.FindFolders(WellKnownFolderName.Inbox, new FolderView(int.MaxValue));
            foreach(Folder fold in ffr) {
                if (fold.DisplayName == name)
                {
                    return fold;
                }
            }
            Folder webtutor = new Folder(service);
            webtutor.DisplayName = strFolderName;
            if (webtutor.IsNew) webtutor.Save(WellKnownFolderName.Inbox);
            return webtutor;
        }
    }
}
