
function SendEmail()
{
    var outlookApp = new ActiveXObject("Outlook.Application");
    var nameSpace = outlookApp.getNameSpace("MAPI");
    mailFolder = nameSpace.getDefaultFolder(6);
    mailItem = mailFolder.Items.add('IPM.Note.FormA');
    mailItem.Subject = "Phase 1.5 training completed";
    mailItem.To = "dalal.aljassem@gmail.com";
    mailItem.HTMLBody = "Hello, this is to confirm that I have complete Phase 1.5 training modules";
    mailItem.display(0);
}