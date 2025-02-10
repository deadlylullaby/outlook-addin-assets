Office.onReady(() => {
    console.log("Outlook Add-in is ready.");
});

function reportPhishing(event) {
    Office.context.mailbox.item.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const item = result.value;
            
            const phishingReport = {
                subject: "[Phishing Report] " + item.subject,
                body: "Un utilisateur a signalé cet email comme hameçonnage.\n\n" +
                      "Expéditeur : " + item.from.emailAddress + "\n" +
                      "Objet : " + item.subject + "\n\n" +
                      "Contenu du message :\n" + item.body
            };

            Office.context.mailbox.item.displayNewMessageForm({
                to: "s.barranco@cyberg.fr",  // destination mail 
                subject: phishingReport.subject,
                body: phishingReport.body
            });
        }
    });

    event.completed();
}

Office.actions.associate("reportPhishing", reportPhishing);
