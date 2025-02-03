function validateRecipients(event) {
    const maxRecipients = 7; // Maksimālais atļautais saņēmēju skaits

    Office.context.mailbox.item.to.getAsync(function (toResult) {
        Office.context.mailbox.item.cc.getAsync(function (ccResult) {
            const totalRecipients = toResult.value.length + ccResult.value.length;

            if (totalRecipients > maxRecipients) {
                event.completed({
                    allowEvent: false,
                    errorMessage: `Maksimālais saņēmēju skaits ir ${maxRecipients}.`
                });
            } else {
                event.completed({ allowEvent: true });
            }
        });
    });
}
