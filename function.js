function validateBeforeSend(event) {
    const maxRecipients = 7;
    const maxEmailsInBody = 15;

    Office.context.mailbox.item.to.getAsync((toResult) => {
        Office.context.mailbox.item.cc.getAsync((ccResult) => {
            const totalRecipients = (toResult.value?.length || 0) + (ccResult.value?.length || 0);

            if (totalRecipients > maxRecipients) {
                event.completed({
                    allowEvent: false,
                    errorMessage: `Nedrīkst sūtīt vairāk par ${maxRecipients} adresātiem TO/CC laukos.`
                });
                return;
            }

            Office.context.mailbox.item.body.getAsync("text", (bodyResult) => {
                if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    const bodyText = bodyResult.value.split("From:")[0];
                    const emailMatches = bodyText.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g);
                    const emailCount = emailMatches ? emailMatches.length : 0;

                    if (emailCount > maxEmailsInBody) {
                        event.completed({
                            allowEvent: false,
                            errorMessage: `E-pasta saturā nedrīkst būt vairāk par ${maxEmailsInBody} e-pasta adresēm (pirms 'From:').`
                        });
                    } else {
                        event.completed({ allowEvent: true });
                    }
                } else {
                    event.completed({ allowEvent: true });
                }
            });
        });
    });
}
