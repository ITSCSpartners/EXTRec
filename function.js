Office.onReady(() => {
    Office.actions.associate("validateRecipients", validateRecipients);
});

async function validateRecipients(event) {
    try {
        let message = Office.context.mailbox.item;
        let toRecipients = await message.to.getAsync();
        let ccRecipients = await message.cc.getAsync();

        let totalRecipients = (toRecipients.value.length + ccRecipients.value.length);

        if (totalRecipients > 7) {
            Office.context.mailbox.item.notificationMessages.addAsync(
                "recipientLimit",
                {
                    type: "error",
                    message: "You cannot send emails to more than 7 recipients in 'To' and 'CC'.",
                },
                () => {
                    event.completed({ allowEvent: false });
                }
            );
        } else {
            event.completed({ allowEvent: true });
        }
    } catch (error) {
        console.error("Error checking recipients: ", error);
        event.completed({ allowEvent: true });
    }
}
