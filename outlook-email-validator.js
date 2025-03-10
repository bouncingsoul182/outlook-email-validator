// outlook-email-validator.js
(function () {
    Office.onReady(function () {
        if (Office.context.mailbox) {
            Office.context.mailbox.item.addHandlerAsync(
                Office.EventType.RecipientsChanged,
                validateEmailDomains
            );
        }
    });

    async function validateEmailDomains(event) {
        const item = Office.context.mailbox.item;
        if (!item) return;

        const recipients = await getRecipients(item);
        const invalidDomains = [];

        for (const email of recipients) {
            const domain = email.split("@")[1];
            if (domain && !(await checkDomain(domain))) {
                invalidDomains.push(email);
            }
        }

        if (invalidDomains.length > 0) {
            item.notificationMessages.addAsync("invalidEmailAlert", {
                type: "error",
                message: "Warning: The following email domains may be invalid: " + invalidDomains.join(", ")
            });
        } else {
            item.notificationMessages.removeAsync("invalidEmailAlert");
        }
    }

    async function getRecipients(item) {
        return new Promise((resolve) => {
            item.getRecipientsAsync((result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const emails = result.value.map((recipient) => recipient.emailAddress);
                    resolve(emails);
                } else {
                    resolve([]);
                }
            });
        });
    }

    async function checkDomain(domain) {
        try {
            const response = await fetch(`https://dns.google/resolve?name=${domain}&type=MX`);
            const data = await response.json();
            return data.Answer && data.Answer.length > 0;
        } catch (error) {
            return false;
        }
    }
})();
