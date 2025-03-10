(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the add-in
            initializeAddIn();
        });
    };

    // Set up the add-in
    function initializeAddIn() {
        // Add event listener for recipients changed
        Office.context.mailbox.item.addHandlerAsync(
            Office.EventType.RecipientsChanged,
            handleRecipientsChanged,
            function(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showStatus("Event handler registered successfully.");
                } else {
                    showStatus("Error registering event handler: " + result.error.message);
                }
            }
        );
        
        showStatus("Email Domain Validator is ready!");
    }

    // Handle the event when recipients are changed
    function handleRecipientsChanged(eventArgs) {
        // Get all recipients (To, CC, BCC)
        getAllRecipients()
            .then(validateDomains)
            .catch(function(error) {
                showStatus("Error: " + error.message);
            });
    }

    // Get all recipients from the mail item
    function getAllRecipients() {
        return new Promise(function(resolve, reject) {
            try {
                const item = Office.context.mailbox.item;
                const recipients = {
                    to: [],
                    cc: [],
                    bcc: []
                };
                
                // Get To recipients
                item.to.getAsync(function(result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        recipients.to = result.value;
                        
                        // Get CC recipients
                        item.cc.getAsync(function(result) {
                            if (result.status === Office.AsyncResultStatus.Succeeded) {
                                recipients.cc = result.value;
                                
                                // Get BCC recipients
                                item.bcc.getAsync(function(result) {
                                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                                        recipients.bcc = result.value;
                                        resolve(recipients);
                                    } else {
                                        reject(result.error);
                                    }
                                });
                            } else {
                                reject(result.error);
                            }
                        });
                    } else {
                        reject(result.error);
                    }
                });
            } catch (error) {
                reject(error);
            }
        });
    }

    // Validate all email domains
    function validateDomains(recipients) {
        const allRecipients = [].concat(
            recipients.to || [],
            recipients.cc || [],
            recipients.bcc || []
        );
        
        if (allRecipients.length === 0) {
            showStatus("No recipients to validate.");
            clearResults();
            return;
        }
        
        const domainsToValidate = [];
        const recipientsByDomain = {};
        
        // Extract domains from email addresses
        allRecipients.forEach(function(recipient) {
            if (recipient.emailAddress) {
                const email = recipient.emailAddress;
                const atIndex = email.lastIndexOf('@');
                
                if (atIndex > 0 && atIndex < email.length - 1) {
                    const domain = email.substring(atIndex + 1).toLowerCase();
                    
                    // Track which recipients use each domain
                    if (!recipientsByDomain[domain]) {
                        recipientsByDomain[domain] = [];
                        domainsToValidate.push(domain);
                    }
                    
                    recipientsByDomain[domain].push(recipient);
                }
            }
        });
        
        // Update status
        showStatus(`Validating ${domainsToValidate.length} domain(s)...`);
        
        // Validate each domain using DNS lookup
        const validationPromises = domainsToValidate.map(function(domain) {
            return validateDomain(domain).then(function(isValid) {
                return {
                    domain: domain,
                    isValid: isValid,
                    recipients: recipientsByDomain[domain]
                };
            });
        });
        
        Promise.all(validationPromises)
            .then(displayValidationResults)
            .catch(function(error) {
                showStatus("Error during validation: " + error.message);
            });
    }

    // Validate a single domain using DNS lookup
    function validateDomain(domain) {
        return new Promise(function(resolve) {
            // Using a simple API to check if the domain exists
            // You may want to use a more reliable service in production
            fetch(`https://dns-api.org/MX/${domain}`)
                .then(function(response) {
                    // If we get a 200 response, the domain likely exists
                    resolve(response.ok);
                })
                .catch(function() {
                    // If there's an error, we'll mark the domain as potentially invalid
                    // but you may want to handle this differently
                    resolve(false);
                });
        });
    }

    // Display validation results
    function displayValidationResults(results) {
        const invalidDomains = results.filter(function(result) {
            return !result.isValid;
        });
        
        if (invalidDomains.length === 0) {
            showStatus("All email domains are valid.");
            clearResults();
            return;
        }
        
        showStatus(`Found ${invalidDomains.length} potentially invalid domain(s).`);
        
        const resultsDiv = document.getElementById("results");
        resultsDiv.innerHTML = "";
        
        // Create a list of invalid domains with their recipients
        const list = document.createElement("ul");
        
        invalidDomains.forEach(function(result) {
            const listItem = document.createElement("li");
            listItem.className = "invalid-domain";
            
            const domainSpan = document.createElement("div");
            domainSpan.className = "domain-name";
            domainSpan.textContent = `Invalid domain: ${result.domain}`;
            
            const recipientsList = document.createElement("ul");
            result.recipients.forEach(function(recipient) {
                const recipientItem = document.createElement("li");
                recipientItem.textContent = recipient.emailAddress;
                recipientsList.appendChild(recipientItem);
            });
            
            listItem.appendChild(domainSpan);
            listItem.appendChild(recipientsList);
            list.appendChild(listItem);
        });
        
        resultsDiv.appendChild(list);
        
        // Highlight the notification
        resultsDiv.className = "warning-highlight";
    }

    // Helper to show status messages
    function showStatus(message) {
        document.getElementById("status").textContent = message;
    }
    
    // Clear results
    function clearResults() {
        document.getElementById("results").innerHTML = "";
        document.getElementById("results").className = "";
    }
})();

