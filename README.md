# outlook-addon-js

Building Outlook Addons with Javascript: Creates an Addon for Outlook 365 that interact with the Mailbox and get subject and attachments data and finally send it to a backend REST endpoint, using SSL certificates selfsigned.

Based on: https://docs.microsoft.com/en-us/outlook/add-ins/quick-start

Note: It is required for consuming the backend, set the proper keys required for a SSL conection. This should be done by creating a certs/ folder containing the keys.

For troubleshooting I recomend these links:
- https://github.com/OfficeDev/outlook-dev-docs/blob/master/docs/add-ins/troubleshoot-outlook-add-in-activation.md
- https://dev.office.com/docs/add-ins/testing/troubleshoot-manifest