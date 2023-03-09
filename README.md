# nr-synthetics-ms365-sendmail

This repository provides some examples of how to send emails from New Relic Synthetics using Microsoft Graph sendMail API.

## Microsoft Graph sendMail

While Microsoft provides a SDK for accessing MS Graph, it isn’t currently available on the synthetics public minion runtime. Therefore the example code (GitHub repo) makes calls to the [sendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http) API directly.

Steps to try this out yourself (with credit to this excellent [stack overflow answer](https://stackoverflow.com/a/74498169)):
1.  Set up a free Microsoft 365 developer sandbox [subscription](https://learn.microsoft.com/en-us/office/developer-program/microsoft-365-developer-program-get-started), if you don’t have one already
2.  Create an [Azure App registration](https://go.microsoft.com/fwlink/?linkid=2083908) for sending emails
- Create an Azure App registration, no redirct url is required
- Add mail sending permission: Azure App Registration Admin > API permissions > Add permission > Microsoft Graph > Application permissions > Mail.Send

  WARNING: You will want to limit access of the app registration to specific mailboxes using [application access policy](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access). By default its tenant wide - allows impersonation of any user's mailbox!
- Grant Admin Consent to permission to your newly created Mail.Send permission (likely requires Global Admin)
- Create Application Password: Azure App Registration Admin > Certificates and secrets > Client Secrets -> New client secret (note the generated secret)
- Create a reminder/plan to refresh this secret every expiration period (in Azure as well as your own code)
3.  Consider limiting Mail.Send permission based on the Warning above (see Appendix 2 [here](https://stackoverflow.com/a/74498169))
4. Run one of the scripts in Synthetics new runtime as a Scripted API monitor from a single location
- sendmail.js: a minimal example that sends a hello world email to the recipient
- sendmail-report-dashboard.js: a script that creates a snapshot url for a dashboard and embeds the PNG image in the email body
