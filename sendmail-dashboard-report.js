const got = require("got");
// import got from "got";  // Replace the line above with this to run standalone outside of synthetics
const assert = require("assert");
// import assert from "assert";  // Replace the line above with this to run standalone outside of synthetics

/**
 * For local development
 */
if (typeof $env === "undefined" || $env === null) {
  var $secure = new Object();  // These secure credentials are expected to be present in the New Relic account
  $secure.USER_API_KEY = '<your-new-relic-user-key>';
  $secure.SENDER_ADDRESS = '<your-azure-email-address>';
  $secure.AZURE_TENANT = '<your-azure-tenant>';
  $secure.AZURE_CLIENT_ID = '<your-azure-client-id>';
  $secure.AZURE_CLIENT_SECRET = '<your-azure-client-secret>';
}

/** CONFIGURATIONS */
const REGIONS = {
  'us': 'https://api.newrelic.com/graphql',
  'eu': 'https://api.eu.newrelic.com/graphql'
};
const NERDGRAPH_URL = REGIONS['us'];
const DASHBOARD_GUID = "<your-dashboard-guid>>";  // https://docs.newrelic.com/docs/apis/nerdgraph/examples/export-dashboards-pdfpng-using-api/#dash-multiple
const DASHBOARD_URL = '<your-dashboard-page-url>';  // The dashboard page url
const SUBJECT = 'Weekly Report'; // Subject of the email
const RECIPIENTS = [ 
  "<your-first-recipient>",
];  // Email or emails of recipients as array elements, separated by a comma
const SENDER = $secure.SENDER_ADDRESS; // Email you want to appear as the sender
/** END CONFIGURATIONS */


async function getToken() {
  try {
    let response = await got.post('https://login.microsoftonline.com/' + $secure.AZURE_TENANT + '/oauth2/v2.0/token',
      {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        form: {
          client_id: $secure.AZURE_CLIENT_ID,
          client_secret: $secure.AZURE_CLIENT_SECRET,
          grant_type: 'client_credentials',
          scope: 'https://graph.microsoft.com/.default'
        }
      }
    )
    assert.ok(response.statusCode == 200, "Error response code received during oauth " + response.statusCode);
    const responseJson = JSON.parse(response.body);
    return responseJson['access_token'];
  } catch (error) {
    console.error("Error getting token " + error)
  }
}


async function getDashboardURL() {
  const headers = { 'Content-Type': 'application/json', 'API-Key': $secure.USER_API_KEY };
  const query = `mutation {
    dashboardCreateSnapshotUrl(guid: "${DASHBOARD_GUID}")
  }`;
  const options = {
    url: NERDGRAPH_URL,
    headers: headers,
    json: { 'query': query, 'variables': {} }
  };
  
  try {
    const response = await got.post(options);
    if (response.statusCode == 200) {
      const payload = JSON.parse(response.body);
      var dashboardSnapshotUrl = payload.data.dashboardCreateSnapshotUrl;
      dashboardSnapshotUrl = dashboardSnapshotUrl.replace("PDF", "PNG")
    }
    return dashboardSnapshotUrl;
  } catch (error) {
    console.error("Error getting dashboard pdf " + error)
  }  
}


async function sendEmail(token, dashboardSnapshotUrl) {
  let toRecipients = [];
  for (const recipient of RECIPIENTS) {
    toRecipients.push({emailAddress: {address: recipient}})
  }
  const messagePayload = { 
    //Ref: https://learn.microsoft.com/en-us/graph/api/resources/message#properties
    message: {
      subject: SUBJECT,
      body: {
        contentType: 'HTML',
        content: `
                  <head>
                    <meta charset="utf-8">
                    <meta name="viewport" content="width=device-width, initial-scale=1">
                  </head>
                  <body>
                    <h1>Weekly Report</h1>
                    Click <a href="${DASHBOARD_URL}">here</a> to view the dashboard in New Relic<br>
                    Click <a href="${dashboardSnapshotUrl}">here</a> to see the PNG of this New Relic Dashboard<br><br>
                    <img src="${dashboardSnapshotUrl}"/></a>
                    <br>
                  </body>
                `
      },
      toRecipients: toRecipients
    }
  };

  try {
    await got.post({ // Send Email using Microsoft Graph
      url: `https://graph.microsoft.com/v1.0/users/${SENDER}/sendMail`,
      headers: {
        'Authorization': "Bearer " + token,
        'Content-Type': 'application/json'
      },
      json: messagePayload
    })    
  } catch (error) {
    console.error("Error sending email " + error)  
  }  
}

async function main() {
  const token = await getToken();
  const dashboardSnapshotUrl = await getDashboardURL();
  await sendEmail(token, dashboardSnapshotUrl);
}

main()