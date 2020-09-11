// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

module.exports = {
  getUserDetails: async function(accessToken) {
    const client = getAuthenticatedClient(accessToken);

    const user = await client.api('/me').get();
    return user;
  },

  getEmails: async function(accessToken) {
    const client = getAuthenticatedClient(accessToken);
    
    const messages = await client
      .api('/me/messages')
      .select('from,sender,subject,toRecipients')
      .get();
    
    return messages;
  },

  sendEmail: async function(accessToken) {
    const client = getAuthenticatedClient(accessToken);
    console.log('TESTE');
    const sendTest = {
      message: {
        subject: "TESTE EMAIL",
        body: {
          contentType: "HTML",
          content: "EMAIL ENVIADO UTILIZANDO A API DO MICROSOFT GRAPH"
        },
        toRecipients: [
          {
            emailAddress: {
              address: "paulo.salles@embraer.com.br"
            }
          }
        ],
        internetMessageHeaders:[
          {
            name:"x-custom-header-group-name",
            value:"Nevada"
          },
          {
            name:"x-custom-header-group-id",
            value:"NV001"
          }
        ]
      }
    };
    
    let res = await client.api('/me/sendMail')
        .post(sendTest);
  },



  // <GetEventsSnippet>
  getEvents: async function(accessToken) {
    const client = getAuthenticatedClient(accessToken);

    const events = await client
      .api('/me/events')
      .select('subject,organizer,start,end')
      .orderby('createdDateTime DESC')
      .get();

    return events;
  }
  // </GetEventsSnippet>
};

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken);
    }
  });

  return client;
}