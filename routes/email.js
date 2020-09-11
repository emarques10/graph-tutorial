// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var express = require('express');
var router = express.Router();
var tokens = require('../tokens.js');
var graph = require('../graph.js');

/* GET /calendar */
// <GetRouteSnippet>
router.get('/',
  async function(req, res) {
    if (!req.isAuthenticated()) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      let params = {
        active: { email: true }
      };

      // Get the access token
      var accessToken;
      try {
        accessToken = await tokens.getAccessToken(req);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err)
        });
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Get the events
          var emails = await graph.getEmails(accessToken);
          params.emails = emails.value;
          console.log(emails);
        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch emails',
            debug: JSON.stringify(err)
          });
        }
      } else {
        req.flash('error_msg', 'Could not get an access token');
      }

      res.render('email', params);
    }
  }
);
router.post('/',
  async function(req, res) {
    if (!req.isAuthenticated()) {
      // Redirect unauthenticated requests to home page
      res.redirect('/')
    } else {
      let params = {
        active: { email: true }
      };

      // Get the access token
      var accessToken;
      try {
        accessToken = await tokens.getAccessToken(req);
      } catch (err) {
        req.flash('error_msg', {
          message: 'Could not get access token. Try signing out and signing in again.',
          debug: JSON.stringify(err)
        });
      }

      if (accessToken && accessToken.length > 0) {
        try {
          // Get the events
          console.log(req.body);
          await graph.sendEmail(accessToken);
          res.redirect('/')
          //var emails = await graph.sendEmail(accessToken);
          
        } catch (err) {
          req.flash('error_msg', {
            message: 'Could not fetch emails',
            debug: JSON.stringify(err)
          });
        }
      } else {
        req.flash('error_msg', 'Could not get an access token');
      }
    }
  }
);
// </GetRouteSnippet>

module.exports = router;