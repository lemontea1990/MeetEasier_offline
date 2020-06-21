// expose our config directly to our application using module.exports
// this configuration is used for EWS connection
// but this time, we are working on offline O365
// so you can discard this if you are offline O365
module.exports = {
  // this user MUST have full access to all the room accounts
  'exchange' : {
    'username'  : process.env.USERNAME || 'username',
    'password'  : process.env.PASSWORD || 'pwd',
    'uri'       : 'EWS uri'
  },
  // Ex: CONTOSO.COM, Contoso.com, Contoso.co.uk, etc.
  'domain' : process.env.DOMAIN || 'Domain'
};
