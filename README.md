## Installation
`npm install strapi-provider-email-azure-v5`

## Strapi config (config/plugins.js)
```js
module.exports = ({ env }) => ({
  email: {
    config: {
      provider: 'strapi-provider-email-azure-v5',
      providerOptions: {
        endpoint: env('AZURE_ENDPOINT'),  // connection string OR endpoint URL
        useManagedIdentity: env('USE_MANAGED_IDENTITY', false),
        identityClientId: env('IDENTITY_CLIENT_ID'), // optional
      },
      settings: {
        defaultFrom: env('FALLBACK_EMAIL'),
      },
    },
  },
});
```
