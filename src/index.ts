import { EmailClient, EmailMessage, EmailSendResult } from '@azure/communication-email';
import { ManagedIdentityCredential } from '@azure/identity';

class ValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ValidationError';
    this.message = message;
  }
}

const additionalKeys = ['attachments', 'disableUserEngagementTracking', 'headers'] as const;

const emailRegex =
  /^[-!#$%&'*+\/0-9=?A-Z^_a-z{|}~](\.?[-!#$%&'*+\/0-9=?A-Z^_a-z`{|}~])*@[a-zA-Z0-9](-*\.?[a-zA-Z0-9])*\.[a-zA-Z](-?[a-zA-Z0-9])+$/;

const recipentWithEmailRegex = new RegExp(
  `^(?<displayName>.*) <(?<address>${emailRegex.source.slice(1, -1)})>$`,
);

const isEmail = (text: string): boolean => emailRegex.test(text);
const isRecipentWithEmail = (text: string): boolean => recipentWithEmailRegex.test(text);

type Address = { address: string; displayName?: string };
type ProviderOptions = {
  endpoint: string;
  useManagedIdentity?: boolean;
  identityClientId?: string;
};
type Settings = {
  defaultFrom?: string;
};

module.exports = {
  provider: 'azure',
  name: 'Azure Communication Service',

  init: (providerOptions: ProviderOptions = { endpoint: '' }, settings: Settings = {}) => {
    let emailClient: EmailClient;

    if (providerOptions.useManagedIdentity && providerOptions.identityClientId) {
      const credential = new ManagedIdentityCredential(providerOptions.identityClientId);
      emailClient = new EmailClient(providerOptions.endpoint, credential);
    } else {
      emailClient = new EmailClient(providerOptions.endpoint);
    }

    const transformAddress = (
      address: string | Address | (string | Address)[] | undefined,
      defaultAddress?: string,
    ): Address[] | undefined => {
      if (!address)
        return defaultAddress ? [{ address: defaultAddress }] : undefined;

      if (typeof address === 'string') {
        if (isEmail(address)) return [{ address }];
        if (isRecipentWithEmail(address)) {
          const match = address.match(recipentWithEmailRegex);
          if (match?.groups)
            return [{ address: match.groups.address, displayName: match.groups.displayName }];
        }
        throw new ValidationError(`Invalid email address: ${address}`);
      }

      if (!Array.isArray(address) && 'address' in address) return [address as Address];
      return address as Address[];
    };

    return {
      send: async (options: any): Promise<EmailSendResult> => {
        const message: EmailMessage = {
          senderAddress: transformAddress(options.from, settings.defaultFrom)?.[0].address || '',
          replyTo: transformAddress(options.replyTo, settings.defaultFrom),
          content: {
            subject: options.subject || '',
            plainText: options.text || '',
            html: options.html || '',
          },
          recipients: {
            to: transformAddress(options.to),
            cc: transformAddress(options.cc),
            bcc: transformAddress(options.bcc),
          },
          ...Object.fromEntries(
            Object.entries(options).filter(([key]) =>
              (additionalKeys as readonly string[]).includes(key),
            ),
          ),
        };

        const poller = await emailClient.beginSend(message);
        const response = await poller.pollUntilDone();
        return response;
      },
    };
  },
};
