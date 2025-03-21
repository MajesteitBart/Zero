import { driver as googleDriver } from './google'
import { driver as microsoftDriver } from './microsoft'
import { type IConfig, type MailManager } from './types';

const SupportedProviders = {
    google: googleDriver,
    microsoft: microsoftDriver,
};

export const createDriver = async (
    provider: keyof typeof SupportedProviders | string,
    config: IConfig,
): Promise<MailManager> => {
    const factory = SupportedProviders[provider as keyof typeof SupportedProviders];
    if (!factory) throw new Error("Provider not supported");
    switch (provider) {
        case "google":
        case "microsoft":
            return factory(config);
        default:
            throw new Error("Provider not supported");
    }
};
