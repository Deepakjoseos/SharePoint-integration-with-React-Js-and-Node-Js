export interface IAuthEnvProps {
    [key: string]: string;
    custom?: any;
}
export declare const getConfigFromEnvVariables: () => IAuthEnvProps | null;
