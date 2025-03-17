export  interface IUserProfile {
    displayName: string;
    givenName: string;
    surname: string;
    mail: string;
    userPrincipalName: string;
    jobTitle: string;
    department: string;
    officeLocation: string;
    mobilePhone: string;
    businessPhones: string[];
    preferredLanguage: string;
    photo: string;
    id: string;
    companyName:string;
    accountEnabled:boolean;
    presence:string;
}
