export interface ISPListEmployeesItems {
    Title: string;
    addressEmployee: string;
    birthdayEmployee: string;
    fullName: string;
    jobTitle: string;
    levelEmployee: string;
    managerOfEmployee: string;
    statusEmployee: string;
    employeeCard: {
        EMail: string;
        Id: number;
        Title: string;
        Department: string;
        WorkPhone: string;
        MobilePhone: string;
        JobTitle: string;
        Office: string;
    };
    managerCard: {
        Id: number;
        Title: string;
        EMail: string;
    };
}