import * as React from 'react';
import styles from './ABook.module.scss';
import { IABookProps } from './IABookProps';
import { escape, groupBy, findIndex } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { SPHttpClient } from '@microsoft/sp-http';
import 'moment';

import { ISPListEmployeesItems } from './IEmployees'; 
import * as strings from 'ABookWebPartStrings';

import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { filter } from 'lodash';

import { GroupedListBasicExample } from './GroupedListFC';

// import { Icon, initializeIcons } from "office-ui-fabric-react";
// import { Selection, SelectionMode } from 'office-ui-fabric-react/lib/Selection';
// import { GroupedList, IGroup, IGroupRenderProps, IGroupHeaderProps, GroupHeader } from 'office-ui-fabric-react/lib/GroupedList';

const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300 } };

interface ArrayConstructor {
  from(arrayLike: any, mapFn?, thisArg?): Array<any>;
}

const ABookFC: React.FunctionComponent<IABookProps> = (props) => {
  const [employees, setEmployees] = React.useState<any[]>([]);
  const [fullNameQuery, setFullNameQuery] = React.useState('');
  const [jobTitleQuery, setJobTitleQuery] = React.useState('');
  const [managerOfEmployeeQuery, setManagerOfEmployeeQuery] = React.useState('');  
  const [statusQuery, setstatusQuery] = React.useState(['active', 'maternityLeave']);
  const [isHROrAdmin, setIsHROrAdmin] = React.useState(false);

  const _getListOfContacts = () => {
//     Attachments: false
// AuthorId: 20
// ComplianceAssetId: null
// ContentTypeId: "0x0100041ED6490ACF3449B33EC0C7642279FF00CC06B6C503411948820F6A2578D31F78"
// Created: "2021-09-22T05:44:20Z"
// EditorId: 20
// FileSystemObjectType: 0
// GUID: "858ef9cd-5c23-4ee9-8f31-d253fcabb3f6"
// ID: 27
// Id: 27
// Modified: "2021-09-22T06:03:54Z"
// OData__UIVersionString: "8.0"
// ServerRedirectedEmbedUri: null
// ServerRedirectedEmbedUrl: ""
// Title: "Артем Левенец"
// addressEmployee: null
// birthdayEmployee: "02/04/2000 00:00:00"
// employeeCardId: 14
// employeeCardStringId: "14"
// fullName: "Артем Левенец"
// jobTitle: "Менеджер проектов"
// levelEmployee: null
// managerCardId: 14
// managerCardStringId: "14"
// managerOfEmployee: null
// statusEmployee: "active"

// - ПІБ (англ.)
// - Mail
// - Робочий телефон
// - Мобільний телефон
// - Підрозділ (англ.)
// - Посада (англ.)
// - Керівник
// - Місто (англ.)
// - Visa

    sp.web.lists
      .getByTitle('Employees')
      .items
      .select(
        '*', 
        'employeeCard/Id', 
        'employeeCard/Title', 
        'employeeCard/EMail', 
        'employeeCard/Department', 
        'employeeCard/WorkPhone', 
        'employeeCard/MobilePhone', 
        'employeeCard/JobTitle', 
        'employeeCard/Office', 
        'managerCard/Id', 
        'managerCard/Title', 
        'managerCard/EMail'
      )
      .expand('employeeCard', 'managerCard').getAll()
      .then(
        (response: any) => {     
          console.log('response', response);
          if (response) {        
            console.log('response is setting to state');
            setEmployees(response);           
          }        
      })
      .catch((e) => console.log(`getListOfContacts error: ${e.message}`)
      );
  };

  const _checkIfUserInGroups = async(...strGroups: string[]) => {
    let groups = await sp.web.currentUser.groups.get().then((response: any) => {
      response.forEach((userGroups) => {
        if (strGroups.includes(userGroups["Title"])) {
          setIsHROrAdmin(true);
        }
        console.log('groups with Title', userGroups["Title"]);
        console.log('strGroups', strGroups);                
      });
    });
  };

  const filteredByStatus = employees
    .filter((employee) => statusQuery.find(status => status === employee.statusEmployee));

  const filterPersons = persons => persons
  .filter((person) => {
    
    const conditionOfFiltering = (person.fullName
    && person.fullName.toLowerCase().includes(fullNameQuery.toLowerCase()))
    && (person.jobTitle
    && person.jobTitle.toLowerCase().includes(jobTitleQuery.toLowerCase()))
    && (!person.managerCard || 
    person.managerCard 
      && person.managerCard.Title.toLowerCase().includes(managerOfEmployeeQuery.toLowerCase()));
    return conditionOfFiltering;
  })
  ;

  const onChangeValue = React.useCallback(
    (event: React.ChangeEvent<HTMLInputElement>, newValue?: string) => {
      if ((event.target as HTMLInputElement).name === "fullName") {
        setFullNameQuery(newValue || '');
      }
      if ((event.target as HTMLInputElement).name === "jobTitle") {
        setJobTitleQuery(newValue || '');
      } 
      if ((event.target as HTMLInputElement).name === "managerOfEmployee") {
        setManagerOfEmployeeQuery(newValue || '');
      }        
    },
    [],
  ); 

  const handleChangeStatus = React.useCallback(
    (event: React.ChangeEvent<HTMLSelectElement>) => {
      const value: any[] = Array.from(
        event.target.selectedOptions,
        option => option.value,
      );
      console.log("value status", value);
      
      setstatusQuery(value);
    }, []);    

  React.useEffect(() => {
    console.log("useEffect is running - loading employees");
    _checkIfUserInGroups('Servier Ukraine Administrators', 'Servier Ukraine HR');    
    _getListOfContacts(); 
  }, []);
 
  console.log(`fullNameQuery: ${fullNameQuery}`, `jobTitleQuery: ${jobTitleQuery}`, `managerOfEmployeeQuery: ${managerOfEmployeeQuery}`);
  console.log("employees", employees);  
  console.log('isHROrAdmin', isHROrAdmin);  
  console.log('rendering');

    
  return (
    <div className={ styles.aBook }>
      <div className={ styles.container }>
        <div className={ styles.row }>
          <div className={ styles.column }>
            <h1>Search fields</h1>
            <Stack>
              <select
                multiple
                value={statusQuery}
                onChange={handleChangeStatus}
              >
                <option 
                  label="active" 
                  value="active"
                  selected                
                >active</option>
                <option 
                  label="maternityLeave" 
                  value="maternityLeave" 
                  selected
                >maternityLeave</option>
                {isHROrAdmin ? <option 
                  label="fired" 
                  value="fired" 
                  selected={false}
                >fired</option> : ''} 
              </select>
              
              <TextField          
                label="fullName"
                name="fullName" 
                value={fullNameQuery}
                onChange={onChangeValue} 
              />        
              <TextField 
                label="jobTitle"
                name="jobTitle" 
                value={jobTitleQuery}
                onChange={onChangeValue} 
              />    
              <TextField 
                label="managerOfEmployee" 
                name="managerOfEmployee"
                value={managerOfEmployeeQuery}
                onChange={onChangeValue}
              />      
            </Stack> 
          </div>

          <div className={ styles.column }> 
              <GroupedListBasicExample />
              <h1>
                List of employees: 
              </h1>
              {filterPersons(filteredByStatus).map((item:ISPListEmployeesItems) => {
                  return (
                    <ul>
                      <li>
                        <p>
                          {`Name: ${item.fullName}, Job Title ${item.jobTitle}`}  
                        </p>
                        <p>
                          {`Dep: ${item.employeeCard.Department}, Location ${item.employeeCard.Office}`}  
                        </p>
                        <p>
                          {`EMail: ${item.employeeCard.EMail}`}  
                        </p>
                        <p>
                          {item.managerCard && `Manager: ${item.managerCard.Title}`}  
                        </p>
                        <p>
                          {`Title: ${item.Title}, Address ${item.addressEmployee}`}  
                        </p>
                        <p>
                          {`Birthday: ${item.birthdayEmployee}, Level ${item.levelEmployee}`}  
                        </p>
                      </li>
                    </ul>
                  );
                })
              }
          </div>

        </div>
      </div>
    </div>
  );
};

export default React.memo(ABookFC);