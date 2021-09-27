import * as React from 'react';
import styles from './ABook.module.scss';
import { IABookProps } from './IABookProps';
import { escape, groupBy, findIndex } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { SPHttpClient } from '@microsoft/sp-http';

import { ISPListEmployeesItems } from './IEmployees'; 
import * as strings from 'ABookWebPartStrings';

import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { filter } from 'lodash';

import { GroupedList, IGroup, IGroupHeaderProps } from 'office-ui-fabric-react/lib/GroupedList';
import { Toggle, IToggleStyles } from 'office-ui-fabric-react/lib/Toggle';
import { Icon, initializeIcons } from "office-ui-fabric-react";


import { GroupedListBasicExample } from './GroupedListFC';
// import { Icon, initializeIcons } from "office-ui-fabric-react";
// import { Selection, SelectionMode } from 'office-ui-fabric-react/lib/Selection';
// import { GroupedList, IGroup, IGroupRenderProps, IGroupHeaderProps, GroupHeader } from 'office-ui-fabric-react/lib/GroupedList';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IABookHooksWebPartProps } from './IABookHookProps'

const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { width: 300 } };
const persons = [
  {
    "name": "Den",
    "employee": {
      "department": "IT"
    }
  },
  {
    "name": "Greg",
    "employee": {
      "department": "IT"
    }
  },
  {
    "name": "Bob",
    "employee": {
      "department": "Administration"
    }
  }
];
const onRenderHeader = (props?: IGroupHeaderProps): JSX.Element | null => {
  if (props) {
    const toggleCollapse = (): void => {
      props.onToggleCollapse!(props.group!);
    };
    return (
      <div className={styles.groupHeader}>
        <span style={{ cursor: "pointer", fontSize: "18px", color: "grey", margin: "0 5px 0 0" }} onClick={toggleCollapse} >{props.group.name}</span>
        <Icon style={{ fontSize: "14px", cursor: "pointer", color: "grey" }} iconName={props.group!.isCollapsed ? 'CaretLeftSolid8' : "FlickUp"} onClick={toggleCollapse}></Icon>
      </div>
    );
  }
  return null;
};

const groupedListProps = {
  onRenderHeader
};

export const ABookWebPartContext = React.createContext<WebPartContext>(null);

export const ABookFC: React.FunctionComponent<IABookHooksWebPartProps> = (props) => {
  const [employees, setEmployees] = React.useState<any[]>([]);
  const [fullNameQuery, setFullNameQuery] = React.useState('');
  const [jobTitleQuery, setJobTitleQuery] = React.useState('');
  const [managerOfEmployeeQuery, setManagerOfEmployeeQuery] = React.useState('');  
  const [statusQuery, setstatusQuery] = React.useState(['active', 'maternityLeave']);
  const [isHROrAdmin, setIsHROrAdmin] = React.useState(false);

  const toggleStyles: Partial<IToggleStyles> = { root: { marginBottom: '20px' } };
  interface IPersons {
    name: string;
    employee: {
      department: string;
    };  
  }

  const _getListOfContacts = () => {    
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
        // console.log('groups with Title', userGroups["Title"]);
        // console.log('strGroups', strGroups);                
      });
    });
  };


  React.useEffect(() => {
    console.log("useEffect is running - loading employees");
    _checkIfUserInGroups('Servier Ukraine Administrators', 'Servier Ukraine HR');    
    _getListOfContacts(); 
  }, []);

  const filteredByStatus = employees
    .filter((employee) => statusQuery.find(status => status === employee.statusEmployee));

  const filterPersons = React.useCallback(
    (persons: any[]): any[] => {
      return (
        persons.filter((person) => {    
          const conditionOfFiltering = (person.fullName
          && person.fullName.toLowerCase().includes(fullNameQuery.toLowerCase()))
          && (person.jobTitle
          && person.jobTitle.toLowerCase().includes(jobTitleQuery.toLowerCase()))
          && (!person.managerCard || 
          person.managerCard 
            && person.managerCard.Title.toLowerCase().includes(managerOfEmployeeQuery.toLowerCase()));
          return conditionOfFiltering;
      }));
    },
    [],
    ); 



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
      console.log('onChange declaring');              
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


  
  const GroupedListTest: React.FunctionComponent = () => {
    const generateGroups = (sortedPersons: any[]) => {
      const groupedPersons: any = groupBy(sortedPersons, (i: any) => i.employee && i.employee.department && i.employee.department);
      console.log('groupedPersons NEW!!!', groupedPersons);
      let groups: IGroup[] = [];
      for (const person in groupedPersons) {
        groups.push({
          name: person,
          key: person,
          startIndex: findIndex(sortedPersons, (i: any) => i.employee.department == person),
          count: groupedPersons[person].length,
          isCollapsed: sortedPersons.find((i: any) => i.employee.department == person), // 
          // sortedPersons.find((i: any) => i.employee.department == person)
          // findIndex(sortedPersons, (i: any) => i[filterBy] == person) == 0 ? false : true 
        });
      }
      console.log('groups', groups);
      
      return groups;
    };

    const onRenderCell = (nestingDepth?: number, item?: IPersons, itemIndex?: number): JSX.Element => {
      return  (
        <div className={styles.card_container}>
          <div className={styles.container_contacts}>
            <h2>{item.name}</h2>
            {item.employee.department}
          </div>
        </div>
      );
    };

    return (
      <ABookWebPartContext.Provider value={props.context}>
      <div>
          <GroupedList
            items={persons}
            // eslint-disable-next-line react/jsx-no-bind
            groupProps={groupedListProps}
            onRenderCell={onRenderCell}
            groups={generateGroups(persons)}
          />
      </div>
      </ABookWebPartContext.Provider>
    );
  };

 
  // console.log(`fullNameQuery: ${fullNameQuery}`, `jobTitleQuery: ${jobTitleQuery}`, `managerOfEmployeeQuery: ${managerOfEmployeeQuery}`);

  // console.log('isHROrAdmin', isHROrAdmin);  
  console.log('rendering');

    
  return (
    
    
    <div className={ styles.aBook }>
      <div className={ styles.main_container }>
        <div style={{"width": '100%'}}>
             <GroupedListTest
             />
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
          
          <div className={ styles.form }>
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

        </div>
      </div>
    </div>
    
  );
};
 