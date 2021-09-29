import * as React from 'react';
import styles from './ABook.module.scss';
import { IABookProps } from './IABookProps';
import { groupBy, findIndex } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";

import * as strings from 'ABookWebPartStrings';

import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { filter } from 'lodash';
import { GroupedList, IGroup, IGroupHeaderProps } from 'office-ui-fabric-react/lib/GroupedList';
import { Toggle, IToggleStyles } from 'office-ui-fabric-react/lib/Toggle';
import { Icon, initializeIcons } from "office-ui-fabric-react";

import { ABookWebPartContext } from '../utils/context';
import { IABookWebPartProps } from './IABookWebPartProps';
import { CSVLink } from "react-csv";
import { CommandBarButton } from 'office-ui-fabric-react';  


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

export const ABookFC: React.FunctionComponent<IABookWebPartProps> = (props) => {
  const [employees, setEmployees] = React.useState<any[]>([]);
  const [fullNameQuery, setFullNameQuery] = React.useState('');
  const [jobTitleQuery, setJobTitleQuery] = React.useState('');
  const [departmentQuery, setDepartmentQuery] = React.useState('');
  const [mobileQuery, setMobileQuery] = React.useState('');
  const [managerOfEmployeeQuery, setManagerOfEmployeeQuery] = React.useState('');  
  const [statusQuery, setStatusQuery] = React.useState(['active', 'maternityLeave']);
  const [isHROrAdmin, setIsHROrAdmin] = React.useState(false);

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

  const filtered = JSON.parse(JSON.stringify(filteredByStatus
    .filter((person) => {    
      const conditionOfFiltering = (person.fullName
      && person.fullName.toLowerCase().includes(fullNameQuery.toLowerCase()))
      && (person.jobTitle
      && person.jobTitle.toLowerCase().includes(jobTitleQuery.toLowerCase()))
      && (!person.managerCard || 
      person.managerCard 
        && person.managerCard.Title.toLowerCase().includes(managerOfEmployeeQuery.toLowerCase()))
      && person.employeeCard.Department.toLowerCase().includes(departmentQuery.toLowerCase())
      return conditionOfFiltering;
    })))
    .sort((a,b) => a.employeeCard.Department.localeCompare(b.employeeCard.Department))
    .sort((a,b) => {
      if (a.employeeCard.Department.localeCompare(b.employeeCard.Department) < 0) {
        return 0;
      }
      if (a.employeeCard.Department.localeCompare(b.employeeCard.Department) > 0) {
        return 0;
      }
      if (a.employeeCard.Department.localeCompare(b.employeeCard.Department) === 0) {
        if (a.levelEmployee === b.levelEmployee) {
          return a.fullName.localeCompare(b.fullName);
        }
        return a.levelEmployee - b.levelEmployee;
      }
    });    

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
      if ((event.target as HTMLInputElement).name === "Department") {
        setDepartmentQuery(newValue || '');
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
      setStatusQuery(value);
    }, []);    
  
  const GroupedListDep: React.FunctionComponent = () => {
    const generateGroups = (sortedPersons: any[]) => {
      const groupedPersons: any = groupBy(sortedPersons, (i: any) => 
        i.employeeCard && i.employeeCard.Department && i.employeeCard.Department)
      console.log('groupedPersons', groupedPersons);

      let groupedAndSorted = {};
      for (const group in groupedPersons) {
        groupedAndSorted = {
          [group]: JSON.parse(JSON.stringify(groupedPersons[group]))
                    .sort((a,b) => (a.levelEmployee && b.levelEmployee) 
                      ? a.levelEmployee - b.levelEmployee 
                      : a.Title - b.Title),
          ...groupedAndSorted,
        }
      }

      console.log('groupedAndSorted', groupedAndSorted);
      
      let groups: IGroup[] = [];
      for (const person in groupedPersons) {
        groups.push({
          name: person,
          key: person,
          startIndex: findIndex(sortedPersons, (i: any) => i.employeeCard.Department == person),
          count: groupedPersons[person].length,
          isCollapsed: !sortedPersons.find((i: any) => i.employeeCard.Department == person), 
        });
      }
      console.log('groups', groups);
      
      return groups;
    };

    const onRenderCell = (nestingDepth?: number, item?: IABookProps, itemIndex?: number): JSX.Element => {
      return  (
        <div className={styles.card_container}>
        <div className={styles.card}>
          <div className={styles.container_img}>
            <img className={styles.img} src={'/_layouts/15/userphoto.aspx?size=L&accountname=' + item.employeeCard.EMail} alt="some img" />
          </div>
          <div className={styles.container_contacts}>
            <h2 className={styles.title}>{item.Title}</h2>
            <p className={styles.description}>{item.employeeCard.JobTitle}</p>
            <p className={styles.description}>{'Підрозділ: ' + item.employeeCard.Department}</p>
            <p className={styles.description}>{'Мобільний ' + item.employeeCard.WorkPhone}</p>
            <br />
            <p className={styles.description}>{'Email: ' + item.employeeCard.EMail}</p>
            {item.managerCard ? <p className={styles.manager}>{item.managerCard.Title}</p> : ""}
          </div>
        </div>
      </div>
      );
    };
    console.log('filtered', filtered);
    
    return (
      <ABookWebPartContext.Provider value={props.context}>
      <div>
          <GroupedList
            items={filtered}
            // eslint-disable-next-line react/jsx-no-bind
            groupProps={groupedListProps}
            onRenderCell={onRenderCell}
            groups={generateGroups(filtered)}
          />
      </div>
      </ABookWebPartContext.Provider>
    );
  };

  console.log('rendering');

    
  return (   
    
    <div className={ styles.aBook }>
      <div className={ styles.main_container }>
        <div style={{"width": '100%'}}>
          <GroupedListDep
          />
        </div>
        <div className={ styles.form }>
          <Stack className={styles.searchForm}>
            <select
              multiple
              value={statusQuery}
              onChange={handleChangeStatus}
              name="selectStatus"
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
              label="Full Name"
              name="fullName" 
              value={fullNameQuery}
              onChange={onChangeValue} 
            />        
            <TextField 
              label="Job Title"
              name="jobTitle" 
              value={jobTitleQuery}
              onChange={onChangeValue} 
            />    
            <TextField 
              label="Manager" 
              name="managerOfEmployee"
              value={managerOfEmployeeQuery}
              onChange={onChangeValue}
            /> 
            <TextField 
              label="Department" 
              name="Department"
              value={departmentQuery}
              onChange={onChangeValue}
            />       
          </Stack> 
          <br />
          <div className={ styles.button__container }>
          <CSVLink 
            data={filtered.map((item) => {
                    const { 
                      Department, 
                      EMail,
                      Id,
                      MobilePhone,
                      Office,                      
                      WorkPhone  
                    } = item.employeeCard;

                    const employeeJobTitle = item.employeeCard.JobTitle;
                    const employeeTitle = item.employeeCard.Title;

                    let managerEmail: any;
                    let managerId: any;
                    let managerTitle: any;

                    if (item.managerCard) {
                      managerEmail = item.managerCard.EMail;
                      managerId = item.managerCard.Id;
                      managerTitle = item.managerCard.Title;
                    }
                    managerEmail = null;
                    managerId = null;
                    managerTitle = null;
                    
                  return {                   
                    Department, 
                    EMail,
                    Id,
                    MobilePhone,
                    Office,                      
                    WorkPhone,
                    employeeJobTitle,
                    employeeTitle,
                    managerEmail,
                    managerId,
                    managerTitle,
                    ...item,
                  }
                })
              } 
            filename={'UserInformationReport.csv'} 
            className={ styles.button__container }
          >  
            <CommandBarButton 
              className={ styles.button } 
              iconProps={{ iconName: 'ExcelLogoInverse' }} 
              text='' 
            />  
          </CSVLink>  
          </div>
        </div>
        </div>
    </div>    
  );
};
 