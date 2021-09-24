import * as React from 'react';
import { GroupedList, IGroup, IGroupHeaderProps } from 'office-ui-fabric-react/lib/GroupedList';
import { IColumn, DetailsRow } from 'office-ui-fabric-react/lib/DetailsList';
import { Selection, SelectionMode, SelectionZone } from 'office-ui-fabric-react/lib/Selection';
import { Toggle, IToggleStyles } from 'office-ui-fabric-react/lib/Toggle';
import { useBoolean, useConst } from '@uifabric/react-hooks';
import { createListItems, createGroups, IExampleItem } from '@uifabric/example-data';
import { escape, groupBy, findIndex } from '@microsoft/sp-lodash-subset';
import styles from './ABook.module.scss';
import { JsxElement } from 'typescript';
import { Icon, initializeIcons } from "office-ui-fabric-react";

const toggleStyles: Partial<IToggleStyles> = { root: { marginBottom: '20px' } };

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

  /* 
  
    //  {
    //   "name": "Fil",
    //   "department": "Administration"
    // },
    // {
    //   "name": "Joe",
    //   "department": "Marketing"
    // },
    // {
    //   "name": "Emma",
    //   "department": "Marketing"
    // },
    // {
    //   "name": "Clair",
    //   "department": "Marketing"
    // },
    // {
    //   "name": "Irish",
    //   "department": "Marketing"
    // }

  */
  
export interface IPersons {
  name: string;
  employee: {
    department: string;
  };  
}

export const GroupedListBasicExample: React.FunctionComponent = () => {
  const generateGroups = (sortedPersons: any[]) => {
    const groupedPersons: any = groupBy(sortedPersons, (i: any) => i.employee && i.employee.department && i.employee.department);
    console.log('groupedPersons', groupedPersons);
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
    <div>
        <GroupedList
          items={persons}
          // eslint-disable-next-line react/jsx-no-bind
          groupProps={groupedListProps}
          onRenderCell={onRenderCell}
          groups={generateGroups(persons)}
        />
    </div>
  );
};
