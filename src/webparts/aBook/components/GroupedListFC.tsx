import * as React from 'react';
import { GroupedList, IGroup } from 'office-ui-fabric-react/lib/GroupedList';
import { IColumn, DetailsRow } from 'office-ui-fabric-react/lib/DetailsList';
import { Selection, SelectionMode, SelectionZone } from 'office-ui-fabric-react/lib/Selection';
import { Toggle, IToggleStyles } from 'office-ui-fabric-react/lib/Toggle';
import { useBoolean, useConst } from '@uifabric/react-hooks';
import { createListItems, createGroups, IExampleItem } from '@uifabric/example-data';
import { escape, groupBy, findIndex } from '@microsoft/sp-lodash-subset';

const toggleStyles: Partial<IToggleStyles> = { root: { marginBottom: '20px' } };
const groupCount = 3;
const groupDepth = 3;
const persons = [
    {
      "name": "Den",
      "department": "IT"
    },
    {
      "name": "Greg",
      "department": "IT"
    },
    {
      "name": "Bob",
      "department": "Administration"
    },
     {
      "name": "Fil",
      "department": "Administration"
    },
    {
      "name": "Joe",
      "department": "Marketing"
    },
    {
      "name": "Emma",
      "department": "Marketing"
    }
  ]

const items = createListItems(Math.pow(groupCount, groupDepth + 1));
const columns = Object.keys(items[0])
  .slice(0, 3)
  .map(
    (key: string): IColumn => ({
      key: key,
      name: key,
      fieldName: key,
      minWidth: 300,
    }),
  );

const groups = createGroups(groupCount, groupDepth, 0, groupCount);

export const GroupedListBasicExample: React.FunctionComponent = () => {
  const [isCompactMode, { toggle: toggleIsCompactMode }] = useBoolean(false);
  const selection = useConst(() => {
    const s = new Selection();
    s.setItems(items, true);
    return s;
  });

//   const groupItems: any[] = () => {
//     let collection = {};
//     persons.forEach((elem) => {
//       const query = elem.department;
//       if (!Object.keys(collection).includes(query)) {
//         collection[query] = [{
//             key: query,
//             name: query,
//             startIndex: findIndex(items, (i: any) => i.department === query),
//             count: 
//         }];
//       }
//       collection[query].push(elem);
//       console.log(elem);
//     });
//     return Object.values(collection);
//   };

  const generateGroups = (sortedPersons: any[]) => {
    const groupedPersons: any = groupBy(sortedPersons, (i: any) => i.color);
    console.log('groupedPersons', groupedPersons);
    let groups: IGroup[] = [];
    for (const item in groupedPersons) {
      groups.push({
        name: item,
        key: item,
        startIndex: findIndex(sortedPersons, (i: any) => i.color),
        count: groupedPersons[item].length,
        isCollapsed: findIndex(sortedPersons, (i: any) => i.color) == 0 ? false : true 
      })
    }
    console.log('groups', groups);
    
    return groups;
  }

  const onRenderCell = (nestingDepth?: number, item?: IExampleItem, itemIndex?: number): React.ReactNode => {
    return item && typeof itemIndex === 'number' && itemIndex > -1 ? (
      <DetailsRow
        columns={columns}
        groupNestingDepth={nestingDepth}
        item={item}
        itemIndex={itemIndex}
        selection={selection}
        selectionMode={SelectionMode.multiple}
        compact={isCompactMode}
      />
    ) : null;
  };

  return (
    <div>
      <Toggle
        label="Enable compact mode"
        checked={isCompactMode}
        onChange={toggleIsCompactMode}
        onText="Compact"
        offText="Normal"
        styles={toggleStyles}
      />
      <SelectionZone selection={selection} selectionMode={SelectionMode.multiple}>
        <GroupedList
          items={items}
          // eslint-disable-next-line react/jsx-no-bind
          onRenderCell={onRenderCell}
          selection={selection}
          selectionMode={SelectionMode.multiple}
          groups={generateGroups(items)}
          compact={isCompactMode}
        />
      </SelectionZone>
    </div>
  );
};
