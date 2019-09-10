import * as React from 'react';
import { IInputs } from './generated/ManifestTypes';
import { IGroup, DetailsList, IColumn, Text } from 'office-ui-fabric-react';

type Dataset = ComponentFramework.PropertyTypes.DataSet;
type EntityRecord = ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
type Column = ComponentFramework.PropertyHelper.DataSetApi.Column;

export interface GroupListProps {
    context: ComponentFramework.Context<IInputs>;
}

export interface  GroupListState {
    dataset: Dataset;
    groupingColumns: Column[]
}

export class GroupList extends React.Component<GroupListProps, GroupListState> {
    constructor(props: GroupListProps){
        super(props);
        this.state = {
            dataset: props.context.parameters.dataset,
            groupingColumns: []
        };
    }

    render(){
        const columns = this.createColumns(this.state.dataset);
        const {groups, records} = this.groupDataset(this.state.dataset, this.state.groupingColumns);
        return (
            <DetailsList
                columns = {columns}
                items = {records}
                groups = {groups}
                onRenderItemColumn = {this.onRenderItemColumn.bind(this)}
            ></DetailsList>
        )
    }

    private onRenderItemColumn(item?: EntityRecord, index?: number | undefined, column?: IColumn | undefined): React.ReactNode {
        if(item && column){
            return (
                <Text>{item.getFormattedValue(column.key)}</Text>
            )
        }
    }
    
    private groupDataset(dataset: Dataset, columns: Column[]): {
        groups: IGroup[] | undefined;
        records: EntityRecord[];
    }{       
        const records = dataset.sortedRecordIds.map(id => dataset.records[id]);
        let groupRecordsMap: {[groupKey: string]: {name: string; records: EntityRecord[]}} = {};
        
        if(columns.length === 0){
            return {
                groups: undefined,
                records: records
            }
        }

        records.forEach(record => {
            const {key, name} = this.createGroupKeyAndName(record, columns);
            if(groupRecordsMap[key]){
                groupRecordsMap[key].records.push(record);
            }
            else{
                groupRecordsMap[key] = {
                    name: name,
                    records: [record]
                }
            }
        });

        let groupedRecords: EntityRecord[] = [];
        let groups: IGroup[] = [];
        for(let key in groupRecordsMap){
            const group: IGroup = {
                key: key,
                name: groupRecordsMap[key].name,
                startIndex: groupedRecords.length,
                count: groupRecordsMap[key].records.length
            };
            groups.push(group);
            groupRecordsMap[key].records.forEach(record => {
                groupedRecords.push(record);
            });
        }

        return {
            groups: groups,
            records: groupedRecords
        };
    }

    private createGroupKeyAndName(record: EntityRecord, columns: Column[]): {
        key: string;
        name: string;
    }{
        let key = "_";
        let name = "";
        columns.forEach(column => {
            key += record.getValue(column.name).toString();
            name += record.getFormattedValue(column.name) + " ";
        });

        return {
            key: key,
            name: name
        };
    }

    private createColumns(dataset: Dataset): IColumn[] {
        const columns = dataset.columns.map(field => {
            const column: IColumn = {
                name: field.displayName,
                ariaLabel: field.displayName,
                key: field.name,
                fieldName: field.name,
                minWidth: field.visualSizeFactor,
                maxWidth: field.visualSizeFactor,
                isResizable: true,
                data: field,
                onColumnClick: this.onColumnClick.bind(this),
                isGrouped: this.state.groupingColumns.some(column => column.name == field.name)
            };
            return column;
        });

        return columns;
    }

    private onColumnClick(ev: React.MouseEvent<HTMLElement, MouseEvent>, column: IColumn){
        let groupingColumns = this.state.groupingColumns;
        const field = column.data as Column;
        if(groupingColumns.some(column => column.name == field.name)){
            groupingColumns = groupingColumns.filter(column => column.name != field.name);
        }
        else {
            groupingColumns.push(field);
        }
        this.setState({
            groupingColumns: groupingColumns
        });
    }

}
