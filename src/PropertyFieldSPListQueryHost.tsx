/**
 * @file PropertyFieldSPListQueryHost.tsx
 * Renders the controls for PropertyFieldSPListQuery component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 *
 */
import * as React from 'react';
import { IWebPartContext} from '@microsoft/sp-webpart-base';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Slider } from 'office-ui-fabric-react/lib/Slider';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPropertyFieldSPListQueryPropsInternal, PropertyFieldSPListQueryOrderBy } from './PropertyFieldSPListQuery';

import * as strings from 'sp-client-custom-fields/strings';

/**
 * @interface
 * PropertyFieldSPListQueryHost properties interface
 *
 */
export interface IPropertyFieldSPListQueryHostProps extends IPropertyFieldSPListQueryPropsInternal {
}

export interface IFilter {
  field?: string;
  operator?: string;
  value?: string;
}


export interface IPropertyFieldSPListQueryHostState {
  lists: IDropdownOption[];
  fields: IDropdownOption[];
  arranged: IDropdownOption[];
  selectedList?: string;
  selectedField?: string;
  selectedArrange?: string;
  max?: number;
  operators?: IDropdownOption[];
  filters?: IFilter[];
}

/**
 * @class
 * Renders the controls for PropertyFieldSPListQuery component
 */
export default class PropertyFieldSPListQueryHost extends React.Component<IPropertyFieldSPListQueryHostProps, IPropertyFieldSPListQueryHostState> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldSPListQueryHostProps) {
    super(props);
    this.onChangedList = this.onChangedList.bind(this);
    this.onChangedField = this.onChangedField.bind(this);
    this.onChangedArranged = this.onChangedArranged.bind(this);
    this.onChangedMax = this.onChangedMax.bind(this);
    this.loadFields = this.loadFields.bind(this);
    this.onClickAddFilter = this.onClickAddFilter.bind(this);
    this.onClickRemoveFilter = this.onClickRemoveFilter.bind(this);
    this.onChangedFilterField = this.onChangedFilterField.bind(this);
    this.onChangedFilterOperator = this.onChangedFilterOperator.bind(this);
    this.onChangedFilterValue = this.onChangedFilterValue.bind(this);

    this.state = {
			lists: [],
      fields: [],
      arranged: [{key: 'asc', text: 'Asc'}, {key: 'desc', text: 'Desc'}],
      selectedList: '',
      selectedField: '',
      selectedArrange: '',
      operators: [
        {key: 'Eq', text: strings.SPListQueryOperatorEq},
         {key: 'Ne', text: strings.SPListQueryOperatorNe},
          {key: 'startsWith', text: strings.SPListQueryOperatorStartsWith},
           {key: 'substringof', text: strings.SPListQueryOperatorSubstringof},
            {key: 'Lt', text: strings.SPListQueryOperatorLt},
             {key: 'Le', text: strings.SPListQueryOperatorLe},
              {key: 'Gt', text: strings.SPListQueryOperatorGt},
               {key: 'Ge', text: strings.SPListQueryOperatorGe}
      ],
      filters: [],
      max: 20
    };

    this.loadDefaultData();
    this.loadLists();
  }

  private loadDefaultData(): void {
    if (this.props.query == null || this.props.query == '')
      return;
    var indexOfGuid: number = this.props.query.indexOf("lists(guid'");
    if (indexOfGuid > -1) {
      var listId: string = this.props.query.substr(indexOfGuid);
      listId = listId.replace("lists(guid'", "");
      var indexOfEndGuid: number = listId.indexOf("')/items");
      listId = listId.substr(0, indexOfEndGuid);
      this.state.selectedList = listId;
      if (listId != null && listId != '')
        this.loadFields();
    }
    var indexOfOrderBy: number = this.props.query.indexOf("$orderBy=");
    if (indexOfOrderBy > -1) {
      var orderBy: string = this.props.query.substr(indexOfOrderBy);
      orderBy = orderBy.replace("$orderBy=", "");
      var indexOfEndOrderBy: number = orderBy.indexOf("%20");
      var field: string = orderBy.substr(0, indexOfEndOrderBy);
      this.state.selectedField = field;
      var arranged: string = orderBy.substr(indexOfEndOrderBy);
      arranged = arranged.replace("%20", "");
      var indexOfEndArranged: number = arranged.indexOf("&");
      arranged = arranged.substr(0, indexOfEndArranged);
      this.state.selectedArrange = arranged;
    }
    var indexOfTop: number = this.props.query.indexOf("$top=");
    if (indexOfTop > -1) {
      var top: string = this.props.query.substr(indexOfTop);
      top = top.replace("$top=", "");
      var indexOfEndTop: number = top.indexOf("&");
      top = top.substr(0, indexOfEndTop);
      this.state.max = Number(top);
    }
    var indexOfFilters: number = this.props.query.indexOf("$filters=");
    if (indexOfFilters > -1) {
      var filter: string = this.props.query.substr(indexOfFilters);
      filter = filter.replace("$filters=", "");
      var indexOfEndFilter: number = filter.indexOf("&");
      filter = filter.substr(0, indexOfEndFilter);
      if (filter != null && filter != '') {
        var subFilter = filter.split("%20AND%20");
        for (var i = 0; i < subFilter.length; i++) {
          var fieldId: string = subFilter[i].substr(0, subFilter[i].indexOf("%20"));
          var operator: string = subFilter[i].substr(subFilter[i].indexOf("%20"));
          operator = operator.substr(3);
          operator = operator.substr(0, operator.indexOf("%20"));
          var value: string = subFilter[i].substr(subFilter[i].indexOf(operator + "%20"));
          value = value.replace(operator + "%20", "");
          value = value.replace("'", "").replace("'", "").replace("'", "");
          if (value == "undefined")
            value = '';
          var newObj: IFilter = {};
          newObj.field = fieldId;
          newObj.operator = operator;
          newObj.value = value;
          this.state.filters.push(newObj);
        }
      }
    }
  }

  /**
   * @function
   * Loads the list from SharePoint current web site
   */
  private loadLists(): void {
    var listService: SPListPickerService = new SPListPickerService(this.props, this.props.context);
    listService.getLibs().then((response: ISPLists) => {
      this.state.lists = [];
      response.value.map((list: ISPList) => {
        var isSelected: boolean = false;
        if (this.props.selectedList == list.Id) {
          isSelected = true;
          this.state.selectedList = list.Id;
        }
        this.state.lists.push({
          key: list.Id,
          text: list.Title,
          isSelected: isSelected
        });
      });
      this.saveState();
    });
  }

  private loadFields(): void {
    var listService: SPListPickerService = new SPListPickerService(this.props, this.props.context);
    listService.getFields(this.state.selectedList).then((response: ISPFields) => {
      this.state.fields = [];
      response.value.map((list: ISPField) => {
        var isSelected: boolean = false;
        this.state.fields.push({
          key: list.StaticName,
          text: list.Title,
          isSelected: isSelected
        });
      });
      this.saveState();
    });
  }

  private saveState(): void {
      this.setState({
        selectedField: this.state.selectedField,
        selectedArrange: this.state.selectedArrange,
        selectedList: this.state.selectedList,
        lists: this.state.lists,
        fields: this.state.fields,
        max: this.state.max,
        arranged: this.state.arranged,
        operators: this.state.operators,
        filters: this.state.filters
      });
  }

  private saveQuery(): void {
    if (this.props.onPropertyChange) {

      var queryUrl: string = this.props.context.pageContext.web.absoluteUrl;
      queryUrl += "/_api/lists(guid'";
      queryUrl += this.state.selectedList;
      queryUrl += "')/items?";
      if (this.state.selectedField != null && this.state.selectedField != '') {
        queryUrl += "$orderBy=";
        queryUrl += this.state.selectedField;
        queryUrl += "%20";
        queryUrl += this.state.selectedArrange;
        queryUrl += '&';
      }
      if (this.state.max != null) {
        queryUrl += '$top=';
        queryUrl += this.state.max;
        queryUrl += '&';
      }
      if (this.state.filters != null && this.state.filters.length > 0) {
        queryUrl += '$filters=';
        for (var i = 0; i < this.state.filters.length; i++) {
          if (this.state.filters[i].field != null && this.state.filters[i].operator != null) {
            if (i > 0) {
              queryUrl += "%20AND%20";
            }
            queryUrl += this.state.filters[i].field;
            queryUrl += "%20";
            queryUrl += this.state.filters[i].operator;
            queryUrl += "%20'";
            queryUrl += this.state.filters[i].value;
            queryUrl += "'";
          }
        }
        queryUrl += '&';
      }
      this.props.properties[this.props.targetProperty] = queryUrl;
      this.props.onPropertyChange(this.props.targetProperty, this.props.query, queryUrl);
    }
  }

  /**
   * @function
   * Raises when a list has been selected
   */
  private onChangedList(option: IDropdownOption, index?: number): void {
    this.state.selectedList =  option.key as string;
    this.saveQuery();
    this.saveState();
    this.loadFields();
  }

   private onChangedField(option: IDropdownOption, index?: number): void {
    this.state.selectedField =  option.key as string;
    this.saveQuery();
    this.saveState();
  }

   private onChangedArranged(option: IDropdownOption, index?: number): void {
    this.state.selectedArrange =  option.key as string;
    this.saveQuery();
    this.saveState();
  }

  private onChangedMax(newValue?: number): void {
    this.state.max = newValue;
    this.saveQuery();
    this.saveState();
  }

  private onClickAddFilter(elm?: any): void {
    this.state.filters.push({});
    this.saveState();
    this.saveQuery();
  }

  private onClickRemoveFilter(index: number): void {
    if (index > -1) {
      this.state.filters.splice(index, 1);
      this.saveState();
      this.saveQuery();
    }
  }

  private onChangedFilterField(option: IDropdownOption, index?: number, selectedIndex?: number): void {
    this.state.filters[selectedIndex].field = option.key as string;
    this.saveState();
    this.saveQuery();
  }

  private onChangedFilterOperator(option: IDropdownOption, index?: number, selectedIndex?: number): void {
    this.state.filters[selectedIndex].operator = option.key as string;
    this.saveState();
    this.saveQuery();
  }

  private onChangedFilterValue(value?: string, index?: number): void {
    this.state.filters[index].value = value;
    this.saveState();
    this.saveQuery();
  }


  /**
   * @function
   * Renders the SPListpicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {
    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <Dropdown
          label={strings.SPListQueryList}
          onChanged={this.onChangedList}
          options={this.state.lists}
          selectedKey={this.state.selectedList}
          disabled={this.props.disabled}
        />

        {this.props.showOrderBy != false ?
          <div>
            <Dropdown
              label={strings.SPListQueryOrderBy}
              options={this.state.fields}
              selectedKey={this.state.selectedField}
              onChanged={this.onChangedField}
              isDisabled={this.props.disabled === false && this.state.selectedList != null && this.state.selectedList != '' ? false : true }
            />
            <Dropdown
              label={strings.SPListQueryArranged}
              options={this.state.arranged}
              selectedKey={this.state.selectedArrange}
              onChanged={this.onChangedArranged}
              isDisabled={this.props.disabled === false && this.state.selectedList != null && this.state.selectedList != '' ? false : true }
            />
           </div>
          : ''}

        {this.props.showMax != false ?
          <Slider label={strings.SPListQueryMax}
            min={0}
            max={this.props.max == null ? 500 : this.props.max}
            defaultValue={this.state.max}
            onChange={this.onChangedMax}
            disabled={this.props.disabled === false && this.state.selectedList != null && this.state.selectedList != '' ? false : true }
          />
          : ''}

        {this.state.filters.map((value: IFilter, index: number) => {
          return (
            <div>
              <Label>Filter</Label>
              <Dropdown
                label=''
                disabled={this.props.disabled}
                options={this.state.fields}
                selectedKey={value.field}
                onChanged={(option: IDropdownOption, selectIndex?: number) => this.onChangedFilterField(option, selectIndex, index)}
              />
              <Dropdown
                label=''
                disabled={this.props.disabled}
                options={this.state.operators}
                selectedKey={value.operator}
                onChanged={(option: IDropdownOption, selectIndex?: number) => this.onChangedFilterOperator(option, selectIndex, index)}
              />
              <TextField disabled={this.props.disabled} defaultValue={value.value} onChanged={(value2: string) => this.onChangedFilterValue(value2, index)} />
              <Button disabled={this.props.disabled} buttonType={ButtonType.command} onClick={() => this.onClickRemoveFilter(index)} icon="Delete">
                {strings.SPListQueryRemove}
              </Button>
            </div>
          );
        })
        }

        {this.props.showFilters != false ?
          <Button buttonType={ButtonType.command} onClick={this.onClickAddFilter}
          disabled={this.props.disabled === false && this.state.selectedList != null && this.state.selectedList != '' ? false : true } icon="Add">
          {strings.SPListQueryAdd}
          </Button>
          : ''}

      </div>
    );
  }
}

/**
 * @interface
 * Defines a collection of SharePoint lists
 */
interface ISPLists {
  value: ISPList[];
}

/**
 * @interface
 * Defines a SharePoint list
 */
interface ISPList {
  Title: string;
  Id: string;
  BaseTemplate: string;
}

interface ISPField {
  Title: string;
  StaticName: string;
}

interface ISPFields {
  value: ISPField[];
}

/**
 * @class
 * Service implementation to get list & list items from current SharePoint site
 */
class SPListPickerService {

  private context: IWebPartContext;
  private props: IPropertyFieldSPListQueryHostProps;

  /**
   * @function
   * Service constructor
   */
  constructor(_props: IPropertyFieldSPListQueryHostProps, pageContext: IWebPartContext){
      this.props = _props;
      this.context = pageContext;
  }

  public getFields(listId: string): Promise<ISPFields> {
    if (Environment.type === EnvironmentType.Local) {
      //If the running environment is local, load the data from the mock
      return this.getFieldsFromMock();
    }
    else {
      var queryUrl: string = this.context.pageContext.web.absoluteUrl;
      queryUrl += "/_api/lists(guid'";
      queryUrl += listId;
      queryUrl += "')/Fields?$select=Title,StaticName&$orderBy=Title&$filter=Hidden%20eq%20false";
      return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
          return response.json();
      });
    }
  }

  /**
   * @function
   * Gets the collection of libs in the current SharePoint site
   */
  public getLibs(): Promise<ISPLists> {
    if (Environment.type === EnvironmentType.Local) {
      //If the running environment is local, load the data from the mock
      return this.getLibsFromMock();
    }
    else {
      //If the running environment is SharePoint, request the lists REST service
      var queryUrl: string = this.context.pageContext.web.absoluteUrl;
      queryUrl += "/_api/lists?$select=Title,id,BaseTemplate";
      if (this.props.orderBy != null) {
        queryUrl += "&$orderby=";
        if (this.props.orderBy == PropertyFieldSPListQueryOrderBy.Id)
          queryUrl += "Id";
        else if (this.props.orderBy == PropertyFieldSPListQueryOrderBy.Title)
          queryUrl += "Title";
      }
      if (this.props.baseTemplate != null && this.props.baseTemplate) {
        queryUrl += "&$filter=BaseTemplate%20eq%20";
        queryUrl += this.props.baseTemplate;
        if (this.props.includeHidden === false) {
          queryUrl += "%20and%20Hidden%20eq%20false";
        }
      }
      else {
        if (this.props.includeHidden === false) {
          queryUrl += "&$filter=Hidden%20eq%20false";
        }
      }
      return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
          return response.json();
      });
    }
  }

  /**
   * @function
   * Returns 3 fake SharePoint lists for the Mock mode
   */
  private getLibsFromMock(): Promise<ISPLists> {
    return SPListPickerMockHttpClient.getLists(this.context.pageContext.web.absoluteUrl).then(() => {
          const listData: ISPLists = {
              value:
              [
                  { Title: 'Mock List One', Id: '6770c83b-29e8-494b-87b6-468a2066bcc6', BaseTemplate: '109' },
                  { Title: 'Mock List Two', Id: '2ece98f2-cc5e-48ff-8145-badf5009754c', BaseTemplate: '109' },
                  { Title: 'Mock List Three', Id: 'bd5dbd33-0e8d-4e12-b289-b276e5ef79c2', BaseTemplate: '109' }
              ]
          };
          return listData;
      }) as Promise<ISPLists>;
  }

   private getFieldsFromMock(): Promise<ISPFields> {
    return SPListPickerMockHttpClient.getFields(this.context.pageContext.web.absoluteUrl).then(() => {
          const listData: ISPFields = {
              value:
              [
                  { Title: 'ID', StaticName: 'ID'},
                  { Title: 'Title', StaticName: 'Title'},
                  { Title: 'Created', StaticName: 'Created'},
                  { Title: 'Modified', StaticName: 'Modified'}
              ]
          };
          return listData;
      }) as Promise<ISPFields>;
  }

}


/**
 * @class
 * Defines a http client to request mock data to use the web part with the local workbench
 */
class SPListPickerMockHttpClient {

    /**
     * @var
     * Mock SharePoint result sample
     */
    private static _results: ISPLists = { value: []};
    private static _resultsF: ISPFields = { value: []};

    /**
     * @function
     * Mock search People method
     */
    public static getLists(restUrl: string, options?: any): Promise<ISPLists> {
      return new Promise<ISPLists>((resolve) => {
            resolve(SPListPickerMockHttpClient._results);
        });
    }

    public static getFields(restUrl: string, options?: any): Promise<ISPFields> {
      return new Promise<ISPFields>((resolve) => {
            resolve(SPListPickerMockHttpClient._resultsF);
        });
    }

}
