/**
 * @file PropertyFieldSPListPickerHost.tsx
 * Renders the controls for PropertyFieldSPListPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 *
 */
import * as React from 'react';
import { IWebPartContext} from '@microsoft/sp-webpart-base';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { IPropertyFieldSPListPickerPropsInternal, PropertyFieldSPListPickerOrderBy } from './PropertyFieldSPListPicker';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Label } from 'office-ui-fabric-react/lib/Label';

/**
 * @interface
 * PropertyFieldSPListPickerHost properties interface
 *
 */
export interface IPropertyFieldSPListPickerHostProps extends IPropertyFieldSPListPickerPropsInternal {
}

/**
 * @interface
 * PropertyFieldSPListPickerHost state interface
 *
 */
export interface IPropertyFieldFontPickerHostState {
  results: IDropdownOption[];
  selectedKey: string;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldSPListPicker component
 */
export default class PropertyFieldSPListPickerHost extends React.Component<IPropertyFieldSPListPickerHostProps, IPropertyFieldFontPickerHostState> {

  private options: IDropdownOption[] = [];
  private selectedKey: string;

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldSPListPickerHostProps) {
    super(props);

    this.onChanged = this.onChanged.bind(this);
    this.state = {
			results: this.options,
      selectedKey: this.selectedKey,
      errorMessage: ''
    };

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);

    this.loadLists();
  }

  /**
   * @function
   * Loads the list from SharePoint current web site
   */
  private loadLists(): void {
    var listService: SPListPickerService = new SPListPickerService(this.props, this.props.context);
    listService.getLibs().then((response: ISPLists) => {
      response.value.map((list: ISPList) => {
        var isSelected: boolean = false;
        if (this.props.selectedList == list.Id) {
          isSelected = true;
          this.selectedKey = list.Id;
        }
        this.options.push({
          key: list.Id,
          text: list.Title,
          isSelected: isSelected
        });
      });
      this.setState({results: this.options, selectedKey: this.selectedKey});
    });
  }

  /**
   * @function
   * Raises when a list has been selected
   */
  private onChanged(option: IDropdownOption, index?: number): void {
    var newValue: string = option.key as string;
    this.delayedValidate(newValue);
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.selectedList, value);
      return;
    }

    if (this.latestValidateValue === value)
      return;
    this.latestValidateValue = value;

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.selectedList, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.selectedList, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.selectedList, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string, newValue: string) {
    if (this.props.onPropertyChange && newValue != null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
      if (!this.props.disableReactivePropertyChanges && this.props.render != null)
        this.props.render();
    }
  }

  /**
   * @function
   * Called when the component will unmount
   */
  public componentWillUnmount() {
    if (this.async !== undefined)
      this.async.dispose();
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
          disabled={this.props.disabled}
          label=''
          onChanged={this.onChanged}
          options={this.options}
          selectedKey={this.selectedKey}
        />
        { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div style={{paddingBottom: '8px'}}><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
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

/**
 * @class
 * Service implementation to get list & list items from current SharePoint site
 */
class SPListPickerService {

  private context: IWebPartContext;
  private props: IPropertyFieldSPListPickerHostProps;

  /**
   * @function
   * Service constructor
   */
  constructor(_props: IPropertyFieldSPListPickerHostProps, pageContext: IWebPartContext){
      this.props = _props;
      this.context = pageContext;
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
        if (this.props.orderBy == PropertyFieldSPListPickerOrderBy.Id)
          queryUrl += "Id";
        else if (this.props.orderBy == PropertyFieldSPListPickerOrderBy.Title)
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

    /**
     * @function
     * Mock search People method
     */
    public static getLists(restUrl: string, options?: any): Promise<ISPLists> {
      return new Promise<ISPLists>((resolve) => {
            resolve(SPListPickerMockHttpClient._results);
        });
    }

}
