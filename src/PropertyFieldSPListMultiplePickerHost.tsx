/**
 * @file PropertyFieldSPListMultiplePickerHost.tsx
 * Renders the controls for PropertyFieldSPListMultiplePicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 *
 */
import * as React from 'react';
import { IWebPartContext} from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import GuidHelper from './GuidHelper';
import { IPropertyFieldSPListMultiplePickerPropsInternal, PropertyFieldSPListMultiplePickerOrderBy } from './PropertyFieldSPListMultiplePicker';

/**
 * @interface
 * PropertyFieldSPListMultiplePickerHost properties interface
 *
 */
export interface IPropertyFieldSPListMultiplePickerHostProps extends IPropertyFieldSPListMultiplePickerPropsInternal {
}

/**
 * @interface
 * PropertyFieldSPListMultiplePickerHost state interface
 *
 */
export interface IPropertyFieldSPListMultiplePickerHostState {
  results: IChoiceGroupOption[];
  selectedKeys: string[];
  loaded: boolean;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldSPListMultiplePicker component
 */
export default class PropertyFieldSPListMultiplePickerHost extends React.Component<IPropertyFieldSPListMultiplePickerHostProps, IPropertyFieldSPListMultiplePickerHostState> {

  private options: IChoiceGroupOption[] = [];
  private selectedKeys: string[] = [];
  private loaded: boolean = false;
  private async: Async;
  private delayedValidate: (value: string[]) => void;
  private _key: string;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldSPListMultiplePickerHostProps) {
    super(props);

    this._key = GuidHelper.getGuid();
    this.onChanged = this.onChanged.bind(this);
    this.state = {
			results: this.options,
      selectedKeys: this.selectedKeys,
      loaded: this.loaded,
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
    //Builds the SharePoint List service
    var listService: SPListPickerService = new SPListPickerService(this.props, this.props.context);
    //Gets the libs
    listService.getLibs().then((response: ISPLists) => {
      response.value.map((list: ISPList) => {
        var isSelected: boolean = false;
        var indexInExisting: number = -1;
        //Defines if the current list must be selected by default
        if ( this.props.selectedLists)
          indexInExisting = this.props.selectedLists.indexOf(list.Id);
        if (indexInExisting > -1) {
          isSelected = true;
          this.selectedKeys.push(list.Id);
        }
        //Add the option to the list
        this.options.push({
          key: list.Id,
          text: list.Title,
          isChecked: isSelected
        });
      });
      this.loaded = true;
      this.setState({results: this.options, selectedKeys: this.selectedKeys, loaded: true});
    });
  }

  /**
   * @function
   * Remove a string from the selected keys
   */
  private removeSelected(element: string): void {
    var res = [];
    for (var i = 0; i < this.selectedKeys.length; i++) {
      if (this.selectedKeys[i] !== element)
        res.push(this.selectedKeys[i]);
    }
    this.selectedKeys = res;
  }

  /**
   * @function
   * Raises when a list has been selected
   */
  private onChanged(element: any): void {
    if (element) {
      var isChecked: boolean = element.currentTarget.checked;
      var value: string = element.currentTarget.value;

      if (isChecked === false) {
        this.removeSelected(value);
      }
      else {
        this.selectedKeys.push(value);
      }
      this.delayedValidate(this.selectedKeys);
    }
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string[]): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.selectedLists, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || []);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.selectedLists, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.selectedLists, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.selectedLists, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: string[], newValue: string[]) {
    if (this.props.onPropertyChange && newValue != null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
    }
  }

  /**
   * @function
   * Called when the component will unmount
   */
  public componentWillUnmount() {
    this.async.dispose();
  }

  /**
   * @function
   * Renders the SPListMultiplePicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    if (this.loaded === false) {
       return (
        <div>
          <Label>{this.props.label}</Label>
          <Spinner type={ SpinnerType.normal } />
        </div>
    );
    }
    else
    {
        var styleOfLabel: any = {
          color: this.props.disabled === true ? '#A6A6A6' : 'auto'
        };
        //Renders content
        return (
          <div>
            <Label>{this.props.label}</Label>
            {this.options.map((item: IChoiceGroupOption, index: number) => {
              var uniqueKey = this.props.targetProperty + '-' + item.key;
              return (
                <div className="ms-ChoiceField" key={this._key + '-multiplelistpicker-' + index}>
                  <input disabled={this.props.disabled} id={uniqueKey} style={{width: '18px', height: '18px'}} value={item.key} name={uniqueKey} onClick={this.onChanged} defaultChecked={item.isChecked} aria-checked={item.isChecked} type="checkbox" role="checkbox" />
                  <label htmlFor={uniqueKey}><span className="ms-Label" style={styleOfLabel}>{item.text}</span></label>
                </div>
              );
            })
            }
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
  private props: IPropertyFieldSPListMultiplePickerHostProps;

  /**
   * @function
   * Service constructor
   */
  constructor(_props: IPropertyFieldSPListMultiplePickerHostProps, pageContext: IWebPartContext){
      this.props = _props;
      this.context = pageContext;
  }

  /**
   * @function
   * Gets the collection of SP libs in the current SharePoint site
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
        if (this.props.orderBy == PropertyFieldSPListMultiplePickerOrderBy.Id)
          queryUrl += "Id";
        else if (this.props.orderBy == PropertyFieldSPListMultiplePickerOrderBy.Title)
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
