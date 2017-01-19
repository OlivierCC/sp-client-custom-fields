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
import { SPHttpClientConfigurations } from "@microsoft/sp-http";
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { IPropertyFieldSPListMultiplePickerPropsInternal, PropertyFieldSPListMultiplePickerOrderBy } from './PropertyFieldSPListMultiplePicker';

/**
 * @interface
 * PropertyFieldSPListMultiplePickerHost properties interface
 *
 */
export interface IPropertyFieldSPListMultiplePickerHostProps extends IPropertyFieldSPListMultiplePickerPropsInternal {
}

/**
 * @class
 * Renders the controls for PropertyFieldSPListMultiplePicker component
 */
export default class PropertyFieldSPListMultiplePickerHost extends React.Component<IPropertyFieldSPListMultiplePickerHostProps, {}> {

  private options: IChoiceGroupOption[] = [];
  private selectedKeys: string[] = [];
  private loaded: boolean = false;

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldSPListMultiplePickerHostProps) {
    super(props);

    this.onChanged = this.onChanged.bind(this);
    this.state = {
			results: this.options,
      selectedKeys: this.selectedKeys,
      loaded: this.loaded
    };
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
    if (this.props.onPropertyChange && element) {
      var isChecked: boolean = element.currentTarget.checked;
      var value: string = element.currentTarget.value;

      if (isChecked === false) {
        this.removeSelected(value);
      }
      else {
        this.selectedKeys.push(value);
      }
      this.props.properties[this.props.targetProperty] = this.selectedKeys;
      this.props.onPropertyChange(this.props.targetProperty, this.props.selectedLists, this.selectedKeys);
    }
  }

  /**
   * @function
   * Renders the SPListpicker controls with Office UI  Fabric
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
        //Renders content
        return (
          <div>
            <Label>{this.props.label}</Label>
            {this.options.map((item: IChoiceGroupOption, index: number) => {
              var uniqueKey = this.props.targetProperty + '-' + item.key;
              return (
                <div className="ms-ChoiceField">
                  <input id={uniqueKey} style={{width: '18px', height: '18px'}} value={item.key} name={uniqueKey} onClick={this.onChanged} defaultChecked={item.isChecked} aria-checked={item.isChecked} type="checkbox" role="checkbox" />
                  <label htmlFor={uniqueKey}><span className="ms-Label">{item.text}</span></label>
                </div>
              );
            })
            }
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
      return this.context.spHttpClient.get(queryUrl, SPHttpClientConfigurations.v1).then((response: Response) => {
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
