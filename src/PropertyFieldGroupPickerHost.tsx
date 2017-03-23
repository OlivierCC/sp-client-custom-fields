/**
 * @file PropertyFieldGroupPickerHost.tsx
 * Renders the controls for PropertyFieldGroupPicker component
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 *
 */
import * as React from 'react';
import { IPropertyFieldGroupPickerPropsInternal } from './PropertyFieldGroupPicker';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from "@microsoft/sp-http";
import { EnvironmentType, Environment } from '@microsoft/sp-core-library';
import { IPropertyFieldGroup, IGroupType } from './PropertyFieldGroupPicker';
import { NormalPeoplePicker, IBasePickerSuggestionsProps } from 'office-ui-fabric-react/lib/Pickers';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IPersonaProps, PersonaPresence, PersonaInitialsColor } from 'office-ui-fabric-react/lib/Persona';
import { Async } from 'office-ui-fabric-react/lib/Utilities';

import * as strings from 'sp-client-custom-fields/strings';


/**
 * @interface
 * PropertyFieldGroupPickerHost properties interface
 *
 */
export interface IPropertyFieldGroupPickerHostProps extends IPropertyFieldGroupPickerPropsInternal {
}

/**
 * @interface
 * Defines the state of the component
 *
 */
export interface IPeoplePickerState {
  resultsPeople?: Array<IPropertyFieldGroup>;
  resultsPersonas?: Array<IPersonaProps>;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldGroupPicker component
 */
export default class PropertyFieldGroupPickerHost extends React.Component<IPropertyFieldGroupPickerHostProps, IPeoplePickerState> {

  private searchService: PropertyFieldSearchService;
  private intialPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();
  private resultsPeople: Array<IPropertyFieldGroup> = new Array<IPropertyFieldGroup>();
  private resultsPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();
  private selectedPeople: Array<IPropertyFieldGroup> = new Array<IPropertyFieldGroup>();
  private selectedPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();
  private async: Async;
  private delayedValidate: (value: IPropertyFieldGroup[]) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldGroupPickerHostProps) {
    super(props);

    this.searchService = new PropertyFieldSearchService(props.context);
    this.onSearchFieldChanged = this.onSearchFieldChanged.bind(this);
    this.onItemChanged = this.onItemChanged.bind(this);

    this.createInitialPersonas();

    this.state = {
      resultsPeople: this.resultsPeople,
      resultsPersonas: this.resultsPersonas,
      errorMessage: ''
    };

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  /**
   * @function
   * Renders the PeoplePicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    var suggestionProps: IBasePickerSuggestionsProps = {
      suggestionsHeaderText: strings.PeoplePickerSuggestedContacts,
      noResultsFoundText: strings.PeoplePickerNoResults,
      loadingText: strings.PeoplePickerLoading,
    };

    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <NormalPeoplePicker
          pickerSuggestionsProps={suggestionProps}
          onResolveSuggestions={this.onSearchFieldChanged}
          onChange={this.onItemChanged}
          defaultSelectedItems={this.intialPersonas}
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

  /**
   * @function
   * A search field change occured
   */
  private onSearchFieldChanged(searchText: string, currentSelected: IPersonaProps[]): Promise<IPersonaProps> | IPersonaProps[] {
    if (searchText.length > 2) {
      //Clear the suggestions list
      this.setState({ resultsPeople: this.resultsPeople, resultsPersonas: this.resultsPersonas });
      //Request the search service
      var result = this.searchService.searchGroups(searchText, this.props.groupType).then((response: IPropertyFieldGroup[]) => {
        this.resultsPeople = [];
        this.resultsPersonas = [];
        //If allowDuplicate == false, so remove duplicates from results
        if (this.props.allowDuplicate === false)
          response = this.removeDuplicates(response);
        response.map((element: IPropertyFieldGroup, index: number) => {
          //Fill the results Array
          this.resultsPeople.push(element);
          //Transform the response in IPersonaProps object
          this.resultsPersonas.push(this.getPersonaFromGroup(element, index));
        });
        //Refresh the component's state
        this.setState({ resultsPeople: this.resultsPeople, resultsPersonas: this.resultsPersonas });
        return this.resultsPersonas;
      });
      return result;
    }
    else {
      return [];
    }
  }

  /**
   * @function
   * Remove the duplicates if property allowDuplicate equals false
   */
  private removeDuplicates(responsePeople: IPropertyFieldGroup[]): IPropertyFieldGroup[] {
    if (this.selectedPeople == null || this.selectedPeople.length == 0)
      return responsePeople;
    var res: IPropertyFieldGroup[] = [];
    responsePeople.map((element: IPropertyFieldGroup) => {
      var found: boolean = false;
      for (var i: number = 0; i < this.selectedPeople.length; i++) {
        var responseItem: IPropertyFieldGroup = this.selectedPeople[i];
        if (responseItem.id == element.id) {
          found = true;
          break;
        }
      }
      if (found === false)
        res.push(element);
    });
    return res;
  }

  /**
   * @function
   * Creates the collection of initial personas from initial IPropertyFieldGroup collection
   */
  private createInitialPersonas(): void {
    if (this.props.initialData == null || typeof (this.props.initialData) != typeof Array<IPropertyFieldGroup>())
      return;
    this.props.initialData.map((element: IPropertyFieldGroup, index: number) => {
      var persona: IPersonaProps = this.getPersonaFromGroup(element, index);
      this.intialPersonas.push(persona);
      this.selectedPersonas.push(persona);
      this.selectedPeople.push(element);
    });
  }

  /**
   * @function
   * Generates a IPersonaProps object from a IPropertyFieldGroup object
   */
  private getPersonaFromGroup(element: IPropertyFieldGroup, index: number): IPersonaProps {
    return {
      primaryText: element.fullName, secondaryText: element.description
    };
  }


  /**
   * @function
   * Refreshes the web part properties
   */
  private refreshWebPartProperties(): void {
    this.delayedValidate(this.selectedPeople);
  }

   /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: IPropertyFieldGroup[]): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialData, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || []);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialData, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialData, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.initialData, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: IPropertyFieldGroup[], newValue: IPropertyFieldGroup[]) {
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
   * Event raises when the user changed people from hte PeoplePicker component
   */

  private onItemChanged(selectedItems: IPersonaProps[]): void {
    if (selectedItems.length > 0) {
      if (selectedItems.length > this.selectedPersonas.length) {
        var index: number = this.resultsPersonas.indexOf(selectedItems[selectedItems.length - 1]);
        if (index > -1) {
          var people: IPropertyFieldGroup = this.resultsPeople[index];
          this.selectedPeople.push(people);
          this.selectedPersonas.push(this.resultsPersonas[index]);
          this.refreshWebPartProperties();
        }
      } else {
        this.selectedPersonas.map((person, index2) => {
            var selectedItemIndex: number = selectedItems.indexOf(person);
            if (selectedItemIndex === -1) {
              this.selectedPersonas.splice(index2, 1);
              this.selectedPeople.splice(index2, 1);
            }
          });
      }

    } else {
      this.selectedPersonas.splice(0, this.selectedPersonas.length);
      this.selectedPeople.splice(0, this.selectedPeople.length);
    }

    this.refreshWebPartProperties();
  }

}

/**
 * @interface
 * Service interface definition
 */
interface IPropertyFieldSearchService {
  /**
   * @function
   * Search Groups from a query
   */
  searchGroups(query: string, type: IGroupType): Promise<Array<IPropertyFieldGroup>>;
}

/**
 * @class
 * Service implementation to search people in SharePoint
 */
class PropertyFieldSearchService implements IPropertyFieldSearchService {

  private context: IWebPartContext;

  /**
   * @function
   * Service constructor
   */
  constructor(pageContext: IWebPartContext) {
    this.context = pageContext;
  }

  /**
   * @function
   * Search groups from the SharePoint database
   */
  public searchGroups(query: string, type: IGroupType): Promise<Array<IPropertyFieldGroup>> {
    if (Environment.type === EnvironmentType.Local) {
      //If the running environment is local, load the data from the mock
      return this.searchGroupsFromMock(query);
    }
    else {
      //If the running env is SharePoint, loads from the peoplepicker web service
      var contextInfoUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/contextinfo";
      var userRequestUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser";
      var httpPostOptions: ISPHttpClientOptions = {
        headers: {
          "accept": "application/json",
          "content-type": "application/json"
        }
      };
      return this.context.spHttpClient.post(contextInfoUrl, SPHttpClient.configurations.v1, httpPostOptions).then((response: SPHttpClientResponse) => {
        return response.json().then((jsonResponse: any) => {
          var formDigestValue: string = jsonResponse.FormDigestValue;
          var data = {
            'queryParams': {
              //'__metadata': {
              //    'type': 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters'
              //},
              'AllowEmailAddresses': true,
              'AllowMultipleEntities': false,
              'AllUrlZones': false,
              'MaximumEntitySuggestions': 20,
              'PrincipalSource': 15,
              //PrincipalType controls the type of entities that are returned in the results.
              //Choices are All - 15, Distribution List - 2 , Security Groups - 4,
              //SharePoint Groups &ndash; 8, User &ndash; 1. These values can be combined
              'PrincipalType': type === IGroupType.SharePoint ? 8 : 4,
              'QueryString': query
              //'Required':false,
              //'SharePointGroupID':null,
              //'UrlZone':null,
              //'UrlZoneSpecified':false,
            }
          };
          httpPostOptions = {
            headers: {
              'accept': 'application/json',
              'content-type': 'application/json',
              "X-RequestDigest": formDigestValue
            },
            body: JSON.stringify(data)
          };
          return this.context.spHttpClient.post(userRequestUrl, SPHttpClient.configurations.v1, httpPostOptions).then((searchResponse: SPHttpClientResponse) => {
            return searchResponse.json().then((usersResponse: any) => {
              var res: IPropertyFieldGroup[] = [];
              var values: any = JSON.parse(usersResponse.value);
              values.map(element => {
                var persona: IPropertyFieldGroup = {
                  fullName: element.DisplayText,
                  login: type === IGroupType.SharePoint ? element.EntityData.AccountName : element.ProviderName,
                  id : type === IGroupType.SharePoint ? element.EntityData.SPGroupID : element.Key,
                  description: element.Description
                };
                res.push(persona);
              });
              return res;
            });
          });
        });
      });
    }
  }


  /**
   * @function
   * Returns fake people results for the Mock mode
   */
  private searchGroupsFromMock(query: string): Promise<Array<IPropertyFieldGroup>> {
    return PeoplePickerMockHttpClient.searchGroups(this.context.pageContext.web.absoluteUrl).then(() => {
      const results: IPropertyFieldGroup[] = [
        { id: '1', fullName: "Members", login: "Members", description: 'Members' },
        { id: '2', fullName: "Viewers", login: "Viewers", description: 'Viewers' },
        { id: '3', fullName: "Excel Services Viewers", login: "Excel Services Viewers", description: 'Excel Services Viewers' }
      ];
      return results;
    }) as Promise<Array<IPropertyFieldGroup>>;
  }
}

/**
 * @class
 * Defines a http client to request mock data to use the web part with the local workbench
 */
class PeoplePickerMockHttpClient {

  /**
   * @var
   * Mock SharePoint result sample
   */
  private static _results: IPropertyFieldGroup[] = [];

  /**
   * @function
   * Mock search People method
   */
  public static searchGroups(restUrl: string, options?: any): Promise<IPropertyFieldGroup[]> {
    return new Promise<IPropertyFieldGroup[]>((resolve) => {
      resolve(PeoplePickerMockHttpClient._results);
    });
  }

}