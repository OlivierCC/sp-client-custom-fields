/**
 * @file PropertyFieldPeoplePickerHost.tsx
 * Renders the controls for PropertyFieldPeoplePicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 *
 */
import * as React from 'react';
import { IPropertyFieldPeoplePickerPropsInternal } from './PropertyFieldPeoplePicker';
import { PeoplePicker, IPeoplePickerProps, PeoplePickerType } from 'office-ui-fabric-react/lib/PeoplePicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { IPersonaProps, PersonaPresence, PersonaInitialsColor} from 'office-ui-fabric-react/lib/Persona';
import { IPropertyFieldPeople } from './PropertyFieldPeoplePicker';
import { IWebPartContext } from '@microsoft/sp-client-preview';
import { EnvironmentType, IHttpClientOptions } from '@microsoft/sp-client-base';

import * as strings from 'propertyFieldsStrings';


/**
 * @interface
 * PropertyFieldPeoplePickerHost properties interface
 *
 */
export interface IPropertyFieldPeoplePickerHostProps extends IPropertyFieldPeoplePickerPropsInternal {
}

/**
 * @class
 * Defines the labels of the DatePicker control (as months, days, etc.)
 *
 */
class PeoplePickerProperties implements IPeoplePickerProps {
    /**
     * @var
     * Kind of peoplepicker component
     */
    public type: PeoplePickerType = PeoplePickerType.normal;
    /**
     * @var
     * Suggested contacts text
     */
    public searchCategoryName: string = strings.PeoplePickerSuggestedContacts;
    /**
     * @var
     * No results text
     */
    public noResultsText: string = strings.PeoplePickerNoResults;
    /**
     * @var
     * Loading contacts text
     */
    public loadingContacts: string = strings.PeoplePickerLoading;
}

/**
 * @interface
 * Defines the state of the component
 *
 */
interface IPeoplePickerState {
	resultsPeople?: Array<IPropertyFieldPeople>;
  resultsPersonas?: Array<IPersonaProps>;
  loading?: boolean;
}

/**
 * @class
 * Renders the controls for PropertyFieldPeoplePicker component
 */
export default class PropertyFieldPeoplePickerHost extends React.Component<IPropertyFieldPeoplePickerHostProps, {}> {

  private peoplePickerProperties: PeoplePickerProperties;
  private searchService: PropertyFieldSearchService;
  private intialPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();
  private resultsPeople: Array<IPropertyFieldPeople> = new Array<IPropertyFieldPeople>();
  private resultsPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();
  private selectedPeople: Array<IPropertyFieldPeople> = new Array<IPropertyFieldPeople>();
  private selectedPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldPeoplePickerHostProps) {
    super(props);

    this.peoplePickerProperties = new PeoplePickerProperties();
    this.searchService = new PropertyFieldSearchService(props.context);
    this.onSearchFieldChanged = this.onSearchFieldChanged.bind(this);
    this.onItemAdded = this.onItemAdded.bind(this);
    this.onItemRemoved = this.onItemRemoved.bind(this);
    this.createInitialPersonas();

    this.state = {
			resultsPeople: this.resultsPeople,
      resultsPersonas: this.resultsPersonas,
      loading: false
		};
  }

  /**
   * @function
   * Renders the PeoplePicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    //Checks if we need to display the loading message
    var labelToDisplay: string = this.peoplePickerProperties.noResultsText;
    if (this.state['loading'] === true)
      labelToDisplay = this.peoplePickerProperties.loadingContacts;

    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <PeoplePicker suggestions={this.resultsPersonas}
        type={this.peoplePickerProperties.type} isConnected={true}
        canSearchMore={false}
        onSearchFieldChanged={this.onSearchFieldChanged}
        onItemAdded={this.onItemAdded}
        onItemRemoved={this.onItemRemoved}
        searchCategoryName={this.peoplePickerProperties.searchCategoryName}
        initialItems={this.intialPersonas}
        noResultsText={labelToDisplay}
        />
      </div>
    );
  }

  /**
   * @function
   * A search field change occured
   */
  private onSearchFieldChanged(newValue: any): void {
    //Cleat the suggestions list
    this.setState({ resultsPeople: this.resultsPeople, resultsPersonas: this.resultsPersonas, loading: true });
    //Request the search service
    this.searchService.searchPeople(newValue).then((response: IPropertyFieldPeople[]) => {
      this.resultsPeople = [];
      this.resultsPersonas = [];
      //If allowDuplicate == false, so remove duplicates from results
      if (this.props.allowDuplicate === false)
        response = this.removeDuplicates(response);
      response.map((element: IPropertyFieldPeople, index: number) => {
        //Fill the results Array
        this.resultsPeople.push(element);
        //Transform the response in IPersonaProps object
        this.resultsPersonas.push(this.getPersonaFromPeople(element, index));
      });
      //Refresh the component's state
      this.setState({ resultsPeople: this.resultsPeople, resultsPersonas: this.resultsPersonas, loading: false});
    });
  }

  /**
   * @function
   * Remove the duplicates if property allowDuplicate equals false
   */
  private removeDuplicates(responsePeople: IPropertyFieldPeople[]): IPropertyFieldPeople[] {
    if (this.selectedPeople == null || this.selectedPeople.length == 0)
      return responsePeople;
    var res: IPropertyFieldPeople[] = [];
    responsePeople.map((element: IPropertyFieldPeople) => {
      var found: boolean = false;
      for (var i: number = 0; i < this.selectedPeople.length; i++) {
        var responseItem: IPropertyFieldPeople = this.selectedPeople[i];
        if (responseItem.login == element.login) {
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
   * Creates the collection of initial personas from initial IPropertyFieldPeople collection
   */
  private createInitialPersonas(): void {
    if (this.props.initialData == null || typeof (this.props.initialData) != typeof Array<IPropertyFieldPeople>())
      return;
    this.props.initialData.map((element: IPropertyFieldPeople, index: number) => {
      var persona: IPersonaProps = this.getPersonaFromPeople(element, index);
      this.intialPersonas.push(persona);
      this.selectedPersonas.push(persona);
      this.selectedPeople.push(element);
    });
  }

  /**
   * @function
   * Generates a IPersonaProps object from a IPropertyFieldPeople object
   */
  private getPersonaFromPeople(element: IPropertyFieldPeople, index: number): IPersonaProps {
      return {
        primaryText: element.fullName, secondaryText: element.jobTitle, imageUrl: element.imageUrl,
        imageInitials: element.initials, presence: PersonaPresence.none, initialsColor: this.getRandomInitialsColor(index)
      };
  }

  /**
   * @function
   * Refreshes the web part properties
   */
  private refreshWebPartProperties(): void {
    if (this.props.onPropertyChange) {
      this.props.onPropertyChange(this.props.targetProperty, this.selectedPeople);
    }
  }

  /**
   * @function
   * Event raises when the user is selected new people from hte PeoplePicker component
   */
  private onItemAdded(item: IPersonaProps): void {
    var index: number = this.resultsPersonas.indexOf(item);
    if (index > -1) {
      var people: IPropertyFieldPeople = this.resultsPeople[index];
      this.selectedPeople.push(people);
      this.selectedPersonas.push(this.resultsPersonas[index]);
      this.refreshWebPartProperties();
    }
  }

  /**
   * @function
   * Generates a IPersonaProps object from a IPropertyFieldPeople object
   */
  private onItemRemoved(item: IPersonaProps): void {
    var index: number = this.selectedPersonas.indexOf(item);
    if (index > -1) {
      this.selectedPersonas.splice(index, 1);
      this.selectedPeople.splice(index, 1);
      this.refreshWebPartProperties();
    }
  }

  /**
   * @function
   * Generate a PersonaInitialsColor from the item position in the collection
   */
  private getRandomInitialsColor(index: number): PersonaInitialsColor {
    var num: number = index % 13;
    switch (num) {
      case 0: return PersonaInitialsColor.blue;
      case 1: return PersonaInitialsColor.darkBlue;
      case 2: return PersonaInitialsColor.teal;
      case 3: return PersonaInitialsColor.lightGreen;
      case 4: return PersonaInitialsColor.green;
      case 5: return PersonaInitialsColor.darkGreen;
      case 6: return PersonaInitialsColor.lightPink;
      case 7: return PersonaInitialsColor.pink;
      case 8: return PersonaInitialsColor.magenta;
      case 9: return PersonaInitialsColor.purple;
      case 10: return PersonaInitialsColor.black;
      case 11: return PersonaInitialsColor.orange;
      case 12: return PersonaInitialsColor.red;
      case 13: return PersonaInitialsColor.darkRed;
      default: return PersonaInitialsColor.blue;
    }
  }

}

/**
 * @interface
 * Service interface definition
 */
interface IPropertyFieldSearchService {
  /**
   * @function
   * Search People from a query
   */
  searchPeople(query: string): Promise<Array<IPropertyFieldPeople>>;
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
  constructor(pageContext: IWebPartContext){
      this.context = pageContext;
  }

  /**
   * @function
   * Search people from the SharePoint People database
   */
  public searchPeople(query: string): Promise<Array<IPropertyFieldPeople>> {
    if (this.context.environment.type === EnvironmentType.Local) {
      //If the running environment is local, load the data from the mock
      return this.searchPeopleFromMock(query);
    }
    else {
      //If the running env is SharePoint, loads from the peoplepicker web service
      var contextInfoUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/contextinfo";
      var userRequestUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser";
      var httpPostOptions: IHttpClientOptions = {
            headers: {
               "accept": "application/json",
               "content-type": "application/json"
            }
      };
      return this.context.httpClient.post(contextInfoUrl, httpPostOptions).then((response: Response) => {
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
                    'PrincipalType': 1,
                    'QueryString': query
                    //'Required':false,
                    //'SharePointGroupID':null,
                    //'UrlZone':null,
                    //'UrlZoneSpecified':false,
                }
            };
          httpPostOptions = {
            headers: {
               'accept':'application/json',
               'content-type':'application/json',
               "X-RequestDigest": formDigestValue
            },
            body: JSON.stringify(data)
          };
          return this.context.httpClient.post(userRequestUrl, httpPostOptions).then((searchResponse: Response) => {
            return searchResponse.json().then((usersResponse: any) => {
              var res: IPropertyFieldPeople[] = [];
              var values: any = JSON.parse(usersResponse.value);
              values.map(element => {
                var persona: IPropertyFieldPeople = {fullName: element.DisplayText, login: element.Description};
                persona.email = element.EntityData.Email;
                persona.jobTitle = element.EntityData.Title;
                persona.initials = this.getFullNameInitials(persona.fullName);
                persona.imageUrl = this.getUserPhotoUrl(persona.email, this.context.pageContext.web.absoluteUrl);
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
   * Generates Initials from a full name
   */
  private getFullNameInitials(fullName: string): string {
    if (fullName == null)
      return fullName;
    var words: string[] = fullName.split(" ");
    if (words.length == 0) {
      return "";
    }
    else if (words.length == 1) {
      return words[0].charAt(0);
    }
    else {
      return (words[0].charAt(0) + words[1].charAt(0));
    }
  }

  /**
   * @function
   * Gets the user photo url
   */
  private getUserPhotoUrl(userEmail: string, siteUrl: string): string {
    return `${siteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${userEmail}`;
  }


  /**
   * @function
   * Returns fake people results for the Mock mode
   */
  private searchPeopleFromMock(query: string): Promise<Array<IPropertyFieldPeople>> {
    return PeoplePickerMockHttpClient.searchPeople(this.context.pageContext.web.absoluteUrl).then(() => {
          const results: IPropertyFieldPeople[] = [
                { fullName: "Olivier Carpentier", initials: "OC", jobTitle: "Architect", email: "olivierc@contoso.com", login: "olivierc@contoso.com"},
                { fullName: "Katie Jordan", initials: "KJ", jobTitle: "VIP Marketing", email: "katiej@contoso.com", login: "katiej@contoso.com"},
                { fullName: "Gareth Fort", initials: "GF", jobTitle: "Sales Lead", email: "garethf@contoso.com", login: "garethf@contoso.com"},
                { fullName: "Sara Davis", initials: "SD", jobTitle: "Assistant", email: "sarad@contoso.com", login: "sarad@contoso.com"},
                { fullName: "John Doe", initials: "JD", jobTitle: "Developer", email: "johnd@contoso.com", login: "johnd@contoso.com"}
              ];
          return results;
      }) as Promise<Array<IPropertyFieldPeople>>;
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
    private static _results: IPropertyFieldPeople[] = [];

    /**
     * @function
     * Mock search People method
     */
    public static searchPeople(restUrl: string, options?: any): Promise<IPropertyFieldPeople[]> {
      return new Promise<IPropertyFieldPeople[]>((resolve) => {
            resolve(PeoplePickerMockHttpClient._results);
        });
    }

}
