/**
 * @file PropertyFieldTermSetPickerHost.tsx
 * Renders the controls for PropertyFieldTermSetPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 *
 */
import * as React from 'react';
import { IWebPartContext} from '@microsoft/sp-webpart-base';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  IPropertyFieldTermSetPickerPropsInternal,
  ISPTermStores, ISPTermStore,
  ISPTermGroups, ISPTermGroup,
  ISPTermSets, ISPTermSet,
  ISPTermObject
} from './PropertyFieldTermSetPicker';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { TextField } from 'office-ui-fabric-react/lib/TextField';

require('react-ui-tree-draggable/dist/react-ui-tree.css');
var Tree: any = require('react-ui-tree-draggable/dist/react-ui-tree');

/**
 * @interface
 * PropertyFieldTermSetPickerHost properties interface
 *
 */
export interface IPropertyFieldTermSetPickerHostProps extends IPropertyFieldTermSetPickerPropsInternal {
}

/**
 * @interface
 * PropertyFieldTermSetPickerHost state interface
 *
 */
export interface IPropertyFieldFontPickerHostState {
  termStores: ISPTermStores;
  errorMessage?: string;
  openPanel: boolean;
  loaded: boolean;
  activeNodes: ISPTermSets;
}

/**
 * @class
 * Renders the controls for PropertyFieldTermSetPicker component
 */
export default class PropertyFieldTermSetPickerHost extends React.Component<IPropertyFieldTermSetPickerHostProps, IPropertyFieldFontPickerHostState> {

  private async: Async;
  private delayedValidate: (value: ISPTermSets) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldTermSetPickerHostProps) {
    super(props);

    this.state = {
      activeNodes: this.props.initialValues !== undefined ? this.props.initialValues : [],
      termStores: [],
      loaded: false,
      openPanel: false,
      errorMessage: ''
    };

    this.onOpenPanel = this.onOpenPanel.bind(this);
    this.onClosePanel = this.onClosePanel.bind(this);
    this.renderNode = this.renderNode.bind(this);
    this.onClickNode = this.onClickNode.bind(this);
    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  /**
   * @function
   * Loads the list from SharePoint current web site
   */
  private loadTermStores(): void {
    var termsService: SPTermStorePickerService = new SPTermStorePickerService(this.props, this.props.context);
    termsService.getTermStores().then((response: ISPTermStores) => {
      this.state.termStores = response;
      this.state.loaded = true;
      this.setState(this.state);
      response.map((termStore: ISPTermStore, index: number) => {
        termsService.getTermStoresGroups(termStore).then((groupsResponse: ISPTermGroups) => {
          termStore.children = groupsResponse;
          this.setState(this.state);
          groupsResponse.map((group: ISPTermGroup) => {
            termsService.getTermSets(termStore, group).then((termSetsResponse: ISPTermSets) => {
              group.children = termSetsResponse;
              this.setState(this.state);
            });
          });
        });
      });
    });
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: ISPTermSets): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialValues, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || []);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialValues, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValues, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.initialValues, value);
    }
  }

  /**
   * @function
   * Notifies the parent Web Part of a property value change
   */
  private notifyAfterValidate(oldValue: ISPTermSets, newValue: ISPTermSets) {
    if (this.props.onPropertyChange && newValue != null) {
      this.props.properties[this.props.targetProperty] = newValue;
      this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
    }
  }

  /**
   * @function
   * Open the right Panel
   */
  private onOpenPanel(): void {
    if (this.props.disabled === true)
      return;
    this.state.openPanel = true;
    this.state.loaded = false;
    this.loadTermStores();
    this.setState(this.state);
  }

  /**
   * @function
   * Close the panel
   */
  private onClosePanel(): void {
    this.state.openPanel = false;
    this.state.loaded = false;
    this.setState(this.state);
  }

  /**
   * clicks on a node
   * @param node
   */
  private onClickNode(node: ISPTermSet): void {
    if (node.children !== undefined && node.children.length != 0)
      return;
    if (this.props.allowMultipleSelections === false) {
      this.state.activeNodes = [node];
    }
    else {
      var index = this.getSelectedNodePosition(node);
      if (index != -1)
        this.state.activeNodes.splice(index, 1);
      else
        this.state.activeNodes.push(node);
    }
    this.setState(this.state);
    this.delayedValidate(this.state.activeNodes);
  }

  /**
   * @function
   * Gets the given node position in the active nodes collection
   * @param node
   */
  private getSelectedNodePosition(node: ISPTermSet): number {
    for (var i = 0; i < this.state.activeNodes.length; i++) {
      if (node.Guid === this.state.activeNodes[i].Guid)
        return i;
    }
    return -1;
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
   * Renders the given node
   * @param node
   */
  private renderNode(node: ISPTermObject): JSX.Element {
    var style: any = { padding: '4px 5px', width: '100%', display: 'flex'};
    var selected: boolean = false;
    var isFolder: boolean = false;
    if (node.leaf === false || (node.children !== undefined && node.children.length != 0))
      isFolder = true;
    var checkBoxAvailable: boolean = true;
    if (isFolder === true)
      checkBoxAvailable = false;
    var picUrl: string = '';
    if (node.type === "TermStore") {
      picUrl = this.props.context.pageContext.web.absoluteUrl + '/_layouts/15/Images/EMMRoot.png';
    }
    else if (node.type === "TermGroup") {
      picUrl = this.props.context.pageContext.web.absoluteUrl + '/_layouts/15/Images/EMMGroup.png';
    }
    else if (node.type === "TermSet") {
      picUrl = this.props.context.pageContext.web.absoluteUrl + '/_layouts/15/Images/EMMTermSet.png';
      selected = this.getSelectedNodePosition(node as ISPTermSet) != -1;
      if (selected === true) {
        style.backgroundColor = '#EAEAEA';
      }
    }
    return (
        <div style={style} onClick={this.onClickNode.bind(null, node)} name={node.Guid} id={node.Guid} role="menuitem">
          { checkBoxAvailable ?
              <div style={{marginRight: '5px'}}>
                <Checkbox
                    checked={selected}
                    disabled={this.props.disabled}
                    label=''
                    onChange={this.onClickNode.bind(null, node)}
                  />
              </div>
            : ''
          }
          <div style={{paddingTop: '7px'}}>
          {
            picUrl !== undefined && picUrl != '' ?
              <img src={picUrl} width="18" height="18" style={{paddingRight: '5px'}} alt={node.Name}/>
            : ''
          }
          { node.type === "TermStore" ? <strong>{node.Name}</strong> : node.Name }
          </div>
        </div>
    );
  }

  /**
   * @function
   * Renders the SPListpicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {
    var termSetsString: string = '';
    if (this.state.activeNodes !== undefined) {
      this.state.activeNodes.map((termSet: ISPTermSet, index: number) => {
        if (index > 0)
          termSetsString += '; ';
        termSetsString += termSet.Name;
      });
    }
    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <table style={{width: '100%', borderSpacing: 0}}>
          <tbody>
            <tr>
              <td width="*">
                <TextField
                  disabled={this.props.disabled}
                  style={{width:'100%'}}
                  onChanged={null}
                  readOnly={true}
                  value={termSetsString}
                />
              </td>
              <td width="32">
                <Button disabled={this.props.disabled} buttonType={ButtonType.icon} icon="Tag" onClick={this.onOpenPanel} />
              </td>
            </tr>
          </tbody>
        </table>
        { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div style={{paddingBottom: '8px'}}><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
            : ''}

        <Panel
          isOpen={this.state.openPanel} hasCloseButton={true} onDismiss={this.onClosePanel}
          isLightDismiss={true} type={PanelType.medium}
          headerText={this.props.panelTitle}>

          { this.state.loaded === false ? <Spinner type={ SpinnerType.normal } /> : '' }
          { this.state.loaded === true ? this.state.termStores.map((rootNode: ISPTermStore, index: number) => {
              return (
                <Tree
                  paddingLeft={15}
                  tree={rootNode}
                  isNodeCollapsed={false}
                  renderNode={this.renderNode}
                  draggable={false}
                  key={'termRootNode-' + index}
                />
              );
            })
            : ''
          }
        </Panel>

      </div>
    );
  }
}

/**
 * @class
 * Service implementation to manage term stores in SharePoint
 */
class SPTermStorePickerService {

  private context: IWebPartContext;
  private props: IPropertyFieldTermSetPickerHostProps;
  private taxonomySession: string;
  private formDigest: string;

  /**
   * @function
   * Service constructor
   */
  constructor(_props: IPropertyFieldTermSetPickerHostProps, pageContext: IWebPartContext){
      this.props = _props;
      this.context = pageContext;
  }

  /**
   * @function
   * Gets the collection of term stores in the current SharePoint env
   */
  public getTermStores(): Promise<ISPTermStores> {
    if (Environment.type === EnvironmentType.Local) {
      //If the running environment is local, load the data from the mock
      return this.getTermStoresFromMock();
    }
    else {
      //First gets the FORM DIGEST VALUE
      var contextInfoUrl: string = this.context.pageContext.web.absoluteUrl + "/_api/contextinfo";
      var httpPostOptions: ISPHttpClientOptions = {
        headers: {
          "accept": "application/json",
          "content-type": "application/json"
        }
      };
      return this.context.spHttpClient.post(contextInfoUrl, SPHttpClient.configurations.v1, httpPostOptions).then((response: SPHttpClientResponse) => {
        return response.json().then((jsonResponse: any) => {
          this.formDigest = jsonResponse.FormDigestValue;

          //Build the Client Service Request
          var clientServiceUrl = this.context.pageContext.web.absoluteUrl + '/_vti_bin/client.svc/ProcessQuery';
          var data = '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="Javascript Library"><Actions><ObjectPath Id="1" ObjectPathId="0" /><ObjectIdentityQuery Id="2" ObjectPathId="0" /><Query Id="3" ObjectPathId="0"><Query SelectAllProperties="true"><Properties /></Query></Query><ObjectPath Id="5" ObjectPathId="4" /><Query Id="6" ObjectPathId="4"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="0" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Property Id="4" ParentId="0" Name="TermStores" /></ObjectPaths></Request>';
          httpPostOptions = {
            headers: {
              'accept': 'application/json',
              'content-type': 'application/json',
              "X-RequestDigest": this.formDigest
            },
            body: data
          };
          return this.context.spHttpClient.post(clientServiceUrl, SPHttpClient.configurations.v1, httpPostOptions).then((serviceResponse: SPHttpClientResponse) => {
            return serviceResponse.json().then((serviceJSONResponse: any) => {

              //Construct results
              var res: ISPTermStores = [];
              serviceJSONResponse.map((child: any) => {
                if (child != null && child['_ObjectType_'] !== undefined) {
                  var objType = child['_ObjectType_'];
                  if (objType === "SP.Taxonomy.TaxonomySession") {
                    this.taxonomySession = child['_ObjectIdentity_'];
                  }
                  else if (objType === "SP.Taxonomy.TermStoreCollection") {
                    var childTermStores = child['_Child_Items_'];
                    childTermStores.map((childTerm: any) => {
                      var newTermStore: ISPTermStore = {
                        Name: childTerm['Name'] !== undefined ? childTerm['Name'] : '',
                        Guid: childTerm['Id'] !== undefined ? this.cleanGuid(childTerm['Id']): '',
                        Identity: childTerm['_ObjectIdentity_'] !== undefined ? childTerm['_ObjectIdentity_'] : '',
                        IsOnline: childTerm['IsOnline'] !== undefined ? childTerm['IsOnline'] : '',
                        WorkingLanguage: childTerm['WorkingLanguage'] !== undefined ? childTerm['WorkingLanguage'] : '',
                        DefaultLanguage: childTerm['DefaultLanguage'] !== undefined ? childTerm['DefaultLanguage'] : '',
                        Languages: childTerm['Languages'] !== undefined ? childTerm['Languages'] : [],
                        leaf: false,
                        type: 'TermStore'
                      };
                      if (!(this.props.excludeOfflineTermStores === true && newTermStore.IsOnline === false))
                        res.push(newTermStore);
                    });
                  }
                }
              });
              return res;

            });
          });

        });
      });
    }
  }

  public getTermStoresGroups(termStore: ISPTermStore): Promise<ISPTermGroups> {
    if (Environment.type === EnvironmentType.Local) {
      //If the running environment is local, load the data from the mock
      return this.getTermStoresGroupsFromMock(termStore.Identity);
    }
    else {
      //Build the Client Service Request
      var clientServiceUrl = this.context.pageContext.web.absoluteUrl + '/_vti_bin/client.svc/ProcessQuery';
      var data = '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="Javascript Library"><Actions><ObjectPath Id="16" ObjectPathId="15" /><Query Id="17" ObjectPathId="15"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Property Id="15" ParentId="5" Name="Groups" /><Identity Id="5" Name="' + termStore.Identity + '" /></ObjectPaths></Request>';
      var httpPostOptions = {
        headers: {
              'accept': 'application/json',
              'content-type': 'application/json',
              "X-RequestDigest": this.formDigest
        },
        body: data
      };
      return this.context.spHttpClient.post(clientServiceUrl, SPHttpClient.configurations.v1, httpPostOptions).then((serviceResponse: SPHttpClientResponse) => {
        return serviceResponse.json().then((serviceJSONResponse: any) => {
          var res: ISPTermGroups = [];
          serviceJSONResponse.map((child: any) => {
            var objType = child['_ObjectType_'];
            if (objType === "SP.Taxonomy.TermGroupCollection") {
              if (child['_Child_Items_'] !== undefined) {
                child['_Child_Items_'].map((childGroup: any) => {
                  var objGroup: ISPTermGroup = {
                    Name: childGroup['Name'] !== undefined ? childGroup['Name'] : '',
                    Guid: childGroup['Id'] !== undefined ? this.cleanGuid(childGroup['Id']) : '',
                    Identity: childGroup['_ObjectIdentity_'] !== undefined ? childGroup['_ObjectIdentity_'] : '',
                    IsSiteCollectionGroup: childGroup['IsSiteCollectionGroup'] !== undefined ? childGroup['IsSiteCollectionGroup'] : '',
                    IsSystemGroup: childGroup['IsSystemGroup'] !== undefined ? childGroup['IsSystemGroup'] : '',
                    CreatedDate: childGroup['CreatedDate'] !== undefined ? childGroup['CreatedDate'] : '',
                    LastModifiedDate: childGroup['LastModifiedDate'] !== undefined ? childGroup['LastModifiedDate'] : '',
                    leaf: false,
                    type: 'TermGroup'
                  };
                  if (this.props.excludeSystemGroup === true) {
                    if (objGroup.IsSystemGroup !== true)
                      res.push(objGroup);
                  }
                  else {
                    res.push(objGroup);
                  }
                });
              }
            }
          });
          return res;
        });
      });
    }
  }

  public getTermSets(termStore: ISPTermStore, group: ISPTermGroup): Promise<ISPTermSets> {
    if (Environment.type === EnvironmentType.Local) {
      //If the running environment is local, load the data from the mock
      return this.getTermSetsFromMock(termStore.Identity, group.Guid);
    }
    else {
      //Build the Client Service Request
      var clientServiceUrl = this.context.pageContext.web.absoluteUrl + '/_vti_bin/client.svc/ProcessQuery';
      var data = '<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="Javascript Library"><Actions><ObjectPath Id="26" ObjectPathId="25" /><ObjectIdentityQuery Id="27" ObjectPathId="25" /><ObjectPath Id="29" ObjectPathId="28" /><Query Id="30" ObjectPathId="28"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Method Id="25" ParentId="15" Name="GetById"><Parameters><Parameter Type="String">' + group.Guid + '</Parameter></Parameters></Method><Property Id="28" ParentId="25" Name="TermSets" /><Property Id="15" ParentId="5" Name="Groups" /><Identity Id="5" Name="' + termStore.Identity + '" /></ObjectPaths></Request>';
      var httpPostOptions = {
        headers: {
              'accept': 'application/json',
              'content-type': 'application/json',
              "X-RequestDigest": this.formDigest
        },
        body: data
      };
      return this.context.spHttpClient.post(clientServiceUrl, SPHttpClient.configurations.v1, httpPostOptions).then((serviceResponse: SPHttpClientResponse) => {
        return serviceResponse.json().then((serviceJSONResponse: any) => {
          var res: ISPTermSets = [];
          serviceJSONResponse.map((child: any) => {
            var objType = child['_ObjectType_'];
            if (objType === "SP.Taxonomy.TermSetCollection") {
              if (child['_Child_Items_'] !== undefined) {
                child['_Child_Items_'].map((childGroup: any) => {
                  var objGroup: ISPTermSet = {
                    Name: childGroup['Name'] !== undefined ? childGroup['Name'] : '',
                    Guid: childGroup['Id'] !== undefined ? this.cleanGuid(childGroup['Id']) : '',
                    Identity: childGroup['_ObjectIdentity_'] !== undefined ? childGroup['_ObjectIdentity_'] : '',
                    CustomSortOrder: childGroup['CustomSortOrder'] !== undefined ? childGroup['CustomSortOrder'] : '',
                    IsAvailableForTagging: childGroup['IsAvailableForTagging'] !== undefined ? childGroup['IsAvailableForTagging'] : '',
                    Owner: childGroup['Owner'] !== undefined ? childGroup['Owner'] : '',
                    Contact: childGroup['Contact'] !== undefined ? childGroup['Contact'] : '',
                    Description: childGroup['Description'] !== undefined ? childGroup['Description'] : '',
                    IsOpenForTermCreation: childGroup['IsOpenForTermCreation'] !== undefined ? childGroup['IsOpenForTermCreation'] : '',
                    TermStoreGuid: termStore.Guid,
                    leaf: true,
                    type: 'TermSet'
                  };
                  if (this.props.displayOnlyTermSetsAvailableForTagging === true) {
                    if (objGroup.IsAvailableForTagging === true)
                      res.push(objGroup);
                  }
                  else {
                    res.push(objGroup);
                  }
                });
              }
            }
          });
          return res;
        });
      });
    }
  }

  /**
   * @function
   * Clean the Guid from the Web Service response
   * @param guid
   */
  private cleanGuid(guid: string): string {
    if (guid !== undefined)
      return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
    else
      return '';
  }

  /**
   * @function
   * Returns 3 fake SharePoint lists for the Mock mode
   */
  private getTermStoresFromMock(): Promise<ISPTermStores> {
    return SPTermStoreMockHttpClient.getTermStores(this.context.pageContext.web.absoluteUrl).then(() => {
          const mockData: ISPTermStores = [
            { Name: 'Taxonomy_jHIKWt45FAQsxsbHfZ3r1Q==', Guid: '/Guid(8ca33abb-2ee5-42d4-acb6-bd138adec078)/',
              Identity: '8ca33abb-2ee5-42d4-acb6-bd138adec078',
              IsOnline: true, WorkingLanguage: '1033',
              DefaultLanguage: '1033', Languages:[],
              leaf: false, type: 'TermStore'
            }
          ];
          return mockData;
      }) as Promise<ISPTermStores>;
  }

  private getTermStoresGroupsFromMock(termStoreIdentity: string): Promise<ISPTermGroups> {
    return SPTermStoreMockHttpClient.getTermStoresGroups(this.context.pageContext.web.absoluteUrl).then(() => {
          const mockData: ISPTermGroups = [
            {
              Name: 'People', Guid: '/Guid(8ca33abb-2ee5-42d4-acb6-bd138adec078)/',
              Identity: '8ca33abb-2ee5-42d4-acb6-bd138adec078',
              IsSiteCollectionGroup: false,
              IsSystemGroup: false,
              CreatedDate: '',
              LastModifiedDate: '',
              leaf: false, type: 'TermGroup'
            },
            {
              Name: 'Search Dictionaries', Guid: '/Guid(8ca33acc-2ee5-42d4-acb6-bd138adec078)/',
              Identity: '8ca33abb-2ee5-42d4-acb6-bd138adec078',
              IsSiteCollectionGroup: false,
              IsSystemGroup: false,
              CreatedDate: '',
              LastModifiedDate: '',
              leaf: false, type: 'TermGroup'
            }
          ];
          return mockData;
      }) as Promise<ISPTermGroups>;
  }

  private getTermSetsFromMock(termStoreIdentity: string, groupGuid: string): Promise<ISPTermSets> {
    return SPTermStoreMockHttpClient.getTermSetsFromMock(this.context.pageContext.web.absoluteUrl).then(() => {
          const mockData: ISPTermSets = [
            {
              Name: 'People', Guid: '/Guid(8ca44acc-2ee5-42d4-acb6-bd138adec078)/',
              Identity: '8ca44acc-2ee5-42d4-acb6-bd138adec078',
              CustomSortOrder: '',
              IsAvailableForTagging: true,
              Owner: '',
              Contact: '',
              Description: '',
              IsOpenForTermCreation: true,
              TermStoreGuid: '8ca33abb-2ee5-42d4-acb6-bd138adec078',
              leaf: true, type: 'TermSet'
            },
            {
              Name: 'Job Title', Guid: '/Guid(8ca44acc-2ff4-42d4-acb6-bd138adec078)/',
              Identity: '8ca44acc-2ff4-42d4-acb6-bd138adec078',
              CustomSortOrder: '',
              IsAvailableForTagging: true,
              Owner: '',
              Contact: '',
              Description: '',
              IsOpenForTermCreation: true,
              TermStoreGuid: '8ca33abb-2ee5-42d4-acb6-bd138adec078',
              leaf: true, type: 'TermSet'
            }
          ];
          return mockData;
      }) as Promise<ISPTermSets>;
  }

}


/**
 * @class
 * Defines a http client to request mock data to use the web part with the local workbench
 */
class SPTermStoreMockHttpClient {

    /**
     * @var
     * Mock SharePoint result sample
     */
    private static _mockTermStores: ISPTermStores = [];

    private static _mockTermStoresGroups: ISPTermGroups = [];

    private static _mockTermSets: ISPTermSets = [];

    /**
     * @function
     * Mock search People method
     */
    public static getTermStores(restUrl: string, options?: any): Promise<ISPTermStores> {
      return new Promise<ISPTermStores>((resolve) => {
            resolve(SPTermStoreMockHttpClient._mockTermStores);
        });
    }

    public static getTermStoresGroups(restUrl: string, options?: any): Promise<ISPTermGroups> {
      return new Promise<ISPTermGroups>((resolve) => {
            resolve(SPTermStoreMockHttpClient._mockTermStoresGroups);
        });
    }

    public static getTermSetsFromMock(restUrl: string, options?: any): Promise<ISPTermSets> {
      return new Promise<ISPTermSets>((resolve) => {
            resolve(SPTermStoreMockHttpClient._mockTermSets);
        });
    }

}
