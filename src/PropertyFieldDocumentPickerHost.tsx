/**
 * @file PropertyFieldDocumentPickerHost.tsx
 * Renders the controls for PropertyFieldDocumentPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IPropertyFieldDocumentPickerPropsInternal } from './PropertyFieldDocumentPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Async } from 'office-ui-fabric-react/lib/Utilities';

import * as strings from 'sp-client-custom-fields/strings';

/**
 * @interface
 * PropertyFieldDocumentPickerHost properties interface
 *
 */
export interface IPropertyFieldDocumentPickerHostProps extends IPropertyFieldDocumentPickerPropsInternal {
}

export interface IPropertyFieldDocumentPickerHostState {
  openPanel?: boolean;
  openRecent?: boolean;
  openSite?: boolean;
  openUpload?: boolean;
  recentImages?: string[];
  selectedImage: string;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldDocumentPicker component
 */
export default class PropertyFieldDocumentPickerHost extends React.Component<IPropertyFieldDocumentPickerHostProps, IPropertyFieldDocumentPickerHostState> {

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldDocumentPickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onOpenPanel = this.onOpenPanel.bind(this);
    this.onClosePanel = this.onClosePanel.bind(this);
    this.onImageRender = this.onImageRender.bind(this);
    this.onClickRecent = this.onClickRecent.bind(this);
    this.onClickSite = this.onClickSite.bind(this);
    this.onClickUpload = this.onClickUpload.bind(this);
    this.mouseEnterImage = this.mouseEnterImage.bind(this);
    this.mouseLeaveImage = this.mouseLeaveImage.bind(this);
    this.handleIframeData = this.handleIframeData.bind(this);
    this.onEraseButton = this.onEraseButton.bind(this);

    //Inits the state
    this.state = {
      selectedImage: this.props.initialValue,
      openPanel: false,
      openRecent: false,
      openSite: true,
      openUpload: false,
      recentImages: [],
      errorMessage: ''
    };

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);

    //Load recent images
    this.LoadRecentImages();
  }


 /**
  * @function
  * Save the image value
  *
  */
  private saveImageProperty(imageUrl: string): void {
    this.delayedValidate(imageUrl);
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialValue, value);
      return;
    }

    if (this.latestValidateValue === value)
      return;
    this.latestValidateValue = value;

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialValue, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValue, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.initialValue, value);
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
    }
  }

  /**
  * @function
  * Click on erase button
  *
  */
  private onEraseButton(): void {
    this.state.selectedImage = '';
    this.setState(this.state);
    this.saveImageProperty('');
  }

  /**
  * @function
  * Open the panel
  *
  */
  private onOpenPanel(element?: any): void {
    this.state.openPanel = true;
    this.setState(this.state);
  }

  /**
  * @function
  * Close the panel
  *
  */
  private onClosePanel(element?: any): void {
    this.state.openPanel = false;
    this.setState(this.state);
  }

  private onClickRecent(element?: any): void {
    //this.state.openRecent = true;
    //this.state.openSite = false;
    //this.state.openUpload = false;
    //this.setState(this.state);
  }

  /**
  * @function
  * Intercepts the iframe onedrive messages
  *
  */
  private handleIframeData(element?: any) {
    if (this.state.openSite != true || this.state.openPanel != true)
      return;
    var data: string = element.data;
    var indexOfPicker = data.indexOf("[OneDrive-FromPicker]");
    if (indexOfPicker != -1) {
      var message = data.replace("[OneDrive-FromPicker]", "");
      var messageObject = JSON.parse(message);
      if (messageObject.type == "cancel") {
        this.onClosePanel();
      } else if (messageObject.type == "success") {
        var imageUrl = messageObject.items[0].sharePoint.url;
        if (imageUrl.indexOf(".doc") > -1 || imageUrl.indexOf(".docx") > -1 || imageUrl.indexOf(".ppt") > -1 ||
         imageUrl.indexOf(".pptx") > -1 || imageUrl.indexOf(".xls") > -1 || imageUrl.indexOf(".xlsx") > -1 ||
         imageUrl.indexOf(".pdf") > -1  || imageUrl.indexOf(".txt") > -1) {
          this.state.selectedImage = imageUrl;
          this.setState(this.state);
          this.saveImageProperty(imageUrl);
          this.onClosePanel();
         }
      }
    }
  }

  /**
  * @function
  * When component is mount, attach the iframe event watcher
  *
  */
  public componentDidMount() {
    window.addEventListener('message', this.handleIframeData, false);
  }

  /**
  * @function
  * Releases the watcher
  *
  */
  public componentWillUnmount() {
    window.removeEventListener('message', this.handleIframeData, false);
    this.async.dispose();
  }

  private onClickSite(element?: any): void {
    this.state.openRecent = false;
    this.state.openSite = true;
    this.state.openUpload = false;
    this.setState(this.state);
  }

  private onClickUpload(element?: any): void {
    this.state.openRecent = false;
    this.state.openSite = false;
    this.state.openUpload = true;
    this.setState(this.state);
  }

  private LoadRecentImages(): void {
    //var folderService: SPFolderPickerService = new SPFolderPickerService(this.props.context);
    //folderService.getFolders(this.state.currentSPFolder, this.currentPage, this.pageItemCount).then((response: ISPFolders) => {
      //Binds the results
      //this.state.childrenFolders = response;
      //this.setState({ openRecent: this.state.openRecent,openSite: this.state.openSite, openUpload: this.state.openUpload, loading: false, selectedFolder: this.state.selectedFolder, currentSPFolder: this.state.currentSPFolder, childrenFolders: this.state.childrenFolders });
    //});
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    var iframeUrl = this.props.context.pageContext.web.absoluteUrl;
    iframeUrl += '/_layouts/15/onedrive.aspx?picker=';
    iframeUrl += '%7B%22sn%22%3Afalse%2C%22v%22%3A%22files%22%2C%22id%22%3A%221%22%2C%22o%22%3A%22';
    iframeUrl += encodeURI(this.props.context.pageContext.web.absoluteUrl.replace(this.props.context.pageContext.web.serverRelativeUrl, ""));
    iframeUrl += "%22%7D&id=";
    iframeUrl += encodeURI(this.props.context.pageContext.web.serverRelativeUrl);
    iframeUrl += '&view=2&typeFilters=';
    iframeUrl += encodeURI('folder,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.pdf,.txt');
    iframeUrl += '&p=2';

    var previewUrl = this.props.context.pageContext.web.absoluteUrl;
    previewUrl += '/_layouts/15/getpreview.ashx?path=';
    previewUrl += encodeURI(this.state.selectedImage);

    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>
        <Button disabled={this.props.disabled} onClick={this.onOpenPanel}>{strings.DocumentPickerButtonSelect}</Button>
        <Button onClick={this.onEraseButton} disabled={this.props.disabled === false && (this.state.selectedImage != null && this.state.selectedImage != '') ? false: true}>
        {strings.DocumentPickerButtonReset}</Button>

        { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
            : ''}

        {this.state.selectedImage != null && this.state.selectedImage != '' ?
        <div style={{marginTop: '7px'}}>
          <img src={previewUrl} width="225px" height="225px" alt="Preview" />
        </div>
        : ''}

        { this.state.openPanel === true ?
        <Panel
          isOpen={this.state.openPanel} hasCloseButton={true} onDismiss={this.onClosePanel}
          isLightDismiss={true} type={PanelType.large}
          headerText={strings.DocumentPickerTitle}>

          <div style={{backgroundColor: '#F4F4F4', width: '100%', height:'80vh', paddingTop: '0px', display: 'inline-flex'}}>

            <div style={{width: '206px', backgroundColor: 'white'}}>
              <div style={{width: '260px', backgroundColor: '#F4F4F4', height:'40px', marginBottom:'70px'}}>
              </div>

              <div style={{paddingLeft: '20px', paddingTop: '10px', color:'#A6A6A6', paddingBottom: '10px',
              borderLeftWidth: '1px',
              borderLeftStyle: 'solid',
              borderLeftColor: this.state.openRecent === true ? 'blue' : 'white',
              backgroundColor: this.state.openRecent === true ? '#F4F4F4' : '#FFFFFF'
              }} onClick={this.onClickRecent} role="menuitem">
                <i className="ms-Icon ms-Icon--Clock" style={{fontSize: '30px'}}></i>
                &nbsp;{strings.DocumentPickerRecent}
              </div>
              <div style={{cursor: 'pointer', paddingLeft: '20px', paddingTop: '10px', paddingBottom: '10px',
              borderLeftWidth: '1px',
              borderLeftStyle: 'solid',
              borderLeftColor: this.state.openSite === true ? 'blue' : 'white',
              backgroundColor: this.state.openSite === true ? '#F4F4F4' : '#FFFFFF'
              }} onClick={this.onClickSite} role="menuitem">
                <i className="ms-Icon ms-Icon--Globe" style={{fontSize: '30px'}}></i>
                &nbsp;{strings.DocumentPickerSite}
              </div>
          </div>

          {this.state.openRecent == true ?
          <div id="recent" style={{marginLeft: '2px', width:'100%', backgroundColor: 'white'}}>
            <div style={{width: '100%', backgroundColor: '#F4F4F4', height:'40px', marginBottom:'20px'}}>
              </div>
            <div style={{paddingLeft: '30px'}}>
              <h1 className="ms-font-xl">Recent images</h1>

                {["1", "2", "1", "2", "1", "2", "1", "2", "1", "2"].map((element?: any, index?: any) => {
                  return this.onImageRender(element, index);
                })}

             </div>
          </div>
          : '' }

          <div id="site" style={{marginLeft: '2px',paddingLeft: '0px', paddingTop:'0px', backgroundColor: 'white', visibility: this.state.openSite === true ? 'visible' : 'hidden', width: this.state.openSite === true ? '100%' : '0px', height: this.state.openSite === true ? '80vh' : '0px',}}>

            <iframe ref="filePickerIFrame" style={{width: this.state.openSite === true ? '100%':'0px', height: this.state.openSite === true ?'80vh':'0px', borderWidth:'0'}} className="filePickerIFrame_d791363d" role="application" title="Select files from site picker view. Use toolbaar menu to perform operations, breadcrumbs to navigate between folders and arrow keys to navigate within the list"
            src={iframeUrl}></iframe>

          </div>


          </div>


          {this.state.openSite === false ?
          <div style={{
                position: 'absolute',
    bottom: '0',
    right: '0',
    marginBottom: '20px',
    marginRight: '20px'
          }}>
            <Button buttonType={ButtonType.primary}> Open </Button>
            <Button buttonType={ButtonType.normal} onClick={this.onClosePanel}> Cancel </Button>
          </div>
          : ''}

        </Panel>
        : ''}

      </div>
    );
  }


  private mouseEnterImage(element?: any): void {
    element.currentTarget.style.backgroundColor = 'grey';
    element.currentTarget.children[0].children[0].style.visibility = 'visible';
  }

  private mouseLeaveImage(element?: any): void {
    element.currentTarget.style.backgroundColor = 'white';
    element.currentTarget.children[0].children[0].style.visibility = 'hidden';
  }

    private onImageRender(item?: any, index?: number): React.ReactNode {
    return (
      <div style={{padding: '2px', width: '191px', height: '191px', display:'inline-block'}} onMouseEnter={this.mouseEnterImage} onMouseLeave={this.mouseLeaveImage}>
        <div style={{cursor: 'pointer',width: '187px', height: '187px',
          backgroundSize: 'cover',
          marginRight: '0px', marginBottom: '0px', paddingTop: '0px', paddingLeft: '0'
          }}>
        </div>
      </div>
    );
  }

}


/**
 * @interface
 * Defines a collection of SharePoint folders
 */
export interface ISPFolders {
  value: ISPFolder[];
}

/**
 * @interface
 * Defines a SharePoint folder
 */
export interface ISPFolder {
  Name: string;
  ServerRelativeUrl: string;
}

/**
 * @class
 * Service implementation to get folders from current SharePoint site
 */
class SPFolderPickerService {

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
   * Gets the collection of sub folders of the given folder
   */
  public getFolders(parentFolderServerRelativeUrl?: string, currentPage?: number, pageItemCount?: number): Promise<ISPFolders> {
    if (Environment.type === EnvironmentType.Local) {
      //If the running environment is local, load the data from the mock
      return this.getFoldersMock(parentFolderServerRelativeUrl);
    }
    else {
      //If the running environment is SharePoint, request the folders REST service
      var queryUrl: string = this.context.pageContext.web.absoluteUrl;
      var skipNumber = currentPage * pageItemCount;
      if (parentFolderServerRelativeUrl == null || parentFolderServerRelativeUrl == '' || parentFolderServerRelativeUrl == '/') {
        //The folder is the web root site
        queryUrl += "/_api/web/folders?$select=Name,ServerRelativeUrl&$orderBy=Name&$top=";
        queryUrl += pageItemCount;
        queryUrl += "&$skip=";
        queryUrl += skipNumber;
      }
      else {
        //Loads sub folders
        queryUrl += "/_api/web/GetFolderByServerRelativeUrl('";
        queryUrl += parentFolderServerRelativeUrl;
        queryUrl += "')/folders?$select=Name,ServerRelativeUrl&$orderBy=Name&$top=";
        queryUrl += pageItemCount;
        queryUrl += "&$skip=";
        queryUrl += skipNumber;
      }

      return this.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
          return response.json();
      });
    }
  }

  /**
   * @function
   * Returns 3 fake SharePoint folders for the Mock mode
   */
  private getFoldersMock(parentFolderServerRelativeUrl?: string): Promise<ISPFolders> {
    return SPFolderPickerMockHttpClient.getFolders(this.context.pageContext.web.absoluteUrl).then(() => {
          const listData: ISPFolders = {
              value:
              [
                  { Name: 'Mock Folder One', ServerRelativeUrl: '/mockfolderone' },
                  { Name: 'Mock Folder Two', ServerRelativeUrl: '/mockfoldertwo' },
                  { Name: 'Mock Folder Three', ServerRelativeUrl: '/mockfolderthree' }
              ]
          };
          return listData;
      }) as Promise<ISPFolders>;
  }

}


/**
 * @class
 * Defines a http client to request mock data to use the web part with the local workbench
 */
class SPFolderPickerMockHttpClient {

    /**
     * @var
     * Mock SharePoint result sample
     */
    private static _results: ISPFolders = { value: []};

    /**
     * @function
     * Mock get folders method
     */
    public static getFolders(restUrl: string, options?: any): Promise<ISPFolders> {
      return new Promise<ISPFolders>((resolve) => {
            resolve(SPFolderPickerMockHttpClient._results);
        });
    }

}
