/**
 * @file PropertyFieldPicturePickerHost.tsx
 * Renders the controls for PropertyFieldPicturePicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { EnvironmentType } from '@microsoft/sp-client-base';
import { IWebPartContext } from '@microsoft/sp-client-preview';
import { IPropertyFieldPicturePickerPropsInternal } from './PropertyFieldPicturePicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';

import * as strings from 'sp-client-custom-fields/strings';

/**
 * @interface
 * PropertyFieldPicturePickerHost properties interface
 *
 */
export interface IPropertyFieldPicturePickerHostProps extends IPropertyFieldPicturePickerPropsInternal {
}

export interface IPropertyFieldPicturePickerHostState {
  openPanel?: boolean;
  openRecent?: boolean;
  openSite?: boolean;
  openUpload?: boolean;
  recentImages?: string[];
  selectedImage: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldPicturePicker component
 */
export default class PropertyFieldPicturePickerHost extends React.Component<IPropertyFieldPicturePickerHostProps, IPropertyFieldPicturePickerHostState> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldPicturePickerHostProps) {
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
      recentImages: []
    };

    //Load recent images
    this.LoadRecentImages();
  }


 /**
  * @function
  * Save the image value
  *
  */
  private saveImageProperty(imageUrl: string): void {
    if (this.props.onPropertyChange) {
      this.props.onPropertyChange(this.props.targetProperty, imageUrl);
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
        if (imageUrl.indexOf(".jpg") > -1 || imageUrl.indexOf(".png") > -1 || imageUrl.indexOf(".jpeg") > -1 ||
         imageUrl.indexOf(".gif") > -1 || imageUrl.indexOf(".tiff") > -1) {
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
    iframeUrl += encodeURI('folder,.gif,.jpg,.jpeg,.bmp,.dib,.tif,.tiff,.ico,.png');
    iframeUrl += '&p=2';

    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>
        <Button onClick={this.onOpenPanel}>{strings.PicturePickerButtonSelect}</Button>
        <Button onClick={this.onEraseButton} disabled={this.state.selectedImage != null && this.state.selectedImage != '' ? false: true}>
        {strings.PicturePickerButtonReset}</Button>
        {this.state.selectedImage != null && this.state.selectedImage != '' ?
        <div style={{marginTop: '7px'}}>
          <img src={this.state.selectedImage} width="225px" height="225px" />
        </div>
        : ''}

        { this.state.openPanel === true ?
        <Panel
          isOpen={this.state.openPanel} hasCloseButton={true} onDismiss={this.onClosePanel}
          isLightDismiss={true} type={PanelType.large}
          headerText={strings.PicturePickerTitle}>

          <div style={{backgroundColor: '#F4F4F4', width: '100%', height:'80vh', paddingTop: '0px', display: 'inline-flex'}}>

            <div style={{width: '206px', backgroundColor: 'white'}}>
              <div style={{width: '260px', backgroundColor: '#F4F4F4', height:'40px', marginBottom:'70px'}}>
              </div>

              <div style={{paddingLeft: '20px', paddingTop: '10px', color:'#A6A6A6', paddingBottom: '10px',
              borderLeftWidth: '1px',
              borderLeftStyle: 'solid',
              borderLeftColor: this.state.openRecent === true ? 'blue' : 'white',
              backgroundColor: this.state.openRecent === true ? '#F4F4F4' : '#FFFFFF'
              }} onClick={this.onClickRecent}>
                <i className="ms-Icon ms-Icon--Clock" style={{fontSize: '30px'}}></i>
                &nbsp;{strings.PicturePickerRecent}
              </div>
              <div style={{cursor: 'pointer', paddingLeft: '20px', paddingTop: '10px', paddingBottom: '10px',
              borderLeftWidth: '1px',
              borderLeftStyle: 'solid',
              borderLeftColor: this.state.openSite === true ? 'blue' : 'white',
              backgroundColor: this.state.openSite === true ? '#F4F4F4' : '#FFFFFF'
              }} onClick={this.onClickSite}>
                <i className="ms-Icon ms-Icon--Globe" style={{fontSize: '30px'}}></i>
                &nbsp;{strings.PicturePickerSite}
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
        : '' }

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
          backgroundImage: "url('https://ocarpenmsdn.sharepoint.com/sites/devcenter/Pictures/09.jpg')",
          backgroundSize: 'cover',
          marginRight: '0px', marginBottom: '0px', paddingTop: '0px', paddingLeft: '0'
          }}>

          <div className="ms-ItemTile-checkCircle" style={{
                position: 'absolute', top: '0', right: '0', marginTop: '5px', marginRight: '5px', visibility: 'hidden'
          }}>
            <svg className="ms-CheckCircle is-checked" height="20" width="20">
              <circle style={{
                    fill: '#ffffff', stroke: '#ffffff', strokeWidth: '1px'
              }} cx="10" cy="10" r="9" strokeWidth="1" ></circle>
              <polyline style={{ stroke: '#ffffff'}} points="6.3,10.3 9,13 13.3,7.5" strokeWidth="1.5" fill="none"></polyline>
            </svg>
          </div>
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
    if (this.context.environment.type === EnvironmentType.Local) {
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
      return this.context.httpClient.get(queryUrl).then((response: Response) => {
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
