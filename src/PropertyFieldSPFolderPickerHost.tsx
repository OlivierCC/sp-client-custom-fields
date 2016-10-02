/**
 * @file PropertyFieldSPFolderPickerHost.tsx
 * Renders the controls for PropertyFieldSPFolderPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { EnvironmentType } from '@microsoft/sp-client-base';
import { IWebPartContext } from '@microsoft/sp-client-preview';
import { IPropertyFieldSPFolderPickerPropsInternal } from './PropertyFieldSPFolderPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { List } from 'office-ui-fabric-react/lib/List';

import * as strings from 'propertyFieldsStrings';

/**
 * @interface
 * PropertyFieldSPFolderPickerHost properties interface
 *
 */
export interface IPropertyFieldSPFolderPickerHostProps extends IPropertyFieldSPFolderPickerPropsInternal {
}

/**
 * @interface
 * Interface to define the state of the rendering control
 *
 */
export interface IPropertyFieldSPFolderPickerHostState {
  isOpen: boolean;
  loading: boolean;
  currentSPFolder?: string;
  childrenFolders?: ISPFolders;
  selectedFolder?: string;
  confirmFolder?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldSPFolderPicker component
 */
export default class PropertyFieldSPFolderPickerHost extends React.Component<IPropertyFieldSPFolderPickerHostProps, IPropertyFieldSPFolderPickerHostState> {

  private currentPage: number = 0;
  private pageItemCount: number = 6;

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldSPFolderPickerHostProps) {
    super(props);
    //Bind the current object to the external called methods
    this.onBrowseClick = this.onBrowseClick.bind(this);
    this.onDismiss = this.onDismiss.bind(this);
    this.onRenderCell = this.onRenderCell.bind(this);
    this.onClickNext = this.onClickNext.bind(this);
    this.onClickPrevious = this.onClickPrevious.bind(this);
    this.onClickLink = this.onClickLink.bind(this);
    this.onClickParent = this.onClickParent.bind(this);
    this.onFolderChecked = this.onFolderChecked.bind(this);
    this.onClickSelect = this.onClickSelect.bind(this);
    this.onClearSelectionClick = this.onClearSelectionClick.bind(this);

    //Inits the intial folders
    var initialFolder: string;
    var currentSPFolder: string = '';
    if (props.baseFolder != null)
      currentSPFolder = props.baseFolder;
    if (props.initialFolder != null && props.initialFolder != '') {
      initialFolder = props.initialFolder;
      currentSPFolder = this.getParentFolder(initialFolder);
    }
    //Inits the state
    this.state = {
      isOpen: false,
      loading: true,
      currentSPFolder: currentSPFolder,
      confirmFolder: initialFolder,
      selectedFolder: initialFolder,
      childrenFolders: { value: [] }
    };
  }

  /**
   * @function
   * Function called when the user wants to browse folders
   */
  private onBrowseClick(): void {
    this.currentPage = 0;
    this.LoadChildrenFolders();
  }

  /**
   * @function
   * Function called when the user erase the current selection
   */
  private onClearSelectionClick(): void {
    this.state.confirmFolder = '';
    this.state.currentSPFolder = '';
    if (this.props.baseFolder != null)
      this.state.currentSPFolder = this.props.baseFolder;
    this.currentPage = 0;
    this.setState({ isOpen: false, loading: true, selectedFolder: this.state.selectedFolder, currentSPFolder: this.state.currentSPFolder, childrenFolders: this.state.childrenFolders });
    if (this.props.onPropertyChange) {
      this.props.onPropertyChange(this.props.targetProperty, this.state.confirmFolder);
    }
  }

  /**
   * @function
   * Loads the sub folders from the current
   */
  private LoadChildrenFolders(): void {
    //Loading
    this.state.childrenFolders = { value: [] };
    this.setState({ isOpen: true, loading: true, selectedFolder: this.state.selectedFolder, currentSPFolder: this.state.currentSPFolder, childrenFolders: this.state.childrenFolders });
    //Inits the service
    var folderService: SPFolderPickerService = new SPFolderPickerService(this.props.context);
    folderService.getFolders(this.state.currentSPFolder, this.currentPage, this.pageItemCount).then((response: ISPFolders) => {
      //Binds the results
      this.state.childrenFolders = response;
      this.setState({ isOpen: true, loading: false, selectedFolder: this.state.selectedFolder, currentSPFolder: this.state.currentSPFolder, childrenFolders: this.state.childrenFolders });
    });
  }

  /**
   * @function
   * User clicks on the previous button
   */
   private onClickPrevious(): void {
     this.currentPage = this.currentPage - 1;
     this.state.selectedFolder = '';
     if (this.currentPage < 0)
      this.currentPage = 0;
     this.LoadChildrenFolders();
  }

  /**
   * @function
   * User clicks on the next button
   */
  private onClickNext(): void {
    this.state.selectedFolder = '';
    this.currentPage = this.currentPage + 1;
    this.LoadChildrenFolders();
  }

  /**
   * @function
   * User clicks on a sub folder
   */
  private onClickLink(element?: any): void {
    this.currentPage = 0;
    this.state.selectedFolder = '';
    this.state.currentSPFolder = element.currentTarget.value;
    this.LoadChildrenFolders();
  }

  /**
   * @function
   * User clicks on the go-to parent button
   */
  private onClickParent(): void {
    var parentFolder: string = this.getParentFolder(this.state.currentSPFolder);
    if (parentFolder == this.props.context.pageContext.web.serverRelativeUrl)
      parentFolder = '';
    this.currentPage = 0;
    this.state.selectedFolder = '';
    this.state.currentSPFolder = parentFolder;
    this.LoadChildrenFolders();
  }

  /**
   * @function
   * Gets the parent folder server relative url from a folder url
   */
  private getParentFolder(folderUrl: string): string {
    var splitted = folderUrl.split('/');
    var parentFolder: string = '';
    for (var i = 0; i < splitted.length -1; i++) {
      var node: string = splitted[i];
      if (node != null && node != '') {
        parentFolder += '/';
        parentFolder += splitted[i];
      }
    }
    return parentFolder;
  }

  /**
   * @function
   * Occurs when the selected folder changed
   */
  private onFolderChecked(element?: any): void {
    this.state.selectedFolder = element.currentTarget.value;
    this.setState({ isOpen: true, loading: false, selectedFolder: this.state.selectedFolder, currentSPFolder: this.state.currentSPFolder, childrenFolders: this.state.childrenFolders });
  }

  /**
   * @function
   * User clicks on Select button
   */
  private onClickSelect(): void {
    this.state.confirmFolder = this.state.selectedFolder;
    this.setState({ isOpen: false, loading: false, selectedFolder: this.state.selectedFolder,
      confirmFolder: this.state.selectedFolder,
      currentSPFolder: this.state.currentSPFolder,
      childrenFolders: this.state.childrenFolders });
    if (this.props.onPropertyChange) {
      this.props.onPropertyChange(this.props.targetProperty, this.state.confirmFolder);
    }
  }

  /**
   * @function
   * User close the dialog wihout saving
   */
  private onDismiss(ev?: React.MouseEvent): any {
    this.setState({ isOpen: false, loading: false, selectedFolder: this.state.selectedFolder, currentSPFolder: this.state.currentSPFolder, childrenFolders: this.state.childrenFolders });
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    var currentFolderisRoot: boolean = false;
    if (this.state.currentSPFolder == null || this.state.currentSPFolder == '' || this.state.currentSPFolder == this.props.baseFolder)
      currentFolderisRoot = true;

    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <div style={{display:'flex'}}>
          <TextField style={{width:'220px'}} readOnly={true} value={this.state.confirmFolder} />
          <Button buttonType={ButtonType.icon} icon="FolderSearch" onClick={this.onBrowseClick} />
          <Button buttonType={ButtonType.icon} icon="Delete" onClick={this.onClearSelectionClick} />
        </div>

        <Dialog type={DialogType.close} title={strings.SPFolderPickerDialogTitle} isOpen={this.state.isOpen} isDarkOverlay={true} isBlocking={false} onDismiss={this.onDismiss}>

            <div style={{ height: '330px'}}>
                { this.state.loading ? <div><Spinner type={ SpinnerType.normal } /></div> : null }

                { this.state.loading === false && currentFolderisRoot === false ? <Button buttonType={ButtonType.icon} onClick={this.onClickParent} icon="Reply">...</Button> : null }

                <List items={this.state.childrenFolders.value}  onRenderCell={this.onRenderCell} />
                { this.state.loading === false ?
                <Button buttonType={ButtonType.icon} icon="CaretLeft8" onClick={this.onClickPrevious}
                  disabled={ this.currentPage > 0 ? false : true }
                  />
                : null }
                { this.state.loading === false ?
                <Button buttonType={ButtonType.icon} icon="CaretRight8" onClick={this.onClickNext}
                  disabled={ this.state.childrenFolders.value.length < this.pageItemCount ? true : false }
                   />
                : null }
            </div>

            <div style={{marginTop: '20px'}}>

              <Button buttonType={ButtonType.primary} disabled={this.state.selectedFolder != null && this.state.selectedFolder != '' ? false : true }
                onClick={this.onClickSelect}>{strings.SPFolderPickerSelectButton}</Button>
              <Button buttonType={ButtonType.normal} onClick={this.onDismiss}>{strings.SPFolderPickerCancelButton}</Button>
            </div>

        </Dialog>
      </div>
    );
  }

  /**
   * @function
   * Renders a list cell
   */
  private onRenderCell(item?: any, index?: number): React.ReactNode {
    var idUnique: string = 'radio-' + item.ServerRelativeUrl;
    return (
      <div style={{fontSize: '14px', padding: '4px'}}>
        <div className="ms-ChoiceField">
          <input id={idUnique} style={{width: '18px', height: '18px'}}
            defaultChecked={item.ServerRelativeUrl === this.state.confirmFolder ? true: false}
            onChange={this.onFolderChecked} type="radio" name="radio1" value={item.ServerRelativeUrl}/>
          <label htmlFor={idUnique} >
            <span className="ms-Label">
              <i className="ms-Icon ms-Icon--FolderFill" style={{color: '#0062AF', fontSize: '22px'}}></i>
              <span style={{paddingLeft: '5px'}}>
                <button style={{paddingBottom: '0', height: '27px'}} className="ms-Button ms-Button--command" value={item.ServerRelativeUrl} onClick={this.onClickLink}>
                  <span className="ms-Button-label">
                    {item.Name}
                  </span>
                </button>
              </span>
            </span>
          </label>
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
