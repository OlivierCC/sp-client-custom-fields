/**
 * @file PropertyFieldOfficeVideoPickerHost.tsx
 * Renders the controls for PropertyFieldOfficeVideoPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldOfficeVideoPickerPropsInternal } from './PropertyFieldOfficeVideoPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import { Spinner, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import GuidHelper from './GuidHelper';
import { TextField } from 'office-ui-fabric-react/lib/TextField';

/**
 * @interface
 * PropertyFieldOfficeVideoPickerHost properties interface
 *
 */
export interface IPropertyFieldOfficeVideoPickerHostProps extends IPropertyFieldOfficeVideoPickerPropsInternal {
}

export interface IPropertyFieldOfficeVideoPickerHostState {
  openPanel?: boolean;
  openRecent?: boolean;
  openSite?: boolean;
  openUpload?: boolean;
  recentImages?: string[];
  selectedVideo: string;
  errorMessage?: string;
  iframeLoaded: boolean;
}

/**
 * @class
 * Renders the controls for PropertyFieldOfficeVideoPicker component
 */
export default class PropertyFieldOfficeVideoPickerHost extends React.Component<IPropertyFieldOfficeVideoPickerHostProps, IPropertyFieldOfficeVideoPickerHostState> {

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;
  private guid: string;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldOfficeVideoPickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onTextFieldChanged = this.onTextFieldChanged.bind(this);
    this.onOpenPanel = this.onOpenPanel.bind(this);
    this.onClosePanel = this.onClosePanel.bind(this);
    this.onEraseButton = this.onEraseButton.bind(this);
    this.iFrameLoaded = this.iFrameLoaded.bind(this);
    this.iFrameValidation = this.iFrameValidation.bind(this);

    //Inits the state
    this.state = {
      iframeLoaded: false,
      selectedVideo: this.props.initialValue,
      openPanel: false,
      openRecent: false,
      openSite: true,
      openUpload: false,
      recentImages: [],
      errorMessage: ''
    };

    this.guid = GuidHelper.getGuid();
    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }


 /**
  * @function
  * Save the image value
  *
  */
  private saveVideoProperty(imageUrl: string): void {
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
    this.state.selectedVideo = '';
    this.setState(this.state);
    this.saveVideoProperty('');
  }

 /**
  * @function
  * Open the panel
  *
  */
  private onOpenPanel(element?: any): void {
    this.state.openPanel = true;
    this.state.iframeLoaded = false;
    this.setState(this.state);
  }

 /**
  * @function
  * The text field value changed
  *
  */
  private onTextFieldChanged(newValue: string): void {
    this.state.selectedVideo = newValue;
    this.setState(this.state);
    this.saveVideoProperty(newValue);
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

  public componentDidUpdate(prevProps: any, prevState: any, prevContext: any): void {
    var iframe: any = document.getElementById(this.guid);

    if (iframe != null && iframe != undefined) {
      if (iframe.addEventListener)
        iframe.addEventListener("load", this.iFrameLoaded, false);
      else
        iframe.attachEvent("onload", this.iFrameLoaded);
    }
  }

  private iFrameLoaded(): void {
    var okButton = window.frames[this.guid].document.getElementById("ctl00_OkButton");
    okButton.onclick = '';
    okButton.addEventListener("click", this.iFrameValidation, false);
    var cancelButton = window.frames[this.guid].document.getElementById("CancelButton");
    cancelButton.onclick = '';
    cancelButton.addEventListener("click", this.onClosePanel, false);

    this.state.iframeLoaded = true;
    this.setState(this.state);
  }

  private iFrameValidation(): void {
    var dialogResult = window.frames[this.guid].window.dialogResult;
    if (dialogResult == null)
      return;
    if (dialogResult.Url == null) {
      this.onClosePanel();
      return;
    }
    var vidUrl = dialogResult.Url;
    this.state.selectedVideo = vidUrl;
    this.setState(this.state);
    this.saveVideoProperty(vidUrl);
    this.onClosePanel();
  }

  /**
  * @function
  * When component is mount, attach the iframe event watcher
  *
  */
  public componentDidMount() {
  }

  /**
  * @function
  * Releases the watcher
  *
  */
  public componentWillUnmount() {
    if (this.async !== undefined)
      this.async.dispose();
  }


  /**
   * @function
   * Renders the controls
   */
  public render(): JSX.Element {

    var iframeUrl = this.props.context.pageContext.web.absoluteUrl;
    iframeUrl += '/portals/hub/_layouts/15/VideoAssetDialog.aspx?list=&IsDlg=1';

    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>
        <table style={{width: '100%', borderSpacing: 0}}>
          <tbody>
            <tr>
              <td width="*">
                <TextField
                  disabled={this.props.disabled}
                  value={this.state.selectedVideo}
                  style={{width:'100%'}}
                  onChanged={this.onTextFieldChanged}
                  readOnly={this.props.readOnly}
                />
              </td>
              <td width="64">
                <Button disabled={this.props.disabled} buttonType={ButtonType.icon} icon="FolderSearch" onClick={this.onOpenPanel} />
                <Button disabled={this.props.disabled === false && (this.state.selectedVideo != null && this.state.selectedVideo != '') ? false: true} buttonType={ButtonType.icon} icon="Delete" onClick={this.onEraseButton} />
              </td>
            </tr>
          </tbody>
        </table>
        { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
            : ''}

        { this.state.openPanel === true ?

        <Panel
          isOpen={this.state.openPanel} hasCloseButton={true} onDismiss={this.onClosePanel}
          isLightDismiss={true} type={PanelType.large}
          headerText={this.props.panelTitle}>

          <div style={{visibility: this.state.iframeLoaded === false ? 'visible': 'hidden',
            display: this.state.iframeLoaded === false ? 'block': 'none',
            height: this.state.iframeLoaded === false ? 'auto': '0px'}}>
             <Spinner type={ SpinnerType.normal } />
          </div>

          <div id="site" style={{width: '100%', height: '700px'}}>

            <iframe ref="filePickerIFrame" style={{
              width: '100%', borderWidth:'0',
              visibility: this.state.iframeLoaded === true ? 'visible': 'hidden',
              display: this.state.iframeLoaded === true ? 'block': 'none',
              height: this.state.iframeLoaded === true ? '650px': '0px'}}
            role="application"
            src={iframeUrl}
            id={this.guid}
            name={this.guid}></iframe>

          </div>

        </Panel>
        : '' }

      </div>
    );
  }

}
