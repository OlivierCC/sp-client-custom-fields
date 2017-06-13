/**
 * @file PropertyFieldColorPickerMiniHost.tsx
 * Renders the controls for PropertyFieldColorPickerMini component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldColorPickerMiniPropsInternal } from './PropertyFieldColorPickerMini';
import { ColorPicker } from 'office-ui-fabric-react/lib/ColorPicker';
import { Button } from 'office-ui-fabric-react/lib/Button';
import { Callout } from 'office-ui-fabric-react/lib/Callout';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';

/**
 * @interface
 * PropertyFieldColorPickerMiniHost properties interface
 *
 */
export interface IPropertyFieldColorPickerMiniHostProps extends IPropertyFieldColorPickerMiniPropsInternal {
}

export interface IPropertyFieldColorPickerMiniHostState {
  color?: string;
  calloutVisible: boolean;
  isHover: boolean;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldColorPickerMini component
 */
export default class PropertyFieldColorPickerMiniHost extends React.Component<IPropertyFieldColorPickerMiniHostProps, IPropertyFieldColorPickerMiniHostState> {

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;
  private menuButtonElement: HTMLElement;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldColorPickerMiniHostProps) {
    super(props);

    //Inits state
    var defaultColor: string = '#FFFFFF';
    if (this.props.initialColor && this.props.initialColor != '')
      defaultColor = this.props.initialColor;
    this.state = {
        color: defaultColor,
        calloutVisible: false,
        isHover: false,
        errorMessage: ''
    };

    this.onClickButton = this.onClickButton.bind(this);
    this.onMouseEnterButton = this.onMouseEnterButton.bind(this);
    this.onMouseLeaveButton = this.onMouseLeaveButton.bind(this);
    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);

    //Bind the current object to the external called onSelectDate method
    this.onColorChanged = this.onColorChanged.bind(this);
  }

  /**
   * @function
   * Function called when the ColorPicker Office UI Fabric component selected color changed
   */
  private onColorChanged(color: string): void {
    this.state.color = color;
    this.setState(this.state);
    this.delayedValidate(color);
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: string): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialColor, value);
      return;
    }

    if (this.latestValidateValue === value)
      return;
    this.latestValidateValue = value;

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialColor, value);
        this.state.errorMessage = result;
        this.setState(this.state);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialColor, value);
          this.state.errorMessage = errorMessage;
          this.setState(this.state);
        });
      }
    }
    else {
      this.notifyAfterValidate(this.props.initialColor, value);
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
   * Called when the color button is clicked
   */
  private onClickButton(): void {
    if (this.props.disabled === true)
      return;
    this.state.calloutVisible = !this.state.calloutVisible;
    this.setState(this.state);
  }

  private onMouseEnterButton(): void {
    if (this.props.disabled === true)
      return;
    this.state.isHover = true;
    this.setState(this.state);
  }

  private onMouseLeaveButton(): void {
    if (this.props.disabled === true)
      return;
    this.state.isHover = false;
    this.setState(this.state);
  }

  /**
   * @function
   * Renders the control
   */
  public render(): JSX.Element {
    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <div ref={ (menuButton) => this.menuButtonElement = menuButton }
          style={{width: '70px', height: '35px',
          backgroundColor: this.state.isHover ? '#eaeaea' : '#F4F4F4',
          padding: '6px',
          cursor: this.props.disabled === true ? 'default' : 'pointer',
          borderBottomLeftRadius: '5px', borderBottomRightRadius: '5px', borderTopLeftRadius: '5px', borderTopRightRadius: '5px'}}
          onClick={this.onClickButton}
          onMouseEnter={this.onMouseEnterButton}
          onMouseLeave={this.onMouseLeaveButton}
          role="button">
          <div style={{ width: '100%', height: '100%', backgroundColor: this.state.color}}>
          </div>
        </div>
        { this.state.calloutVisible && (
          <Callout
              className='ms-CalloutExample-callout'
              gapSpace={ 0 }
              targetElement={ this.menuButtonElement }
              setInitialFocus={ true }
              onDismiss={this.onClickButton}
            >
            <ColorPicker
              color={this.state.color}
              onColorChanged={this.onColorChanged}
            />
          </Callout>
        )}
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