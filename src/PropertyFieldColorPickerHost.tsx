/**
 * @file PropertyFieldColorPickerHost.tsx
 * Renders the controls for PropertyFieldColorPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldColorPickerPropsInternal } from './PropertyFieldColorPicker';
import { ColorPicker } from 'office-ui-fabric-react/lib/ColorPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';

/**
 * @interface
 * PropertyFieldColorPickerHost properties interface
 *
 */
export interface IPropertyFieldColorPickerHostProps extends IPropertyFieldColorPickerPropsInternal {
}

export interface IPropertyFieldColorPickerHostState {
  color?: string;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldColorPicker component
 */
export default class PropertyFieldColorPickerHost extends React.Component<IPropertyFieldColorPickerHostProps, IPropertyFieldColorPickerHostState> {

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldColorPickerHostProps) {
    super(props);

    //Inits state
    var defaultColor: string = '#FFFFFF';
    if (this.props.initialColor && this.props.initialColor != '')
      defaultColor = this.props.initialColor;
    this.state = {
        color: defaultColor,
        errorMessage: ''
    };

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
   * Renders the control
   */
  public render(): JSX.Element {
    //Renders content
    return (
      <div>
        <Label>{this.props.label}</Label>
        <ColorPicker
          color={this.state.color}
          onColorChanged={this.onColorChanged}
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
}