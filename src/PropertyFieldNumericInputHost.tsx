/**
 * @file PropertyFieldNumericInputHost.tsx
 * Renders the controls for PropertyFieldNumericInput component
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldNumericInputPropsInternal } from './PropertyFieldNumericInput';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
var NumericInput: any = require('react-numeric-input');

/**
 * @interface
 * PropertyFieldNumericInputHost properties interface
 *
 */
export interface IPropertyFieldNumericInputHostProps extends IPropertyFieldNumericInputPropsInternal {
}

export interface IPropertyFieldNumericInputState {
  currentValue?: number;
  errorMessage: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldNumericInput component
 */
export default class PropertyFieldNumericInputHost extends React.Component<IPropertyFieldNumericInputHostProps, IPropertyFieldNumericInputState> {

  private async: Async;
  private delayedValidate: (value: number) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldNumericInputHostProps) {
    super(props);

    this.async = new Async(this);
    this.state = ({ errorMessage: '', currentValue: this.props.initialValue} as IPropertyFieldNumericInputState);

    //Bind the current object to the external called onSelectDate method
    this.onValueChanged = this.onValueChanged.bind(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  /**
   * @function
   * Function called when the component value changed
   */
  private onValueChanged(value: number): void {
    this.state.currentValue = value;
    this.setState(this.state);
    this.delayedValidate(value);
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: number): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialValue, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || 0);
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialValue, value);
        this.setState({ errorMessage: result} as IPropertyFieldNumericInputState);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValue, value);
          this.setState({ errorMessage } as IPropertyFieldNumericInputState);
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
  private notifyAfterValidate(oldValue: number, newValue: number) {
    this.props.properties[this.props.targetProperty] = newValue;
    this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
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
   * Renders the controls
   */
  public render(): JSX.Element {
    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>
        <NumericInput
          className="ms-TextField-field"
          size={this.props.size}
          disabled={this.props.disabled}
          onChange={this.onValueChanged}
          min={this.props.min}
          max={this.props.max}
          value={this.state.currentValue}
          step={this.props.step}
          precision={this.props.precision}
        />
        { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
          <div><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
          <span>
            <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
          </span>
          </div>
        : ''}
      </div>
    );
  }
}
