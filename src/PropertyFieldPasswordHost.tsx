/**
 * @file PropertyFieldPasswordHost.tsx
 * Renders the controls for PropertyFieldPassword component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldPasswordPropsInternal } from './PropertyFieldPassword';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import { TextField } from 'office-ui-fabric-react/lib/TextField';

/**
 * @interface
 * PropertyFieldPasswordHost properties interface
 *
 */
export interface IPropertyFieldPasswordHostProps extends IPropertyFieldPasswordPropsInternal {
}

export interface IPropertyFieldPasswordState {
  currentValue?: string;
  errorMessage: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldPassword component
 */
export default class PropertyFieldPasswordHost extends React.Component<IPropertyFieldPasswordHostProps, IPropertyFieldPasswordState> {

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldPasswordHostProps) {
    super(props);

    this.async = new Async(this);
    this.state = ({ errorMessage: '', currentValue: this.props.initialValue} as IPropertyFieldPasswordState);

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
  private onValueChanged(newValue: any): void {
    this.state.currentValue = newValue;
    this.setState(this.state);
    this.delayedValidate(newValue);
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
        this.setState({ errorMessage: result} as IPropertyFieldPasswordState);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValue, value);
          this.setState({ errorMessage } as IPropertyFieldPasswordState);
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
        <TextField
          disabled={this.props.disabled}
          aria-multiline="false"
          placeholder={this.props.placeHolder !== undefined ? this.props.placeHolder: ''}
          type="password"
          value={this.state.currentValue !== undefined ? this.state.currentValue.toString():''}
          onChanged={this.onValueChanged}
          aria-invalid={ !!this.state.errorMessage }
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
