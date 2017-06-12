/**
 * @file PropertyFieldDimensionPickerHost.tsx
 * Renders the controls for PropertyFieldDimensionPicker component
 *
 * @copyright 2017 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldDimensionPickerPropsInternal, IPropertyFieldDimension } from './PropertyFieldDimensionPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Async } from 'office-ui-fabric-react/lib/Utilities';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import GuidHelper from './GuidHelper';

import * as strings from 'sp-client-custom-fields/strings';

/**
 * @interface
 * PropertyFieldDimensionPickerHost properties interface
 *
 */
export interface IPropertyFieldDimensionPickerHostProps extends IPropertyFieldDimensionPickerPropsInternal {
}

export interface IPropertyFieldDimensionPickerState {
  width?: number;
  height?: number;
  widthUnit?: string;
  heightUnit?: string;
  conserveRatio?: boolean;
  errorMessage: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldDimensionPicker component
 */
export default class PropertyFieldDimensionPickerHost extends React.Component<IPropertyFieldDimensionPickerHostProps, IPropertyFieldDimensionPickerState> {

  private async: Async;
  private delayedValidate: (value: IPropertyFieldDimension) => void;
  private _key: string;
  private units: IDropdownOption[] = [
    { key: 'px', text: 'px'},
    { key: '%', text: '%'}
  ];

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldDimensionPickerHostProps) {
    super(props);

    this._key = GuidHelper.getGuid();
    this.async = new Async(this);

    this.state = ({
      errorMessage: '',
      width: 0,
      height: 0,
      widthUnit: 'px',
      heightUnit: 'px',
      conserveRatio: this.props.preserveRatio
    } as IPropertyFieldDimensionPickerState);
    this.loadDefaultData();

    //Bind the current object to the external called onSelectDate method
    this.onWidthChanged = this.onWidthChanged.bind(this);
    this.onHeightChanged = this.onHeightChanged.bind(this);
    this.onWidthUnitChanged = this.onWidthUnitChanged.bind(this);
    this.onHeightUnitChanged = this.onHeightUnitChanged.bind(this);
    this.onRatioChanged = this.onRatioChanged.bind(this);
    this.saveDimension = this.saveDimension.bind(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  /**
   * @function
   * Function called to load data from the initialValue
   */
  private loadDefaultData(): void {
    if (this.props.initialValue != null && this.props.initialValue !== undefined) {
      if (this.props.initialValue.width != null && this.props.initialValue.width !== undefined) {
        if (this.props.initialValue.width.indexOf('px') > -1) {
          this.state.widthUnit = 'px';
          this.state.width = Math.round(+this.props.initialValue.width.replace('px', ''));
        }
        else if (this.props.initialValue.height.indexOf('%') > -1) {
          this.state.widthUnit = '%';
          this.state.width = Math.round(+this.props.initialValue.width.replace('%', ''));
        }
      }
      if (this.props.initialValue.height != null && this.props.initialValue.height !== undefined) {
        if (this.props.initialValue.height.indexOf('px') > -1) {
          this.state.heightUnit = 'px';
          this.state.height = Math.round(+this.props.initialValue.height.replace('px', ''));
        }
        else if (this.props.initialValue.height.indexOf('%') > -1) {
          this.state.heightUnit = '%';
          this.state.height = Math.round(+this.props.initialValue.height.replace('%', ''));
        }
      }
    }
  }

  /**
   * @function
   * Function called when the width changed
   */
  private onWidthChanged(newValue: any): void {
    if (this.state.widthUnit === this.state.heightUnit && this.state.conserveRatio === true && this.props.preserveRatioEnabled === true) {
      if (this.state.width != 0)
        this.state.height = Math.round((this.state.height / this.state.width) * +newValue);
    }
    this.state.width = Math.round(+newValue);
    this.setState(this.state);
    this.saveDimension();
  }

  /**
   * @function
   * Function called when the height changed
   */
  private onHeightChanged(newValue: any): void {
    if (this.state.widthUnit === this.state.heightUnit && this.state.conserveRatio === true && this.props.preserveRatioEnabled === true) {
      if (this.state.height != 0)
        this.state.width = Math.round((this.state.width / this.state.height) * +newValue);
    }
    this.state.height = Math.round(+newValue);
    this.setState(this.state);
    this.saveDimension();
  }

  /**
   * @function
   * Function called when the width unit changed
   */
  private onWidthUnitChanged(element?: IDropdownOption): void {
    if (element != null) {
      var newValue: string = element.key.toString();
      this.state.widthUnit = newValue;
      this.setState(this.state);
      this.saveDimension();
    }
  }

  /**
   * @function
   * Function called when the height unit changed
   */
  private onHeightUnitChanged(element?: IDropdownOption): void {
    if (element != null) {
      var newValue: string = element.key.toString();
      this.state.heightUnit = newValue;
      this.setState(this.state);
      this.saveDimension();
    }
  }

  /**
   * @function
   * Function called when the ratio changed
   */
  private onRatioChanged(element: any, isChecked: boolean): void {
     if (element) {
      this.state.conserveRatio = isChecked;
      this.setState(this.state);
     }
  }

  /**
   * @function
   * Saves the dimension
   */
  private saveDimension(): void {
    var dimension: IPropertyFieldDimension = {
      width: this.state.width + this.state.widthUnit,
      height: this.state.height + this.state.heightUnit
    };
    this.delayedValidate(dimension);
  }

  /**
   * @function
   * Validates the new custom field value
   */
  private validate(value: IPropertyFieldDimension): void {
    if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
      this.notifyAfterValidate(this.props.initialValue, value);
      return;
    }

    var result: string | PromiseLike<string> = this.props.onGetErrorMessage(value || '');
    if (result !== undefined) {
      if (typeof result === 'string') {
        if (result === undefined || result === '')
          this.notifyAfterValidate(this.props.initialValue, value);
        this.setState({ errorMessage: result} as IPropertyFieldDimensionPickerState);
      }
      else {
        result.then((errorMessage: string) => {
          if (errorMessage === undefined || errorMessage === '')
            this.notifyAfterValidate(this.props.initialValue, value);
          this.setState({ errorMessage } as IPropertyFieldDimensionPickerState);
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
  private notifyAfterValidate(oldValue: IPropertyFieldDimension, newValue: IPropertyFieldDimension) {
    this.props.properties[this.props.targetProperty] = newValue;
    this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
  }

  /**
   * @function
   * Called when the component will unmount
   */
  public componentWillUnmount() {
    if (this.async != null)
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

        <table style={{paddingTop: '10px'}}>
          <tbody>
            <tr>
              <td style={{verticalAlign: 'top', minWidth: '55px'}}>
                <Label disabled={this.props.disabled}>{strings.DimensionWidth}</Label>
              </td>
              <td style={{verticalAlign: 'top', width: '80px'}}>
                <TextField
                  disabled={this.props.disabled}
                  role="textbox"
                  aria-multiline="false"
                  type="number"
                  min='0'
                  value={this.state.width !== undefined ? this.state.width.toString():''}
                  onChanged={this.onWidthChanged}
                />
              </td>
              <td style={{verticalAlign: 'top'}}>
                <Dropdown
                  label="" options={this.units} selectedKey={this.state.widthUnit}
                  disabled={this.props.disabled}
                  onChanged={this.onWidthUnitChanged}
                />
              </td>
            </tr>
            <tr>
              <td style={{verticalAlign: 'top', minWidth: '55px'}}>
                <Label disabled={this.props.disabled}>{strings.DimensionHeight}</Label>
              </td>
              <td style={{verticalAlign: 'top', width: '80px'}}>
                <TextField
                  disabled={this.props.disabled}
                  role="textbox"
                  aria-multiline="false"
                  type="number"
                  min='0'
                  value={this.state.height !== undefined ? this.state.height.toString():''}
                  onChanged={this.onHeightChanged}
                />
              </td>
              <td style={{verticalAlign: 'top'}}>
                <Dropdown
                  label="" options={this.units} selectedKey={this.state.heightUnit}
                  disabled={this.props.disabled}
                  onChanged={this.onHeightUnitChanged}
                />
              </td>
            </tr>
            { this.props.preserveRatioEnabled === true ?
            <tr>
              <td></td>
              <td colSpan={2}>
                <div className="ms-ChoiceField" style={{paddingLeft: '0px'}}>
                  <Checkbox
                    checked={this.state.conserveRatio}
                    disabled={this.props.disabled}
                    label={strings.DimensionRatio}
                    onChange={this.onRatioChanged}
                  />
                </div>
              </td>
            </tr>
            : ''}
          </tbody>
        </table>

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
