
/**
 * @file PropertyFieldMapPickerHost.tsx
 * Renders the controls for PropertyFieldMapPicker component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldMapPickerPropsInternal } from './PropertyFieldMapPicker';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Async } from 'office-ui-fabric-react/lib/Utilities';

var MapComponent: any = require('react-cartographer/lib/components/Map');

/**
 * @interface
 * PropertyFieldMapPickerHost properties interface
 *
 */
export interface IPropertyFieldMapPickerHostProps extends IPropertyFieldMapPickerPropsInternal {
}

export interface IPropertyFieldMapPickerHostState {
  longitude: string;
  latitude: string;
  isOpen: boolean;
  errorMessage?: string;
}

/**
 * @class
 * Renders the controls for PropertyFieldMapPicker component
 */
export default class PropertyFieldMapPickerHost extends React.Component<IPropertyFieldMapPickerHostProps, IPropertyFieldMapPickerHostState> {

  private latestValidateValue: string;
  private async: Async;
  private delayedValidate: (value: string) => void;

  /**
   * @function
   * Constructor
   */
  constructor(props: IPropertyFieldMapPickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onClickChevron = this.onClickChevron.bind(this);
    this.onLongitudeChange = this.onLongitudeChange.bind(this);
    this.onLatitudeChange = this.onLatitudeChange.bind(this);
    this.onGetCurrentLocation = this.onGetCurrentLocation.bind(this);
    this.showPosition = this.showPosition.bind(this);

    this.state = {
      longitude: this.props.longitude,
      latitude: this.props.latitude,
      isOpen: true,
      errorMessage: ''
    };

    this.async = new Async(this);
    this.validate = this.validate.bind(this);
    this.notifyAfterValidate = this.notifyAfterValidate.bind(this);
    this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
  }

  private onClickChevron(element: any): void {
    this.state.isOpen = !this.state.isOpen;
    this.setState(this.state);
  }

  private onGetCurrentLocation(element: any): void {
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(this.showPosition);
    }
  }

  private showPosition(position: any): void {
    this.state.latitude = position.coords.latitude;
    this.state.longitude = position.coords.longitude;
    this.setState(this.state);

    var newValue: string = this.state.longitude + ',' + this.state.latitude;
    this.delayedValidate(newValue);
  }

  private onLongitudeChange(element: any): void {
    var value = element.currentTarget.value;
    this.state.longitude = value;
    this.setState(this.state);

    var newValue: string = this.state.longitude + ',' + this.state.latitude;
    this.delayedValidate(newValue);
  }

  private onLatitudeChange(element: any): void {
    var value = element.currentTarget.value;
    this.state.latitude = value;
    this.setState(this.state);

    var newValue: string = this.state.longitude + ',' + this.state.latitude;
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

    if (this.state.isOpen == true) {

    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>

        <div style={{paddingTop: '10px'}}>

        <div style={{width:'90px', float: 'left', paddingRight: '10px'}}>
          <span style={{paddingBottom:'6px', display:'block', fontFamily: '"Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif',fontSize: '12px', fontWeight: 400}}>
          Longitude
          </span>
          <input id="longitude" style={{width:'80px', borderRadius: '0px',
    border: '1px solid rgb(200, 200, 200)',
    display: 'block',
    overflow: 'hidden',
    whiteSpace: 'nowrap',
    position: 'relative',
    height: '26px',
    lineHeight: '26px',
    padding: '0px 0px 0px 8px',
    color: 'rgb(68, 68, 68)',
    textDecoration: 'none',
    backgroundColor: 'rgb(255, 255, 255)',
    backgroundClip: 'padding-box'}} value={this.state.longitude} disabled={this.props.disabled}  onChange={this.onLongitudeChange} />
        </div>

        <div style={{width:'90px', display: 'inline' }}>
          <span style={{paddingBottom:'6px', display:'block', fontFamily: '"Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif',fontSize: '12px', fontWeight: 400}}>
          Latitude
          </span>
          <input id="latitude" style={{width:'80px', borderRadius: '0px',
    border: '1px solid rgb(200, 200, 200)',
    display: 'block',
    overflow: 'hidden',
    whiteSpace: 'nowrap',
    position: 'relative',
    height: '26px',
    lineHeight: '26px',
    padding: '0px 0px 0px 8px',
    color: 'rgb(68, 68, 68)',
    textDecoration: 'none',
    backgroundColor: 'rgb(255, 255, 255)',
    backgroundClip: 'padding-box'}} value={this.state.latitude} onChange={this.onLatitudeChange} disabled={this.props.disabled}/>
        </div>
        <div style={{width:'80px', float: 'right',top: '-30px', position: 'relative' }}>
          <div style={{float: 'left' }}><Button buttonType={ButtonType.icon} icon="MapPin" disabled={this.props.disabled} onClick={this.onGetCurrentLocation}  /></div>
          <div style={{display:'inline'}}>
          <Button buttonType={ButtonType.icon} disabled={this.props.disabled} icon="ChevronUpSmall"  onClick={this.onClickChevron}  /></div>
        </div>
        </div>
        { this.state.errorMessage != null && this.state.errorMessage != '' && this.state.errorMessage != undefined ?
              <div><div aria-live='assertive' className='ms-u-screenReaderOnly' data-automation-id='error-message'>{  this.state.errorMessage }</div>
              <span>
                <p className='ms-TextField-errorMessage ms-u-slideDownIn20'>{ this.state.errorMessage }</p>
              </span>
              </div>
            : ''}
        <div style={{position: 'relative', top: '0px', paddingBottom: '30px'}}>
          <MapComponent
              provider='bing'
              providerKey='Ag3-9ixwWbFk4BdNzkj6MCnFN2_pQiL2hedXxiiuaF_DSuzDqAVp2mW9wPE0coeL'
              mapId='map'
              latitude={+this.state.latitude}
              longitude={+this.state.longitude}
              zoom={15}
              height={250}
              width={283}
              />
        </div>
      </div>
    );

    }
    else {
return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>

        <div style={{paddingTop: '10px'}}>
        <div style={{width:'90px', float: 'left', paddingRight: '10px'}}>
          <span style={{paddingBottom:'6px', display:'block', fontFamily: '"Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif',fontSize: '12px', fontWeight: 400}}>
          Longitude
          </span>
          <input id="longitude" style={{width:'80px', borderRadius: '0px',
    border: '1px solid rgb(200, 200, 200)',
    display: 'block',
    overflow: 'hidden',
    whiteSpace: 'nowrap',
    position: 'relative',
    height: '26px',
    lineHeight: '26px',
    padding: '0px 0px 0px 8px',
    color: 'rgb(68, 68, 68)',
    textDecoration: 'none',
    backgroundColor: 'rgb(255, 255, 255)',
    backgroundClip: 'padding-box'}} value={this.state.longitude} onChange={this.onLongitudeChange} disabled={this.props.disabled} />
        </div>
          <div style={{width:'90px', display: 'inline' }}>
          <span style={{paddingBottom:'6px', display:'block', fontFamily: '"Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif',fontSize: '12px'}}>
          Latitude
          </span>
          <input id="latitude" style={{width:'80px', borderRadius: '0px',
    border: '1px solid rgb(200, 200, 200)',
    display: 'block',
    overflow: 'hidden',
    whiteSpace: 'nowrap',
    position: 'relative',
    height: '26px',
    lineHeight: '26px',
    padding: '0px 0px 0px 8px',
    color: 'rgb(68, 68, 68)',
    textDecoration: 'none',
    backgroundColor: 'rgb(255, 255, 255)',
    backgroundClip: 'padding-box'}} value={this.state.latitude}  onChange={this.onLatitudeChange} disabled={this.props.disabled} />
        </div>
        <div style={{width:'80px', float: 'right',top: '-30px', position: 'relative' }}>
          <div style={{float: 'left' }}><Button buttonType={ButtonType.icon} icon="MapPin" onClick={this.onGetCurrentLocation} /></div>
          <div style={{display:'inline'}}>
          <Button buttonType={ButtonType.icon} icon="ChevronDownSmall"  onClick={this.onClickChevron}  /></div>
        </div>
        </div>
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
}
