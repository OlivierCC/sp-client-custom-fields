
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

var MapComponent    = require('react-cartographer/lib/components/Map');

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
}

/**
 * @class
 * Renders the controls for PropertyFieldMapPicker component
 */
export default class PropertyFieldMapPickerHost extends React.Component<IPropertyFieldMapPickerHostProps, IPropertyFieldMapPickerHostState> {

  /**
   * @function
   * Contructor
   */
  constructor(props: IPropertyFieldMapPickerHostProps) {
    super(props);
    //Bind the current object to the external called onSelectDate method
    this.onValueChanged = this.onValueChanged.bind(this);
    this.onClickChevron = this.onClickChevron.bind(this);
    this.onLongitudeChange = this.onLongitudeChange.bind(this);
    this.onLatitudeChange = this.onLatitudeChange.bind(this);
    this.onGetCurrentLocation = this.onGetCurrentLocation.bind(this);
    this.showPosition = this.showPosition.bind(this);

    this.state = {
      longitude: this.props.longitude,
      latitude: this.props.latitude,
      isOpen: true
    };
    this.setState(this.state);
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

     if (this.props.onPropertyChange) {
        this.props.properties[this.props.targetProperty] = this.state.longitude + ',' + this.state.latitude;
        this.props.onPropertyChange(this.props.targetProperty, this.props.initialValue, this.state.longitude + ',' + this.state.latitude);
      }
  }

  /**
   * @function
   * Function called when the ColorPicker Office UI Fabric component selected color changed
   */
  private onValueChanged(element: any): void {
    //Checks if there is a method to called
    if (this.props.onPropertyChange && element != null) {
      this.props.properties[this.props.targetProperty] = element.currentTarget.value;
      this.props.onPropertyChange(this.props.targetProperty, this.props.initialValue, element.currentTarget.value);
    }
  }

  private onLongitudeChange(element: any): void {
      var value = element.currentTarget.value;
      this.state.longitude = value;
      this.setState(this.state);

      if (this.props.onPropertyChange && element != null) {
        this.props.properties[this.props.targetProperty] = this.state.longitude + ',' + this.state.latitude;
        this.props.onPropertyChange(this.props.targetProperty, this.props.initialValue, this.state.longitude + ',' + this.state.latitude);
      }
  }

   private onLatitudeChange(element: any): void {
    var value = element.currentTarget.value;
      this.state.latitude = value;
      this.setState(this.state);

      if (this.props.onPropertyChange && element != null) {
        this.props.properties[this.props.targetProperty] = this.state.longitude + ',' + this.state.latitude;
        this.props.onPropertyChange(this.props.targetProperty, this.props.initialValue, this.state.longitude + ',' + this.state.latitude);
      }
  }

  /**
   * @function
   * Renders the datepicker controls with Office UI  Fabric
   */
  public render(): JSX.Element {

    if (this.state.isOpen == true) {

    //Renders content
    return (
      <div style={{ marginBottom: '8px'}}>
        <Label>{this.props.label}</Label>

        <div style={{paddingTop: '10px'}}>

        <div style={{width:'90px', float: 'left', paddingRight: '10px'}}>
          <span style={{paddingBottom:'6px', display:'block', fontFamily: '"Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif',fontSize: '12px', fontWeight: '400'}}>
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
    backgroundClip: 'padding-box'}} value={this.state.longitude}  onChange={this.onLongitudeChange} />
        </div>

        <div style={{width:'90px', display: 'inline' }}>
          <span style={{paddingBottom:'6px', display:'block', fontFamily: '"Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif',fontSize: '12px', fontWeight: '400'}}>
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
    backgroundClip: 'padding-box'}} value={this.state.latitude} onChange={this.onLatitudeChange}/>
        </div>
        <div style={{width:'80px', float: 'right',top: '-30px', position: 'relative' }}>
          <div style={{float: 'left' }}><Button buttonType={ButtonType.icon} icon="MapPin" onClick={this.onGetCurrentLocation}  /></div>
          <div style={{display:'inline'}}>
          <Button buttonType={ButtonType.icon} icon="ChevronUpSmall"  onClick={this.onClickChevron}  /></div>
        </div>
        </div>
        <div style={{position: 'relative', top: '0px', paddingBottom: '30px'}}>
          <MapComponent
              provider='bing'
              providerKey='Ag3-9ixwWbFk4BdNzkj6MCnFN2_pQiL2hedXxiiuaF_DSuzDqAVp2mW9wPE0coeL'
              mapId='map'
              latitude={this.state.latitude}
              longitude={this.state.longitude}
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
          <span style={{paddingBottom:'6px', display:'block', fontFamily: '"Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif',fontSize: '12px', fontWeight: '400'}}>
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
    backgroundClip: 'padding-box'}} value={this.state.longitude} onChange={this.onLongitudeChange} />
        </div>
          <div style={{width:'90px', display: 'inline' }}>
          <span style={{paddingBottom:'6px', display:'block', fontFamily: '"Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif',fontSize: '12px', fontWeight: '400'}}>
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
    backgroundClip: 'padding-box'}} value={this.state.latitude}  onChange={this.onLatitudeChange} />
        </div>
        <div style={{width:'80px', float: 'right',top: '-30px', position: 'relative' }}>
          <div style={{float: 'left' }}><Button buttonType={ButtonType.icon} icon="MapPin" onClick={this.onGetCurrentLocation} /></div>
          <div style={{display:'inline'}}>
          <Button buttonType={ButtonType.icon} icon="ChevronDownSmall"  onClick={this.onClickChevron}  /></div>
        </div>
        </div>

      </div>
    );
    }
  }
}
