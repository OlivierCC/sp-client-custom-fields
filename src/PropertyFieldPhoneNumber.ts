/**
 * @file PropertyFieldPhoneNumber.ts
 * Define a custom field of type PropertyFieldPhoneNumber for
 * the SharePoint Framework (SPfx)
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  IPropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import PropertyFieldPhoneNumberHost, { IPropertyFieldPhoneNumberHostProps } from './PropertyFieldPhoneNumberHost';

export enum IPhoneNumberFormat {
  UnitedStates = 0,
  UK = 1,
  France = 2,
  Mexico = 3,
  Australia = 4,
  Denmark = 6,
  Iceland = 7,
  Canada = 8,
  Quebec = 9,
  NorwayLandLine = 10,
  NorwayMobile = 11,
  Portugal = 12,
  PolandLandLine = 13,
  PolandMobile = 14,
  Spain = 15,
  Switzerland = 16,
  Turkey = 17,
  Russian = 18,
  Germany = 19,
  BelgiumLandLine = 20,
  BelgiumMobile = 21,
  Pakistan = 22,
  IndiaLandLine = 23,
  IndiaMobile = 24,
  ChinaLandLine = 25,
  ChinaMobile = 26,
  HongKong = 27,
  Japan = 28,
  Malaysia = 29,
  Philippines = 30,
  Singapore = 31,
  TaiwanLandLine = 32,
  TaiwanMobile = 33,
  SouthKoreaMobile = 34,
  NewZealand = 35,
  CostaRica = 36,
  ElSalvador = 37,
  Guatemala = 38,
  HondurasLandLine = 39,
  HondurasMobile = 40,
  BrazilLandLine = 41,
  BrazilMobile = 42,
  Peru = 43
}

/**
 * @interface
 * Public properties of the PropertyFieldPhoneNumber custom field
 *
 */
export interface IPropertyFieldPhoneNumberProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Initial value
   */
  initialValue?: string;
  /**
   * @var
   * Phone number format
   */
  phoneNumberFormat?: IPhoneNumberFormat;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected Color changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    /**
   * @var
   * Parent Web Part properties
   */
  properties: any;
}

/**
 * @interface
 * Private properties of the PropertyFieldPhoneNumber custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldPhoneNumber.
 *
 */
export interface IPropertyFieldPhoneNumberPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: string;
  phoneNumberFormat?: IPhoneNumberFormat;
  targetProperty: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
}

/**
 * @interface
 * Represents a PropertyFieldPhoneNumber object
 *
 */
class PropertyFieldPhoneNumberBuilder implements IPropertyPaneField<IPropertyFieldPhoneNumberPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: IPropertyPaneFieldType = 1;//IPropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldPhoneNumberPropsInternal;

  //Custom properties
  private label: string;
  private initialValue: string;
  private phoneNumberFormat: IPhoneNumberFormat;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldPhoneNumberPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.phoneNumberFormat = _properties.phoneNumberFormat;
    this.initialValue = _properties.initialValue;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.customProperties = _properties.properties;
  }

  /**
   * @function
   * Renders the ColorPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldPhoneNumberHostProps> = React.createElement(PropertyFieldPhoneNumberHost, {
      label: this.label,
      initialValue: this.initialValue,
      phoneNumberFormat: this.phoneNumberFormat,
      targetProperty: this.targetProperty,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties
    });
    //Calls the REACT content generator
    ReactDom.render(element, elem);
  }

  /**
   * @function
   * Disposes the current object
   */
  private dispose(elem: HTMLElement): void {

  }

}

/**
 * @function
 * Helper method to create a Color Picker on the PropertyPane.
 * @param targetProperty - Target property the Color picker is associated to.
 * @param properties - Strongly typed Color Picker properties.
 */
export function PropertyFieldPhoneNumber(targetProperty: string, properties: IPropertyFieldPhoneNumberProps): IPropertyPaneField<IPropertyFieldPhoneNumberPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldPhoneNumberPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      phoneNumberFormat: properties.phoneNumberFormat,
      initialValue: properties.initialValue,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldPhoneNumber builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldPhoneNumberBuilder(targetProperty, newProperties);
}


