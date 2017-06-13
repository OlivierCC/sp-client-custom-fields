/**
 * @file PropertyFieldDateTimePicker.ts
 * Define a custom field of type PropertyFieldDateTimePicker for
 * the SharePoint Framework (SPfx)
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import PropertyFieldDateTimePickerHost, { IPropertyFieldDateTimePickerHostProps } from './PropertyFieldDateTimePickerHost';

/**
  * @enum
  * Time convention
  */
export enum ITimeConvention {
  /**
   * The 12-hour clock is a time convention in which the 24 hours of the day are
   * divided into two periods: a.m. and p.m.
   */
  Hours12 = 0,
  /**
   * The 24-hour clock is the convention of time keeping in which the day runs from midnight to
   * midnight and is divided into 24 hours, indicated by the hours passed since midnight, from 0 to 23
   */
  Hours24 = 1
}

/**
 * @interface
 * Public properties of the PropertyFieldDateTimePicker custom field
 *
 */
export interface IPropertyFieldDateTimePickerProps {
  /**
   * @var
   * Property field label displayed on top
   */
  label: string;
  /**
   * @var
   * Initial date of the control
   */
  initialDate?: string;
  /**
   * @function
   * Defines a formatDate function to display the date of the custom Field.
   * By defaut date.toDateString() is used.
   */
  formatDate?: (date: Date) => string;
  /**
   * @var
   * Defines the time convention to use. The default value is the 24-hour clock convention.
   */
  timeConvention?: ITimeConvention;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected date changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  /**
   * @function
   * This API is called to render the web part.
   * Normally this function must be always defined with the 'this.render.bind(this)'
   * method of the web part object.
   */
  render(): void;
  /**
   * This property is used to indicate the web part's PropertyPane interaction mode: Reactive or NonReactive.
   * The default behaviour is Reactive.
   */
  disableReactivePropertyChanges?: boolean;
  /**
   * @var
   * Parent Web Part properties
   */
  properties: any;
  /**
   * @var
   * An UNIQUE key indicates the identity of this control
   */
  key?: string;
  /**
   * The method is used to get the validation error message and determine whether the input value is valid or not.
   *
   *   When it returns string:
   *   - If valid, it returns empty string.
   *   - If invalid, it returns the error message string and the text field will
   *     show a red border and show an error message below the text field.
   *
   *   When it returns Promise<string>:
   *   - The resolved value is display as error message.
   *   - The rejected, the value is thrown away.
   *
   */
   onGetErrorMessage?: (value: string) => string | Promise<string>;
   /**
    * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
    * Default value is 200.
    */
   deferredValidationTime?: number;
}

/**
 * @interface
 * Private properties of the PropertyFieldDateTimePicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldDateTimePicker.
 *
 */
export interface IPropertyFieldDateTimePickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialDate?: string;
  targetProperty: string;
  formatDate?: (date: Date) => string;
  timeConvention?: ITimeConvention;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  render(): void;
  disableReactivePropertyChanges?: boolean;
  properties: any;
  onGetErrorMessage?: (value: string) => string | Promise<string>;
  deferredValidationTime?: number;
}

/**
 * @interface
 * Represents a PropertyFieldDateTimePicker object
 *
 */
class PropertyFieldDateTimePickerBuilder implements IPropertyPaneField<IPropertyFieldDateTimePickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldDateTimePickerPropsInternal;

  //Custom properties
  private label: string;
  private initialDate: string;
  private formatDate: (date: Date) => string;
  private timeConvention: ITimeConvention;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;
  private onGetErrorMessage: (value: string) => string | Promise<string>;
  private deferredValidationTime: number = 200;
  private renderWebPart: () => void;
  private disableReactivePropertyChanges: boolean = false;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldDateTimePickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialDate = _properties.initialDate;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
    this.formatDate = _properties.formatDate;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
    this.onGetErrorMessage = _properties.onGetErrorMessage;
    if (_properties.deferredValidationTime !== undefined)
      this.deferredValidationTime = _properties.deferredValidationTime;
    if (_properties.timeConvention !== undefined)
      this.timeConvention = _properties.timeConvention;
    else
      this.timeConvention = ITimeConvention.Hours24;
    this.renderWebPart = _properties.render;
    if (_properties.disableReactivePropertyChanges !== undefined && _properties.disableReactivePropertyChanges != null)
      this.disableReactivePropertyChanges = _properties.disableReactivePropertyChanges;
  }

  /**
   * @function
   * Renders the DatePicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldDateTimePickerHostProps> = React.createElement(PropertyFieldDateTimePickerHost, {
      label: this.label,
      initialDate: this.initialDate,
      targetProperty: this.targetProperty,
      formatDate: this.formatDate,
      timeConvention: this.timeConvention,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties,
      key: this.key,
      onGetErrorMessage: this.onGetErrorMessage,
      deferredValidationTime: this.deferredValidationTime,
      render: this.renderWebPart,
      disableReactivePropertyChanges: this.disableReactivePropertyChanges
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
 * Helper method to create the customer field on the PropertyPane.
 * @param targetProperty - Target property the custom field is associated to.
 * @param properties - Strongly typed custom field properties.
 */
export function PropertyFieldDateTimePicker(targetProperty: string, properties: IPropertyFieldDateTimePickerProps): IPropertyPaneField<IPropertyFieldDateTimePickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldDateTimePickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialDate: properties.initialDate,
      timeConvention: properties.timeConvention,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      formatDate: properties.formatDate,
      onDispose: null,
      onRender: null,
      key: properties.key,
      onGetErrorMessage: properties.onGetErrorMessage,
      deferredValidationTime: properties.deferredValidationTime,
      render: properties.render,
      disableReactivePropertyChanges: properties.disableReactivePropertyChanges
    };
    //Calls the PropertyFieldDateTimePicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldDateTimePickerBuilder(targetProperty, newProperties);
}


