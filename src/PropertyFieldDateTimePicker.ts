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
  IPropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import PropertyFieldDateTimePickerHost, { IPropertyFieldDateTimePickerHostProps } from './PropertyFieldDateTimePickerHost';

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
   * @function
   * Defines a onPropertyChange function to raise when the selected date changed.
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
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
}

/**
 * @interface
 * Represents a PropertyFieldDateTimePicker object
 *
 */
class PropertyFieldDateTimePickerBuilder implements IPropertyPaneField<IPropertyFieldDateTimePickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: IPropertyPaneFieldType = 1;//IPropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldDateTimePickerPropsInternal;

  //Custom properties
  private label: string;
  private initialDate: string;
  private formatDate: (date: Date) => string;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;

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
 * Helper method to create a Date Picker on the PropertyPane.
 * @param targetProperty - Target property the date picker is associated to.
 * @param properties - Strongly typed Date Picker properties.
 */
export function PropertyFieldDateTimePicker(targetProperty: string, properties: IPropertyFieldDateTimePickerProps): IPropertyPaneField<IPropertyFieldDateTimePickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldDateTimePickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialDate: properties.initialDate,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      formatDate: properties.formatDate,
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldDateTimePicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldDateTimePickerBuilder(targetProperty, newProperties);
}


