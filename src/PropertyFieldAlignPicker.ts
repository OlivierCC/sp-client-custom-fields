/**
 * @file PropertyFieldAlignPicker.ts
 * Define a custom field of type PropertyFieldAlignPicker for
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
import PropertyFieldAlignPickerHost, { IPropertyFieldAlignPickerHostProps } from './PropertyFieldAlignPickerHost';

/**
 * @interface
 * Public properties of the PropertyFieldAlignPicker custom field
 *
 */
export interface IPropertyFieldAlignPickerProps {
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
   * @function
   * Defines a onPropertyChange function to raise when the selected Color changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChanged(propertyPath: string, oldValue: any, newValue: any): void;
  /**
   * @var
   * Parent Web Part properties
   */
  properties: any;

  /**
   * @var
   * Key to help React identify which items have changed, are added, or are removed.
   */
  key?: string;
}

/**
 * @interface
 * Private properties of the PropertyFieldAlignPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldAlignPicker.
 *
 */
export interface IPropertyFieldAlignPickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: string;
  targetProperty: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChanged(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
  key: string;
}

/**
 * @interface
 * Represents a PropertyFieldAlignPicker object
 *
 */
class PropertyFieldAlignPickerBuilder implements IPropertyPaneField<IPropertyFieldAlignPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldAlignPickerPropsInternal;

  //Custom properties
  private label: string;
  private initialValue: string;
  private onPropertyChanged: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;
  private key: string;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldAlignPickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialValue = _properties.initialValue;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChanged = _properties.onPropertyChanged;
    this.customProperties = _properties.properties;
    this.key = _properties.key;
  }

  /**
   * @function
   * Renders the ColorPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldAlignPickerHostProps> = React.createElement(PropertyFieldAlignPickerHost, {
      label: this.label,
      initialValue: this.initialValue,
      targetProperty: this.targetProperty,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChanged: this.onPropertyChanged,
      properties: this.customProperties,
      key: this.key,
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
export function PropertyFieldAlignPicker(targetProperty: string, properties: IPropertyFieldAlignPickerProps): IPropertyPaneField<IPropertyFieldAlignPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldAlignPickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialValue: properties.initialValue,
      onPropertyChanged: properties.onPropertyChanged,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: properties.key
    };
    //Calles the PropertyFieldAlignPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldAlignPickerBuilder(targetProperty, newProperties);
}


