/**
 * @file PropertyFieldIconPicker.ts
 * Define a custom field of type PropertyFieldIconPicker for
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
import PropertyFieldIconPickerHost, { IPropertyFieldIconPickerHostProps } from './PropertyFieldIconPickerHost';

/**
 * @interface
 * Public properties of the PropertyFieldIconPicker custom field
 *
 */
export interface IPropertyFieldIconPickerProps {
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
   * Uses MSDN order or alphabetical for icons
   */
  orderAlphabetical?: boolean;
  /**
   * @var
   * Preview the fonts in the dropdown control (default is true)
   */
  preview?: boolean;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected Font changed.
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
 * Private properties of the PropertyFieldIconPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldIconPicker.
 *
 */
export interface IPropertyFieldIconPickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: string;
  orderAlphabetical?: boolean;
  preview?: boolean;
  targetProperty: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
}

/**
 * @interface
 * Represents a PropertyFieldIconPicker object
 *
 */
class PropertyFieldIconPickerBuilder implements IPropertyPaneField<IPropertyFieldIconPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldIconPickerPropsInternal;

  //Custom properties
  private label: string;
  private initialValue: string;
  private orderAlphabetical: boolean;
  private preview: boolean;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldIconPickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialValue = _properties.initialValue;
    this.orderAlphabetical = _properties.orderAlphabetical;
    this.preview = _properties.preview;
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
    const element: React.ReactElement<IPropertyFieldIconPickerHostProps> = React.createElement(PropertyFieldIconPickerHost, {
      label: this.label,
      initialValue: this.initialValue,
      orderAlphabetical: this.orderAlphabetical,
      preview: this.preview,
      targetProperty: this.targetProperty,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange,
      properties: this.customProperties,
      key: this.targetProperty,
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
 * Helper method to create a Font Picker on the PropertyPane.
 * @param targetProperty - Target property the Font picker is associated to.
 * @param properties - Strongly typed Font Picker properties.
 */
export function PropertyFieldIconPicker(targetProperty: string, properties: IPropertyFieldIconPickerProps): IPropertyPaneField<IPropertyFieldIconPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldIconPickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialValue: properties.initialValue,
      orderAlphabetical: properties.orderAlphabetical,
      preview: properties.preview,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null,
      key: targetProperty,
    };
    //Calles the PropertyFieldIconPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldIconPickerBuilder(targetProperty, newProperties);
}


