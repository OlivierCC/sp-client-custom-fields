/**
 * @file PropertyFieldFontPicker.ts
 * Define a custom field of type PropertyFieldFontPicker for
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
import PropertyFieldFontPickerHost, { IPropertyFieldFontPickerHostProps } from './PropertyFieldFontPickerHost';

/**
 * @interface
 * Public properties of the PropertyFieldFontPicker custom field
 *
 */
export interface IPropertyFieldFontPickerProps {
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
   * Uses web safe font or exact name (default is true)
   */
  useSafeFont?: boolean;
  /**
   * @var
   * Preview the fonts in the dropdown control (default is true)
   */
  previewFonts?: boolean;
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
 * Private properties of the PropertyFieldFontPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldFontPicker.
 *
 */
export interface IPropertyFieldFontPickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: string;
  useSafeFont?: boolean;
  previewFonts?: boolean;
  targetProperty: string;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
  properties: any;
}

/**
 * @interface
 * Represents a PropertyFieldFontPicker object
 *
 */
class PropertyFieldFontPickerBuilder implements IPropertyPaneField<IPropertyFieldFontPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: IPropertyPaneFieldType = 1;//IPropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldFontPickerPropsInternal;

  //Custom properties
  private label: string;
  private initialValue: string;
  private useSafeFont: boolean;
  private previewFonts: boolean;
  private onPropertyChange: (propertyPath: string, oldValue: any, newValue: any) => void;
  private customProperties: any;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldFontPickerPropsInternal) {
    this.render = this.render.bind(this);
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialValue = _properties.initialValue;
    this.useSafeFont = _properties.useSafeFont;
    this.previewFonts = _properties.previewFonts;
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
    const element: React.ReactElement<IPropertyFieldFontPickerHostProps> = React.createElement(PropertyFieldFontPickerHost, {
      label: this.label,
      initialValue: this.initialValue,
      useSafeFont: this.useSafeFont,
      previewFonts: this.previewFonts,
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
 * Helper method to create a Font Picker on the PropertyPane.
 * @param targetProperty - Target property the Font picker is associated to.
 * @param properties - Strongly typed Font Picker properties.
 */
export function PropertyFieldFontPicker(targetProperty: string, properties: IPropertyFieldFontPickerProps): IPropertyPaneField<IPropertyFieldFontPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldFontPickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialValue: properties.initialValue,
      useSafeFont: properties.useSafeFont,
      previewFonts: properties.previewFonts,
      onPropertyChange: properties.onPropertyChange,
      properties: properties.properties,
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldFontPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldFontPickerBuilder(targetProperty, newProperties);
}


