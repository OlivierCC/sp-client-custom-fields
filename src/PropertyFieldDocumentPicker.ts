/**
 * @file PropertyFieldDocumentPicker.ts
 * Define a custom field of type PropertyFieldDocumentPicker for
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
} from '@microsoft/sp-client-preview';
import PropertyFieldDocumentPickerHost, { IPropertyFieldDocumentPickerHostProps } from './PropertyFieldDocumentPickerHost';
import { IWebPartContext } from '@microsoft/sp-client-preview';

/**
 * @interface
 * Public properties of the PropertyFieldDocumentPicker custom field
 *
 */
export interface IPropertyFieldDocumentPickerProps {
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
   * Parent web part context
   */
  context: IWebPartContext;
  /**
   * @function
   * Defines a onPropertyChange function to raise when the selected Color changed.
   * Normally this function must be always defined with the 'this.onPropertyChange'
   * method of the web part object.
   */
  onPropertyChange(propertyPath: string, newValue: any): void;
}

/**
 * @interface
 * Private properties of the PropertyFieldDocumentPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldDocumentPicker.
 *
 */
export interface IPropertyFieldDocumentPickerPropsInternal extends IPropertyPaneCustomFieldProps {
  label: string;
  initialValue?: string;
  targetProperty: string;
  context: IWebPartContext;
  onRender(elem: HTMLElement): void;
  onDispose(elem: HTMLElement): void;
  onPropertyChange(propertyPath: string, newValue: any): void;
}

/**
 * @interface
 * Represents a PropertyFieldDocumentPicker object
 *
 */
class PropertyFieldDocumentPickerBuilder implements IPropertyPaneField<IPropertyFieldDocumentPickerPropsInternal> {

  //Properties defined by IPropertyPaneField
  public type: IPropertyPaneFieldType = IPropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyFieldDocumentPickerPropsInternal;

  //Custom properties
  private label: string;
  private initialValue: string;
  private context: IWebPartContext;
  private onPropertyChange: (propertyPath: string, newValue: any) => void;

  /**
   * @function
   * Ctor
   */
  public constructor(_targetProperty: string, _properties: IPropertyFieldDocumentPickerPropsInternal) {
    this.targetProperty = _properties.targetProperty;
    this.properties = _properties;
    this.label = _properties.label;
    this.initialValue = _properties.initialValue;
    this.context = _properties.context;
    this.properties.onDispose = this.dispose;
    this.properties.onRender = this.render;
    this.onPropertyChange = _properties.onPropertyChange;
  }

  /**
   * @function
   * Renders the ColorPicker field content
   */
  private render(elem: HTMLElement): void {
    //Construct the JSX properties
    const element: React.ReactElement<IPropertyFieldDocumentPickerHostProps> = React.createElement(PropertyFieldDocumentPickerHost, {
      label: this.label,
      initialValue: this.initialValue,
      context: this.context,
      targetProperty: this.targetProperty,
      onDispose: this.dispose,
      onRender: this.render,
      onPropertyChange: this.onPropertyChange
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
export function PropertyFieldDocumentPicker(targetProperty: string, properties: IPropertyFieldDocumentPickerProps): IPropertyPaneField<IPropertyFieldDocumentPickerPropsInternal> {

    //Create an internal properties object from the given properties
    var newProperties: IPropertyFieldDocumentPickerPropsInternal = {
      label: properties.label,
      targetProperty: targetProperty,
      initialValue: properties.initialValue,
      onPropertyChange: properties.onPropertyChange,
      context: properties.context,
      onDispose: null,
      onRender: null
    };
    //Calles the PropertyFieldDocumentPicker builder object
    //This object will simulate a PropertyFieldCustom to manage his rendering process
    return new PropertyFieldDocumentPickerBuilder(targetProperty, newProperties);
}


